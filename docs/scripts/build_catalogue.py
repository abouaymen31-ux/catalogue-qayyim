"""Récupère le catalogue public de qayyim.fr (API WooCommerce Store),
le range par rayons / sous-catégories et produit :
  - data/catalogue.json  (utilisé par la page)
  - data/catalogue.csv   (à ouvrir dans Excel / Sheets)
  - img/<id>.jpg         (vignettes 240 px)
Usage : python3 scripts/build_catalogue.py [--offline DOSSIER_JSON]
"""
import csv, datetime, html, io, json, os, re, sys, time, unicodedata, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API = "https://qayyim.fr/wp-json/wc/store/v1/products?per_page=100&page={}"
UA = {"User-Agent": "Mozilla/5.0 (catalogue-builder)"}

# Ordre d'affichage : rayon -> sous-catégories
RAYONS = [
    ("livres", "Livres", ["Coran", "Enfants", "Apprendre l'Islam", "Apprendre l'arabe",
        "Croyance", "Prière & jurisprudence", "Prophète & biographies", "Spiritualité",
        "Invocations & petits livres", "Femmes & mariage", "Mort & au-delà", "Roqya, rêves & santé"]),
    ("jeux", "Jeux éducatifs", ["Jeux de société & quiz", "Cartes d'arabe", "Puzzles"]),
    ("senteurs", "Musc & senteurs", ["Muscs en huile", "Muscs solides", "Parfums voiture",
        "Brumes & parfums d'intérieur", "Encens & bakhour"]),
    ("miel", "Miel & alimentation", ["Miels", "Dattes"]),
    ("nigelle", "Gamme nigelle", ["Huiles & gélules", "Hygiène & soins"]),
    ("soins", "Soins & hygiène", ["Dentifrices & miswak", "Savons & shampoings",
        "Baumes & huiles", "Beauté marocaine"]),
    ("medecine", "Médecine prophétique", ["Hijama", "Plantes & poudres"]),
    ("priere", "Prière & dévotion", ["Tasbih", "Tapis de prière", "Accessoires & décoration"]),
    ("omra", "Omra & Hajj", ["Omra & Hajj"]),
]

NAME_FIX = {3979: " (version fille)", 3975: " (version garçon)"}


def norm(s):
    s = unicodedata.normalize("NFD", s.lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def classify(p):
    cats = {html.unescape(c["name"]) for c in p["categories"]}
    n = norm(html.unescape(p["name"]))
    has = cats.__contains__
    # ---- hors livres
    if has("Omra et Hajj") or has("Packs & Coffrets") or "omra" in n or "ihram" in n:
        return "omra", "Omra & Hajj"
    if has("Jeux"):
        if has("Apprendre l’arabe"):
            return "jeux", "Cartes d'arabe"
        if "puzzle" in n:
            return "jeux", "Puzzles"
        return "jeux", "Jeux de société & quiz"
    if has("Muscs voiture"): return "senteurs", "Parfums voiture"
    if has("Muscs solides"): return "senteurs", "Muscs solides"
    if has("Parfums intérieurs"): return "senteurs", "Brumes & parfums d'intérieur"
    if has("Bakhour") or has("Tabkhira"): return "senteurs", "Encens & bakhour"
    if has("Muscs") or (has("Senteurs") and "musc" in n): return "senteurs", "Muscs en huile"
    if has("Miel"): return "miel", "Miels"
    if "datte" in n: return "miel", "Dattes"
    if has("Gamme nigelle"):
        if "huile" in n or "gelule" in n: return "nigelle", "Huiles & gélules"
        return "nigelle", "Hygiène & soins"
    if has("Hijama") and not has("Livres"): return "medecine", "Hijama"
    if has("Herbes et minéraux") or "costus" in n: return "medecine", "Plantes & poudres"
    if "dentifrice" in n or "miswak" in n or "siwak" in n: return "soins", "Dentifrices & miswak"
    if has("Produits beautés marocains") or "henne" in n: return "soins", "Beauté marocaine"
    if "savon" in n or "shampoing" in n: return "soins", "Savons & shampoings"
    if "baume" in n or has("Huile"): return "soins", "Baumes & huiles"
    if has("Tasbih"): return "priere", "Tasbih"
    if has("Tapis de prière") or has("Sans boussole") or "tapis" in n: return "priere", "Tapis de prière"
    if has("Accessoires de Prière") or has("Décorations") or "horloge" in n or "sticker" in n:
        return "priere", "Accessoires & décoration"
    # ---- livres (sous-catégorie la plus parlante d'abord)
    if has("Coran") and not has("Enfants"): return "livres", "Coran"
    if has("Enfants"): return "livres", "Enfants"
    if has("Apprendre l’arabe"): return "livres", "Apprendre l'arabe"
    if has("Mort et Monde Invisible"): return "livres", "Mort & au-delà"
    if has("Femmes") or has("Mariage"): return "livres", "Femmes & mariage"
    if has("Biographies") or "prophete" in n: return "livres", "Prophète & biographies"
    if has("Croyance"): return "livres", "Croyance"
    if has("Apprendre l’Islam") and not has("Jurisprudence"): return "livres", "Apprendre l'Islam"
    if has("Jurisprudence") or "priere" in n: return "livres", "Prière & jurisprudence"
    if has("Petits Livres") or "invocation" in n: return "livres", "Invocations & petits livres"
    if has("Spiritualité") or has("Développement personnel"): return "livres", "Spiritualité"
    if has("Médecine prophétique"): return "livres", "Roqya, rêves & santé"
    return "livres", "Roqya, rêves & santé"


def strip(s):
    s = re.sub(r"<[^>]+>", " ", html.unescape(s or ""))
    return re.sub(r"\s+", " ", s).strip()


def fetch(url, tries=4):
    """Lit une adresse du site, en réessayant si le site ne répond pas tout de suite."""
    for k in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            return urllib.request.urlopen(req, timeout=60).read()
        except Exception as e:
            if k == tries - 1: raise
            print("nouvel essai :", url, e)
            time.sleep(5 * (k + 1))


def load(offline, extra=""):
    out = []
    for page in range(1, 40):
        if offline:
            f = os.path.join(offline, f"{'v' if extra else 'p'}{page}.json")
            if not os.path.exists(f): break
            data = json.load(open(f, encoding="utf-8"))
        else:
            data = json.loads(fetch(API.format(page) + extra))
        if not data: break
        out += data
        if len(data) < 100: break
    return out


def variant_prices(variations):
    """Prix par variante, seulement quand ils diffèrent d'une variante à l'autre.
    Si une seule option (ex. la taille) décide du prix, on ne garde que celle-ci."""
    rows = []
    for v in variations:
        pr = v["prices"]
        if not pr.get("price") or int(pr["price"]) <= 0: continue
        unit = 10 ** pr.get("currency_minor_unit", 2)
        label = html.unescape(v.get("variation") or "").strip()
        parts = dict(x.split(": ", 1) for x in label.split(", ") if ": " in x)
        rows.append((label, parts, int(pr["price"]) / unit, bool(v["is_in_stock"])))
    if len({r[2] for r in rows}) < 2: return None
    for attr in (rows[0][1] if rows else {}):
        by = {}
        for _, parts, price, _ in rows:
            by.setdefault(parts.get(attr), set()).add(price)
        if None not in by and len(by) > 1 and all(len(s) == 1 for s in by.values()):
            seen, out = set(), []
            for _, parts, price, _ in rows:
                if parts[attr] in seen: continue
                seen.add(parts[attr])
                out.append({"v": parts[attr], "p": price,
                            "s": any(r[3] for r in rows if r[1].get(attr) == parts[attr])})
            return {"n": attr, "l": sorted(out, key=lambda x: x["p"])}
    out = [{"v": " · ".join(parts.values()) or label, "p": price, "s": st} for label, parts, price, st in rows]
    return {"n": "", "l": sorted(out, key=lambda x: x["p"])[:16]}


def thumb(p):
    if not p["images"]: return None
    img = p["images"][0]
    cands = re.findall(r"(\S+) (\d+)w", img.get("srcset") or "")
    cands = sorted(((int(w), u) for u, w in cands), key=lambda x: x[0])
    for w, u in cands:
        if w >= 280: return u
    return img["thumbnail"] or img["src"]


def main():
    offline = sys.argv[2] if len(sys.argv) > 2 and sys.argv[1] == "--offline" else None
    raw = load(offline)
    # catalogue précédent : il fixe la position de chaque produit (utilisée par les QR codes)
    cat_path = os.path.join(ROOT, "data", "catalogue.json")
    old = json.load(open(cat_path, encoding="utf-8"))["products"] if os.path.exists(cat_path) else []
    old_live = sum(1 for x in old if not x.get("gone"))
    if not raw or len(raw) < old_live * 0.6:
        sys.exit(f"Arrêt : le site n'a renvoyé que {len(raw)} produits (contre {old_live} avant). "
                 "Le catalogue n'est pas modifié.")
    by_parent = {}
    for v in load(offline, "&type=variation"):
        by_parent.setdefault(v.get("parent"), []).append(v)
    order = {k: i for i, (k, _, _) in enumerate(RAYONS)}
    sub_order = {(k, s): j for k, _, subs in RAYONS for j, s in enumerate(subs)}
    fresh = {}
    for p in sorted(raw, key=lambda p: p["id"]):
        r, s = classify(p)
        assert (r, s) in sub_order, (r, s, p["name"])
        pr = p["prices"]
        unit = 10 ** pr.get("currency_minor_unit", 2)
        price = (int(pr["price"]) / unit) if pr.get("price") and int(pr["price"]) > 0 else None
        rng = pr.get("price_range")
        reg = int(pr["regular_price"]) / unit if pr.get("regular_price") else None
        opts = [{"n": a["name"], "v": [t["name"] for t in a["terms"]]} for a in p["attributes"]]
        low = p.get("low_stock_remaining")
        fresh[p["id"]] = {
            "id": p["id"],
            "name": html.unescape(p["name"]).strip() + NAME_FIX.get(p["id"], ""),
            "r": r, "s": s,
            "price": price,
            "priceMax": int(rng["max_amount"]) / unit if rng else None,
            # prix avant promotion, seulement si le produit est en promotion
            "was": reg if p.get("on_sale") and reg and price and reg > price else None,
            "stock": bool(p["is_in_stock"]),
            # nombre restant, indiqué par le site seulement quand le stock est bas
            "low": low if isinstance(low, int) and low > 0 else None,
            "vp": variant_prices(by_parent.get(p["id"], [])),
            "opts": opts,
            "desc": strip(p["short_description"])[:220],
            "url": p["permalink"],
            "sku": p.get("sku") or "",
            "siteCats": [html.unescape(c["name"]) for c in p["categories"]],
            "_img": thumb(p),
        }
    # Les positions ne changent jamais : un produit retiré du site garde sa place (marqué "gone"),
    # les nouveaux s'ajoutent à la fin. Les QR codes déjà imprimés restent donc valables.
    prods, known = [], set()
    for o in old:
        known.add(o["id"])
        if o["id"] in fresh:
            prods.append(fresh[o["id"]])
        else:
            o = dict(o, gone=True, stock=False, _img=None)
            prods.append(o)
    added = [fresh[i] for i in sorted(fresh) if i not in known]
    prods += added
    gone_n = sum(1 for x in prods if x.get("gone"))
    # images
    from PIL import Image
    os.makedirs(os.path.join(ROOT, "img"), exist_ok=True)
    for x in prods:
        dst = os.path.join(ROOT, "img", f"{x['id']}.jpg")
        if x.get("gone"): continue
        x["img"] = f"img/{x['id']}.jpg" if x["_img"] else None
        if not x["_img"] or os.path.exists(dst): continue
        try:
            im = Image.open(io.BytesIO(fetch(x["_img"])))
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, (255, 255, 255)); bg.paste(im, mask=im.split()[3])
            bg.thumbnail((240, 240))
            bg.save(dst, "JPEG", quality=78, optimize=True, progressive=True)
        except Exception as e:  # une image manquante ne doit pas bloquer la mise à jour
            print("image ignorée :", x["name"], e)
            x["img"] = None
    for x in prods: x.pop("_img", None)
    # la date ne change que si le contenu change, pour ne pas créer une modification chaque nuit
    stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if os.path.exists(cat_path):
        prev = json.load(open(cat_path, encoding="utf-8"))
        if prev.get("products") == prods and prev.get("updated"):
            stamp = prev["updated"]
    data = {
        "source": "https://qayyim.fr",
        "updated": stamp,
        "rayons": [{"k": k, "name": n, "subs": s} for k, n, s in RAYONS],
        "products": prods,  # l'ordre (id croissant) sert d'index pour les codes de sélection
    }
    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    json.dump(data, open(cat_path, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    names = {k: n for k, n, _ in RAYONS}
    fr = lambda n: "" if n is None else f"{n:.2f}".replace(".", ",")
    with open(os.path.join(ROOT, "data", "catalogue.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Rayon", "Sous-catégorie", "Produit", "Prix public (€)", "Prix maximum (€)", "Prix avant promotion (€)",
                    "Variantes", "Prix par variante", "En stock", "Reste (stock bas)", "Référence", "Lien"])
        live = [x for x in prods if not x.get("gone")]
        for x in sorted(live, key=lambda x: (order[x["r"]], sub_order[(x["r"], x["s"])], norm(x["name"]))):
            vp = x.get("vp")
            w.writerow([names[x["r"]], x["s"], x["name"], fr(x["price"]),
                        fr(x["priceMax"]) if x["priceMax"] and x["priceMax"] != x["price"] else "",
                        fr(x.get("was")),
                        " | ".join(f"{o['n']} : {', '.join(o['v'])}" for o in x["opts"]),
                        " | ".join(f"{v['v']} : {fr(v['p'])}" for v in vp["l"]) if vp else "",
                        "oui" if x["stock"] else "rupture", x.get("low") or "", x["sku"], x["url"]])
    tpl = os.path.join(ROOT, "src", "page.html")
    if os.path.exists(tpl):
        page = open(tpl, encoding="utf-8").read()
        blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        page = page.replace("/*__DATA__*/null", blob)
        head = ('<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
                '<style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
                'body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n')
        # le <title>, les polices et le <style> de la page vont dans <head>, le reste dans <body>
        split = page.index("</style>") + len("</style>")
        doc = head + page[:split] + "\n</head>\n<body>\n" + page[split:] + "\n</body>\n</html>\n"
        open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(doc)
    from collections import Counter
    c = Counter((x["r"], x["s"]) for x in prods if not x.get("gone"))
    for k, n, subs in RAYONS:
        print(n, sum(c[(k, s)] for s in subs), {s: c[(k, s)] for s in subs})
    print("total", len(prods) - gone_n, "| nouveaux", len(added), "| retirés du site", gone_n)
    for x in added: print("  + ", x["name"])


if __name__ == "__main__":
    main()
