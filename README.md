# Catalogue Qayyim

Tous les produits publics de https://qayyim.fr sur une seule page, rangés en 9 rayons et 37 sous-catégories.

**La page : https://abouaymen31-ux.github.io/catalogue-qayyim/**

Elle sert à voir tout le catalogue d'un coup d'œil, à choisir des produits, à changer leurs prix si besoin, et à créer une carte de visite 85 × 55 mm avec un QR code. Le QR code ouvre une page client qui montre la sélection.

## Mise à jour automatique

Chaque nuit, GitHub relit le site qayyim.fr et met la page à jour : nouveaux produits, prix, promotions, ruptures de stock. Personne n'a rien à faire. La date de la dernière mise à jour est affichée en haut de la page.

Pour lancer une mise à jour tout de suite : onglet **Actions** du dépôt → « Mise à jour du catalogue » → **Run workflow**.

Si le site ne répond pas ou renvoie beaucoup moins de produits que d'habitude, la mise à jour s'arrête sans rien modifier.

## Ce que montre la page

- Le prix public, ou la fourchette « de … à … » quand le prix dépend de la variante, avec le prix de chaque variante.
- Les promotions (ancien prix barré).
- Les ruptures de stock, et « Plus que N » quand le site signale un stock bas.
- La référence du produit, quand elle est renseignée sur le site.
- Une recherche, un filtre par rayon et un tri (A à Z, prix, nouveautés).

Les données viennent de la partie publique du site : la page ne connaît donc ni les quantités exactes en stock, ni les prix d'achat, ni les produits en brouillon.

## Fichiers

| Fichier | Contenu |
|---|---|
| `docs/index.html` + `docs/img/` | La page publiée. |
| `docs/data/catalogue.csv` | Le catalogue en tableau, à ouvrir dans Excel (séparateur `;`). |
| `docs/data/catalogue.json` | Les mêmes données, utilisées par la page. |
| `docs/src/page.html` | Le modèle de la page (le programme y insère les données). |
| `docs/scripts/build_catalogue.py` | Le programme qui relit qayyim.fr et régénère tout. |
| `.github/workflows/maj-catalogue.yml` | Le réglage de la mise à jour automatique. |

## Le code dans le QR

Le lien du QR code se termine par `#c-…` (ou `#p-…` quand les prix sont affichés). Il contient la liste des produits cochés, le nom, le téléphone et les prix modifiés.

Chaque produit est repéré par sa position dans `catalogue.json`. Cette position ne change jamais : un produit retiré du site garde sa place (il est seulement masqué), et les nouveaux s'ajoutent à la fin. Les cartes déjà imprimées restent donc valables.

## Origine

Ce projet reprend et prolonge https://github.com/SachaRM7/cfff.
