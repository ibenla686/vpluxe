# VPLuxe — plateforme de vente privée de luxe

Site statique (HTML, CSS, JavaScript) hébergé gratuitement sur GitHub Pages. Aucun serveur, aucune base de données : le catalogue est un simple fichier Excel.

## Structure

```
index.html              Accueil (ventes en cours, sélection, catégories, maisons)
ventes.html             Ventes en cours, à venir, terminées
maisons.html            Index A–Z des 50 maisons
maison.html?m=<slug>    Espace dédié à une maison (présentation, vente, pièces)
catalogue.html          Catalogue complet avec filtres (maison, catégorie, genre, prix…)
produit.html?ref=<réf>  Fiche produit
panier.html             Panier et demande de commande
invitation.html         Demande d'invitation (accès membre)
contact.html            Contact
a-propos.html · authenticite.html · mentions-legales.html · cgv.html · confidentialite.html

data/site.json          Nom, accroche, coordonnées, textes de livraison/retours, informations légales
data/maisons.json       Les 50 maisons (nom, pays, univers, présentation)
data/catalogue.xlsx     LE CATALOGUE : onglets Catalogue, Maisons, Ventes, Aide
data/catalogue.csv      Copie de secours (lue si le fichier Excel est illisible)
data/catalogue.json     Copie de secours (lue si le CSV est illisible)
data/ventes.json        Ventes (copie de secours de l'onglet Ventes)
assets/img/produits/    Vos photos produits
assets/img/placeholders Visuels d'attente à l'identité VPLuxe (par type de pièce)
assets/img/identite/    Monogramme, favicon, motif, visuel de couverture
assets/css/style.css    Charte graphique
assets/js/app.js        Application (lecture du catalogue, filtres, fiches, panier)
```

## Mettre à jour le catalogue

1. Ouvrez `data/catalogue.xlsx` dans Excel.
2. Onglet **Catalogue** : une ligne par pièce. L'onglet **Aide** décrit chaque colonne. Les lignes actuelles sont des exemples de démonstration : supprimez-les et saisissez vos pièces réelles.
3. Onglet **Ventes** : une ligne par vente événementielle (code, maison, titre, dates, remise). Le site calcule automatiquement « en cours », « à venir » ou « terminée ».
4. Enregistrez, puis remplacez le fichier dans le dossier `data/` du dépôt GitHub (bouton **Add file → Upload files**). Le site est à jour en une à deux minutes.

Photos : déposez vos fichiers (JPG ou WebP, format 4:5 conseillé, 1 500 px de côté, moins de 500 Ko) dans `assets/img/produits/` et indiquez leurs noms dans la colonne **Images** (séparés par des points-virgules). Sans photo, le site affiche automatiquement un visuel d'attente VPLuxe adapté au type de pièce.

Paiement : la colonne **Lien paiement** accepte un lien de paiement (Stripe Payment Link, PayPal, SumUp…) ; un bouton « Acheter maintenant » apparaît alors sur la fiche. Sans lien, le client adresse une demande de commande (e-mail ou WhatsApp pré-remplis) et vous lui envoyez le lien de paiement.

## Personnaliser le site

- `data/site.json` : nom, signature, accroche, e-mail, téléphone, WhatsApp, adresse, Instagram, textes de livraison et de retours, informations légales de la société (rubrique `societe`). Ces informations alimentent automatiquement le pied de page, la page Contact, les mentions légales et les CGV.
- `data/maisons.json` : ajouter ou retirer une maison, modifier sa présentation.
- `assets/css/style.css` : couleurs (variables en tête de fichier), typographies.
- Formulaires : par défaut, les formulaires ouvrent la messagerie du visiteur avec un message pré-rempli. Pour recevoir les demandes sans passer par la messagerie, créez un formulaire sur un service comme Formspree et indiquez son adresse dans `formulaire_action` (`data/site.json`).

## Mise en ligne sur GitHub Pages

1. Dépôt GitHub public contenant ces fichiers à la racine.
2. **Settings → Pages → Build and deployment** : Source « Deploy from a branch », branche `main`, dossier `/ (root)`.
3. Le site est accessible à l'adresse `https://<utilisateur>.github.io/<dépôt>/` après une à deux minutes. Un nom de domaine personnalisé peut être ajouté dans la même page de réglages.

## Avant la mise en ligne réelle

- Remplacer les lignes de démonstration du catalogue par vos pièces réelles et vos photos.
- Renseigner les informations légales dans `data/site.json` et faire relire les CGV, les mentions légales et la politique de confidentialité par votre conseil.
- Compléter la page Authenticité & garanties (statut d'approvisionnement, maison par maison).
- Vérifier la disponibilité du nom VPLuxe auprès de l'INPI.

Les noms des maisons sont cités uniquement pour identifier les pièces proposées et restent la propriété de leurs titulaires. Aucun visuel ni texte des maisons n'est reproduit sur ce site ; les visuels d'attente sont des créations originales à l'identité VPLuxe.
