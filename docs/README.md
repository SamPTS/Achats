# Documentation — Achats — Contrats fournisseurs

Ce dossier contient trois documents HTML autonomes (captures d'écran incluses), à ouvrir directement dans un
navigateur :

- **`guide-utilisateur.html`** — mode d'emploi pas à pas, écran par écran.
- **`documentation-fonctionnelle.html`** — règles de gestion et comportement attendu de chaque écran.
- **`documentation-technique.html`** — architecture d'implémentation (services, modèle de données, provisioning,
  permissions).

## ⚠️ Version documentée

Ces trois documents décrivent la variante **SPFx** de l'application, avec ses fonctionnalités avancées : fichier de
conditions commerciales unique sans historique, liaison à un fichier externe (jamais copié), colonne **Marché**
optionnelle pour désambiguïser un code sous-segment, et permissions basées sur les niveaux d'autorisation SharePoint
(Propriétaires / Membres).

La branche actuelle de ce dépôt (`client/` + `server/`) a été revertée vers le **prototype Node/React d'origine**,
qui ne contient pas ces fonctionnalités (historique multi-versions des conditions, pas de lien externe, pas de champ
Marché, pas de rôles) ni le code SPFx. Les captures d'écran de ces documents sont des reproductions fidèles de
l'interface SPFx (mêmes composants, mêmes styles), avec des données fictives à titre d'illustration — elles ne
correspondent pas à l'écran du prototype actuellement présent dans ce dépôt.
