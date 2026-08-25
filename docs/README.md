# Documentation — Achats — Contrats fournisseurs

- **[Guide utilisateur](guide-utilisateur.md)** — mode d'emploi pas à pas, écran par écran.
- **[Documentation fonctionnelle](documentation-fonctionnelle.md)** — règles de gestion et comportement attendu de
  chaque écran.
- **[Documentation technique](documentation-technique.md)** — architecture d'implémentation (services, modèle de
  données, provisioning, permissions).

Ces trois documents décrivent la variante **SPFx** de l'application (`spfx/achats-contrats`) : fichier de
conditions commerciales unique sans historique, liaison à un fichier externe (jamais copié), colonne **Marché**
optionnelle pour désambiguïser un code sous-segment, et permissions basées sur les niveaux d'autorisation SharePoint
(Propriétaires / Membres). Les captures d'écran (`images/`) sont des reproductions fidèles de cette interface (mêmes
composants, mêmes styles), avec des données fictives à titre d'illustration.

Le prototype Node/React (`client/` + `server/`, à la racine du dépôt) reste une variante indépendante de
l'application, avec un modèle de données antérieur (historique multi-versions des conditions, pas de lien externe,
pas de champ Marché, pas de rôles) — il n'est pas couvert par ces documents.
