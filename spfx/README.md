# Achats — Contrats fournisseurs (web part SharePoint Framework)

Portage en **application SPFx native** de l'application autonome (`client/` +
`server/`) : plus aucun serveur — tout tourne dans le navigateur, dans une
page SharePoint, avec les données stockées dans des **listes et
bibliothèques SharePoint** plutôt que dans des fichiers JSON/disque locaux.

> Les deux versions (autonome et SPFx) coexistent dans ce dépôt et sont
> indépendantes l'une de l'autre — aucune migration de données automatique
> entre les deux.

## Ce qui a changé par rapport à l'application autonome

| Aspect | Application autonome | Web part SPFx |
|---|---|---|
| Exécution | Serveur Node/Bun (exécutable) | Uniquement dans le navigateur, hébergé par SharePoint |
| Authentification | Aucune | Héritée de SharePoint (Azure AD), automatique |
| Données (conditions, templates, mapping, journal) | Fichiers JSON sur disque | 4 listes SharePoint (voir ci-dessous) |
| Fichiers déposés (.xlsx, .docx) | `server/storage/` sur disque | 2 bibliothèques de documents SharePoint |
| Parsing Excel / fusion Word | Node.js (`exceljs`, `jszip`) | Les mêmes bibliothèques, tournant côté navigateur — logique inchangée |
| Navigation entre les 3 modules | 3 pages (routes React Router) | 3 onglets (Pivot Fluent UI) dans un seul web part |
| Déploiement | `.exe` à double-cliquer | `.sppkg` à charger dans l'App Catalog SharePoint |

Le cœur métier (détection des colonnes du fichier de conditions,
extraction/fusion des balises `{{Variable}}` dans les templates Word,
détection des mappings devenus invalides) est **porté à l'identique** — voir
`src/webparts/achatsContrats/services/excel.ts` et `docx.ts`, copies quasi
conformes de `server/src/utils/excel.ts` et `docx.ts` (ces fichiers
n'utilisaient déjà aucune API Node.js).

## Auto-provisionnement des listes SharePoint

Le web part crée automatiquement, à son premier chargement sur un site, les
listes et bibliothèques dont il a besoin (voir
`src/webparts/achatsContrats/services/provisioning.ts`) :

- Listes : `ConditionsVersions`, `TemplatesContrats`, `MappingsTemplate`, `Generations`
- Bibliothèques de documents : `ConditionsFichiers`, `TemplatesFichiers`

**Important** : la première personne qui ouvre le web part sur le site doit
disposer des droits de gestion de listes (Propriétaire, ou Membre avec
droits de conception) pour que cette création automatique réussisse. Une
fois les listes créées, les permissions standard du site s'appliquent
normalement pour tous les usages suivants.

## Déploiement

1. **Build du paquet** (déjà fait, `sharepoint/solution/achats-contrats.sppkg`
   généré ; pour le regénérer après une modification du code) :
   ```bash
   cd spfx/achats-contrats
   npm install
   npm run build   # heft test --clean --production && heft package-solution --production
   ```
2. **Charger le paquet** :
   - Sur le site `CONTRATHEQUE_GRP-PROJETCONTRATSIA2` (ou tout autre site),
     ouvrir **Contenu du site → App Catalog du site** (à créer si elle
     n'existe pas encore : Paramètres du site → App Catalog du site → Créer),
     ou utiliser l'App Catalog du tenant si vous préférez un déploiement
     centralisé pour plusieurs sites.
   - Glisser `achats-contrats.sppkg` dans l'App Catalog, confirmer la
     confiance ("Faire confiance à cette application ?").
3. **Ajouter le web part à une page** :
   - Sur une page moderne du site, **+ Ajouter un composant** → rechercher
     **« AchatsContrats »** → l'ajouter.
   - Ouvrir la page : le web part provisionne les listes au premier
     chargement (voir ci-dessus), puis affiche directement les 3 onglets.

`skipFeatureDeployment: true` (voir `config/package-solution.json`) : une
fois chargé dans l'App Catalog du site, le web part est immédiatement
disponible sur ce site sans étape d'activation supplémentaire.

## ⚠️ Limite importante : non testé contre un tenant SharePoint réel

Ce portage a été développé et **vérifié par une compilation et un
empaquetage SPFx complets et réussis** (`heft build` puis
`heft package-solution`, sans erreur), mais **sans accès à un tenant
SharePoint réel pour le tester en conditions réelles** — contrairement à
l'application autonome, testée de bout en bout par de vrais appels HTTP.
Les points suivants sont donc à vérifier en priorité lors du premier
déploiement réel :

- Le nommage exact des champs internes créés par PnPjs (`fields.addText`,
  `addBoolean`, etc.) peut différer légèrement du nom demandé si un conflit
  existe déjà sur le site — à vérifier dans les paramètres de chaque liste
  si une erreur de champ introuvable apparaît.
- Les droits nécessaires au premier chargement (création des listes) sont
  documentés ci-dessus mais non vérifiés en pratique.
- Le comportement de `sp.web.getFileByServerRelativePath(...).getBuffer()`
  sur de gros fichiers Excel (fichiers de conditions volumineux) n'a pas été
  mesuré en performance réelle.

Merci de remonter tout écart constaté lors du premier déploiement — il sera
corrigé rapidement.
