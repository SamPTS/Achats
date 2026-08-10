# Achats — Contrats fournisseurs (solution SPFx)

Solution SharePoint Framework (SPFx) implémentant l'application de génération de
contrats fournisseurs, hébergée nativement dans SharePoint. Cette solution
remplace le prototype Node/Express + React livré précédemment (voir `../server`
et `../client`, conservés à titre d'historique mais non maintenus).

## Principe

Un unique web part (« Achats — Contrats fournisseurs ») avec une navigation
interne à 3 onglets, correspondant aux 3 modules de la spécification :

1. **Conditions commerciales** — dépôt/historisation des fichiers Excel.
2. **Templates de contrats** — dépôt/historisation des templates Word +
   mapping variables ↔ colonnes.
3. **Générer un contrat** — sélection + recherche + génération du `.docx`.

Le web part fonctionne uniquement dans le contexte SharePoint : pas de backend
séparé, pas d'authentification à gérer (l'utilisateur est déjà authentifié via
SharePoint/Azure AD), et l'accès aux données/fichiers respecte les permissions
du site SharePoint où le web part est ajouté.

## Stockage SharePoint

Au premier chargement, le web part crée automatiquement (si elles n'existent
pas déjà) :

- 4 **listes** de métadonnées :
  - `AchatsContrats_ConditionsVersions`
  - `AchatsContrats_TemplatesContrats`
  - `AchatsContrats_MappingsTemplate`
  - `AchatsContrats_Generations` (journal des générations)
- 1 **bibliothèque de documents** `AchatsContrats_Fichiers`, avec deux
  dossiers `Conditions` et `Templates`, contenant les fichiers bruts
  (`.xlsx` / `.docx`) déposés. Chaque dépôt crée un nouveau fichier horodaté
  (`{type}_{libelle}_{YYYYMMDD-HHMM}.{ext}`) — aucun fichier n'est jamais
  écrasé, conformément à la règle d'historisation permanente.

Les noms des listes/bibliothèque sont configurables dans le panneau de
propriétés du web part (utile si plusieurs instances doivent coexister sur le
même site, ou pour s'aligner sur une convention de nommage existante).

**Droits requis** : la création automatique des listes nécessite les droits
"Créer des listes" sur le site (généralement Propriétaire/Membre avec droits
de conception). Si l'utilisateur courant n'a pas ces droits, un administrateur
doit charger la page une première fois (ou provisionner les listes/bibliothèque
manuellement avec les mêmes noms/colonnes) ; les utilisateurs suivants n'ont
besoin que des droits de lecture/écriture standard sur les listes.

## Traitement Excel/Word — 100 % côté client

Contrainte SPFx : pas de serveur Node séparé, tout s'exécute dans le
navigateur. La logique de traitement des fichiers a donc été portée sans
changement fonctionnel :

- **Excel** (`utils/excel.ts`) : lecture des en-têtes/lignes, détection des
  doublons de code sous-segment, génération du fichier de mapping vierge
  avec liste déroulante — via `exceljs`, qui fonctionne aussi bien dans le
  navigateur que côté Node.
- **Word** (`utils/docx.ts`) : détection des variables `{{Variable}}` en
  fusionnant les runs XML (robuste si Word scinde une variable entre
  plusieurs balises `<w:t>`), et remplacement du texte préservant
  intégralement la mise en forme — via `jszip`, en manipulant directement
  `word/document.xml`.

## Démarrage / déploiement

```bash
npm install

# Aperçu local (SharePoint Workbench local, sans données réelles)
npm run serve   # ou: gulp serve

# Aperçu contre un vrai tenant (workbench hébergé) :
# gulp serve --nobrowser
# puis ouvrir https://<votre-tenant>.sharepoint.com/_layouts/15/workbench.aspx

# Empaqueter pour déploiement
gulp bundle --ship
gulp package-solution --ship
# Le paquet .sppkg est généré dans ./sharepoint/solution/achats-contrats.sppkg
```

Déploiement : charger le `.sppkg` dans l'App Catalog du tenant (ou du site),
puis ajouter l'app/web part sur une page SharePoint (moderne) du site
Achats/Import.

## Structure du code

```
src/webparts/achatsContrats/
  AchatsContratsWebPart.ts       # point d'entrée SPFx, panneau de propriétés
  components/                    # App.tsx (onglets) + 1 composant par module
  services/
    spClient.ts                 # instance PnPjs partagée
    provisioning.ts              # création idempotente des listes/bibliothèque
    conditionsService.ts         # Module 1
    templatesService.ts          # Module 2
    generateService.ts           # Module 3
  utils/
    excel.ts, docx.ts, naming.ts # logique métier, portée du prototype Node
  models/types.ts                # types partagés
```
