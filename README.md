# Achats — Générateur de contrats fournisseurs

Application interne (sans authentification) pour centraliser les fichiers de
conditions commerciales, les templates de contrats Word, et générer des
contrats remplis automatiquement à partir d'un code sous-segment.

## Modules

1. **Conditions commerciales** — dépôt et historisation des fichiers Excel de
   conditions commerciales fournisseurs.
2. **Templates de contrats** — dépôt et historisation des templates Word
   (`.docx`, variables au format `{{Variable}}`), avec mapping
   variable ↔ colonne du fichier de conditions actif.
3. **Générer un contrat** — sélection template + fichier de conditions +
   code sous-segment → formulaire pré-rempli et éditable → génération du
   `.docx` final.

## Stack technique

- **Backend**: Node.js / TypeScript / Express, base de données relationnelle
  **SQLite** (`better-sqlite3`) pour les métadonnées, stockage des fichiers
  bruts sur disque (`server/storage/`). Aucun fichier déposé n'est jamais
  écrasé : chaque dépôt crée une nouvelle version horodatée.
- **Traitement Excel**: `exceljs` (lecture des en-têtes/lignes, génération du
  fichier de mapping vierge avec liste déroulante de validation).
- **Traitement Word**: manipulation directe du XML interne du `.docx`
  (`word/document.xml`) via `jszip` — détection des motifs `{{Variable}}` en
  fusionnant les runs Word, puis remplacement du texte des variables en place
  sans toucher à la mise en forme, aux styles, ni aux autres éléments du
  document (voir `server/src/utils/docx.ts`).
- **Frontend**: React + Vite + TypeScript, 3 écrans (un par module).

Ce choix de stack correspond à l'alternative "Node.js + stockage fichier
local" évoquée dans la spécification (section 6.2), pour rester
autonome (pas de dépendance à un service externe type Supabase).

## Démarrage

```bash
# Backend (API sur http://localhost:4000)
cd server
npm install
npm run dev

# Frontend (UI sur http://localhost:5173, proxy /api vers le backend)
cd client
npm install
npm run dev
```

La base SQLite (`server/data/app.db`) et les fichiers stockés
(`server/storage/`) sont créés automatiquement au premier démarrage.

## Modèle de données

Voir `server/src/db.ts` pour le schéma SQL complet :
`conditions_versions`, `templates_contrats`, `mappings_template`,
`generations`.
