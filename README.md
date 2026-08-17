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

- **Backend**: Node.js / TypeScript / Express, métadonnées persistées dans
  des fichiers JSON (`server/data/`, sans dépendance native — voir plus bas),
  fichiers bruts stockés sur disque (`server/storage/`). Aucun fichier
  déposé n'est jamais écrasé : chaque dépôt crée une nouvelle version
  horodatée.
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

## Exécutable Windows autonome (.exe)

Un exécutable Windows autonome peut être généré avec
[Bun](https://bun.sh) (`bun build --compile`, cross-compilation depuis
Linux/Mac vers Windows possible) : il embarque son propre runtime, sans rien
installer sur la machine cible. Bun doit être installé sur la machine qui
compile (pas sur la machine cible) : voir https://bun.sh/docs/installation.

```bash
cd client && npm install && npm run build && cd ..
cd server && npm install && npm run build:exe
```

Produit `release/AchatsContrats.exe` (~115 Mo) et `release/public/`
(l'interface buildée, copiée depuis `client/dist`). **Les deux doivent rester
ensemble** : contrairement à un empaquetage `pkg` classique, Bun n'embarque
pas les fichiers statiques dans le binaire — voir `LISEZ-MOI-EXECUTABLE.txt`
à distribuer avec le dossier `release/`.

Double-clic sur l'exe → l'application démarre et s'ouvre sur
`http://localhost:4000` dans le navigateur par défaut. La base de données
(fichiers JSON) et les fichiers déposés sont créés dans des dossiers `data/`
et `storage/` **à côté de l'exécutable** — copie-le donc dans un dossier
dédié plutôt qu'à la racine d'un disque partagé.

Contrainte technique : le stockage utilise des fichiers JSON plutôt qu'une
base SQLite, afin d'éviter toute dépendance native incompatible avec
l'empaquetage en exécutable (voir `server/src/jsonStore.ts`).

Un empaquetage alternatif avec `pkg` (Node.js, embarque les fichiers
statiques dans le binaire — un seul fichier à distribuer, mais ~65 Mo au lieu
de ~115 Mo) reste disponible via `npm run build:exe:pkg-legacy` dans
`server/`.

## Hébergement partagé et intégration dans SharePoint Online

L'application reste une architecture classique client/serveur (pas une
application SPFx native) : elle ne "tourne" pas dans SharePoint, mais peut
être **affichée depuis** une page SharePoint via un lien ou un web part
« Intégrer », à condition d'être joignable en HTTPS.

- **Usage individuel** (chacun lance l'exécutable sur son poste) : rien à
  faire, `http://localhost:4000` bénéficie d'une exception navigateur pour
  le contenu mixte et s'affiche normalement dans une page SharePoint (HTTPS)
  consultée depuis ce même poste.
- **Instance unique partagée par toute l'équipe** (un poste/serveur du
  réseau fait tourner l'application en continu) : un navigateur bloque par
  défaut l'affichage d'un contenu `http://` non-local dans une page
  SharePoint (`https://`) — un certificat est nécessaire. Deux variables
  d'environnement activent le HTTPS natif du serveur :

  ```bash
  HOST=0.0.0.0 \
  SSL_CERT_PATH=C:\chemin\vers\certificat.pem \
  SSL_KEY_PATH=C:\chemin\vers\cle-privee.pem \
  AchatsContrats.exe
  ```

  `HOST` élargit l'écoute au-delà de `127.0.0.1` (nécessaire pour être
  joignable par les autres postes) ; `SSL_CERT_PATH`/`SSL_KEY_PATH` pointent
  vers un certificat et sa clé privée (au format PEM) obtenus auprès de
  l'autorité de certification interne de l'organisation, ou de tout
  fournisseur de certificat reconnu par les navigateurs des utilisateurs — un
  certificat auto-signé fonctionne techniquement mais affiche un
  avertissement de sécurité à chaque utilisateur. Sans certificat, le
  serveur avertit dans sa console et reste en HTTP.

  Côté SharePoint : sur une page moderne, ajouter le web part **« Intégrer »**
  (Embed) ou **« Lien »**, et y coller l'URL HTTPS de l'application. Aucune
  configuration côté tenant SharePoint (App Catalog, permissions API) n'est
  nécessaire — ce n'est qu'une intégration visuelle d'iframe/lien, pas une
  application SPFx installée.

## Démarrage rapide (Windows) — pour tester sans taper de commande

Double-clique sur **`lancer-application.bat`** (à la racine du dépôt). Le
script installe les dépendances si besoin, construit l'interface, démarre
l'application sur `http://localhost:4000` et ouvre cette page dans ton
navigateur. Nécessite [Node.js](https://nodejs.org/) (version LTS) installé
au préalable. Laisse la fenêtre noire ouverte pendant le test ; ferme-la pour
arrêter l'application.

## Démarrage manuel (développement)

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

Le serveur sert aussi automatiquement l'interface buildée
(`client/dist`, générée par `npm run build`) sur son propre port
(`http://localhost:4000`) si elle existe — c'est ce que fait le script
`lancer-application.bat`, pour n'avoir qu'un seul port à ouvrir.

Les fichiers de données JSON (`server/data/*.json`) et les fichiers stockés
(`server/storage/`) sont créés automatiquement au premier démarrage.

## Modèle de données

Voir `server/src/db.ts` pour les 4 tables (fichiers JSON) : `conditions_versions`,
`templates_contrats`, `mappings_template`, `generations`. La persistance est
gérée par `server/src/jsonStore.ts` (table JSON minimaliste, sans dépendance
native — voir la section « Exécutable Windows » ci-dessus pour le pourquoi).
