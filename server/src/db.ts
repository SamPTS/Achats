import Database from 'better-sqlite3';
import fs from 'fs';
import path from 'path';

const DATA_DIR = path.join(__dirname, '..', 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

export const db = new Database(path.join(DATA_DIR, 'app.db'));
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
CREATE TABLE IF NOT EXISTS conditions_versions (
  id TEXT PRIMARY KEY,
  nom_fichier TEXT NOT NULL,
  date_depot TEXT NOT NULL,
  chemin_stockage TEXT NOT NULL,
  colonnes_detectees TEXT NOT NULL, -- JSON string[]
  colonne_code_sous_segment TEXT,
  est_active INTEGER NOT NULL DEFAULT 0,
  depose_par TEXT,
  nb_lignes INTEGER NOT NULL DEFAULT 0,
  nb_lignes_vides INTEGER NOT NULL DEFAULT 0,
  nb_colonnes INTEGER NOT NULL DEFAULT 0,
  doublons_detectes INTEGER NOT NULL DEFAULT 0,
  archive INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS templates_contrats (
  id TEXT PRIMARY KEY,
  group_id TEXT NOT NULL, -- lie les versions successives d'un même template
  libelle TEXT NOT NULL,
  departement TEXT,
  date_depot TEXT NOT NULL,
  chemin_stockage TEXT NOT NULL,
  version INTEGER NOT NULL,
  nom_fichier TEXT NOT NULL,
  variables_detectees TEXT NOT NULL, -- JSON string[]
  depose_par TEXT,
  archive INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS mappings_template (
  id TEXT PRIMARY KEY,
  template_id TEXT NOT NULL REFERENCES templates_contrats(id),
  variable TEXT NOT NULL,
  colonne_correspondante TEXT,
  statut TEXT NOT NULL DEFAULT 'manquante' -- 'mappee' | 'libre' | 'manquante'
);

CREATE TABLE IF NOT EXISTS generations (
  id TEXT PRIMARY KEY,
  template_id TEXT NOT NULL,
  conditions_version_id TEXT NOT NULL,
  code_sous_segment TEXT NOT NULL,
  date_generation TEXT NOT NULL,
  traite_par TEXT
);
`);

export default db;
