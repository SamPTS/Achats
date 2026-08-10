import fs from 'fs';
import path from 'path';
import db from './db';
import { DIR_CONDITIONS } from './storage';
import { parseConditionsFile } from './utils/excel';

export interface ConditionsVersionRow {
  id: string;
  nom_fichier: string;
  chemin_stockage: string;
  colonnes_detectees: string;
  colonne_code_sous_segment: string | null;
  est_active: number;
}

export function getConditionsVersion(id: string): ConditionsVersionRow | undefined {
  return db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(id) as
    | ConditionsVersionRow
    | undefined;
}

export function getActiveConditionsVersion(): ConditionsVersionRow | undefined {
  return db
    .prepare('SELECT * FROM conditions_versions WHERE est_active = 1 AND archive = 0 LIMIT 1')
    .get() as ConditionsVersionRow | undefined;
}

/** Relit le fichier Excel stocké sur disque et retourne les lignes de données (pas de cache). */
export async function readConditionsRows(version: ConditionsVersionRow) {
  const filePath = path.join(DIR_CONDITIONS, version.chemin_stockage);
  const buffer = fs.readFileSync(filePath);
  return parseConditionsFile(buffer);
}
