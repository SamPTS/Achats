import fs from 'fs';
import path from 'path';
import { conditionsVersionsTable, ConditionsVersionRow } from './db';
import { DIR_CONDITIONS } from './storage';
import { parseConditionsFile } from './utils/excel';

export function getConditionsVersion(id: string): ConditionsVersionRow | undefined {
  return conditionsVersionsTable.getById(id);
}

export function getActiveConditionsVersion(): ConditionsVersionRow | undefined {
  return conditionsVersionsTable.findOne((v) => v.estActive && !v.archive);
}

/** Relit le fichier Excel stocké sur disque et retourne les lignes de données (pas de cache). */
export async function readConditionsRows(version: ConditionsVersionRow) {
  const filePath = path.join(DIR_CONDITIONS, version.cheminStockage);
  const buffer = fs.readFileSync(filePath);
  return parseConditionsFile(buffer);
}
