import fs from 'fs';
import path from 'path';
import { JsonTable, Row } from './jsonStore';
import { getAppDir } from './runtimePaths';

const DATA_DIR = path.join(getAppDir(), 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

export interface ConditionsVersionRow extends Row {
  nomFichier: string;
  dateDepot: string;
  cheminStockage: string;
  colonnes: string[];
  colonneCodeSousSegment: string | null;
  estActive: boolean;
  deposePar: string | null;
  nbLignes: number;
  nbLignesVides: number;
  nbColonnes: number;
  doublonsDetectes: number;
  archive: boolean;
}

export interface TemplateRow extends Row {
  groupId: string;
  libelle: string;
  departement: string | null;
  dateDepot: string;
  cheminStockage: string;
  version: number;
  nomFichier: string;
  variables: string[];
  deposePar: string | null;
  archive: boolean;
}

export type MappingStatut = 'mappee' | 'libre' | 'manquante';

export interface MappingRow extends Row {
  templateId: string;
  variable: string;
  colonneCorrespondante: string | null;
  statut: MappingStatut;
}

export interface GenerationRow extends Row {
  templateId: string;
  conditionsVersionId: string;
  codeSousSegment: string;
  dateGeneration: string;
  traitePar: string | null;
}

export const conditionsVersionsTable = new JsonTable<ConditionsVersionRow>(DATA_DIR, 'conditions_versions');
export const templatesTable = new JsonTable<TemplateRow>(DATA_DIR, 'templates_contrats');
export const mappingsTable = new JsonTable<MappingRow>(DATA_DIR, 'mappings_template');
export const generationsTable = new JsonTable<GenerationRow>(DATA_DIR, 'generations');
