import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './spClient';
import { ensureProvisioned, LISTS } from './provisioning';
import { getConditionsVersion, readConditionsRows } from './conditionsService';
import { getTemplate, listTemplates, downloadTemplateFile } from './templatesService';
import { fillDocxTemplate } from './docx';
import { safeFileNamePart } from './storage';
import type { GenerateTemplateOption, GenerationLog, SearchMatch } from '../model/types';

interface GenerationItem {
  Id: number;
  TemplateId: string | null;
  TemplateLibelle: string | null;
  ConditionsVersionId: string | null;
  ConditionsNomFichier: string | null;
  CodeSousSegment: string;
  DateGeneration: string;
  TraitePar: string | null;
}

const GENERATION_SELECT = [
  'Id',
  'TemplateId',
  'TemplateLibelle',
  'ConditionsVersionId',
  'ConditionsNomFichier',
  'CodeSousSegment',
  'DateGeneration',
  'TraitePar',
];

function generationsList() {
  return getSP().web.lists.getByTitle(LISTS.generations);
}

/** Templates disponibles pour la génération (dernière version par groupe), avec l'indicateur
 * mappingComplet — le statut de mapping des templates est déjà calculé contre le fichier de
 * conditions actif (voir templatesService), donc identique à ce que ferait un calcul séparé. */
export async function listGenerableTemplates(): Promise<GenerateTemplateOption[]> {
  const templates = await listTemplates();
  return templates
    .map((t) => ({
      id: t.id,
      libelle: t.libelle,
      departement: t.departement,
      version: t.version,
      mappingComplet: t.statutMapping === 'complet',
    }))
    .sort((a, b) => a.libelle.localeCompare(b.libelle));
}

export async function searchCode(conditionsVersionId: string, code: string): Promise<SearchMatch[]> {
  const version = await getConditionsVersion(conditionsVersionId);
  if (!version) throw new Error('Fichier de conditions introuvable.');
  if (!version.colonneCodeSousSegment) {
    throw new Error('La colonne "code sous-segment" n\'est pas définie pour ce fichier.');
  }
  const { rows } = await readConditionsRows(version);
  const normalized = code.trim().toLowerCase();
  const colonne = version.colonneCodeSousSegment;

  const matches: SearchMatch[] = [];
  rows.forEach((row, idx) => {
    if ((row[colonne] ?? '').trim().toLowerCase() === normalized) {
      const previewKeys = Object.keys(row).slice(0, 5);
      const preview: Record<string, string> = {};
      for (const k of previewKeys) preview[k] = row[k];
      matches.push({ rowIndex: idx, preview });
    }
  });
  if (matches.length === 0) throw new Error('Aucun code sous-segment correspondant.');
  return matches;
}

export interface MappedValuesResult {
  values: Record<string, string>;
  colonnesManquantes: string[];
}

export async function getMappedValues(conditionsVersionId: string, templateId: string, rowIndex: number): Promise<MappedValuesResult> {
  const version = await getConditionsVersion(conditionsVersionId);
  if (!version) throw new Error('Fichier de conditions introuvable.');
  const { rows } = await readConditionsRows(version);
  const row = rows[rowIndex];
  if (!row) throw new Error('Ligne introuvable.');

  const template = await getTemplate(templateId);
  if (!template) throw new Error('Template introuvable.');

  const values: Record<string, string> = {};
  const colonnesManquantes: string[] = [];
  for (const m of template.mappings) {
    if (m.statut === 'mappee' && m.colonneCorrespondante) {
      if (!version.colonnes.includes(m.colonneCorrespondante)) {
        colonnesManquantes.push(m.variable);
        values[m.variable] = '';
        continue;
      }
      // Convention "zéro interprétation" : colonne vide -> champ vide, jamais de valeur devinée.
      values[m.variable] = row[m.colonneCorrespondante] ?? '';
    } else {
      values[m.variable] = '';
    }
  }
  return { values, colonnesManquantes };
}

export interface GenerateContractResult {
  blob: Blob;
  filename: string;
}

export async function generateContract(params: {
  templateId: string;
  conditionsVersionId: string;
  codeSousSegment: string;
  values: Record<string, string>;
  traitePar: string;
}): Promise<GenerateContractResult> {
  await ensureProvisioned();
  const template = await getTemplate(params.templateId);
  if (!template) throw new Error('Template introuvable.');
  const version = await getConditionsVersion(params.conditionsVersionId);

  const buffer = await downloadTemplateFile(template);
  const filled = await fillDocxTemplate(buffer, params.values);

  await generationsList().items.add({
    Title: `${template.libelle} — ${params.codeSousSegment}`,
    TemplateId: params.templateId,
    TemplateLibelle: `${template.libelle} (v${template.version})`,
    ConditionsVersionId: params.conditionsVersionId,
    ConditionsNomFichier: version ? version.nomFichier : null,
    CodeSousSegment: params.codeSousSegment,
    DateGeneration: new Date().toISOString(),
    TraitePar: params.traitePar || null,
  });

  const filename = `${safeFileNamePart(template.nomFichier.replace(/\.docx$/i, ''))}_${safeFileNamePart(params.codeSousSegment)}.docx`;
  const blob = new Blob([filled], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  return { blob, filename };
}

export async function listGenerations(limit = 50): Promise<GenerationLog[]> {
  await ensureProvisioned();
  const items = (await generationsList()
    .items.select(...GENERATION_SELECT)
    .top(limit)()) as GenerationItem[];
  return items
    .map((i) => ({
      id: String(i.Id),
      templateId: i.TemplateId,
      templateLibelle: i.TemplateLibelle,
      conditionsVersionId: i.ConditionsVersionId,
      conditionsNomFichier: i.ConditionsNomFichier,
      codeSousSegment: i.CodeSousSegment,
      dateGeneration: i.DateGeneration,
      traitePar: i.TraitePar,
    }))
    .sort((a, b) => (a.dateGeneration < b.dateGeneration ? 1 : -1));
}
