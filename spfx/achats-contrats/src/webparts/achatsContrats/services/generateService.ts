import '@pnp/sp/lists';
import '@pnp/sp/items';
import JSZip from 'jszip';
import { getSP } from './spClient';
import { ensureProvisioned, withListRecovery, LISTS } from './provisioning';
import { getConditionsVersion, readConditionsRows } from './conditionsService';
import { getTemplate, listTemplates, downloadTemplateFile } from './templatesService';
import { fillDocxTemplate } from './docx';
import { safeFileNamePart, timestampTag } from './storage';
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

export interface GenerateBatchResult {
  blob: Blob;
  filename: string;
  /** Un message par contrat qui a échoué (le reste du lot continue malgré une erreur isolée). */
  erreurs: { codeSousSegment: string; message: string }[];
}

/** Génère plusieurs contrats en une fois (un par ligne relue et validée dans l'écran de
 * génération en lot) et les regroupe dans une seule archive ZIP à télécharger — plutôt que
 * déclencher N téléchargements séparés, souvent bloqués par le navigateur au-delà de quelques-uns
 * d'un coup. Le fichier de conditions et le template sont partagés par tout le lot : leur mise en
 * cache (readConditionsRows, downloadTemplateFile) évite de les retélécharger à chaque contrat. */
export async function generateContractsBatch(
  items: { templateId: string; conditionsVersionId: string; codeSousSegment: string; values: Record<string, string>; traitePar: string }[],
): Promise<GenerateBatchResult> {
  const zip = new JSZip();
  const erreurs: { codeSousSegment: string; message: string }[] = [];
  const usedNames = new Set<string>();

  for (const item of items) {
    try {
      const { blob, filename } = await generateContract(item);
      // Deux codes différents ne devraient jamais produire le même nom de fichier (le code en
      // fait partie), mais on se protège malgré tout d'une collision plutôt que d'écraser
      // silencieusement une entrée du zip par une autre.
      let finalName = filename;
      let i = 2;
      while (usedNames.has(finalName)) {
        finalName = filename.replace(/\.docx$/i, ` (${i++}).docx`);
      }
      usedNames.add(finalName);
      zip.file(finalName, await blob.arrayBuffer());
    } catch (e) {
      erreurs.push({ codeSousSegment: item.codeSousSegment, message: e instanceof Error ? e.message : String(e) });
    }
  }

  const zipBuffer = await zip.generateAsync({ type: 'arraybuffer' });
  return {
    blob: new Blob([zipBuffer], { type: 'application/zip' }),
    filename: `contrats_${timestampTag()}.zip`,
    erreurs,
  };
}

export async function listGenerations(limit = 50): Promise<GenerationLog[]> {
  await ensureProvisioned();
  const items = (await withListRecovery(() =>
    generationsList().items.select(...GENERATION_SELECT).top(limit)(),
  )) as GenerationItem[];
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
