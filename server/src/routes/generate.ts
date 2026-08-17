import { Router } from 'express';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import { templatesTable, mappingsTable, generationsTable, TemplateRow, MappingRow } from '../db';
import { DIR_TEMPLATES } from '../storage';
import { fillDocxTemplate } from '../utils/docx';
import { getActiveConditionsVersion, getConditionsVersion, readConditionsRows } from '../conditionsService';

const router = Router();

function isMappingComplete(templateId: string): boolean {
  const mappings = mappingsTable.find((m) => m.templateId === templateId);
  if (mappings.length === 0) return false;
  // Une colonne mappée mais absente du fichier de conditions actif ne compte
  // pas comme "complet" : le mapping doit être refait avant de proposer le
  // template pour la génération (voir server/src/routes/templates.ts).
  const active = getActiveConditionsVersion();
  const colonnesRef = active ? active.colonnes : null;
  return mappings.every((m) => {
    if (m.statut === 'libre') return true;
    if (m.statut !== 'mappee' || !m.colonneCorrespondante) return false;
    if (colonnesRef && !colonnesRef.includes(m.colonneCorrespondante)) return false;
    return true;
  });
}

// Tous les templates déposés (dernière version par groupe), avec leur statut de mapping.
// On ne filtre plus les templates au mapping incomplet : ils doivent rester visibles dans
// le sélecteur (sinon un template qu'on vient de déposer semble avoir disparu), mais restent
// non sélectionnables tant que leur mapping n'est pas complet — voir mappingComplet.
router.get('/templates', (_req, res) => {
  const rows = templatesTable.find((t) => !t.archive);
  const latestByGroup = new Map<string, TemplateRow>();
  for (const row of rows) {
    const current = latestByGroup.get(row.groupId);
    if (!current || row.version > current.version) latestByGroup.set(row.groupId, row);
  }
  const all = Array.from(latestByGroup.values()).sort((a, b) => a.libelle.localeCompare(b.libelle));

  res.json(
    all.map((r) => ({
      id: r.id,
      libelle: r.libelle,
      departement: r.departement,
      version: r.version,
      mappingComplet: isMappingComplete(r.id),
    }))
  );
});

// Recherche d'un code sous-segment dans le fichier de conditions choisi.
router.get('/search', async (req, res) => {
  try {
    const conditionsVersionId = req.query.conditionsVersionId as string;
    const code = (req.query.code as string) || '';
    if (!conditionsVersionId) return res.status(400).json({ error: 'conditionsVersionId requis.' });
    const version = getConditionsVersion(conditionsVersionId);
    if (!version) return res.status(404).json({ error: 'Fichier de conditions introuvable.' });
    if (!version.colonneCodeSousSegment) {
      return res.status(400).json({ error: 'La colonne "code sous-segment" n\'est pas définie pour ce fichier.' });
    }
    const { rows } = await readConditionsRows(version);
    const normalized = code.trim().toLowerCase();
    const colonne = version.colonneCodeSousSegment;

    const matches: { rowIndex: number; preview: Record<string, string> }[] = [];
    rows.forEach((row, idx) => {
      if ((row[colonne] ?? '').trim().toLowerCase() === normalized) {
        const previewKeys = Object.keys(row).slice(0, 5);
        const preview: Record<string, string> = {};
        for (const k of previewKeys) preview[k] = row[k];
        matches.push({ rowIndex: idx, preview });
      }
    });

    if (matches.length === 0) {
      return res.status(404).json({ error: 'Aucun code sous-segment correspondant.' });
    }
    res.json({ matches });
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors de la recherche.' });
  }
});

// Valeurs pré-remplies pour un template donné, à partir d'une ligne de conditions résolue.
router.get('/mapped-values', async (req, res) => {
  try {
    const conditionsVersionId = req.query.conditionsVersionId as string;
    const templateId = req.query.templateId as string;
    const rowIndex = Number(req.query.rowIndex);
    if (!conditionsVersionId || !templateId || Number.isNaN(rowIndex)) {
      return res.status(400).json({ error: 'conditionsVersionId, templateId et rowIndex requis.' });
    }
    const version = getConditionsVersion(conditionsVersionId);
    if (!version) return res.status(404).json({ error: 'Fichier de conditions introuvable.' });
    const { rows } = await readConditionsRows(version);
    const row = rows[rowIndex];
    if (!row) return res.status(404).json({ error: 'Ligne introuvable.' });

    const mappings: MappingRow[] = mappingsTable.find((m) => m.templateId === templateId);

    const values: Record<string, string> = {};
    // Variables dont la colonne mappée n'existe pas (ou plus) dans CE fichier
    // de conditions : la valeur est vide, mais il ne faut pas laisser croire
    // que la colonne était simplement vide dans les données.
    const colonnesManquantes: string[] = [];
    for (const m of mappings) {
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
    res.json({ values, colonnesManquantes });
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors de la lecture des valeurs.' });
  }
});

// Génération finale du contrat Word rempli + journalisation.
router.post('/download', async (req, res) => {
  try {
    const { templateId, conditionsVersionId, codeSousSegment, values, traitePar } = req.body as {
      templateId: string;
      conditionsVersionId: string;
      codeSousSegment: string;
      values: Record<string, string>;
      traitePar?: string;
    };
    if (!templateId || !conditionsVersionId || !codeSousSegment || !values) {
      return res.status(400).json({ error: 'Paramètres manquants pour la génération.' });
    }
    const template = templatesTable.getById(templateId);
    if (!template) return res.status(404).json({ error: 'Template introuvable.' });

    const filePath = path.join(DIR_TEMPLATES, template.cheminStockage);
    if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier template manquant sur le disque.' });

    const buffer = fs.readFileSync(filePath);
    const filled = await fillDocxTemplate(buffer, values);

    generationsTable.insert({
      id: uuid(),
      templateId,
      conditionsVersionId,
      codeSousSegment,
      dateGeneration: new Date().toISOString(),
      traitePar: traitePar || null,
    });

    // codeSousSegment vient du champ de recherche (saisie libre) : un guillemet ou un
    // caractère de contrôle briserait l'en-tête Content-Disposition construit à la main
    // (contrairement à res.download() utilisé ailleurs, qui échappe correctement via la
    // bibliothèque content-disposition). On le nettoie explicitement ici.
    const safeCode = codeSousSegment.replace(/[^\p{L}\p{N}_-]+/gu, '-');
    const outName = `${path.parse(template.nomFichier).name}_${safeCode}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(filled);
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors de la génération du contrat.' });
  }
});

// Journal des générations (traçabilité).
router.get('/log', (_req, res) => {
  const rows = generationsTable.all().sort((a, b) => (a.dateGeneration < b.dateGeneration ? 1 : -1));
  res.json(rows.slice(0, 200));
});

export default router;
