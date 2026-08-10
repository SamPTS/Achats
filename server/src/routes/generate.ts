import { Router } from 'express';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import db from '../db';
import { DIR_TEMPLATES } from '../storage';
import { fillDocxTemplate } from '../utils/docx';
import { getConditionsVersion, readConditionsRows } from '../conditionsService';

const router = Router();

interface TemplateRow {
  id: string;
  group_id: string;
  libelle: string;
  departement: string | null;
  date_depot: string;
  chemin_stockage: string;
  version: number;
  nom_fichier: string;
  variables_detectees: string;
  archive: number;
}

interface MappingRow {
  variable: string;
  colonne_correspondante: string | null;
  statut: 'mappee' | 'libre' | 'manquante';
}

function isMappingComplete(templateId: string): boolean {
  const mappings = db
    .prepare('SELECT * FROM mappings_template WHERE template_id = ?')
    .all(templateId) as MappingRow[];
  if (mappings.length === 0) return false;
  return mappings.every((m) => m.statut === 'mappee' || m.statut === 'libre');
}

// Templates disponibles pour la génération (mapping complet uniquement), dernière version par groupe.
router.get('/templates', (_req, res) => {
  const rows = db
    .prepare(
      `SELECT t.* FROM templates_contrats t
       WHERE t.archive = 0 AND t.version = (
         SELECT MAX(t2.version) FROM templates_contrats t2 WHERE t2.group_id = t.group_id AND t2.archive = 0
       )
       ORDER BY t.libelle ASC`
    )
    .all() as TemplateRow[];
  const usable = rows.filter((r) => isMappingComplete(r.id));
  res.json(
    usable.map((r) => ({
      id: r.id,
      libelle: r.libelle,
      departement: r.departement,
      version: r.version,
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
    if (!version.colonne_code_sous_segment) {
      return res.status(400).json({ error: 'La colonne "code sous-segment" n\'est pas définie pour ce fichier.' });
    }
    const { rows } = await readConditionsRows(version);
    const normalized = code.trim().toLowerCase();
    const colonne = version.colonne_code_sous_segment;

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

    const mappings = db
      .prepare('SELECT * FROM mappings_template WHERE template_id = ?')
      .all(templateId) as MappingRow[];

    const values: Record<string, string> = {};
    for (const m of mappings) {
      // Convention "zéro interprétation" : colonne vide -> champ vide, jamais de valeur devinée.
      values[m.variable] = m.statut === 'mappee' && m.colonne_correspondante ? row[m.colonne_correspondante] ?? '' : '';
    }
    res.json({ values });
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
    const template = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(templateId) as
      | TemplateRow
      | undefined;
    if (!template) return res.status(404).json({ error: 'Template introuvable.' });

    const filePath = path.join(DIR_TEMPLATES, template.chemin_stockage);
    if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier template manquant sur le disque.' });

    const buffer = fs.readFileSync(filePath);
    const filled = await fillDocxTemplate(buffer, values);

    db.prepare(
      `INSERT INTO generations (id, template_id, conditions_version_id, code_sous_segment, date_generation, traite_par)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).run(uuid(), templateId, conditionsVersionId, codeSousSegment, new Date().toISOString(), traitePar || null);

    const outName = `${path.parse(template.nom_fichier).name}_${codeSousSegment}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(filled);
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors de la génération du contrat.' });
  }
});

// Journal des générations (traçabilité).
router.get('/log', (_req, res) => {
  const rows = db.prepare('SELECT * FROM generations ORDER BY date_generation DESC LIMIT 200').all();
  res.json(rows);
});

export default router;
