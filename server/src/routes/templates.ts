import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import db from '../db';
import { DIR_TEMPLATES, buildStoredFilename } from '../storage';
import { extractVariablesFromDocx } from '../utils/docx';
import { buildBlankMappingWorkbook, parseMappingFile } from '../utils/excel';
import { getActiveConditionsVersion, getConditionsVersion } from '../conditionsService';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
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
  depose_par: string | null;
  archive: number;
}

interface MappingRow {
  id: string;
  template_id: string;
  variable: string;
  colonne_correspondante: string | null;
  statut: 'mappee' | 'libre' | 'manquante';
}

function getMappings(templateId: string): MappingRow[] {
  return db.prepare('SELECT * FROM mappings_template WHERE template_id = ?').all(templateId) as MappingRow[];
}

function mappingStatus(mappings: MappingRow[]): 'complet' | 'incomplet' | 'absent' {
  if (mappings.length === 0) return 'absent';
  const allDone = mappings.every((m) => m.statut === 'mappee' || m.statut === 'libre');
  return allDone ? 'complet' : 'incomplet';
}

function templateToApi(row: TemplateRow) {
  const mappings = getMappings(row.id);
  return {
    id: row.id,
    groupId: row.group_id,
    libelle: row.libelle,
    departement: row.departement,
    dateDepot: row.date_depot,
    version: row.version,
    nomFichier: row.nom_fichier,
    variables: JSON.parse(row.variables_detectees) as string[],
    deposePar: row.depose_par,
    mappings: mappings.map((m) => ({
      variable: m.variable,
      colonneCorrespondante: m.colonne_correspondante,
      statut: m.statut,
    })),
    statutMapping: mappingStatus(mappings),
  };
}

// Liste : la dernière version de chaque groupe de template.
router.get('/', (_req, res) => {
  const rows = db
    .prepare(
      `SELECT t.* FROM templates_contrats t
       WHERE t.archive = 0 AND t.version = (
         SELECT MAX(t2.version) FROM templates_contrats t2 WHERE t2.group_id = t.group_id AND t2.archive = 0
       )
       ORDER BY t.date_depot DESC`
    )
    .all() as TemplateRow[];
  res.json(rows.map(templateToApi));
});

router.get('/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(req.params.id) as
    | TemplateRow
    | undefined;
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });
  res.json(templateToApi(row));
});

// Historique des versions d'un même template (groupe).
router.get('/group/:groupId/versions', (req, res) => {
  const rows = db
    .prepare('SELECT * FROM templates_contrats WHERE group_id = ? AND archive = 0 ORDER BY version DESC')
    .all(req.params.groupId) as TemplateRow[];
  res.json(rows.map(templateToApi));
});

// Dépôt d'un template Word (nouveau template, ou nouvelle version d'un groupe existant).
router.post('/', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier fourni.' });
    if (!/\.docx$/i.test(req.file.originalname)) {
      return res.status(400).json({ error: 'Seuls les fichiers .docx sont acceptés.' });
    }
    const libelle = (req.body.libelle as string)?.trim();
    if (!libelle) return res.status(400).json({ error: 'Le libellé du template est requis.' });

    let variables: string[];
    try {
      variables = await extractVariablesFromDocx(req.file.buffer);
    } catch (e: any) {
      return res.status(400).json({ error: `Fichier Word illisible ou invalide : ${e.message}` });
    }

    const groupId: string = (req.body.groupId as string) || uuid();
    const prevVersion = db
      .prepare('SELECT MAX(version) as v FROM templates_contrats WHERE group_id = ?')
      .get(groupId) as { v: number | null };
    const version = (prevVersion.v ?? 0) + 1;

    const storedName = buildStoredFilename('template', libelle, 'docx');
    fs.writeFileSync(path.join(DIR_TEMPLATES, storedName), req.file.buffer);

    const id = uuid();
    db.prepare(
      `INSERT INTO templates_contrats
       (id, group_id, libelle, departement, date_depot, chemin_stockage, version, nom_fichier,
        variables_detectees, depose_par, archive)
       VALUES (@id, @group_id, @libelle, @departement, @date_depot, @chemin_stockage, @version, @nom_fichier,
        @variables_detectees, @depose_par, 0)`
    ).run({
      id,
      group_id: groupId,
      libelle,
      departement: req.body.departement || null,
      date_depot: new Date().toISOString(),
      chemin_stockage: storedName,
      version,
      nom_fichier: req.file.originalname,
      variables_detectees: JSON.stringify(variables),
      depose_par: req.body.deposePar || null,
    });

    const insertMapping = db.prepare(
      `INSERT INTO mappings_template (id, template_id, variable, colonne_correspondante, statut)
       VALUES (?, ?, ?, NULL, 'manquante')`
    );
    const tx = db.transaction(() => {
      for (const v of variables) insertMapping.run(uuid(), id, v);
    });
    tx();

    const created = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(id) as TemplateRow;
    res.status(201).json(templateToApi(created));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du template.' });
  }
});

router.get('/:id/download', (req, res) => {
  const row = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(req.params.id) as
    | TemplateRow
    | undefined;
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });
  const filePath = path.join(DIR_TEMPLATES, row.chemin_stockage);
  if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier manquant sur le disque.' });
  res.download(filePath, row.nom_fichier);
});

// Génère le fichier de mapping vierge (Variable | Colonne correspondante | Colonnes disponibles).
router.get('/:id/mapping/blank', async (req, res) => {
  const row = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(req.params.id) as
    | TemplateRow
    | undefined;
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });

  const conditionsVersionId = (req.query.conditionsVersionId as string) || undefined;
  const conditionsVersion = conditionsVersionId
    ? getConditionsVersion(conditionsVersionId)
    : getActiveConditionsVersion();
  const colonnes = conditionsVersion ? JSON.parse(conditionsVersion.colonnes_detectees) : [];

  const variables: string[] = JSON.parse(row.variables_detectees);
  const buffer = await buildBlankMappingWorkbook(variables, colonnes);
  res.setHeader('Content-Disposition', `attachment; filename="mapping_${row.libelle}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

// Dépôt du fichier de mapping rempli : validation contre le référentiel de colonnes.
router.post('/:id/mapping', upload.single('file'), async (req, res) => {
  try {
    const row = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(req.params.id) as
      | TemplateRow
      | undefined;
    if (!row) return res.status(404).json({ error: 'Template introuvable.' });
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier fourni.' });

    const conditionsVersionId = (req.body.conditionsVersionId as string) || undefined;
    const conditionsVersion = conditionsVersionId
      ? getConditionsVersion(conditionsVersionId)
      : getActiveConditionsVersion();
    const colonnesRef: string[] = conditionsVersion
      ? JSON.parse(conditionsVersion.colonnes_detectees)
      : [];

    let parsedRows;
    try {
      parsedRows = await parseMappingFile(req.file.buffer);
    } catch (e: any) {
      return res.status(400).json({ error: `Fichier de mapping illisible : ${e.message}` });
    }

    const variables: string[] = JSON.parse(row.variables_detectees);
    const seenVariables = new Set<string>();
    const seenColonnes = new Set<string>();
    const erreurs: string[] = [];
    const byVariable = new Map<string, string>();

    for (const r of parsedRows) {
      if (!variables.includes(r.variable)) {
        erreurs.push(`Variable inconnue dans ce template : "${r.variable}".`);
        continue;
      }
      if (seenVariables.has(r.variable)) {
        erreurs.push(`Variable en double dans le mapping : "${r.variable}".`);
      }
      seenVariables.add(r.variable);

      if (r.colonne) {
        if (colonnesRef.length > 0 && !colonnesRef.includes(r.colonne)) {
          erreurs.push(`Colonne inexistante dans le référentiel de conditions : "${r.colonne}".`);
          continue;
        }
        if (seenColonnes.has(r.colonne)) {
          erreurs.push(`Colonne mappée plusieurs fois : "${r.colonne}".`);
        }
        seenColonnes.add(r.colonne);
        byVariable.set(r.variable, r.colonne);
      }
    }

    if (erreurs.length > 0) {
      return res.status(400).json({ error: 'Incohérences détectées dans le mapping.', details: erreurs });
    }

    const tx = db.transaction(() => {
      for (const variable of variables) {
        const colonne = byVariable.get(variable) ?? null;
        const statut = colonne ? 'mappee' : 'manquante';
        db.prepare(
          'UPDATE mappings_template SET colonne_correspondante = ?, statut = ? WHERE template_id = ? AND variable = ?'
        ).run(colonne, statut, row.id, variable);
      }
    });
    tx();

    const updated = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(row.id) as TemplateRow;
    res.json(templateToApi(updated));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du mapping.' });
  }
});

// Édition manuelle d'une ligne de mapping (ex: marquer une variable comme "saisie libre").
router.patch('/:id/mapping/:variable', (req, res) => {
  const row = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(req.params.id) as
    | TemplateRow
    | undefined;
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });

  const { colonneCorrespondante, statut } = req.body as {
    colonneCorrespondante?: string | null;
    statut?: 'mappee' | 'libre' | 'manquante';
  };
  const finalStatut = statut ?? (colonneCorrespondante ? 'mappee' : 'manquante');

  const result = db
    .prepare(
      'UPDATE mappings_template SET colonne_correspondante = ?, statut = ? WHERE template_id = ? AND variable = ?'
    )
    .run(colonneCorrespondante ?? null, finalStatut, row.id, req.params.variable);
  if (result.changes === 0) return res.status(404).json({ error: 'Variable inconnue pour ce template.' });

  const updated = db.prepare('SELECT * FROM templates_contrats WHERE id = ?').get(row.id) as TemplateRow;
  res.json(templateToApi(updated));
});

export default router;
