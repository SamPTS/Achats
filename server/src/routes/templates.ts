import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import { templatesTable, mappingsTable, TemplateRow, MappingRow } from '../db';
import { DIR_TEMPLATES, buildStoredFilename } from '../storage';
import { extractVariablesFromDocx } from '../utils/docx';
import { buildBlankMappingWorkbook, parseMappingFile } from '../utils/excel';
import { getActiveConditionsVersion, getConditionsVersion } from '../conditionsService';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const router = Router();

function getMappings(templateId: string): MappingRow[] {
  return mappingsTable.find((m) => m.templateId === templateId);
}

/**
 * Une ligne "mappee" n'est valide que si sa colonne existe encore dans le
 * référentiel de colonnes fourni (le fichier de conditions actif). Si aucun
 * référentiel n'est fourni (pas de conditions active), on ne peut pas savoir
 * si la colonne existe : on ne pénalise pas ce cas pour ne pas casser
 * l'usage avant tout dépôt de conditions.
 */
function ligneValide(m: MappingRow, colonnesRef: string[] | null): boolean {
  if (m.statut === 'libre') return true;
  if (m.statut !== 'mappee' || !m.colonneCorrespondante) return false;
  if (colonnesRef && !colonnesRef.includes(m.colonneCorrespondante)) return false;
  return true;
}

function mappingStatus(
  mappings: MappingRow[],
  colonnesRef: string[] | null
): 'complet' | 'incomplet' | 'absent' {
  if (mappings.length === 0) return 'absent';
  const allDone = mappings.every((m) => ligneValide(m, colonnesRef));
  return allDone ? 'complet' : 'incomplet';
}

function templateToApi(row: TemplateRow) {
  const mappings = getMappings(row.id);
  const active = getActiveConditionsVersion();
  const colonnesRef = active ? active.colonnes : null;
  return {
    id: row.id,
    groupId: row.groupId,
    libelle: row.libelle,
    departement: row.departement,
    dateDepot: row.dateDepot,
    version: row.version,
    nomFichier: row.nomFichier,
    variables: row.variables,
    deposePar: row.deposePar,
    mappings: mappings.map((m) => ({
      variable: m.variable,
      colonneCorrespondante: m.colonneCorrespondante,
      statut: m.statut,
      // Mappée en base, mais la colonne n'existe plus dans le fichier de
      // conditions actuellement actif : à re-mapper avant génération.
      colonneIntrouvable:
        m.statut === 'mappee' &&
        !!m.colonneCorrespondante &&
        !!colonnesRef &&
        !colonnesRef.includes(m.colonneCorrespondante),
    })),
    statutMapping: mappingStatus(mappings, colonnesRef),
  };
}

// Liste : la dernière version de chaque groupe de template.
router.get('/', (_req, res) => {
  const rows = templatesTable.find((t) => !t.archive);
  const latestByGroup = new Map<string, TemplateRow>();
  for (const row of rows) {
    const current = latestByGroup.get(row.groupId);
    if (!current || row.version > current.version) latestByGroup.set(row.groupId, row);
  }
  const result = Array.from(latestByGroup.values()).sort((a, b) => (a.dateDepot < b.dateDepot ? 1 : -1));
  res.json(result.map(templateToApi));
});

router.get('/:id', (req, res) => {
  const row = templatesTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });
  res.json(templateToApi(row));
});

// Historique des versions d'un même template (groupe).
router.get('/group/:groupId/versions', (req, res) => {
  const rows = templatesTable
    .find((t) => t.groupId === req.params.groupId && !t.archive)
    .sort((a, b) => b.version - a.version);
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
    const previousVersions = templatesTable.find((t) => t.groupId === groupId);
    const version = previousVersions.reduce((max, t) => Math.max(max, t.version), 0) + 1;

    const storedName = buildStoredFilename('template', libelle, 'docx');
    fs.writeFileSync(path.join(DIR_TEMPLATES, storedName), req.file.buffer);

    const id = uuid();
    const created = templatesTable.insert({
      id,
      groupId,
      libelle,
      departement: req.body.departement || null,
      dateDepot: new Date().toISOString(),
      cheminStockage: storedName,
      version,
      nomFichier: req.file.originalname,
      variables,
      deposePar: req.body.deposePar || null,
      archive: false,
    });

    for (const v of variables) {
      mappingsTable.insert({
        id: uuid(),
        templateId: id,
        variable: v,
        colonneCorrespondante: null,
        statut: 'manquante',
      });
    }

    res.status(201).json(templateToApi(created));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du template.' });
  }
});

router.get('/:id/download', (req, res) => {
  const row = templatesTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });
  const filePath = path.join(DIR_TEMPLATES, row.cheminStockage);
  if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier manquant sur le disque.' });
  res.download(filePath, row.nomFichier);
});

// Suppression définitive de cette version de template : supprime aussi ses lignes de mapping et
// le fichier .docx déposé. Les générations déjà journalisées référencent le template par un
// simple identifiant texte (jamais réécrit après coup) et conservent leurs propres libellés
// figés au moment de la génération : les supprimer n'affecte pas le journal existant.
router.delete('/:id', (req, res) => {
  const row = templatesTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });
  const filePath = path.join(DIR_TEMPLATES, row.cheminStockage);
  if (fs.existsSync(filePath)) {
    try {
      fs.unlinkSync(filePath);
    } catch {
      // Fichier verrouillé ou droits insuffisants : on supprime quand même l'enregistrement.
    }
  }
  mappingsTable.deleteWhere((m) => m.templateId === req.params.id);
  templatesTable.deleteById(req.params.id);
  res.json({ ok: true });
});

// Génère le fichier de mapping vierge (Variable | Colonne correspondante | Colonnes disponibles).
router.get('/:id/mapping/blank', async (req, res) => {
  const row = templatesTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });

  const conditionsVersionId = (req.query.conditionsVersionId as string) || undefined;
  const conditionsVersion = conditionsVersionId
    ? getConditionsVersion(conditionsVersionId)
    : getActiveConditionsVersion();
  const colonnes = conditionsVersion ? conditionsVersion.colonnes : [];

  const buffer = await buildBlankMappingWorkbook(row.variables, colonnes);
  res.setHeader('Content-Disposition', `attachment; filename="mapping_${row.libelle}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

// Dépôt du fichier de mapping rempli : validation contre le référentiel de colonnes.
router.post('/:id/mapping', upload.single('file'), async (req, res) => {
  try {
    const row = templatesTable.getById(req.params.id);
    if (!row) return res.status(404).json({ error: 'Template introuvable.' });
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier fourni.' });

    const conditionsVersionId = (req.body.conditionsVersionId as string) || undefined;
    const conditionsVersion = conditionsVersionId
      ? getConditionsVersion(conditionsVersionId)
      : getActiveConditionsVersion();
    const colonnesRef: string[] = conditionsVersion ? conditionsVersion.colonnes : [];

    let parsedRows;
    try {
      parsedRows = await parseMappingFile(req.file.buffer);
    } catch (e: any) {
      return res.status(400).json({ error: `Fichier de mapping illisible : ${e.message}` });
    }

    const variables = row.variables;
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

    // Une variable marquée "saisie libre" avant l'import, et non renseignée dans le fichier de
    // mapping importé (ligne absente ou colonne vide), garde ce statut plutôt que de repasser
    // silencieusement à "manquante" : l'import ne doit modifier que les variables qu'il renseigne
    // effectivement, pas réinitialiser tout le reste du mapping.
    const mappingsAvant = getMappings(row.id);
    for (const variable of variables) {
      const colonne = byVariable.get(variable) ?? null;
      let statut: 'mappee' | 'libre' | 'manquante' = colonne ? 'mappee' : 'manquante';
      if (!colonne) {
        const avant = mappingsAvant.find((m) => m.variable === variable);
        if (avant?.statut === 'libre') statut = 'libre';
      }
      mappingsTable.updateWhere(
        (m) => m.templateId === row.id && m.variable === variable,
        { colonneCorrespondante: colonne, statut }
      );
    }

    const updated = templatesTable.getById(row.id)!;
    res.json(templateToApi(updated));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du mapping.' });
  }
});

// Édition manuelle d'une ligne de mapping (ex: marquer une variable comme "saisie libre").
router.patch('/:id/mapping/:variable', (req, res) => {
  const row = templatesTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Template introuvable.' });

  const { colonneCorrespondante, statut } = req.body as {
    colonneCorrespondante?: string | null;
    statut?: 'mappee' | 'libre' | 'manquante';
  };
  const finalStatut = statut ?? (colonneCorrespondante ? 'mappee' : 'manquante');

  const existing = mappingsTable.findOne((m) => m.templateId === row.id && m.variable === req.params.variable);
  if (!existing) return res.status(404).json({ error: 'Variable inconnue pour ce template.' });

  mappingsTable.update(existing.id, {
    colonneCorrespondante: colonneCorrespondante ?? null,
    statut: finalStatut,
  });

  const updated = templatesTable.getById(row.id)!;
  res.json(templateToApi(updated));
});

export default router;
