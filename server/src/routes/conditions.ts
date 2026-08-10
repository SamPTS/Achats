import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import { conditionsVersionsTable, ConditionsVersionRow } from '../db';
import { DIR_CONDITIONS, buildStoredFilename } from '../storage';
import { parseConditionsFile, findDuplicateCodes } from '../utils/excel';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const router = Router();

function rowToApi(row: ConditionsVersionRow) {
  return {
    id: row.id,
    nomFichier: row.nomFichier,
    dateDepot: row.dateDepot,
    colonnes: row.colonnes,
    colonneCodeSousSegment: row.colonneCodeSousSegment,
    estActive: row.estActive,
    deposePar: row.deposePar,
    nbLignes: row.nbLignes,
    nbLignesVides: row.nbLignesVides,
    nbColonnes: row.nbColonnes,
    doublonsDetectes: row.doublonsDetectes,
  };
}

// Liste des versions historisées, triées par date décroissante.
router.get('/', (_req, res) => {
  const rows = conditionsVersionsTable
    .find((v) => !v.archive)
    .sort((a, b) => (a.dateDepot < b.dateDepot ? 1 : -1));
  res.json(rows.map(rowToApi));
});

router.get('/:id', (req, res) => {
  const row = conditionsVersionsTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  res.json(rowToApi(row));
});

// Dépôt d'un nouveau fichier Excel de conditions commerciales.
router.post('/', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier fourni.' });
    const originalName = req.file.originalname;
    if (!/\.xlsx$/i.test(originalName)) {
      return res.status(400).json({ error: 'Seuls les fichiers .xlsx sont acceptés.' });
    }

    let parsed;
    try {
      parsed = await parseConditionsFile(req.file.buffer);
    } catch (e: any) {
      return res.status(400).json({ error: `Fichier Excel illisible ou invalide : ${e.message}` });
    }

    const colonneCode: string | null =
      (req.body.colonneCodeSousSegment as string) || parsed.colonneCodeCandidate;
    const doublons = colonneCode ? findDuplicateCodes(parsed.rows, colonneCode) : 0;

    const id = uuid();
    const libelle = path.parse(originalName).name;
    const storedName = buildStoredFilename('conditions', libelle, 'xlsx');
    fs.writeFileSync(path.join(DIR_CONDITIONS, storedName), req.file.buffer);

    const estActive = conditionsVersionsTable.find((v) => !v.archive).length === 0;

    const created = conditionsVersionsTable.insert({
      id,
      nomFichier: originalName,
      dateDepot: new Date().toISOString(),
      cheminStockage: storedName,
      colonnes: parsed.colonnes,
      colonneCodeSousSegment: colonneCode,
      estActive,
      deposePar: req.body.deposePar || null,
      nbLignes: parsed.rows.length,
      nbLignesVides: parsed.nbLignesVides,
      nbColonnes: parsed.colonnes.length,
      doublonsDetectes: doublons,
      archive: false,
    });

    res.status(201).json(rowToApi(created));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du fichier.' });
  }
});

// Désigner manuellement la colonne "code sous-segment", si non détectée automatiquement.
router.patch('/:id', (req, res) => {
  const row = conditionsVersionsTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  const { colonneCodeSousSegment } = req.body;
  if (colonneCodeSousSegment && !row.colonnes.includes(colonneCodeSousSegment)) {
    return res.status(400).json({ error: 'Colonne inconnue dans ce fichier.' });
  }
  const updated = conditionsVersionsTable.update(req.params.id, { colonneCodeSousSegment });
  res.json(rowToApi(updated!));
});

// Marquer une version comme active (une seule version active à la fois).
router.post('/:id/activate', (req, res) => {
  const row = conditionsVersionsTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  conditionsVersionsTable.updateWhere(() => true, { estActive: false });
  conditionsVersionsTable.update(req.params.id, { estActive: true });
  res.json({ ok: true });
});

// Archivage (masquage) plutôt que suppression définitive.
router.post('/:id/archive', (req, res) => {
  const row = conditionsVersionsTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  conditionsVersionsTable.update(req.params.id, { archive: true, estActive: false });
  res.json({ ok: true });
});

router.get('/:id/download', (req, res) => {
  const row = conditionsVersionsTable.getById(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  const filePath = path.join(DIR_CONDITIONS, row.cheminStockage);
  if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier manquant sur le disque.' });
  res.download(filePath, row.nomFichier);
});

export default router;
