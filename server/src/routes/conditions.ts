import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import fs from 'fs';
import path from 'path';
import db from '../db';
import { DIR_CONDITIONS, buildStoredFilename } from '../storage';
import { parseConditionsFile, findDuplicateCodes } from '../utils/excel';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const router = Router();

function rowToApi(row: any) {
  return {
    id: row.id,
    nomFichier: row.nom_fichier,
    dateDepot: row.date_depot,
    colonnes: JSON.parse(row.colonnes_detectees),
    colonneCodeSousSegment: row.colonne_code_sous_segment,
    estActive: !!row.est_active,
    deposePar: row.depose_par,
    nbLignes: row.nb_lignes,
    nbLignesVides: row.nb_lignes_vides,
    nbColonnes: row.nb_colonnes,
    doublonsDetectes: row.doublons_detectes,
  };
}

// Liste des versions historisées, triées par date décroissante.
router.get('/', (_req, res) => {
  const rows = db
    .prepare('SELECT * FROM conditions_versions WHERE archive = 0 ORDER BY date_depot DESC')
    .all();
  res.json(rows.map(rowToApi));
});

router.get('/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
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

    const existingCount = db
      .prepare('SELECT COUNT(*) as n FROM conditions_versions WHERE archive = 0')
      .get() as { n: number };
    const estActive = existingCount.n === 0 ? 1 : 0;

    db.prepare(
      `INSERT INTO conditions_versions
       (id, nom_fichier, date_depot, chemin_stockage, colonnes_detectees, colonne_code_sous_segment,
        est_active, depose_par, nb_lignes, nb_lignes_vides, nb_colonnes, doublons_detectes, archive)
       VALUES (@id, @nom_fichier, @date_depot, @chemin_stockage, @colonnes_detectees, @colonne_code_sous_segment,
        @est_active, @depose_par, @nb_lignes, @nb_lignes_vides, @nb_colonnes, @doublons_detectes, 0)`
    ).run({
      id,
      nom_fichier: originalName,
      date_depot: new Date().toISOString(),
      chemin_stockage: storedName,
      colonnes_detectees: JSON.stringify(parsed.colonnes),
      colonne_code_sous_segment: colonneCode,
      est_active: estActive,
      depose_par: req.body.deposePar || null,
      nb_lignes: parsed.rows.length,
      nb_lignes_vides: parsed.nbLignesVides,
      nb_colonnes: parsed.colonnes.length,
      doublons_detectes: doublons,
    });

    const created = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(id);
    res.status(201).json(rowToApi(created));
  } catch (e: any) {
    res.status(500).json({ error: e.message || 'Erreur lors du dépôt du fichier.' });
  }
});

// Désigner manuellement la colonne "code sous-segment", si non détectée automatiquement.
router.patch('/:id', (req, res) => {
  const row: any = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  const colonnes: string[] = JSON.parse(row.colonnes_detectees);
  const { colonneCodeSousSegment } = req.body;
  if (colonneCodeSousSegment && !colonnes.includes(colonneCodeSousSegment)) {
    return res.status(400).json({ error: 'Colonne inconnue dans ce fichier.' });
  }
  db.prepare('UPDATE conditions_versions SET colonne_code_sous_segment = ? WHERE id = ?').run(
    colonneCodeSousSegment,
    req.params.id
  );
  const updated = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
  res.json(rowToApi(updated));
});

// Marquer une version comme active (une seule version active à la fois).
router.post('/:id/activate', (req, res) => {
  const row = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  const tx = db.transaction(() => {
    db.prepare('UPDATE conditions_versions SET est_active = 0').run();
    db.prepare('UPDATE conditions_versions SET est_active = 1 WHERE id = ?').run(req.params.id);
  });
  tx();
  res.json({ ok: true });
});

// Archivage (masquage) plutôt que suppression définitive.
router.post('/:id/archive', (req, res) => {
  const row = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  db.prepare('UPDATE conditions_versions SET archive = 1, est_active = 0 WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

router.get('/:id/download', (req, res) => {
  const row: any = db.prepare('SELECT * FROM conditions_versions WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'Version introuvable.' });
  const filePath = path.join(DIR_CONDITIONS, row.chemin_stockage);
  if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Fichier manquant sur le disque.' });
  res.download(filePath, row.nom_fichier);
});

export default router;
