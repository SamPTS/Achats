import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './spClient';
import { ensureProvisioned, retryOnce, LISTS, LIBRARIES } from './provisioning';
import { buildStoredFilename, uploadToLibrary, downloadFromServerRelativeUrl } from './storage';
import { parseConditionsFile, findDuplicateCodes } from './excel';
import type { ConditionsVersion } from '../model/types';

/** Forme brute d'un élément de la liste ConditionsVersions côté SharePoint. */
interface ConditionsVersionItem {
  Id: number;
  NomFichier: string;
  DateDepot: string;
  CheminStockage: string;
  Colonnes: string; // JSON.stringify(string[])
  ColonneCodeSousSegment: string | null;
  EstActive: boolean;
  DeposePar: string | null;
  NbLignes: number;
  NbLignesVides: number;
  NbColonnes: number;
  DoublonsDetectes: number;
  Archive: boolean;
}

const SELECT_FIELDS = [
  'Id',
  'NomFichier',
  'DateDepot',
  'CheminStockage',
  'Colonnes',
  'ColonneCodeSousSegment',
  'EstActive',
  'DeposePar',
  'NbLignes',
  'NbLignesVides',
  'NbColonnes',
  'DoublonsDetectes',
  'Archive',
];

function toModel(item: ConditionsVersionItem): ConditionsVersion {
  let colonnes: string[] = [];
  try {
    colonnes = JSON.parse(item.Colonnes || '[]');
  } catch {
    colonnes = [];
  }
  return {
    id: String(item.Id),
    nomFichier: item.NomFichier,
    dateDepot: item.DateDepot,
    cheminStockage: item.CheminStockage,
    colonnes,
    colonneCodeSousSegment: item.ColonneCodeSousSegment ?? null,
    estActive: !!item.EstActive,
    deposePar: item.DeposePar ?? null,
    nbLignes: item.NbLignes ?? 0,
    nbLignesVides: item.NbLignesVides ?? 0,
    nbColonnes: item.NbColonnes ?? 0,
    doublonsDetectes: item.DoublonsDetectes ?? 0,
    archive: !!item.Archive,
  };
}

function list() {
  return getSP().web.lists.getByTitle(LISTS.conditionsVersions);
}

/** Liste des versions non archivées, triées par date de dépôt décroissante. */
export async function listConditions(): Promise<ConditionsVersion[]> {
  await ensureProvisioned();
  const items = (await retryOnce(() =>
    list().items.select(...SELECT_FIELDS).filter('Archive eq 0').top(2000)(),
  )) as ConditionsVersionItem[];
  return items.map(toModel).sort((a, b) => (a.dateDepot < b.dateDepot ? 1 : -1));
}

export async function getConditionsVersion(id: string): Promise<ConditionsVersion | undefined> {
  await ensureProvisioned();
  try {
    const item = (await list()
      .items.getById(Number(id))
      .select(...SELECT_FIELDS)()) as ConditionsVersionItem;
    return toModel(item);
  } catch {
    return undefined;
  }
}

export async function getActiveConditionsVersion(): Promise<ConditionsVersion | undefined> {
  await ensureProvisioned();
  const items = (await retryOnce(() =>
    list().items.select(...SELECT_FIELDS).filter('EstActive eq 1 and Archive eq 0').top(1)(),
  )) as ConditionsVersionItem[];
  return items[0] ? toModel(items[0]) : undefined;
}

/** Relit le fichier Excel stocké et retourne les lignes de données (pas de cache, toujours à jour). */
export async function readConditionsRows(version: ConditionsVersion): Promise<{ rows: Record<string, string>[] }> {
  const buffer = await downloadFromServerRelativeUrl(version.cheminStockage);
  const parsed = await parseConditionsFile(buffer);
  return { rows: parsed.rows };
}

export async function uploadConditions(file: File, deposePar: string): Promise<ConditionsVersion> {
  await ensureProvisioned();
  if (!/\.xlsx$/i.test(file.name)) {
    throw new Error('Seuls les fichiers .xlsx sont acceptés.');
  }
  const buffer = await file.arrayBuffer();
  const parsed = await parseConditionsFile(buffer);

  const doublons = parsed.colonneCodeCandidate ? findDuplicateCodes(parsed.rows, parsed.colonneCodeCandidate) : 0;

  const storedName = buildStoredFilename('conditions', file.name.replace(/\.xlsx$/i, ''), 'xlsx');
  const cheminStockage = await uploadToLibrary(LIBRARIES.conditionsFichiers, storedName, buffer);

  // Première version jamais déposée (aucune version non archivée) : activée automatiquement.
  const existing = (await list().items.select('Id').filter('Archive eq 0').top(1)()) as unknown[];
  const estActive = existing.length === 0;

  const iar = await list().items.add({
    Title: file.name,
    NomFichier: file.name,
    DateDepot: new Date().toISOString(),
    CheminStockage: cheminStockage,
    Colonnes: JSON.stringify(parsed.colonnes),
    ColonneCodeSousSegment: parsed.colonneCodeCandidate,
    EstActive: estActive,
    DeposePar: deposePar || null,
    NbLignes: parsed.rows.length,
    NbLignesVides: parsed.nbLignesVides,
    NbColonnes: parsed.colonnes.length,
    DoublonsDetectes: doublons,
    Archive: false,
  });

  const created = (await list()
    .items.getById(iar.data.Id)
    .select(...SELECT_FIELDS)()) as ConditionsVersionItem;
  return toModel(created);
}

export async function setConditionsCodeColumn(id: string, colonneCodeSousSegment: string): Promise<void> {
  await ensureProvisioned();
  const version = await getConditionsVersion(id);
  if (!version) throw new Error('Version introuvable.');
  if (colonneCodeSousSegment && !version.colonnes.includes(colonneCodeSousSegment)) {
    throw new Error('Colonne inconnue dans ce fichier.');
  }
  await list().items.getById(Number(id)).update({ ColonneCodeSousSegment: colonneCodeSousSegment || null });
}

export async function activateConditions(id: string): Promise<void> {
  await ensureProvisioned();
  const actives = (await list().items.select('Id').filter('EstActive eq 1')()) as { Id: number }[];
  await Promise.all(actives.map((a) => list().items.getById(a.Id).update({ EstActive: false })));
  await list().items.getById(Number(id)).update({ EstActive: true });
}

/** Archiver la version active est refusé : voir server/src/routes/conditions.ts pour le pourquoi
 * (une application sans aucune version active désactive silencieusement la détection des
 * mappings devenus invalides). */
export async function archiveConditions(id: string): Promise<void> {
  await ensureProvisioned();
  const version = await getConditionsVersion(id);
  if (!version) throw new Error('Version introuvable.');
  if (version.estActive) {
    throw new Error("Impossible d'archiver la version active : activez une autre version au préalable.");
  }
  await list().items.getById(Number(id)).update({ Archive: true, EstActive: false });
}
