import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './spClient';
import { ensureProvisioned, retryOnce, withListRecovery, LISTS, LIBRARIES } from './provisioning';
import { buildStoredFilename, uploadToLibrary, downloadFromServerRelativeUrl, deleteByServerRelativeUrl, getFileModified } from './storage';
import { parseConditionsFile, findDuplicateCodes } from './excel';
import { odataEscape } from './odata';
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
  FichierModifieLe: string | null;
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
  'FichierModifieLe',
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
  const items = (await withListRecovery(() =>
    retryOnce(() => list().items.select(...SELECT_FIELDS).filter('Archive eq 0').top(2000)()),
  )) as ConditionsVersionItem[];
  return items.map(toModel).sort((a, b) => (a.dateDepot < b.dateDepot ? 1 : -1));
}

// Cache en mémoire des lignes déjà lues, par identifiant de version (voir readConditionsRows plus
// bas). Déclaré avant resyncIfFileChanged, qui l'invalide, pour rester lisible dans l'ordre
// d'exécution.
const rowsCache = new Map<string, Record<string, string>[]>();

export async function getConditionsVersion(id: string): Promise<ConditionsVersion | undefined> {
  await ensureProvisioned();
  try {
    const item = (await withListRecovery(() =>
      list().items.getById(Number(id)).select(...SELECT_FIELDS)(),
    )) as ConditionsVersionItem;
    return toModel(await resyncIfFileChanged(item));
  } catch {
    return undefined;
  }
}

/** Détecte une édition du fichier Excel faite directement dans la bibliothèque de documents
 * (hors dépôt via l'application) en comparant sa date de modification actuelle à celle enregistrée
 * lors du dernier calcul des métadonnées (Colonnes, NbLignes, ...). Si elle diffère, recalcule ces
 * métadonnées à partir du contenu actuel du fichier et les persiste, plutôt que de continuer à
 * afficher des colonnes/statistiques qui ne correspondent plus au fichier réel — voir la discussion
 * avec l'utilisateur : héberger les fichiers dans une bibliothèque SharePoint permet une édition
 * directe, que l'application ne peut pas empêcher, seulement détecter et rattraper. N'échoue jamais
 * l'appelant si la resynchronisation elle-même échoue (droits insuffisants, fichier verrouillé...) :
 * dans ce cas, l'élément d'origine (potentiellement périmé) est retourné tel quel plutôt que de
 * bloquer toute lecture. */
async function resyncIfFileChanged(item: ConditionsVersionItem): Promise<ConditionsVersionItem> {
  const currentModified = await getFileModified(item.CheminStockage);
  if (!currentModified || currentModified === item.FichierModifieLe) return item;
  try {
    const buffer = await downloadFromServerRelativeUrl(item.CheminStockage);
    const parsed = await parseConditionsFile(buffer);
    const doublons = parsed.colonneCodeCandidate ? findDuplicateCodes(parsed.rows, parsed.colonneCodeCandidate) : 0;
    // La colonne de code sous-segment choisie (manuellement ou automatiquement) est conservée si
    // elle existe toujours dans le fichier modifié ; sinon on retombe sur la détection automatique
    // faite sur le contenu actuel (peut-être vide, si aucune colonne candidate n'est reconnue).
    const colonneCodeSousSegment =
      item.ColonneCodeSousSegment && parsed.colonnes.includes(item.ColonneCodeSousSegment)
        ? item.ColonneCodeSousSegment
        : parsed.colonneCodeCandidate;
    const patch = {
      Colonnes: JSON.stringify(parsed.colonnes),
      ColonneCodeSousSegment: colonneCodeSousSegment,
      NbLignes: parsed.rows.length,
      NbLignesVides: parsed.nbLignesVides,
      NbColonnes: parsed.colonnes.length,
      DoublonsDetectes: doublons,
      FichierModifieLe: currentModified,
    };
    await list().items.getById(item.Id).update(patch);
    rowsCache.delete(String(item.Id));
    // eslint-disable-next-line no-console
    console.warn(
      `[achats-contrats] Fichier de conditions modifié directement dans la bibliothèque (hors dépôt via l'application) : métadonnées recalculées pour "${item.NomFichier}".`,
    );
    return { ...item, ...patch };
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error('[achats-contrats] Échec de la resynchronisation automatique des métadonnées de conditions :', e);
    return item;
  }
}

export async function getActiveConditionsVersion(): Promise<ConditionsVersion | undefined> {
  await ensureProvisioned();
  const items = (await withListRecovery(() =>
    retryOnce(() => list().items.select(...SELECT_FIELDS).filter('EstActive eq 1 and Archive eq 0').top(1)()),
  )) as ConditionsVersionItem[];
  return items[0] ? toModel(items[0]) : undefined;
}

/** Relit le fichier Excel stocké et retourne les lignes de données. Mis en cache par version (voir
 * rowsCache) — une version de conditions est normalement immuable une fois créée (un nouveau
 * dépôt crée toujours une nouvelle version, jamais de modification en place), sauf édition directe
 * du fichier dans la bibliothèque : dans ce cas resyncIfFileChanged (déclenché par un appel
 * préalable à getConditionsVersion) invalide l'entrée correspondante avant qu'elle ne soit relue
 * ici. Sans ce cache, observé en conditions réelles : chaque recherche de code ET chaque résolution
 * de ligne retéléchargeaient et reparsaient l'intégralité du fichier Excel depuis SharePoint
 * (potentiellement plusieurs centaines de colonnes) — deux fois par génération de contrat, et à
 * nouveau à chaque nouvelle recherche, sans jamais réutiliser un résultat déjà obtenu. Sur un
 * fichier volumineux, l'aller-retour réseau (télécharger le fichier) domine largement le temps de
 * calcul du parsing lui-même. */
export async function readConditionsRows(version: ConditionsVersion): Promise<{ rows: Record<string, string>[] }> {
  const cached = rowsCache.get(version.id);
  if (cached) return { rows: cached };

  const buffer = await downloadFromServerRelativeUrl(version.cheminStockage);
  const parsed = await parseConditionsFile(buffer);
  rowsCache.set(version.id, parsed.rows);
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
  // Référence de départ pour la détection de modification directe dans la bibliothèque (voir
  // resyncIfFileChanged) — null si indisponible (jamais bloquant, juste une resynchronisation
  // possible dès la prochaine lecture au lieu d'être différée à la modification suivante).
  const fichierModifieLe = await getFileModified(cheminStockage);

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
    FichierModifieLe: fichierModifieLe,
  });

  // items.add() renvoie directement l'élément créé (avec .Id à la racine), pas un objet enveloppé
  // {data: {...}} — cette hypothèse de forme incorrecte (iar.data.Id) provoquait une exception
  // "Cannot read properties of undefined (reading 'Id')" systématique à chaque dépôt, à tort
  // attribuée à un délai de propagation SharePoint (voir historique de ce fichier) : en réalité
  // iar.data était toujours undefined, quel que soit le nombre de tentatives.
  // Défense supplémentaire contre une réponse d'ajout vide/incomplète (observé une fois sur une
  // liste tout juste créée) : recherche par CheminStockage, unique par dépôt (nom horodaté à la
  // minute), plutôt que planter ou retenter l'ajout au risque de créer un doublon.
  const newId =
    iar && typeof iar.Id === 'number'
      ? iar.Id
      : await (async () => {
          const found = (await retryOnce(async () => {
            const rows = (await list()
              .items.select('Id')
              .filter(`CheminStockage eq '${odataEscape(cheminStockage)}'`)
              .top(1)()) as { Id: number }[];
            if (rows.length === 0) throw new Error('Version introuvable après création (réponse vide).');
            return rows[0];
          })) as { Id: number };
          return found.Id;
        })();

  const created = (await retryOnce(() =>
    list().items.getById(newId).select(...SELECT_FIELDS)(),
  )) as ConditionsVersionItem;
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

/** Suppression définitive (contrairement à archiveConditions, qui ne fait que masquer la
 * version) : supprime aussi le fichier Excel déposé. Refusée pour la version active, pour la
 * même raison que l'archivage. */
export async function deleteConditions(id: string): Promise<void> {
  await ensureProvisioned();
  const version = await getConditionsVersion(id);
  if (!version) throw new Error('Version introuvable.');
  if (version.estActive) {
    throw new Error("Impossible de supprimer la version active : activez une autre version au préalable.");
  }
  await deleteByServerRelativeUrl(version.cheminStockage);
  await list().items.getById(Number(id)).delete();
}
