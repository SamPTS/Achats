import '@pnp/sp/files';
import '@pnp/sp/folders';
import { getSP } from './spClient';

/** Horodatage YYYYMMDD-HHMM utilisé pour le nommage des fichiers stockés — porté depuis server/src/storage.ts. */
export function timestampTag(d = new Date()): string {
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
}

const COMBINING_DIACRITICS = new RegExp('[\u0300-\u036f]', 'g');

/**
 * Nettoie un nom pour en faire un segment de chemin/fichier sûr. La cible de compilation SPFx
 * est ES5 : pas de \p{...} Unicode property escapes (nécessitent le flag /u, donc ES2018+) —
 * on retire les diacritiques via la plage des marques combinantes Unicode, comme le fait déjà
 * server/src/utils/docx.ts (sanitizeFileName) dans la version autonome de l'application.
 */
function slug(s: string): string {
  return s
    .normalize('NFD')
    .replace(COMBINING_DIACRITICS, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

/** Construit un nom de fichier stocké: {type}_{libelle}_{YYYYMMDD-HHMM}.{ext} — jamais écrasé, chaque dépôt est unique. */
export function buildStoredFilename(type: string, libelle: string, ext: string): string {
  const safeLibelle = slug(libelle).slice(0, 60) || 'fichier';
  return `${type}_${safeLibelle}_${timestampTag()}.${ext}`;
}

/** Nettoie un segment de nom de fichier issu d'une saisie libre (ex : code sous-segment). */
export function safeFileNamePart(s: string): string {
  return slug(s) || 'x';
}

/** Dépose un fichier dans la bibliothèque indiquée et retourne son chemin relatif au serveur.
 *
 * addUsingPath() peut, comme items.add() (voir conditionsService.ts/templatesService.ts),
 * renvoyer une réponse vide/incomplète (observé en conditions réelles : ServerRelativeUrl absent
 * de la réponse) — le champ « cheminStockage » de l'enregistrement se retrouvait alors vide,
 * cassant silencieusement tout lien de téléchargement vers ce fichier sans qu'aucune erreur ne
 * soit levée au dépôt (un <a href=""> rouvre simplement la page courante). Si la réponse ne
 * contient pas ServerRelativeUrl, on relit le fichier qu'on vient de déposer directement dans la
 * bibliothèque, par son nom (unique par dépôt, horodaté à la minute), plutôt que de stocker une
 * valeur vide sans avertir personne. */
export async function uploadToLibrary(libraryTitle: string, fileName: string, content: ArrayBuffer): Promise<string> {
  const sp = getSP();
  const files = sp.web.lists.getByTitle(libraryTitle).rootFolder.files;
  const info = await files.addUsingPath(fileName, content, { Overwrite: false });
  if (info && info.ServerRelativeUrl) return info.ServerRelativeUrl;

  const reread = (await files.getByUrl(fileName).select('ServerRelativeUrl')()) as { ServerRelativeUrl?: string };
  if (!reread || !reread.ServerRelativeUrl) {
    throw new Error(
      `Le fichier "${fileName}" a été déposé mais son chemin de stockage n'a pas pu être déterminé (réponse SharePoint incomplète).`,
    );
  }
  return reread.ServerRelativeUrl;
}

/** Télécharge le contenu binaire d'un fichier à partir de son chemin relatif au serveur. */
export async function downloadFromServerRelativeUrl(serverRelativeUrl: string): Promise<ArrayBuffer> {
  const sp = getSP();
  return sp.web.getFileByServerRelativePath(serverRelativeUrl).getBuffer();
}

/** Supprime définitivement un fichier à partir de son chemin relatif au serveur. N'échoue pas si
 * le fichier est déjà absent (cohérent avec fs.unlinkSync + existsSync côté standalone : un
 * enregistrement dont le fichier a disparu ne doit pas empêcher sa suppression). */
export async function deleteByServerRelativeUrl(serverRelativeUrl: string): Promise<void> {
  const sp = getSP();
  try {
    await sp.web.getFileByServerRelativePath(serverRelativeUrl).delete();
  } catch {
    // Fichier déjà absent, ou droits insuffisants : on laisse l'appelant supprimer quand même
    // l'enregistrement plutôt que de bloquer sur un fichier orphelin.
  }
}
