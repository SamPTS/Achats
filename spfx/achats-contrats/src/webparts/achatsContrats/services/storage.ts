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

/** Dépose un fichier dans la bibliothèque indiquée et retourne son chemin relatif au serveur. */
export async function uploadToLibrary(libraryTitle: string, fileName: string, content: ArrayBuffer): Promise<string> {
  const sp = getSP();
  const info = await sp.web.lists.getByTitle(libraryTitle).rootFolder.files.addUsingPath(fileName, content, {
    Overwrite: false,
  });
  return info.ServerRelativeUrl;
}

/** Télécharge le contenu binaire d'un fichier à partir de son chemin relatif au serveur. */
export async function downloadFromServerRelativeUrl(serverRelativeUrl: string): Promise<ArrayBuffer> {
  const sp = getSP();
  return sp.web.getFileByServerRelativePath(serverRelativeUrl).getBuffer();
}
