import fs from 'fs';
import path from 'path';

export const STORAGE_ROOT = path.join(__dirname, '..', 'storage');
export const DIR_CONDITIONS = path.join(STORAGE_ROOT, 'conditions');
export const DIR_TEMPLATES = path.join(STORAGE_ROOT, 'templates');
export const DIR_MAPPINGS = path.join(STORAGE_ROOT, 'mappings');
export const DIR_GENERATED = path.join(STORAGE_ROOT, 'generated');

for (const dir of [DIR_CONDITIONS, DIR_TEMPLATES, DIR_MAPPINGS, DIR_GENERATED]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

/** Horodatage YYYYMMDD-HHMM utilisé pour le nommage des fichiers stockés. */
export function timestampTag(d = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
}

/** Construit un nom de fichier stocké: {type}_{libelle}_{YYYYMMDD-HHMM}.{ext} */
export function buildStoredFilename(type: string, libelle: string, ext: string): string {
  const safeLibelle = libelle
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || 'fichier';
  return `${type}_${safeLibelle}_${timestampTag()}.${ext}`;
}
