/** Horodatage YYYYMMDD-HHMM utilisé pour le nommage des fichiers stockés. */
export function timestampTag(d: Date = new Date()): string {
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(
    d.getMinutes()
  )}`;
}

/** Construit un nom de fichier stocké: {type}_{libelle}_{YYYYMMDD-HHMM}.{ext} */
export function buildStoredFilename(type: string, libelle: string, ext: string): string {
  const safeLibelle =
    libelle
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-zA-Z0-9_-]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60) || 'fichier';
  return `${type}_${safeLibelle}_${timestampTag()}.${ext}`;
}
