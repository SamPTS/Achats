import path from 'path';

/**
 * Répertoire de base de l'application : le dossier contenant l'exécutable quand l'application
 * tourne packagée (via `pkg`, ou compilée en exécutable autonome avec `bun build --compile`),
 * sinon le dossier `server/` en développement. Toutes les données écrites sur disque (base
 * JSON, fichiers déposés) doivent être résolues à partir de ce dossier plutôt que de
 * `__dirname`, qui pointe dans le système de fichiers virtuel (lecture seule) de l'exécutable
 * une fois packagé.
 */
export function getAppDir(): string {
  const pkgProcess = process as unknown as { pkg?: unknown };
  // `globalThis.Bun` n'existe que lorsque le code tourne sur le runtime Bun — c'est-à-dire
  // uniquement l'exécutable compilé ici (le développement local utilise ts-node-dev sur
  // Node), donc process.execPath y pointe bien vers l'exécutable lui-même.
  const isBun = typeof (globalThis as { Bun?: unknown }).Bun !== 'undefined';
  if (pkgProcess.pkg || isBun) {
    return path.dirname(process.execPath);
  }
  return path.join(__dirname, '..');
}
