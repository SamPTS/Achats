import path from 'path';

/**
 * Répertoire de base de l'application : le dossier contenant l'exécutable quand l'application
 * tourne packagée (via `pkg`), sinon le dossier `server/` en développement. Toutes les données
 * écrites sur disque (base JSON, fichiers déposés) doivent être résolues à partir de ce dossier
 * plutôt que de `__dirname`, qui pointe dans le système de fichiers virtuel (lecture seule) de
 * l'exécutable une fois packagé.
 */
export function getAppDir(): string {
  const pkgProcess = process as unknown as { pkg?: unknown };
  if (pkgProcess.pkg) {
    return path.dirname(process.execPath);
  }
  return path.join(__dirname, '..');
}
