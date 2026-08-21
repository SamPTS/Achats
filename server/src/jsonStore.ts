import fs from 'fs';
import path from 'path';

export interface Row {
  id: string;
  [key: string]: unknown;
}

/**
 * Table minimaliste persistée dans un fichier JSON. Remplace une vraie base relationnelle pour
 * cette application interne mono-utilisateur : aucune dépendance native, donc compatible avec un
 * packaging en exécutable autonome (`pkg`). L'ensemble de la table est chargé en mémoire et
 * réécrit intégralement à chaque mutation — largement suffisant pour les volumes visés ici.
 */
export class JsonTable<T extends Row> {
  private items: T[] = [];
  private readonly filePath: string;

  constructor(dataDir: string, name: string) {
    this.filePath = path.join(dataDir, `${name}.json`);
    this.load();
  }

  private load(): void {
    if (fs.existsSync(this.filePath)) {
      try {
        const raw = fs.readFileSync(this.filePath, 'utf-8');
        this.items = raw ? (JSON.parse(raw) as T[]) : [];
      } catch (e) {
        // Un fichier tronqué (écriture interrompue par un crash, une coupure de courant, ou le
        // processus tué en cours de persist()) faisait auparavant repartir silencieusement sur une
        // table vide, sans aucun avertissement — perte totale des données de cette table (et des
        // fichiers .docx/.xlsx associés, devenus orphelins) sans que personne ne s'en aperçoive
        // avant longtemps. On renomme le fichier corrompu de côté (préservé pour investigation/
        // récupération manuelle) et on fait échouer le démarrage avec un message explicite plutôt
        // que de continuer sur un état vide qui a l'air normal.
        const corruptedPath = `${this.filePath}.corrompu-${Date.now()}`;
        try {
          fs.renameSync(this.filePath, corruptedPath);
        } catch {
          // Le renommage lui-même peut échouer (droits, disque) : on lève quand même l'erreur
          // d'origine ci-dessous plutôt que de continuer sur un état vide.
        }
        throw new Error(
          `Le fichier de données "${this.filePath}" est corrompu et n'a pas pu être lu. ` +
            `Il a été renommé en "${corruptedPath}" pour investigation. Contactez un administrateur ` +
            `avant de continuer : redémarrer sans ce fichier repartirait sur une table vide. (${
              e instanceof Error ? e.message : String(e)
            })`,
        );
      }
    } else {
      this.persist();
    }
  }

  private persist(): void {
    fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
    // Écrit dans un fichier temporaire puis renomme vers le fichier final (rename atomique sur un
    // même volume), plutôt qu'une écriture directe dans le fichier final : sans ça, tuer le
    // processus (fermeture de la fenêtre console, coupure de courant) pendant l'écriture laisse un
    // JSON tronqué, détecté seulement au redémarrage suivant (voir load() ci-dessus) — avec ce
    // schéma, soit l'ancien fichier reste intact, soit le nouveau est complet ; jamais un état
    // intermédiaire à moitié écrit.
    const tmpPath = `${this.filePath}.tmp-${process.pid}`;
    fs.writeFileSync(tmpPath, JSON.stringify(this.items, null, 2), 'utf-8');
    fs.renameSync(tmpPath, this.filePath);
  }

  all(): T[] {
    return this.items.slice();
  }

  find(predicate: (item: T) => boolean): T[] {
    return this.items.filter(predicate);
  }

  findOne(predicate: (item: T) => boolean): T | undefined {
    return this.items.find(predicate);
  }

  getById(id: string): T | undefined {
    return this.items.find((i) => i.id === id);
  }

  insert(item: T): T {
    this.items.push(item);
    this.persist();
    return item;
  }

  update(id: string, patch: Partial<T>): T | undefined {
    const item = this.getById(id);
    if (!item) return undefined;
    Object.assign(item, patch);
    this.persist();
    return item;
  }

  updateWhere(predicate: (item: T) => boolean, patch: Partial<T>): void {
    let changed = false;
    for (const item of this.items) {
      if (predicate(item)) {
        Object.assign(item, patch);
        changed = true;
      }
    }
    if (changed) this.persist();
  }

  /** Variantes de update()/updateWhere() qui ne persistent pas immédiatement — à combiner avec
   * commit() pour appliquer plusieurs mutations en une seule écriture disque (une seule
   * persist()), plutôt qu'une écriture par mutation. Utile quand plusieurs mutations forment un
   * même changement logique (ex. activer une version ET désactiver les autres) : deux persist()
   * séparés laissent un état intermédiaire visible sur disque si le processus est interrompu entre
   * les deux (ex. aucune version active après un crash au milieu de la séquence). */
  updateNoPersist(id: string, patch: Partial<T>): T | undefined {
    const item = this.getById(id);
    if (!item) return undefined;
    Object.assign(item, patch);
    return item;
  }

  updateWhereNoPersist(predicate: (item: T) => boolean, patch: Partial<T>): void {
    for (const item of this.items) {
      if (predicate(item)) Object.assign(item, patch);
    }
  }

  /** Écrit sur disque l'état courant — à appeler après un ou plusieurs *NoPersist() pour qu'ils
   * soient effectivement persistés, en une seule écriture atomique (voir persist()). */
  commit(): void {
    this.persist();
  }

  /** Supprime définitivement l'élément (contrairement à archive: true, qui ne fait que le masquer). */
  deleteById(id: string): boolean {
    const before = this.items.length;
    this.items = this.items.filter((i) => i.id !== id);
    const changed = this.items.length !== before;
    if (changed) this.persist();
    return changed;
  }

  /** Supprime tous les éléments correspondant au prédicat, retourne le nombre supprimé. */
  deleteWhere(predicate: (item: T) => boolean): number {
    const before = this.items.length;
    this.items = this.items.filter((i) => !predicate(i));
    const removed = before - this.items.length;
    if (removed > 0) this.persist();
    return removed;
  }
}
