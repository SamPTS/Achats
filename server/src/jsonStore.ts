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
      } catch {
        this.items = [];
      }
    } else {
      this.persist();
    }
  }

  private persist(): void {
    fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
    fs.writeFileSync(this.filePath, JSON.stringify(this.items, null, 2), 'utf-8');
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
}
