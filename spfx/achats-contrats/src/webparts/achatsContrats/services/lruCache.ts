/**
 * Cache en mémoire borné en taille (LRU minimal), utilisé pour les caches de contenu de fichier
 * déjà téléchargé (conditionsService.rowsCache, templatesService.templateFileCache). Ce sont des
 * variables de module — leur portée dépasse le cycle de vie des composants React, donc rien ne les
 * bornait ni ne les vidait sur une session longue (l'onglet du navigateur reste ouvert toute une
 * journée) : consulter plusieurs dizaines de versions/templates différents accumulait en mémoire
 * l'intégralité de chaque fichier Excel parsé et chaque .docx téléchargé, sans jamais en libérer
 * aucun — une dérive mémoire progressive. Une Map JS préserve l'ordre d'insertion ; en supprimant
 * puis réinsérant une clé consultée, elle se retrouve en fin d'itération, ce qui donne un LRU
 * simple sans structure de données dédiée.
 */
export class LruCache<K, V> {
  private readonly map = new Map<K, V>();

  constructor(private readonly maxEntries: number) {}

  get(key: K): V | undefined {
    const value = this.map.get(key);
    if (value === undefined) return undefined;
    this.map.delete(key);
    this.map.set(key, value);
    return value;
  }

  set(key: K, value: V): void {
    this.map.delete(key);
    this.map.set(key, value);
    while (this.map.size > this.maxEntries) {
      const oldest = this.map.keys().next().value;
      if (oldest === undefined) break;
      this.map.delete(oldest);
    }
  }

  delete(key: K): void {
    this.map.delete(key);
  }
}
