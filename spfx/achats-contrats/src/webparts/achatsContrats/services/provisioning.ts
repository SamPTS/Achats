import { getSP } from './spClient';

/**
 * Auto-provisionnement des listes et bibliothèques SharePoint nécessaires à
 * l'application, exécuté au premier chargement du web part.
 *
 * Ce choix (plutôt qu'un script de déploiement séparé à exécuter manuellement
 * par un administrateur du site) évite une étape d'installation distincte :
 * la première personne qui ouvre le web part sur le site déclenche la
 * création de ce qui manque. Chaque création est idempotente (vérifie
 * l'existence avant de créer, avale les erreurs "le champ existe déjà").
 *
 * ATTENTION : nécessite que l'utilisateur qui charge le web part pour la
 * première fois dispose des droits de gestion de listes sur le site
 * (Propriétaire/Membre avec droits de conception) — un simple Lecteur ne
 * pourra pas provisionner. Une fois les listes créées, les permissions
 * standard des listes SharePoint (héritées du site) s'appliquent pour tous
 * les usages suivants (lecture/écriture des éléments), sans droit de
 * conception nécessaire.
 */

export const LISTS = {
  conditionsVersions: 'ConditionsVersions',
  templatesContrats: 'TemplatesContrats',
  mappingsTemplate: 'MappingsTemplate',
  generations: 'Generations',
} as const;

export const LIBRARIES = {
  conditionsFichiers: 'ConditionsFichiers',
  templatesFichiers: 'TemplatesFichiers',
} as const;

const GENERIC_LIST_TEMPLATE = 100;
const DOCUMENT_LIBRARY_TEMPLATE = 101;

async function ensureField(add: () => Promise<unknown>): Promise<void> {
  try {
    await add();
  } catch {
    // Le champ existe déjà (ou une autre contrainte bénigne) : on continue.
  }
}

/**
 * Réessaie un appel qui échoue, avec un délai croissant entre chaque tentative. Utilisé par les
 * services pour toute lecture susceptible de suivre de près une écriture (provisionnement à froid
 * d'une liste, mais aussi — observé en conditions réelles sur un tenant — une simple relecture de
 * liste juste après l'ajout d'une ligne, ex. juste après le dépôt d'un fichier de conditions) : la
 * propagation d'une écriture SharePoint (nouvelle liste, nouveau champ, ou simplement une nouvelle
 * ligne) n'est pas toujours immédiatement visible à une lecture filtrée qui suit de très près,
 * d'où l'erreur générique PnPjs "Cannot read properties of undefined (reading 'Id')" observée de
 * façon reproductible juste après un dépôt. Un seul essai de rattrapage (1,5 s) s'est avéré
 * insuffisant dans ce cas précis (délai de propagation plus long qu'après un simple provisionnement
 * de liste) — jusqu'à 3 tentatives au total, délais croissants (1,5 s / 3 s). Ne masque jamais un
 * échec réellement persistant : si toutes les tentatives échouent, l'erreur d'origine de la
 * dernière tentative remonte normalement à l'appelant.
 */
export async function retryOnce<T>(fn: () => Promise<T>, attempts = 3): Promise<T> {
  let lastError: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      lastError = e;
      if (i < attempts - 1) await delay(1500 * (i + 1));
    }
  }
  throw lastError;
}

/**
 * Variante de retryOnce pour le cas où le délai de propagation ne se traduit pas par une erreur,
 * mais par une lecture qui « réussit » avec un résultat incomplet — observé en conditions réelles :
 * juste après le dépôt d'un template, les lignes de mapping tout juste créées peuvent être absentes
 * d'une lecture qui suit de très près, sans qu'aucune exception ne soit levée (contrairement au cas
 * couvert par retryOnce). isValid détermine si le résultat obtenu est exploitable ; si non, on
 * réessaie avec le même délai croissant. Si aucune tentative n'est valide, retourne quand même le
 * dernier résultat obtenu (jamais d'exception artificielle) — un mapping resté vide après toutes
 * les tentatives est un état légitime (template qui n'a simplement aucune variable), pas une erreur.
 */
export async function retryUntilValid<T>(fn: () => Promise<T>, isValid: (v: T) => boolean, attempts = 3): Promise<T> {
  let last: T | undefined;
  let lastError: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      last = await fn();
      if (isValid(last)) return last;
    } catch (e) {
      lastError = e;
    }
    if (i < attempts - 1) await delay(1500 * (i + 1));
  }
  if (last !== undefined) return last;
  throw lastError;
}

let provisioned: Promise<void> | undefined;

/**
 * Détecte l'erreur SharePoint "La liste «X» n'existe pas sur le site associé à l'URL «...»" —
 * observée en conditions réelles alors que les listes avaient bien été créées par le passé sur ce
 * site (provisionnement mémorisé dans le localStorage du navigateur, donc jamais rejoué). Cause la
 * plus probable : une liste supprimée après coup (manuellement, ou par une opération de ménage
 * côté site) sans que le navigateur n'en soit informé. Reconnaît aussi l'équivalent anglais, au
 * cas où la langue du site diffère.
 */
function isListMissingError(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e);
  return /n['’]existe pas sur le site associ/i.test(msg) || /does not exist at site with URL/i.test(msg);
}

/** Oublie le provisionnement mémorisé (mémoire + localStorage), pour forcer une reprovision complète. */
function forgetProvisioning(): void {
  provisioned = undefined;
  try {
    window.localStorage.removeItem(storageKey());
  } catch {
    // Rien de grave si on ne peut pas l'effacer : ensureProvisioned() relira quand même la même
    // clé au prochain appel et échouera pareil si la liste manque toujours, sans boucle infinie
    // puisque withListRecovery ne retente qu'une seule fois.
  }
}

/**
 * Enveloppe une lecture pour la rendre résiliente à une liste supprimée après coup malgré un
 * provisionnement mémorisé comme déjà fait : si l'erreur SharePoint indique clairement que la
 * liste n'existe plus, oublie ce provisionnement, le relance en entier, puis retente une seule
 * fois l'appel d'origine. Ne masque aucune autre erreur (droits insuffisants, réseau, etc.), qui
 * remonte normalement à l'appelant sans nouvelle tentative.
 */
export async function withListRecovery<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (e) {
    if (!isListMissingError(e)) throw e;
    forgetProvisioning();
    await ensureProvisioned();
    return fn();
  }
}

/**
 * Idempotent et mémoïsé pour la durée de vie du web part : à appeler avant tout accès aux
 * données. La mémoïsation en mémoire (`provisioned`) ne couvre que l'onglet de navigateur en
 * cours — sans le repli localStorage ci-dessous, chaque rechargement de page relancerait
 * l'intégralité de la routine de provisionnement (6 sp.web.lists.ensure + une trentaine
 * d'ajouts de champs, chacun un aller-retour réseau vers SharePoint), pour ne rien faire de
 * plus qu'avaler des erreurs "existe déjà" — un coût inutile payé à chaque chargement de page,
 * par chaque utilisateur, indéfiniment, avec un risque réel de throttling SharePoint (429) en
 * cas d'usage simultané par plusieurs personnes.
 */
export function ensureProvisioned(): Promise<void> {
  if (!provisioned) provisioned = doProvisionOnce();
  return provisioned;
}

function storageKey(): string {
  const webUrl = getSP().web.toUrl();
  return `achats-contrats-provisioned-v1:${webUrl}`;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function doProvisionOnce(): Promise<void> {
  const key = storageKey();
  try {
    if (window.localStorage.getItem(key) === '1') return;
  } catch {
    // localStorage indisponible (navigation privée, politique du navigateur) : on ignore le
    // cache et on reprovisionne à chaque chargement, comportement précédent, jamais incorrect.
  }

  await doProvision();

  // Sur un site où les listes viennent tout juste d'être créées (première ouverture jamais faite
  // ici), SharePoint met parfois quelques secondes à propager les métadonnées des listes/champs
  // fraîchement créés. Une lecture immédiate après provisionnement peut alors échouer avec une
  // erreur générique et mal formulée par PnPjs ("Cannot read properties of undefined (reading
  // 'Id')") plutôt qu'un message exploitable — observé en conditions réelles sur un tenant. Ce
  // court délai (seulement lors d'un provisionnement fraîchement effectué, jamais sur les
  // chargements suivants où ensureProvisioned() est instantané) laisse le temps à cette
  // propagation de se terminer avant que l'écran ne tente sa première lecture.
  await delay(3000);

  try {
    window.localStorage.setItem(key, '1');
  } catch {
    // Rien de grave si on ne peut pas persister le flag : au pire, reprovisionnement au
    // prochain chargement (idempotent, juste redondant).
  }
}

async function doProvision(): Promise<void> {
  const sp = getSP();

  // --- ConditionsVersions ---
  {
    const ler = await sp.web.lists.ensure(LISTS.conditionsVersions, 'Versions historisées des fichiers de conditions commerciales', GENERIC_LIST_TEMPLATE);
    await ensureField(() => ler.list.fields.addText('NomFichier', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addDateTime('DateDepot'));
    await ensureField(() => ler.list.fields.addText('CheminStockage', { MaxLength: 500 }));
    await ensureField(() => ler.list.fields.addMultilineText('Colonnes', { NumberOfLines: 10, RichText: false }));
    await ensureField(() => ler.list.fields.addText('ColonneCodeSousSegment', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addBoolean('EstActive'));
    await ensureField(() => ler.list.fields.addText('DeposePar', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addNumber('NbLignes'));
    await ensureField(() => ler.list.fields.addNumber('NbLignesVides'));
    await ensureField(() => ler.list.fields.addNumber('NbColonnes'));
    await ensureField(() => ler.list.fields.addNumber('DoublonsDetectes'));
    await ensureField(() => ler.list.fields.addBoolean('Archive'));
  }

  // --- TemplatesContrats ---
  {
    const ler = await sp.web.lists.ensure(LISTS.templatesContrats, 'Templates de contrats déposés, versionnés par groupe', GENERIC_LIST_TEMPLATE);
    await ensureField(() => ler.list.fields.addText('GroupId', { MaxLength: 100 }));
    await ensureField(() => ler.list.fields.addText('Libelle', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('Departement', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addDateTime('DateDepot'));
    await ensureField(() => ler.list.fields.addNumber('Version'));
    await ensureField(() => ler.list.fields.addText('NomFichier', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('CheminStockage', { MaxLength: 500 }));
    await ensureField(() => ler.list.fields.addMultilineText('Variables', { NumberOfLines: 10, RichText: false }));
    await ensureField(() => ler.list.fields.addText('DeposePar', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addBoolean('Archive'));
  }

  // --- MappingsTemplate ---
  {
    const ler = await sp.web.lists.ensure(LISTS.mappingsTemplate, 'Correspondance variable de template ↔ colonne de conditions', GENERIC_LIST_TEMPLATE);
    await ensureField(() => ler.list.fields.addText('TemplateId', { MaxLength: 100 }));
    await ensureField(() => ler.list.fields.addText('Variable', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('ColonneCorrespondante', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('Statut', { MaxLength: 20 }));
  }

  // --- Generations (journal, jamais modifié après écriture) ---
  {
    const ler = await sp.web.lists.ensure(LISTS.generations, 'Journal des contrats générés', GENERIC_LIST_TEMPLATE);
    await ensureField(() => ler.list.fields.addText('TemplateId', { MaxLength: 100 }));
    await ensureField(() => ler.list.fields.addText('TemplateLibelle', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('ConditionsVersionId', { MaxLength: 100 }));
    await ensureField(() => ler.list.fields.addText('ConditionsNomFichier', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addText('CodeSousSegment', { MaxLength: 255 }));
    await ensureField(() => ler.list.fields.addDateTime('DateGeneration'));
    await ensureField(() => ler.list.fields.addText('TraitePar', { MaxLength: 255 }));
  }

  // --- Bibliothèques de documents ---
  await sp.web.lists.ensure(LIBRARIES.conditionsFichiers, 'Fichiers Excel de conditions commerciales déposés', DOCUMENT_LIBRARY_TEMPLATE);
  await sp.web.lists.ensure(LIBRARIES.templatesFichiers, 'Templates Word déposés', DOCUMENT_LIBRARY_TEMPLATE);
}
