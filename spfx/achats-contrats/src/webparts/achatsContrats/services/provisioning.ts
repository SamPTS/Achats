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
 * Réessaie une fois, après une courte pause, un appel qui échoue. Utilisé par les services pour
 * les toutes premières lectures suivant un provisionnement (voir doProvisionOnce ci-dessous) : la
 * propagation des métadonnées d'une liste/d'un champ tout juste créés par SharePoint peut prendre
 * quelques secondes, et une lecture immédiate peut échouer une fois avec une erreur générique
 * avant de fonctionner normalement à la tentative suivante — observé en conditions réelles sur un
 * tenant. Ne masque pas un échec persistant : si la deuxième tentative échoue aussi, l'erreur
 * d'origine remonte normalement à l'appelant.
 */
export async function retryOnce<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch {
    await delay(1500);
    return fn();
  }
}

let provisioned: Promise<void> | undefined;

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
