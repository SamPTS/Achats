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

async function ensureField(
  list: ReturnType<ReturnType<typeof getSP>['web']['lists']['getByTitle']>,
  add: () => Promise<unknown>,
): Promise<void> {
  try {
    await add();
  } catch {
    // Le champ existe déjà (ou une autre contrainte bénigne) : on continue.
  }
}

let provisioned: Promise<void> | undefined;

/** Idempotent et mémoïsé pour la durée de vie du web part : à appeler avant tout accès aux données. */
export function ensureProvisioned(): Promise<void> {
  if (!provisioned) provisioned = doProvision();
  return provisioned;
}

async function doProvision(): Promise<void> {
  const sp = getSP();

  // --- ConditionsVersions ---
  {
    const ler = await sp.web.lists.ensure(LISTS.conditionsVersions, 'Versions historisées des fichiers de conditions commerciales', GENERIC_LIST_TEMPLATE);
    const list = sp.web.lists.getByTitle(LISTS.conditionsVersions);
    await ensureField(list, () => ler.list.fields.addText('NomFichier', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addDateTime('DateDepot'));
    await ensureField(list, () => ler.list.fields.addText('CheminStockage', { MaxLength: 500 }));
    await ensureField(list, () => ler.list.fields.addMultilineText('Colonnes', { NumberOfLines: 10, RichText: false }));
    await ensureField(list, () => ler.list.fields.addText('ColonneCodeSousSegment', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addBoolean('EstActive'));
    await ensureField(list, () => ler.list.fields.addText('DeposePar', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addNumber('NbLignes'));
    await ensureField(list, () => ler.list.fields.addNumber('NbLignesVides'));
    await ensureField(list, () => ler.list.fields.addNumber('NbColonnes'));
    await ensureField(list, () => ler.list.fields.addNumber('DoublonsDetectes'));
    await ensureField(list, () => ler.list.fields.addBoolean('Archive'));
  }

  // --- TemplatesContrats ---
  {
    const ler = await sp.web.lists.ensure(LISTS.templatesContrats, 'Templates de contrats déposés, versionnés par groupe', GENERIC_LIST_TEMPLATE);
    const list = sp.web.lists.getByTitle(LISTS.templatesContrats);
    await ensureField(list, () => ler.list.fields.addText('GroupId', { MaxLength: 100 }));
    await ensureField(list, () => ler.list.fields.addText('Libelle', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('Departement', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addDateTime('DateDepot'));
    await ensureField(list, () => ler.list.fields.addNumber('Version'));
    await ensureField(list, () => ler.list.fields.addText('NomFichier', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('CheminStockage', { MaxLength: 500 }));
    await ensureField(list, () => ler.list.fields.addMultilineText('Variables', { NumberOfLines: 10, RichText: false }));
    await ensureField(list, () => ler.list.fields.addText('DeposePar', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addBoolean('Archive'));
  }

  // --- MappingsTemplate ---
  {
    const ler = await sp.web.lists.ensure(LISTS.mappingsTemplate, 'Correspondance variable de template ↔ colonne de conditions', GENERIC_LIST_TEMPLATE);
    const list = sp.web.lists.getByTitle(LISTS.mappingsTemplate);
    await ensureField(list, () => ler.list.fields.addText('TemplateId', { MaxLength: 100 }));
    await ensureField(list, () => ler.list.fields.addText('Variable', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('ColonneCorrespondante', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('Statut', { MaxLength: 20 }));
  }

  // --- Generations (journal, jamais modifié après écriture) ---
  {
    const ler = await sp.web.lists.ensure(LISTS.generations, 'Journal des contrats générés', GENERIC_LIST_TEMPLATE);
    const list = sp.web.lists.getByTitle(LISTS.generations);
    await ensureField(list, () => ler.list.fields.addText('TemplateId', { MaxLength: 100 }));
    await ensureField(list, () => ler.list.fields.addText('TemplateLibelle', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('ConditionsVersionId', { MaxLength: 100 }));
    await ensureField(list, () => ler.list.fields.addText('ConditionsNomFichier', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addText('CodeSousSegment', { MaxLength: 255 }));
    await ensureField(list, () => ler.list.fields.addDateTime('DateGeneration'));
    await ensureField(list, () => ler.list.fields.addText('TraitePar', { MaxLength: 255 }));
  }

  // --- Bibliothèques de documents ---
  await sp.web.lists.ensure(LIBRARIES.conditionsFichiers, 'Fichiers Excel de conditions commerciales déposés', DOCUMENT_LIBRARY_TEMPLATE);
  await sp.web.lists.ensure(LIBRARIES.templatesFichiers, 'Templates Word déposés', DOCUMENT_LIBRARY_TEMPLATE);
}
