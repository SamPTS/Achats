import { SPFI } from '@pnp/sp';
import { IListsConfig } from './spClient';

/** Ajoute un champ à une liste sans échouer si celui-ci existe déjà (dépôt idempotent). */
async function safe(action: () => Promise<unknown>): Promise<void> {
  try {
    await action();
  } catch (e: any) {
    const message = String(e?.message ?? e);
    if (!/exists|duplicate|already/i.test(message)) {
      // Une erreur inattendue est tout de même journalisée en console, sans bloquer le reste.
      // eslint-disable-next-line no-console
      console.warn('Provisioning : ', message);
    }
  }
}

async function ensureConditionsList(sp: SPFI, title: string): Promise<void> {
  const { list } = await sp.web.lists.ensure(title, 'Historique des fichiers de conditions commerciales', 100);
  await safe(() => list.fields.addText('FichierUrl', { MaxLength: 400 }));
  await safe(() => list.fields.addMultilineText('ColonnesDetectees', { NumberOfLines: 6 }));
  await safe(() => list.fields.addText('ColonneCodeSousSegment'));
  await safe(() => list.fields.addBoolean('EstActive'));
  await safe(() => list.fields.addText('DeposePar'));
  await safe(() => list.fields.addNumber('NbLignes'));
  await safe(() => list.fields.addNumber('NbLignesVides'));
  await safe(() => list.fields.addNumber('NbColonnes'));
  await safe(() => list.fields.addNumber('DoublonsDetectes'));
  await safe(() => list.fields.addBoolean('Archive'));
}

async function ensureTemplatesList(sp: SPFI, title: string): Promise<void> {
  const { list } = await sp.web.lists.ensure(title, 'Historique des templates de contrats Word', 100);
  await safe(() => list.fields.addText('GroupId'));
  await safe(() => list.fields.addText('Departement'));
  await safe(() => list.fields.addText('FichierUrl', { MaxLength: 400 }));
  await safe(() => list.fields.addNumber('Version'));
  await safe(() => list.fields.addText('NomFichier'));
  await safe(() => list.fields.addMultilineText('VariablesDetectees', { NumberOfLines: 6 }));
  await safe(() => list.fields.addText('DeposePar'));
  await safe(() => list.fields.addBoolean('Archive'));
}

async function ensureMappingsList(sp: SPFI, title: string): Promise<void> {
  const { list } = await sp.web.lists.ensure(title, 'Correspondance variable ↔ colonne, par template', 100);
  await safe(() => list.fields.addNumber('TemplateId'));
  await safe(() => list.fields.addText('Variable'));
  await safe(() => list.fields.addText('ColonneCorrespondante'));
  await safe(() =>
    list.fields.addChoice('Statut', { Choices: ['mappee', 'libre', 'manquante'] })
  );
}

async function ensureGenerationsList(sp: SPFI, title: string): Promise<void> {
  const { list } = await sp.web.lists.ensure(title, 'Journal des générations de contrats', 100);
  await safe(() => list.fields.addNumber('TemplateId'));
  await safe(() => list.fields.addNumber('ConditionsVersionId'));
  await safe(() => list.fields.addText('CodeSousSegment'));
  await safe(() => list.fields.addText('TraitePar'));
}

async function ensureDocumentLibrary(sp: SPFI, title: string): Promise<void> {
  const { list } = await sp.web.lists.ensure(title, 'Fichiers Excel/Word bruts (conditions et templates)', 101);
  await safe(() => list.rootFolder.folders.addUsingPath('Conditions'));
  await safe(() => list.rootFolder.folders.addUsingPath('Templates'));
}

/**
 * Crée les listes, colonnes et la bibliothèque de documents nécessaires si elles n'existent pas déjà.
 * Idempotent : peut être appelé à chaque chargement du web part sans effet de bord si tout est déjà en place.
 */
export async function ensureSharePointStructure(sp: SPFI, config: IListsConfig): Promise<void> {
  await Promise.all([
    ensureConditionsList(sp, config.conditionsListTitle),
    ensureTemplatesList(sp, config.templatesListTitle),
    ensureMappingsList(sp, config.mappingsListTitle),
    ensureGenerationsList(sp, config.generationsListTitle),
    ensureDocumentLibrary(sp, config.documentLibraryTitle),
  ]);
}
