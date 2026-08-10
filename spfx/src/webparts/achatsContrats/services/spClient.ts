import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/site-users/web';

let cachedSp: SPFI | undefined;

/** Instance PnPjs partagée, initialisée une seule fois à partir du contexte du web part. */
export function getSp(context: WebPartContext): SPFI {
  if (!cachedSp) {
    cachedSp = spfi().using(SPFx(context));
  }
  return cachedSp;
}

export interface IListsConfig {
  conditionsListTitle: string;
  templatesListTitle: string;
  mappingsListTitle: string;
  generationsListTitle: string;
  documentLibraryTitle: string;
}

export const DEFAULT_LISTS_CONFIG: IListsConfig = {
  conditionsListTitle: 'AchatsContrats_ConditionsVersions',
  templatesListTitle: 'AchatsContrats_TemplatesContrats',
  mappingsListTitle: 'AchatsContrats_MappingsTemplate',
  generationsListTitle: 'AchatsContrats_Generations',
  documentLibraryTitle: 'AchatsContrats_Fichiers',
};
