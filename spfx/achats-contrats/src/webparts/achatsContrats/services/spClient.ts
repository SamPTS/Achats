import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/files/folder';
import '@pnp/sp/folders';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

let _sp: SPFI | undefined;

/**
 * Point d'entrée unique vers PnPjs, configuré avec le contexte du web part
 * (authentification héritée de SharePoint, aucune gestion de jeton à la
 * main). À initialiser une fois avec le contexte (onInit du web part), puis
 * appelable sans argument ensuite depuis n'importe quel service.
 */
export function getSP(context?: WebPartContext): SPFI {
  if (context) {
    _sp = spfi().using(SPFx(context));
  }
  if (!_sp) {
    throw new Error('getSP(context) doit être appelé une première fois avec le contexte du web part.');
  }
  return _sp;
}
