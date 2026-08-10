import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IListsConfig } from './spClient';
import { parseConditionsFile, findDuplicateCodes } from '../utils/excel';
import { buildStoredFilename } from '../utils/naming';
import { IConditionsVersion, IParsedConditions } from '../models/types';

interface IConditionsListItem {
  Id: number;
  Title: string;
  Created: string;
  FichierUrl: string;
  ColonnesDetectees: string;
  ColonneCodeSousSegment: string | null;
  EstActive: boolean;
  DeposePar: string | null;
  NbLignes: number;
  NbLignesVides: number;
  NbColonnes: number;
  DoublonsDetectes: number;
  Archive: boolean;
}

function toModel(item: IConditionsListItem): IConditionsVersion {
  return {
    id: item.Id,
    nomFichier: item.Title,
    dateDepot: item.Created,
    fichierUrl: item.FichierUrl,
    colonnes: item.ColonnesDetectees ? JSON.parse(item.ColonnesDetectees) : [],
    colonneCodeSousSegment: item.ColonneCodeSousSegment ?? null,
    estActive: !!item.EstActive,
    deposePar: item.DeposePar ?? null,
    nbLignes: item.NbLignes ?? 0,
    nbLignesVides: item.NbLignesVides ?? 0,
    nbColonnes: item.NbColonnes ?? 0,
    doublonsDetectes: item.DoublonsDetectes ?? 0,
    archive: !!item.Archive,
  };
}

export class ConditionsService {
  constructor(private sp: SPFI, private config: IListsConfig) {}

  private get list() {
    return this.sp.web.lists.getByTitle(this.config.conditionsListTitle);
  }

  async listVersions(): Promise<IConditionsVersion[]> {
    const items = await this.list.items
      .select(
        'Id',
        'Title',
        'Created',
        'FichierUrl',
        'ColonnesDetectees',
        'ColonneCodeSousSegment',
        'EstActive',
        'DeposePar',
        'NbLignes',
        'NbLignesVides',
        'NbColonnes',
        'DoublonsDetectes',
        'Archive'
      )
      .filter('Archive eq 0')
      .orderBy('Created', false)
      .top(500)();
    return (items as IConditionsListItem[]).map(toModel);
  }

  async getActiveVersion(): Promise<IConditionsVersion | undefined> {
    const items = await this.list.items.filter('EstActive eq 1 and Archive eq 0').top(1)();
    return items.length > 0 ? toModel(items[0] as IConditionsListItem) : undefined;
  }

  async getVersion(id: number): Promise<IConditionsVersion> {
    const item = await this.list.items.getById(id)();
    return toModel(item as IConditionsListItem);
  }

  /** Relit le fichier Excel stocké dans la bibliothèque de documents. */
  async readRows(version: IConditionsVersion): Promise<IParsedConditions> {
    const buffer: ArrayBuffer = await this.sp.web.getFileByServerRelativePath(version.fichierUrl).getBuffer();
    return parseConditionsFile(buffer);
  }

  /** Dépose un nouveau fichier de conditions commerciales : jamais d'écrasement, nouvelle version historisée. */
  async upload(file: File, deposePar: string): Promise<IConditionsVersion> {
    if (!/\.xlsx$/i.test(file.name)) throw new Error('Seuls les fichiers .xlsx sont acceptés.');

    const buffer = await file.arrayBuffer();
    let parsed: IParsedConditions;
    try {
      parsed = await parseConditionsFile(buffer);
    } catch (e: any) {
      throw new Error(`Fichier Excel illisible ou invalide : ${e.message}`);
    }

    const colonneCode = parsed.colonneCodeCandidate;
    const doublons = colonneCode ? findDuplicateCodes(parsed.rows, colonneCode) : 0;

    const libelle = file.name.replace(/\.xlsx$/i, '');
    const storedName = buildStoredFilename('conditions', libelle, 'xlsx');
    const libraryList = this.sp.web.lists.getByTitle(this.config.documentLibraryTitle);
    const uploadResult = await libraryList.rootFolder.folders
      .getByUrl('Conditions')
      .files.addUsingPath(storedName, buffer, { Overwrite: false });
    const fichierUrl = uploadResult.data.ServerRelativeUrl;

    const existing = await this.list.items.filter('Archive eq 0').top(1)();
    const estActive = existing.length === 0;

    const created = await this.list.items.add({
      Title: file.name,
      FichierUrl: fichierUrl,
      ColonnesDetectees: JSON.stringify(parsed.colonnes),
      ColonneCodeSousSegment: colonneCode,
      EstActive: estActive,
      DeposePar: deposePar || null,
      NbLignes: parsed.rows.length,
      NbLignesVides: parsed.nbLignesVides,
      NbColonnes: parsed.colonnes.length,
      DoublonsDetectes: doublons,
      Archive: false,
    });

    return this.getVersion(created.data.Id);
  }

  async setCodeColumn(id: number, colonneCodeSousSegment: string): Promise<void> {
    const version = await this.getVersion(id);
    if (version.colonnes.indexOf(colonneCodeSousSegment) === -1) {
      throw new Error('Colonne inconnue dans ce fichier.');
    }
    await this.list.items.getById(id).update({ ColonneCodeSousSegment: colonneCodeSousSegment });
  }

  async activate(id: number): Promise<void> {
    const all = await this.list.items.select('Id', 'EstActive').filter('EstActive eq 1')();
    await Promise.all(all.map((it: any) => this.list.items.getById(it.Id).update({ EstActive: false })));
    await this.list.items.getById(id).update({ EstActive: true });
  }

  async archive(id: number): Promise<void> {
    await this.list.items.getById(id).update({ Archive: true, EstActive: false });
  }

  downloadUrl(version: IConditionsVersion): string {
    return `${version.fichierUrl}?download=1`;
  }
}
