import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IListsConfig } from './spClient';
import { extractVariablesFromDocx } from '../utils/docx';
import { buildBlankMappingWorkbook, parseMappingFile } from '../utils/excel';
import { buildStoredFilename } from '../utils/naming';
import { IMappingLine, ITemplate, MappingStatut } from '../models/types';

interface ITemplateListItem {
  Id: number;
  Title: string;
  Created: string;
  GroupId: string;
  Departement: string | null;
  FichierUrl: string;
  Version: number;
  NomFichier: string;
  VariablesDetectees: string;
  DeposePar: string | null;
  Archive: boolean;
}

interface IMappingListItem {
  Id: number;
  TemplateId: number;
  Variable: string;
  ColonneCorrespondante: string | null;
  Statut: MappingStatut;
}

function mappingStatus(mappings: IMappingLine[]): 'complet' | 'incomplet' | 'absent' {
  if (mappings.length === 0) return 'absent';
  return mappings.every((m) => m.statut === 'mappee' || m.statut === 'libre') ? 'complet' : 'incomplet';
}

export class TemplatesService {
  constructor(private sp: SPFI, private config: IListsConfig) {}

  private get list() {
    return this.sp.web.lists.getByTitle(this.config.templatesListTitle);
  }

  private get mappingsList() {
    return this.sp.web.lists.getByTitle(this.config.mappingsListTitle);
  }

  private async toModel(item: ITemplateListItem): Promise<ITemplate> {
    const mappings = await this.getMappings(item.Id);
    return {
      id: item.Id,
      groupId: item.GroupId,
      libelle: item.Title,
      departement: item.Departement ?? null,
      dateDepot: item.Created,
      fichierUrl: item.FichierUrl,
      version: item.Version,
      nomFichier: item.NomFichier,
      variables: item.VariablesDetectees ? JSON.parse(item.VariablesDetectees) : [],
      deposePar: item.DeposePar ?? null,
      archive: !!item.Archive,
      mappings,
      statutMapping: mappingStatus(mappings),
    };
  }

  async getMappings(templateId: number): Promise<IMappingLine[]> {
    const items = await this.mappingsList.items.filter(`TemplateId eq ${templateId}`).top(500)();
    return (items as IMappingListItem[]).map((m) => ({
      id: m.Id,
      templateId: m.TemplateId,
      variable: m.Variable,
      colonneCorrespondante: m.ColonneCorrespondante ?? null,
      statut: m.Statut,
    }));
  }

  /** Dernière version (non archivée) de chaque groupe de template. */
  async listLatestPerGroup(): Promise<ITemplate[]> {
    const items = (await this.list.items
      .select('Id', 'Title', 'Created', 'GroupId', 'Departement', 'FichierUrl', 'Version', 'NomFichier', 'VariablesDetectees', 'DeposePar', 'Archive')
      .filter('Archive eq 0')
      .top(1000)()) as ITemplateListItem[];

    const latestByGroup = new Map<string, ITemplateListItem>();
    for (const item of items) {
      const current = latestByGroup.get(item.GroupId);
      if (!current || item.Version > current.Version) latestByGroup.set(item.GroupId, item);
    }
    const result = await Promise.all(Array.from(latestByGroup.values()).map((i) => this.toModel(i)));
    return result.sort((a, b) => a.libelle.localeCompare(b.libelle));
  }

  async getTemplate(id: number): Promise<ITemplate> {
    const item = (await this.list.items.getById(id)()) as ITemplateListItem;
    return this.toModel(item);
  }

  async getVersionsForGroup(groupId: string): Promise<ITemplate[]> {
    const items = (await this.list.items
      .filter(`GroupId eq '${groupId.replace(/'/g, "''")}' and Archive eq 0`)
      .orderBy('Version', false)()) as ITemplateListItem[];
    return Promise.all(items.map((i) => this.toModel(i)));
  }

  /** Dépose un nouveau template (ou une nouvelle version d'un groupe existant). */
  async upload(opts: {
    file: File;
    libelle: string;
    departement: string;
    groupId?: string;
    deposePar: string;
  }): Promise<ITemplate> {
    if (!/\.docx$/i.test(opts.file.name)) throw new Error('Seuls les fichiers .docx sont acceptés.');
    if (!opts.libelle.trim()) throw new Error('Le libellé du template est requis.');

    const buffer = await opts.file.arrayBuffer();
    let variables: string[];
    try {
      variables = await extractVariablesFromDocx(buffer);
    } catch (e: any) {
      throw new Error(`Fichier Word illisible ou invalide : ${e.message}`);
    }

    const groupId = opts.groupId || generateGroupId();
    let version = 1;
    if (opts.groupId) {
      const previous = (await this.list.items
        .filter(`GroupId eq '${groupId.replace(/'/g, "''")}'`)
        .select('Version')
        .top(1000)()) as Array<{ Version: number }>;
      version = previous.reduce((max, p) => Math.max(max, p.Version), 0) + 1;
    }

    const storedName = buildStoredFilename('template', opts.libelle, 'docx');
    const libraryList = this.sp.web.lists.getByTitle(this.config.documentLibraryTitle);
    const uploadResult = await libraryList.rootFolder.folders
      .getByUrl('Templates')
      .files.addUsingPath(storedName, buffer, { Overwrite: false });
    const fichierUrl = uploadResult.data.ServerRelativeUrl;

    const created = await this.list.items.add({
      Title: opts.libelle,
      GroupId: groupId,
      Departement: opts.departement || null,
      FichierUrl: fichierUrl,
      Version: version,
      NomFichier: opts.file.name,
      VariablesDetectees: JSON.stringify(variables),
      DeposePar: opts.deposePar || null,
      Archive: false,
    });

    await Promise.all(
      variables.map((variable) =>
        this.mappingsList.items.add({
          Title: variable,
          TemplateId: created.data.Id,
          Variable: variable,
          ColonneCorrespondante: null,
          Statut: 'manquante',
        })
      )
    );

    return this.getTemplate(created.data.Id);
  }

  downloadUrl(template: ITemplate): string {
    return `${template.fichierUrl}?download=1`;
  }

  async buildBlankMapping(template: ITemplate, colonnesReference: string[]): Promise<Blob> {
    return buildBlankMappingWorkbook(template.variables, colonnesReference);
  }

  /** Importe et valide un fichier de mapping rempli contre le référentiel de colonnes fourni. */
  async importMapping(templateId: number, file: File, colonnesRef: string[]): Promise<void> {
    const template = await this.getTemplate(templateId);
    const buffer = await file.arrayBuffer();
    let parsedRows;
    try {
      parsedRows = await parseMappingFile(buffer);
    } catch (e: any) {
      throw new Error(`Fichier de mapping illisible : ${e.message}`);
    }

    const seenVariables = new Set<string>();
    const seenColonnes = new Set<string>();
    const erreurs: string[] = [];
    const byVariable = new Map<string, string>();

    for (const r of parsedRows) {
      if (template.variables.indexOf(r.variable) === -1) {
        erreurs.push(`Variable inconnue dans ce template : "${r.variable}".`);
        continue;
      }
      if (seenVariables.has(r.variable)) erreurs.push(`Variable en double dans le mapping : "${r.variable}".`);
      seenVariables.add(r.variable);

      if (r.colonne) {
        if (colonnesRef.length > 0 && colonnesRef.indexOf(r.colonne) === -1) {
          erreurs.push(`Colonne inexistante dans le référentiel de conditions : "${r.colonne}".`);
          continue;
        }
        if (seenColonnes.has(r.colonne)) erreurs.push(`Colonne mappée plusieurs fois : "${r.colonne}".`);
        seenColonnes.add(r.colonne);
        byVariable.set(r.variable, r.colonne);
      }
    }

    if (erreurs.length > 0) {
      throw new Error(`Incohérences détectées dans le mapping : ${erreurs.join(' ; ')}`);
    }

    const mappings = await this.getMappings(templateId);
    await Promise.all(
      template.variables.map((variable) => {
        const colonne = byVariable.get(variable) ?? null;
        const statut: MappingStatut = colonne ? 'mappee' : 'manquante';
        const existing = mappings.find((m) => m.variable === variable);
        if (!existing) return Promise.resolve();
        return this.mappingsList.items.getById(existing.id).update({
          ColonneCorrespondante: colonne,
          Statut: statut,
        });
      })
    );
  }

  async setMappingLine(
    templateId: number,
    variable: string,
    body: { colonneCorrespondante?: string | null; statut?: MappingStatut }
  ): Promise<void> {
    const mappings = await this.getMappings(templateId);
    const existing = mappings.find((m) => m.variable === variable);
    if (!existing) throw new Error('Variable inconnue pour ce template.');
    const statut = body.statut ?? (body.colonneCorrespondante ? 'mappee' : 'manquante');
    await this.mappingsList.items.getById(existing.id).update({
      ColonneCorrespondante: body.colonneCorrespondante ?? null,
      Statut: statut,
    });
  }
}

function generateGroupId(): string {
  return 'tpl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
}
