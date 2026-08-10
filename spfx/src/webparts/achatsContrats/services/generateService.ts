import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import { IListsConfig } from './spClient';
import { ConditionsService } from './conditionsService';
import { TemplatesService } from './templatesService';
import { fillDocxTemplate } from '../utils/docx';
import { IGenerateTemplateOption, ISearchMatch } from '../models/types';

export class GenerateService {
  private conditionsService: ConditionsService;
  private templatesService: TemplatesService;

  constructor(private sp: SPFI, private config: IListsConfig) {
    this.conditionsService = new ConditionsService(sp, config);
    this.templatesService = new TemplatesService(sp, config);
  }

  private get generationsList() {
    return this.sp.web.lists.getByTitle(this.config.generationsListTitle);
  }

  /** Templates disponibles pour la génération : mapping complet uniquement, dernière version par groupe. */
  async listGenerableTemplates(): Promise<IGenerateTemplateOption[]> {
    const templates = await this.templatesService.listLatestPerGroup();
    return templates
      .filter((t) => t.statutMapping === 'complet')
      .map((t) => ({ id: t.id, libelle: t.libelle, departement: t.departement, version: t.version }));
  }

  /** Recherche exacte (insensible casse/espaces) d'un code sous-segment dans le fichier de conditions choisi. */
  async search(conditionsVersionId: number, code: string): Promise<ISearchMatch[]> {
    const version = await this.conditionsService.getVersion(conditionsVersionId);
    if (!version.colonneCodeSousSegment) {
      throw new Error("La colonne \"code sous-segment\" n'est pas définie pour ce fichier.");
    }
    const { rows } = await this.conditionsService.readRows(version);
    const normalized = code.trim().toLowerCase();
    const colonne = version.colonneCodeSousSegment;

    const matches: ISearchMatch[] = [];
    rows.forEach((row, idx) => {
      if ((row[colonne] ?? '').trim().toLowerCase() === normalized) {
        const previewKeys = Object.keys(row).slice(0, 5);
        const preview: Record<string, string> = {};
        for (const k of previewKeys) preview[k] = row[k];
        matches.push({ rowIndex: idx, preview });
      }
    });
    if (matches.length === 0) throw new Error('Aucun code sous-segment correspondant.');
    return matches;
  }

  /** Valeurs pré-remplies pour un template donné, à partir d'une ligne de conditions résolue. */
  async getMappedValues(
    conditionsVersionId: number,
    templateId: number,
    rowIndex: number
  ): Promise<Record<string, string>> {
    const version = await this.conditionsService.getVersion(conditionsVersionId);
    const { rows } = await this.conditionsService.readRows(version);
    const row = rows[rowIndex];
    if (!row) throw new Error('Ligne introuvable.');

    const mappings = await this.templatesService.getMappings(templateId);
    const values: Record<string, string> = {};
    for (const m of mappings) {
      // Convention "zéro interprétation" : colonne vide -> champ vide, jamais de valeur devinée.
      values[m.variable] = m.statut === 'mappee' && m.colonneCorrespondante ? row[m.colonneCorrespondante] ?? '' : '';
    }
    return values;
  }

  /** Génère le contrat rempli et journalise la génération. Retourne le Blob à télécharger + le nom de fichier. */
  async generate(payload: {
    templateId: number;
    conditionsVersionId: number;
    codeSousSegment: string;
    values: Record<string, string>;
    traitePar: string;
  }): Promise<{ blob: Blob; filename: string }> {
    const template = await this.templatesService.getTemplate(payload.templateId);
    const buffer: ArrayBuffer = await this.sp.web.getFileByServerRelativePath(template.fichierUrl).getBuffer();
    const blob = await fillDocxTemplate(buffer, payload.values);

    await this.generationsList.items.add({
      Title: payload.codeSousSegment,
      TemplateId: payload.templateId,
      ConditionsVersionId: payload.conditionsVersionId,
      CodeSousSegment: payload.codeSousSegment,
      TraitePar: payload.traitePar || null,
    });

    const baseName = template.nomFichier.replace(/\.docx$/i, '');
    return { blob, filename: `${baseName}_${payload.codeSousSegment}.docx` };
  }
}
