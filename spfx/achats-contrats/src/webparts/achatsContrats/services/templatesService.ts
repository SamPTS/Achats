import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './spClient';
import { ensureProvisioned, retryUntilValid, LISTS, LIBRARIES } from './provisioning';
import { buildStoredFilename, uploadToLibrary, downloadFromServerRelativeUrl } from './storage';
import { extractVariablesFromDocx } from './docx';
import { buildBlankMappingWorkbook as buildBlankMappingWorkbookXlsx, parseMappingFile } from './excel';
import { getActiveConditionsVersion, getConditionsVersion } from './conditionsService';
import { odataEscape } from './odata';
import type { MappingLine, MappingStatut, Template } from '../model/types';

interface TemplateItem {
  Id: number;
  GroupId: string;
  Libelle: string;
  Departement: string | null;
  DateDepot: string;
  Version: number;
  NomFichier: string;
  CheminStockage: string;
  Variables: string; // JSON.stringify(string[])
  DeposePar: string | null;
  Archive: boolean;
}

interface MappingItem {
  Id: number;
  TemplateId: string;
  Variable: string;
  ColonneCorrespondante: string | null;
  Statut: MappingStatut;
}

const TEMPLATE_SELECT = [
  'Id',
  'GroupId',
  'Libelle',
  'Departement',
  'DateDepot',
  'Version',
  'NomFichier',
  'CheminStockage',
  'Variables',
  'DeposePar',
  'Archive',
];
const MAPPING_SELECT = ['Id', 'TemplateId', 'Variable', 'ColonneCorrespondante', 'Statut'];

function templatesList() {
  return getSP().web.lists.getByTitle(LISTS.templatesContrats);
}
function mappingsList() {
  return getSP().web.lists.getByTitle(LISTS.mappingsTemplate);
}

function parseVariables(json: string): string[] {
  try {
    return JSON.parse(json || '[]');
  } catch {
    return [];
  }
}

async function getMappingsRaw(templateId: string): Promise<MappingItem[]> {
  return mappingsList()
    .items.select(...MAPPING_SELECT)
    .filter(`TemplateId eq '${odataEscape(templateId)}'`)
    .top(2000)() as Promise<MappingItem[]>;
}

/** Récupère en un seul appel les mappings de plusieurs templates (évite un appel réseau par
 * template — voir listTemplates, qui construisait auparavant une requête par élément affiché). */
async function getAllMappingsRaw(): Promise<MappingItem[]> {
  return mappingsList().items.select(...MAPPING_SELECT).top(5000)() as Promise<MappingItem[]>;
}

/** Une ligne "mappee" n'est valide que si sa colonne existe encore dans le référentiel de
 * colonnes fourni (le fichier de conditions actif). Si aucun référentiel n'est fourni (pas de
 * conditions active), on ne pénalise pas ce cas — voir server/src/routes/templates.ts. */
function ligneValide(m: MappingItem, colonnesRef: string[] | null): boolean {
  if (m.Statut === 'libre') return true;
  if (m.Statut !== 'mappee' || !m.ColonneCorrespondante) return false;
  if (colonnesRef && !colonnesRef.includes(m.ColonneCorrespondante)) return false;
  return true;
}

function computeStatutMapping(mappings: MappingItem[], colonnesRef: string[] | null): Template['statutMapping'] {
  if (mappings.length === 0) return 'absent';
  return mappings.every((m) => ligneValide(m, colonnesRef)) ? 'complet' : 'incomplet';
}

function toModelSync(item: TemplateItem, mappingsRaw: MappingItem[], colonnesRef: string[] | null): Template {
  const mappings: MappingLine[] = mappingsRaw.map((m) => ({
    variable: m.Variable,
    colonneCorrespondante: m.ColonneCorrespondante,
    statut: m.Statut,
    colonneIntrouvable:
      m.Statut === 'mappee' && !!m.ColonneCorrespondante && !!colonnesRef && !colonnesRef.includes(m.ColonneCorrespondante),
  }));

  return {
    id: String(item.Id),
    groupId: item.GroupId,
    libelle: item.Libelle,
    departement: item.Departement ?? null,
    dateDepot: item.DateDepot,
    version: item.Version,
    nomFichier: item.NomFichier,
    cheminStockage: item.CheminStockage,
    variables: parseVariables(item.Variables),
    deposePar: item.DeposePar ?? null,
    archive: !!item.Archive,
    mappings,
    statutMapping: computeStatutMapping(mappingsRaw, colonnesRef),
  };
}

/** Dernière version de chaque groupe de template, non archivée.
 *
 * Récupère la version de conditions active et l'ensemble des lignes de mapping en deux appels
 * au total, plutôt qu'un appel par template affiché (ancienne version : 2N+1 appels réseau pour
 * afficher N templates — chaque appel étant un aller-retour HTTP complet vers SharePoint,
 * contrairement à l'équivalent côté serveur de l'application autonome, où le même calcul était
 * une simple recherche en mémoire). */
export async function listTemplates(): Promise<Template[]> {
  await ensureProvisioned();
  const [{ items, allMappings }, active] = await Promise.all([
    // Voir retryUntilValid : juste après le dépôt d'un template, ses lignes de mapping tout juste
    // créées peuvent être absentes d'une première lecture sans qu'aucune erreur ne soit levée (pas
    // couvert par retryOnce seul) — on revérifie que chaque template a bien au moins autant de
    // lignes de mapping que de variables déclarées avant d'accepter le résultat.
    retryUntilValid(
      async () => {
        const [items, allMappings] = await Promise.all([
          templatesList().items.select(...TEMPLATE_SELECT).filter('Archive eq 0').top(2000)() as Promise<TemplateItem[]>,
          getAllMappingsRaw(),
        ]);
        return { items, allMappings };
      },
      ({ items, allMappings }) => {
        const counts = new Map<string, number>();
        for (const m of allMappings) counts.set(m.TemplateId, (counts.get(m.TemplateId) ?? 0) + 1);
        return items.every((it) => parseVariables(it.Variables).length <= (counts.get(String(it.Id)) ?? 0));
      },
    ),
    getActiveConditionsVersion(),
  ]);
  const colonnesRef = active ? active.colonnes : null;

  const mappingsByTemplate = new Map<string, MappingItem[]>();
  for (const m of allMappings) {
    const arr = mappingsByTemplate.get(m.TemplateId);
    if (arr) arr.push(m);
    else mappingsByTemplate.set(m.TemplateId, [m]);
  }

  const latestByGroup = new Map<string, TemplateItem>();
  for (const it of items) {
    const current = latestByGroup.get(it.GroupId);
    if (!current || it.Version > current.Version) latestByGroup.set(it.GroupId, it);
  }
  return Array.from(latestByGroup.values())
    .sort((a, b) => (a.DateDepot < b.DateDepot ? 1 : -1))
    .map((it) => toModelSync(it, mappingsByTemplate.get(String(it.Id)) ?? [], colonnesRef));
}

export async function getTemplate(id: string): Promise<Template | undefined> {
  await ensureProvisioned();
  try {
    const [{ item, mappingsRaw }, active] = await Promise.all([
      // Voir listTemplates : même garde contre une lecture réussie mais incomplète juste après le
      // dépôt du template (mapping tout juste créé pas encore visible).
      retryUntilValid(
        async () => {
          const [item, mappingsRaw] = await Promise.all([
            templatesList().items.getById(Number(id)).select(...TEMPLATE_SELECT)() as Promise<TemplateItem>,
            getMappingsRaw(id),
          ]);
          return { item, mappingsRaw };
        },
        ({ item, mappingsRaw }) => parseVariables(item.Variables).length <= mappingsRaw.length,
      ),
      getActiveConditionsVersion(),
    ]);
    return toModelSync(item, mappingsRaw, active ? active.colonnes : null);
  } catch {
    return undefined;
  }
}

export async function uploadTemplate(opts: {
  file: File;
  libelle: string;
  departement: string;
  groupId?: string;
  deposePar: string;
}): Promise<Template> {
  await ensureProvisioned();
  if (!/\.docx$/i.test(opts.file.name)) throw new Error('Seuls les fichiers .docx sont acceptés.');
  const libelle = opts.libelle.trim();
  if (!libelle) throw new Error('Le libellé du template est requis.');

  const buffer = await opts.file.arrayBuffer();
  const variables = await extractVariablesFromDocx(buffer);

  const groupId = opts.groupId || crypto.randomUUID();
  const previousVersions = (await templatesList()
    .items.select('Version')
    .filter(`GroupId eq '${odataEscape(groupId)}'`)()) as { Version: number }[];
  const version = previousVersions.reduce((max, t) => Math.max(max, t.Version), 0) + 1;

  const storedName = buildStoredFilename('template', libelle, 'docx');
  const cheminStockage = await uploadToLibrary(LIBRARIES.templatesFichiers, storedName, buffer);

  const iar = await templatesList().items.add({
    Title: libelle,
    GroupId: groupId,
    Libelle: libelle,
    Departement: opts.departement || null,
    DateDepot: new Date().toISOString(),
    Version: version,
    NomFichier: opts.file.name,
    CheminStockage: cheminStockage,
    Variables: JSON.stringify(variables),
    DeposePar: opts.deposePar || null,
    Archive: false,
  });
  const templateId = String(iar.data.Id);

  await Promise.all(
    variables.map((v) =>
      mappingsList().items.add({
        Title: v,
        TemplateId: templateId,
        Variable: v,
        ColonneCorrespondante: null,
        Statut: 'manquante' as MappingStatut,
      }),
    ),
  );

  const created = await getTemplate(templateId);
  if (!created) throw new Error('Le template déposé est introuvable après création.');
  return created;
}

export async function downloadTemplateFile(template: Template): Promise<ArrayBuffer> {
  return downloadFromServerRelativeUrl(template.cheminStockage);
}

export async function buildBlankMappingWorkbook(templateId: string, conditionsVersionId?: string): Promise<ArrayBuffer> {
  await ensureProvisioned();
  const template = await getTemplate(templateId);
  if (!template) throw new Error('Template introuvable.');
  const version = conditionsVersionId ? await getConditionsVersion(conditionsVersionId) : await getActiveConditionsVersion();
  return buildBlankMappingWorkbookXlsx(template.variables, version ? version.colonnes : []);
}

export interface ImportMappingResult {
  errors: string[];
}

export async function importMapping(templateId: string, file: File, conditionsVersionId?: string): Promise<ImportMappingResult> {
  await ensureProvisioned();
  const template = await getTemplate(templateId);
  if (!template) throw new Error('Template introuvable.');
  const version = conditionsVersionId ? await getConditionsVersion(conditionsVersionId) : await getActiveConditionsVersion();
  const colonnesRef = version ? version.colonnes : [];

  const buffer = await file.arrayBuffer();
  const parsedRows = await parseMappingFile(buffer);

  const variables = template.variables;
  const seenVariables = new Set<string>();
  const seenColonnes = new Set<string>();
  const errors: string[] = [];
  const byVariable = new Map<string, string>();

  for (const r of parsedRows) {
    if (!variables.includes(r.variable)) {
      errors.push(`Variable inconnue dans ce template : "${r.variable}".`);
      continue;
    }
    if (seenVariables.has(r.variable)) {
      errors.push(`Variable en double dans le mapping : "${r.variable}".`);
    }
    seenVariables.add(r.variable);

    if (r.colonne) {
      if (colonnesRef.length > 0 && !colonnesRef.includes(r.colonne)) {
        errors.push(`Colonne inexistante dans le référentiel de conditions : "${r.colonne}".`);
        continue;
      }
      if (seenColonnes.has(r.colonne)) {
        errors.push(`Colonne mappée plusieurs fois : "${r.colonne}".`);
      }
      seenColonnes.add(r.colonne);
      byVariable.set(r.variable, r.colonne);
    }
  }

  if (errors.length > 0) return { errors };

  const mappingsRaw = await getMappingsRaw(templateId);
  await Promise.all(
    variables.map((variable) => {
      const colonne = byVariable.get(variable) ?? null;
      const statut: MappingStatut = colonne ? 'mappee' : 'manquante';
      const row = mappingsRaw.find((m) => m.Variable === variable);
      if (!row) return Promise.resolve();
      return mappingsList().items.getById(row.Id).update({ ColonneCorrespondante: colonne, Statut: statut });
    }),
  );

  return { errors: [] };
}

export async function setMappingLine(
  templateId: string,
  variable: string,
  patch: { colonneCorrespondante?: string | null; statut?: MappingStatut },
): Promise<void> {
  await ensureProvisioned();
  const mappingsRaw = await getMappingsRaw(templateId);
  const row = mappingsRaw.find((m) => m.Variable === variable);
  if (!row) throw new Error('Variable inconnue pour ce template.');
  const finalStatut: MappingStatut = patch.statut ?? (patch.colonneCorrespondante ? 'mappee' : 'manquante');
  await mappingsList().items.getById(row.Id).update({
    ColonneCorrespondante: patch.colonneCorrespondante ?? null,
    Statut: finalStatut,
  });
}
