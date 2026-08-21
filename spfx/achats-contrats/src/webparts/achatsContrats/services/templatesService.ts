import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './spClient';
import { ensureProvisioned, retryOnce, retryUntilValid, withListRecovery, isNotFoundError, LISTS, LIBRARIES } from './provisioning';
import { buildStoredFilename, uploadToLibrary, downloadFromServerRelativeUrl, deleteByServerRelativeUrl, getFileModified } from './storage';
import { extractVariablesFromDocx } from './docx';
import { LruCache } from './lruCache';
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
  FichierModifieLe: string | null;
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
  'FichierModifieLe',
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
    withListRecovery(() =>
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

// Cache en mémoire du contenu déjà téléchargé, par identifiant de template (cette version précise
// d'un template est immuable une fois déposée — sauf édition directe dans la bibliothèque, voir
// resyncTemplateIfFileChanged juste après, qui invalide l'entrée correspondante dans ce cas).
// Évite de retélécharger le même .docx à chaque contrat généré lors d'une génération en lot
// (plusieurs codes sur le même template) — voir conditionsService.ts (rowsCache) pour le même
// principe côté fichier de conditions. Déclaré avant getTemplate/resyncTemplateIfFileChanged, qui
// le référencent, pour rester lisible dans l'ordre d'exécution. Borné à 10 templates (LruCache) :
// voir rowsCache dans conditionsService.ts pour la même raison (dérive mémoire sur une session
// longue).
const templateFileCache = new LruCache<string, ArrayBuffer>(10);

export async function getTemplate(id: string): Promise<Template | undefined> {
  await ensureProvisioned();
  try {
    const [{ item, mappingsRaw }, active] = await Promise.all([
      // Voir listTemplates : même garde contre une lecture réussie mais incomplète juste après le
      // dépôt du template (mapping tout juste créé pas encore visible).
      withListRecovery(() =>
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
      ),
      getActiveConditionsVersion(),
    ]);
    const resynced = await resyncTemplateIfFileChanged(item, id);
    // Le recalcul (s'il a eu lieu) a pu ajouter/retirer des lignes de mapping : on relit dans ce
    // cas seulement, plutôt que de garder mappingsRaw potentiellement périmé.
    const finalMappingsRaw = resynced === item ? mappingsRaw : await getMappingsRaw(id);
    return toModelSync(resynced, finalMappingsRaw, active ? active.colonnes : null);
  } catch (e) {
    // Voir conditionsService.getConditionsVersion : ne traduit en "introuvable" qu'une vraie
    // absence (404), pas un échec transitoire (réseau, throttling, droits) qui doit remonter.
    if (isNotFoundError(e)) return undefined;
    throw e;
  }
}

/** Équivalent, pour un template, de resyncIfFileChanged dans conditionsService.ts : détecte une
 * édition du .docx faite directement dans la bibliothèque de documents et recalcule les variables
 * détectées à partir du contenu actuel. Les lignes de mapping des variables toujours présentes
 * sont conservées telles quelles (statut et colonne correspondante) ; celles des variables
 * disparues sont supprimées ; une ligne "manquante" est créée pour chaque variable nouvellement
 * apparue dans le document. N'échoue jamais l'appelant : en cas d'erreur, l'élément d'origine
 * (potentiellement périmé) est retourné tel quel. */
async function resyncTemplateIfFileChanged(item: TemplateItem, templateId: string): Promise<TemplateItem> {
  const currentModified = await getFileModified(item.CheminStockage);
  if (!currentModified || currentModified === item.FichierModifieLe) return item;
  try {
    const buffer = await downloadFromServerRelativeUrl(item.CheminStockage);
    const variables = await extractVariablesFromDocx(buffer);
    const existingVariables = parseVariables(item.Variables);
    const added = variables.filter((v) => !existingVariables.includes(v));
    const removed = existingVariables.filter((v) => !variables.includes(v));

    if (added.length > 0 || removed.length > 0) {
      const mappingsRaw = await getMappingsRaw(templateId);
      await Promise.all([
        ...added.map((v) =>
          mappingsList().items.add({
            Title: v,
            TemplateId: templateId,
            Variable: v,
            ColonneCorrespondante: null,
            Statut: 'manquante' as MappingStatut,
          }),
        ),
        ...removed.map((v) => {
          const row = mappingsRaw.find((m) => m.Variable === v);
          return row ? mappingsList().items.getById(row.Id).delete() : Promise.resolve();
        }),
      ]);
    }

    const patch = { Variables: JSON.stringify(variables), FichierModifieLe: currentModified };
    await templatesList().items.getById(item.Id).update(patch);
    templateFileCache.delete(templateId);
    // eslint-disable-next-line no-console
    console.warn(
      `[achats-contrats] Template modifié directement dans la bibliothèque (hors dépôt via l'application) : variables recalculées pour "${item.Libelle}" (v${item.Version}).`,
    );
    return { ...item, ...patch };
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error('[achats-contrats] Échec de la resynchronisation automatique des variables du template :', e);
    return item;
  }
}

/** Retrouve l'Id d'un template tout juste créé quand la réponse de items.add() était vide/incomplète
 * — voir uploadTemplate. GroupId+Version est unique au sein d'un même groupe de template. */
async function recoverTemplateId(groupId: string, version: number): Promise<string> {
  const found = (await retryOnce(async () => {
    const rows = (await templatesList()
      .items.select('Id')
      .filter(`GroupId eq '${odataEscape(groupId)}' and Version eq ${version}`)
      .top(1)()) as { Id: number }[];
    if (rows.length === 0) throw new Error('Template introuvable après création (réponse vide).');
    return rows[0];
  })) as { Id: number };
  return String(found.Id);
}

// Voir MAX_UPLOAD_SIZE dans conditionsService.ts pour la même raison : éviter qu'un .docx
// volumineux (ou une "zip bomb") ne gèle l'onglet du navigateur en le chargeant/décompressant
// entièrement avant tout autre traitement.
const MAX_TEMPLATE_UPLOAD_SIZE = 50 * 1024 * 1024;

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
  if (opts.file.size > MAX_TEMPLATE_UPLOAD_SIZE) {
    throw new Error(`Fichier trop volumineux (${Math.round(opts.file.size / 1024 / 1024)} Mo, maximum 50 Mo).`);
  }

  const buffer = await opts.file.arrayBuffer();
  const variables = await extractVariablesFromDocx(buffer);

  const groupId = opts.groupId || crypto.randomUUID();
  const previousVersions = (await templatesList()
    .items.select('Version')
    .filter(`GroupId eq '${odataEscape(groupId)}'`)()) as { Version: number }[];
  const version = previousVersions.reduce((max, t) => Math.max(max, t.Version), 0) + 1;

  const storedName = buildStoredFilename('template', libelle, 'docx');
  const cheminStockage = await uploadToLibrary(LIBRARIES.templatesFichiers, storedName, buffer);
  // Référence de départ pour la détection de modification directe dans la bibliothèque — voir
  // resyncTemplateIfFileChanged.
  const fichierModifieLe = await getFileModified(cheminStockage);

  // items.add() renvoie directement l'élément créé (.Id à la racine), jamais {data: {...}}.
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
    FichierModifieLe: fichierModifieLe,
  });
  // Défense contre une réponse d'ajout vide/incomplète (observé une fois sur une liste tout juste
  // créée) : plutôt que planter avec "Cannot read properties of undefined (reading 'Id')", ou
  // retenter l'ajout au risque de créer un doublon (l'élément a peut-être bien été créé côté
  // SharePoint malgré une réponse malformée côté client), on retrouve l'élément par sa combinaison
  // GroupId+Version, unique au sein d'un même groupe.
  const templateId =
    iar && typeof iar.Id === 'number'
      ? String(iar.Id)
      : await recoverTemplateId(groupId, version);

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
  const cached = templateFileCache.get(template.id);
  if (cached) return cached;
  const buffer = await downloadFromServerRelativeUrl(template.cheminStockage);
  templateFileCache.set(template.id, buffer);
  return buffer;
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

  // Une variable marquée "saisie libre" avant l'import, et non renseignée dans le fichier importé
  // (ligne absente ou colonne vide), garde ce statut plutôt que de repasser silencieusement à
  // "manquante" : l'import ne doit modifier que les variables qu'il renseigne effectivement.
  const mappingsRaw = await getMappingsRaw(templateId);
  await Promise.all(
    variables.map((variable) => {
      const colonne = byVariable.get(variable) ?? null;
      const avant = mappingsRaw.find((m) => m.Variable === variable);
      let statut: MappingStatut = colonne ? 'mappee' : 'manquante';
      if (!colonne && avant?.Statut === 'libre') statut = 'libre';
      if (!avant) return Promise.resolve();
      return mappingsList().items.getById(avant.Id).update({ ColonneCorrespondante: colonne, Statut: statut });
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

/** Suppression définitive de cette version de template : supprime aussi ses lignes de mapping et
 * le fichier .docx déposé. Les générations déjà journalisées référencent le template par un
 * simple identifiant texte figé au moment de la génération : les supprimer n'affecte pas le
 * journal existant (voir server/src/routes/templates.ts pour l'équivalent standalone). */
export async function deleteTemplate(id: string): Promise<void> {
  await ensureProvisioned();
  const template = await getTemplate(id);
  if (!template) throw new Error('Template introuvable.');

  await deleteByServerRelativeUrl(template.cheminStockage);

  const mappingsRaw = await getMappingsRaw(id);
  await Promise.all(mappingsRaw.map((m) => mappingsList().items.getById(m.Id).delete()));

  await templatesList().items.getById(Number(id)).delete();
}
