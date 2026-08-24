import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/webs';
import { Web } from '@pnp/sp/webs';
import { SPQueryable, spGet } from '@pnp/sp/spqueryable';
import { getSP } from './spClient';
import { odataEscape } from './odata';

/** Horodatage YYYYMMDD-HHMM utilisé pour le nommage des fichiers stockés — porté depuis server/src/storage.ts. */
export function timestampTag(d = new Date()): string {
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
}

const COMBINING_DIACRITICS = new RegExp('[\u0300-\u036f]', 'g');

/**
 * Nettoie un nom pour en faire un segment de chemin/fichier sûr. La cible de compilation SPFx
 * est ES5 : pas de \p{...} Unicode property escapes (nécessitent le flag /u, donc ES2018+) —
 * on retire les diacritiques via la plage des marques combinantes Unicode, comme le fait déjà
 * server/src/utils/docx.ts (sanitizeFileName) dans la version autonome de l'application.
 */
function slug(s: string): string {
  return s
    .normalize('NFD')
    .replace(COMBINING_DIACRITICS, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

/** Courte chaîne aléatoire pour éviter une collision de nom entre deux dépôts du même libellé
 * dans la même minute (deux utilisateurs, deux onglets) — sans elle, buildStoredFilename produit
 * exactement le même nom et addUsingPath({Overwrite:false}) lève une erreur SharePoint brute, non
 * spécifiquement gérée, affichée telle quelle (message technique, non localisé) plutôt que d'être
 * simplement évitée. */
function randomSuffix(): string {
  return Math.random().toString(36).slice(2, 8);
}

/** Construit un nom de fichier stocké: {type}_{libelle}_{YYYYMMDD-HHMM}-{alea}.{ext} — jamais
 * écrasé, chaque dépôt est unique. */
export function buildStoredFilename(type: string, libelle: string, ext: string): string {
  const safeLibelle = slug(libelle).slice(0, 60) || 'fichier';
  return `${type}_${safeLibelle}_${timestampTag()}-${randomSuffix()}.${ext}`;
}

/** Nettoie un segment de nom de fichier issu d'une saisie libre (ex : code sous-segment). */
export function safeFileNamePart(s: string): string {
  return slug(s) || 'x';
}

/** Dépose un fichier dans la bibliothèque indiquée et retourne son chemin relatif au serveur.
 *
 * addUsingPath() peut, comme items.add() (voir conditionsService.ts/templatesService.ts),
 * renvoyer une réponse vide/incomplète (observé en conditions réelles : ServerRelativeUrl absent
 * de la réponse) — le champ « cheminStockage » de l'enregistrement se retrouvait alors vide,
 * cassant silencieusement tout lien de téléchargement vers ce fichier sans qu'aucune erreur ne
 * soit levée au dépôt (un <a href=""> rouvre simplement la page courante). Si la réponse ne
 * contient pas ServerRelativeUrl, on relit le fichier qu'on vient de déposer directement dans la
 * bibliothèque, par son nom (unique par dépôt, horodaté à la minute), plutôt que de stocker une
 * valeur vide sans avertir personne. */
export async function uploadToLibrary(libraryTitle: string, fileName: string, content: ArrayBuffer): Promise<string> {
  const sp = getSP();
  const files = sp.web.lists.getByTitle(libraryTitle).rootFolder.files;
  const info = await files.addUsingPath(fileName, content, { Overwrite: false });
  if (info && info.ServerRelativeUrl) return info.ServerRelativeUrl;

  const reread = (await files.getByUrl(fileName).select('ServerRelativeUrl')()) as { ServerRelativeUrl?: string };
  if (!reread || !reread.ServerRelativeUrl) {
    throw new Error(
      `Le fichier "${fileName}" a été déposé mais son chemin de stockage n'a pas pu être déterminé (réponse SharePoint incomplète).`,
    );
  }
  return reread.ServerRelativeUrl;
}

/** Télécharge le contenu binaire d'un fichier à partir de son chemin relatif au serveur. */
export async function downloadFromServerRelativeUrl(serverRelativeUrl: string): Promise<ArrayBuffer> {
  const sp = getSP();
  return sp.web.getFileByServerRelativePath(serverRelativeUrl).getBuffer();
}

/** Date de dernière modification (horodatage ISO SharePoint) d'un fichier, utilisée pour détecter
 * une édition faite directement dans la bibliothèque (hors dépôt via l'application) — voir
 * resyncIfFileChanged dans conditionsService.ts/templatesService.ts. Retourne null plutôt que de
 * faire échouer l'appelant si le fichier a disparu ou si les droits manquent : dans ce cas, la
 * détection de changement est simplement sautée, jamais bloquante. */
export async function getFileModified(serverRelativeUrl: string): Promise<string | null> {
  try {
    const sp = getSP();
    const info = (await sp.web.getFileByServerRelativePath(serverRelativeUrl).select('TimeLastModified')()) as {
      TimeLastModified?: string;
    };
    return info?.TimeLastModified ?? null;
  } catch {
    return null;
  }
}

/** Supprime définitivement un fichier à partir de son chemin relatif au serveur. N'échoue pas si
 * le fichier est déjà absent (cohérent avec fs.unlinkSync + existsSync côté standalone : un
 * enregistrement dont le fichier a disparu ne doit pas empêcher sa suppression). */
export async function deleteByServerRelativeUrl(serverRelativeUrl: string): Promise<void> {
  const sp = getSP();
  try {
    await sp.web.getFileByServerRelativePath(serverRelativeUrl).delete();
  } catch {
    // Fichier déjà absent, ou droits insuffisants : on laisse l'appelant supprimer quand même
    // l'enregistrement plutôt que de bloquer sur un fichier orphelin.
  }
}

/** Extrait l'identifiant "sourcedoc" d'un lien de partage SharePoint moderne (menu "Copier le
 * lien"), de la forme ".../_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D&...". Ce type de lien ne
 * pointe pas directement vers le chemin du fichier (impossible à résoudre via son URL seule) mais
 * référence le document par son identifiant unique, résolvable via sp.web.getFileById. Retourne
 * null si le lien n'est pas de ce format (URL directe vers le fichier, ou simple chemin). */
function extractSourceDocId(input: string): string | null {
  try {
    const raw = new URL(input).searchParams.get('sourcedoc');
    return raw ? raw.replace(/[{}]/g, '') : null;
  } catch {
    return null;
  }
}

/** Convertit une entrée en URL absolue : la laisse telle quelle si elle l'est déjà, sinon la
 * traite comme un chemin relatif au serveur et la préfixe par l'origine du site courant. */
function toAbsoluteUrl(input: string, currentWebUrl: string): string {
  try {
    return new URL(input).toString();
  } catch {
    const origin = new URL(currentWebUrl).origin;
    return origin + (input.startsWith('/') ? input : `/${input}`);
  }
}

/** Convertit une URL absolue en chemin relatif au serveur ; laisse un chemin déjà relatif tel
 * quel (après décodage des caractères encodés type %20). */
function toServerRelativePath(input: string): string {
  try {
    return decodeURIComponent(new URL(input).pathname);
  } catch {
    return decodeURIComponent(input);
  }
}

/** SharePoint refuse getFileByServerRelativePath dès que le chemin donné n'est pas sous le
 * préfixe du web interrogé ("Server relative urls must start with SPWeb.ServerRelativeUrl") — cas
 * fréquent en pratique : le fichier lié vit dans un autre site/sous-site que celui où
 * l'application est installée. Reconnaît ce message précis pour déclencher la résolution
 * cross-site ci-dessous plutôt que d'échouer directement (voir resolveDirectPath). */
function isCrossWebPathError(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e);
  return /must start with/i.test(msg) && /ServerRelativeUrl/i.test(msg);
}

/** Résout l'URL absolue du site/sous-site propriétaire d'une page ou d'un fichier donné par son
 * URL absolue — méthode SharePoint standard (SP.Web.GetWebUrlFromPageUrl), qui fonctionne pour
 * n'importe quel site/sous-site accessible depuis le même web application, pas seulement le site
 * courant. Appelée via le point d'entrée bas niveau de PnPjs (SPQueryable/spGet) car ce n'est pas
 * une méthode modélisée par les types haut niveau de @pnp/sp. */
async function resolveOwnerWebUrl(absoluteUrl: string): Promise<string> {
  const sp = getSP();
  const q = SPQueryable(sp.web, 'getWebUrlFromPageUrl(@v)');
  q.query.set('@v', `'${odataEscape(absoluteUrl)}'`);
  const result = (await spGet(q)) as string | { value?: string } | { GetWebUrlFromPageUrl?: string };
  const url =
    typeof result === 'string'
      ? result
      : (result as { value?: string }).value ?? (result as { GetWebUrlFromPageUrl?: string }).GetWebUrlFromPageUrl;
  if (!url) throw new Error("Impossible de déterminer le site propriétaire de ce fichier.");
  return url;
}

/** Interroge un fichier par son chemin relatif au serveur, en retombant sur le site/sous-site
 * réellement propriétaire du fichier si celui-ci n'est pas sous le web courant (voir
 * isCrossWebPathError/resolveOwnerWebUrl) plutôt que d'échouer directement — sans ce repli, lier
 * un fichier situé ailleurs que dans le site où l'application est installée était impossible,
 * alors que c'est justement l'usage visé ("un fichier existant ailleurs sur le site"). */
async function resolveDirectPath(
  serverRelativePath: string,
  originalInput: string,
): Promise<{ ServerRelativeUrl: string; Name: string }> {
  const sp = getSP();
  try {
    return (await sp.web.getFileByServerRelativePath(serverRelativePath).select('ServerRelativeUrl', 'Name')()) as {
      ServerRelativeUrl: string;
      Name: string;
    };
  } catch (e) {
    if (!isCrossWebPathError(e)) throw e;
    const absoluteUrl = toAbsoluteUrl(originalInput, sp.web.toUrl());
    const ownerWebUrl = await resolveOwnerWebUrl(absoluteUrl);
    // Web([sp.web, url]) construit une requête vers CE web précis (url absolue), en réutilisant
    // les comportements déjà configurés (authentification SPFx) de sp.web plutôt que d'en
    // reconstruire une nouvelle instance depuis zéro.
    const ownerWeb = Web([sp.web, ownerWebUrl]);
    return (await ownerWeb.getFileByServerRelativePath(serverRelativePath).select('ServerRelativeUrl', 'Name')()) as {
      ServerRelativeUrl: string;
      Name: string;
    };
  }
}

/** Résout une référence vers un fichier EXISTANT ailleurs sur le site (ou un autre site/sous-site
 * accessible) — pour lier un fichier de conditions commerciales déjà en place plutôt que d'en
 * déposer une copie (voir conditionsService.linkExternalConditions). Accepte soit un lien de
 * partage SharePoint moderne ("Copier le lien", contenant sourcedoc=<GUID>), soit une URL absolue
 * ou un chemin relatif pointant directement vers le fichier — y compris dans un site/sous-site
 * différent de celui où l'application est installée (voir resolveDirectPath). Le fichier n'est
 * jamais copié : seul son chemin est enregistré, exactement comme un cheminStockage classique —
 * tout le reste (téléchargement, détection de modification via resyncIfFileChanged) fonctionne
 * ensuite de façon identique. */
export async function resolveFileReference(input: string): Promise<{ serverRelativeUrl: string; nom: string }> {
  const trimmed = input.trim();
  if (!trimmed) throw new Error('Lien ou chemin du fichier requis.');

  const sp = getSP();
  const sourceDocId = extractSourceDocId(trimmed);
  try {
    const info = sourceDocId
      ? ((await sp.web.getFileById(sourceDocId).select('ServerRelativeUrl', 'Name')()) as {
          ServerRelativeUrl: string;
          Name: string;
        })
      : await resolveDirectPath(toServerRelativePath(trimmed), trimmed);
    return { serverRelativeUrl: info.ServerRelativeUrl, nom: info.Name };
  } catch (e) {
    throw new Error(
      "Fichier introuvable à partir de ce lien : vérifiez qu'il pointe bien vers le fichier lui-même " +
        '(ouvrez-le dans SharePoint puis utilisez "Copier le lien", ou collez l\'adresse de la page du ' +
        `fichier) et que vous y avez accès. (${e instanceof Error ? e.message : String(e)})`,
    );
  }
}
