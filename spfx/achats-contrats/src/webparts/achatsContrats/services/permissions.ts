/**
 * Gestion des droits par onglet : Conditions commerciales et Templates de contrats réservés aux
 * Propriétaires du site, Générer un contrat ouvert aux Membres (et Propriétaires, qui héritent de
 * tout).
 *
 * Le contrôle se fait par NIVEAU D'AUTORISATION SharePoint (PermissionKind), pas par nom de
 * groupe : les groupes par défaut ("Propriétaires du site X", "Membres du site X") peuvent être
 * renommés, remplacés par des groupes AD/Entra ID, ou l'utilisateur peut avoir ces droits via un
 * chemin différent (permission unique sur le site, appartenance à plusieurs groupes...). Vérifier
 * le niveau d'autorisation réel (ManageWeb = Contrôle total, typiquement seuls les Propriétaires
 * l'ont ; AddListItems = Modifier, typiquement les Membres l'ont, pas les Visiteurs en lecture
 * seule) reste correct quel que soit ce chemin.
 *
 * IMPORTANT — portée de ce contrôle : ceci gouverne uniquement ce qui s'affiche DANS ce web part.
 * Un Membre qui n'a pas accès à l'onglet Conditions commerciales ici pourrait toujours, s'il le
 * voulait, accéder directement aux listes SharePoint sous-jacentes (ConditionsVersions,
 * TemplatesContrats...) via leur propre interface, exactement comme discuté par ailleurs pour
 * l'invariant "une seule version active" — masquer un onglet n'est pas une mesure de sécurité au
 * niveau des données, seulement au niveau de l'expérience utilisateur de ce web part. Pour une
 * restriction réellement imposée par SharePoint, il faut casser l'héritage des permissions sur ces
 * listes et n'y accorder l'accès qu'au groupe Propriétaires — une action distincte, faite au
 * niveau du site, pas du code de l'application.
 */
import '@pnp/sp/security/web';
import { PermissionKind } from '@pnp/sp/security';
import { getSP } from './spClient';

export interface AppPermissions {
  /** Contrôle total (ManageWeb) — typiquement les Propriétaires du site. Donne accès aux onglets
   * Conditions commerciales et Templates de contrats. */
  isOwner: boolean;
  /** Peut ajouter des éléments (AddListItems) — typiquement les Membres (niveau Modifier) et
   * au-dessus, pas les Visiteurs en lecture seule. Donne accès à l'onglet Générer un contrat
   * (qui journalise chaque génération dans la liste Generations). Toujours vrai si isOwner. */
  canGenerate: boolean;
}

let cached: Promise<AppPermissions> | undefined;

/** Mémoïsé pour la durée de vie du web part : les droits de l'utilisateur courant ne changent pas
 * en cours de session, inutile de les revérifier à chaque changement d'onglet. */
export function getAppPermissions(): Promise<AppPermissions> {
  if (!cached) cached = computePermissions();
  return cached;
}

async function computePermissions(): Promise<AppPermissions> {
  const web = getSP().web;
  const [isOwner, canAddItems] = await Promise.all([
    web.currentUserHasPermissions(PermissionKind.ManageWeb),
    web.currentUserHasPermissions(PermissionKind.AddListItems),
  ]);
  return { isOwner, canGenerate: isOwner || canAddItems };
}
