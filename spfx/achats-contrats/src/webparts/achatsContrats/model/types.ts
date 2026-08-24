/**
 * Modèle de données de l'application, porté depuis la version serveur
 * (server/src/db.ts) vers des listes SharePoint. Les identifiants sont les
 * ID numériques SharePoint des éléments de liste, convertis en chaîne pour
 * rester compatibles avec le reste du code porté depuis le client React
 * d'origine (qui manipule des uuid v4 sous forme de chaînes).
 */

export interface ConditionsVersion {
  id: string;
  nomFichier: string;
  dateDepot: string; // ISO 8601
  /** URL relative au serveur du fichier — dans la bibliothèque "ConditionsFichiers" pour un dépôt
   * classique, ou vers son emplacement d'origine ailleurs sur le site pour un lien externe (voir
   * externe ci-dessous). */
  cheminStockage: string;
  colonnes: string[];
  colonneCodeSousSegment: string | null;
  /** Colonne "marché", utilisée avec colonneCodeSousSegment pour désambiguïser la recherche
   * (un même code sous-segment peut correspondre à plusieurs lignes selon le marché) — voir
   * generateService.searchCode. null si non détectée/désignée : dans ce cas la recherche se fait
   * uniquement par code sous-segment, comme avant l'introduction du marché. */
  colonneMarche: string | null;
  estActive: boolean;
  deposePar: string | null;
  nbLignes: number;
  nbLignesVides: number;
  nbColonnes: number;
  doublonsDetectes: number;
  archive: boolean;
  /** true si cette version référence un fichier existant ailleurs sur le site (lié via son URL,
   * jamais copié) plutôt qu'un fichier déposé dans "ConditionsFichiers". Le fichier peut alors être
   * édité directement à son emplacement d'origine — la resynchronisation automatique (voir
   * conditionsService.resyncIfFileChanged) recalcule les métadonnées à chaque lecture si le
   * fichier a changé, sans jamais nécessiter de nouveau dépôt. Suppression : seul le lien est
   * retiré, le fichier d'origine n'est jamais supprimé (voir conditionsService.deleteConditions). */
  externe: boolean;
}

export type MappingStatut = 'mappee' | 'libre' | 'manquante';

export interface MappingLine {
  variable: string;
  colonneCorrespondante: string | null;
  statut: MappingStatut;
  /** Calculé côté client : la colonne mappée n'existe plus dans le référentiel actif. */
  colonneIntrouvable?: boolean;
}

export interface Template {
  id: string;
  groupId: string;
  libelle: string;
  departement: string | null;
  dateDepot: string;
  version: number;
  nomFichier: string;
  /** URL relative au serveur du fichier dans la bibliothèque "TemplatesFichiers". */
  cheminStockage: string;
  variables: string[];
  deposePar: string | null;
  archive: boolean;
  mappings: MappingLine[];
  statutMapping: 'complet' | 'incomplet' | 'absent';
}

export interface GenerateTemplateOption {
  id: string;
  libelle: string;
  departement: string | null;
  version: number;
  mappingComplet: boolean;
}

export interface SearchMatch {
  rowIndex: number;
  preview: Record<string, string>;
}

export interface GenerationLog {
  id: string;
  templateId: string | null;
  templateLibelle: string | null;
  conditionsVersionId: string | null;
  conditionsNomFichier: string | null;
  codeSousSegment: string;
  dateGeneration: string;
  traitePar: string | null;
}
