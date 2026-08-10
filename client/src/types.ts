export interface ConditionsVersion {
  id: string;
  nomFichier: string;
  dateDepot: string;
  colonnes: string[];
  colonneCodeSousSegment: string | null;
  estActive: boolean;
  deposePar: string | null;
  nbLignes: number;
  nbLignesVides: number;
  nbColonnes: number;
  doublonsDetectes: number;
}

export type MappingStatut = 'mappee' | 'libre' | 'manquante';

export interface MappingLine {
  variable: string;
  colonneCorrespondante: string | null;
  statut: MappingStatut;
}

export interface Template {
  id: string;
  groupId: string;
  libelle: string;
  departement: string | null;
  dateDepot: string;
  version: number;
  nomFichier: string;
  variables: string[];
  deposePar: string | null;
  mappings: MappingLine[];
  statutMapping: 'complet' | 'incomplet' | 'absent';
}

export interface GenerateTemplateOption {
  id: string;
  libelle: string;
  departement: string | null;
  version: number;
}

export interface SearchMatch {
  rowIndex: number;
  preview: Record<string, string>;
}
