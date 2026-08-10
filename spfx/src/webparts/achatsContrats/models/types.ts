export interface IConditionsVersion {
  id: number;
  nomFichier: string;
  dateDepot: string;
  fichierUrl: string;
  colonnes: string[];
  colonneCodeSousSegment: string | null;
  estActive: boolean;
  deposePar: string | null;
  nbLignes: number;
  nbLignesVides: number;
  nbColonnes: number;
  doublonsDetectes: number;
  archive: boolean;
}

export type MappingStatut = 'mappee' | 'libre' | 'manquante';

export interface IMappingLine {
  id: number;
  templateId: number;
  variable: string;
  colonneCorrespondante: string | null;
  statut: MappingStatut;
}

export interface ITemplate {
  id: number;
  groupId: string;
  libelle: string;
  departement: string | null;
  dateDepot: string;
  fichierUrl: string;
  version: number;
  nomFichier: string;
  variables: string[];
  deposePar: string | null;
  archive: boolean;
  mappings: IMappingLine[];
  statutMapping: 'complet' | 'incomplet' | 'absent';
}

export interface IGenerateTemplateOption {
  id: number;
  libelle: string;
  departement: string | null;
  version: number;
}

export interface ISearchMatch {
  rowIndex: number;
  preview: Record<string, string>;
}

export interface IParsedConditions {
  colonnes: string[];
  rows: Record<string, string>[];
  nbLignesVides: number;
  colonneCodeCandidate: string | null;
}

export interface IParsedMappingRow {
  variable: string;
  colonne: string;
}
