/**
 * Lecture/écriture des fichiers Excel — porté tel quel depuis
 * server/src/utils/excel.ts : cette logique n'utilisait déjà aucune API
 * Node.js (uniquement ExcelJS, qui fonctionne aussi bien dans le
 * navigateur), seul le type d'entrée change (ArrayBuffer plutôt que
 * Buffer, ce que reçoit le navigateur après lecture d'un fichier ou d'un
 * flux SharePoint).
 */
import ExcelJS from 'exceljs';

export interface ParsedConditions {
  colonnes: string[];
  rows: Record<string, string>[]; // valeurs en texte, cellules vides = ''
  nbLignesVides: number;
  colonneCodeCandidate: string | null;
  colonneMarcheCandidate: string | null;
}

const CODE_HEADER_HINTS = [
  'code sous segment',
  'code sous-segment',
  'code_sous_segment',
  'sous segment',
  'sous-segment',
  'code segment',
  'codesoussegment',
];

const MARCHE_HEADER_HINTS = ['marché', 'marche', 'marchés', 'marches'];

/** Parmi les en-têtes correspondant au prédicat, retourne le plus court plutôt que le premier
 * trouvé. Un en-tête reconstruit en concaténant plusieurs lignes (voir sheetToMatrix — certains
 * fichiers ont une ligne d'instruction complète au-dessus du vrai libellé, non filtrée par
 * LIBELLE_GENERIQUE_RE car ce n'est pas un mot isolé comme "AUTO"/"SAISIE") peut accidentellement
 * contenir le mot recherché au sein d'une phrase bien plus longue que le vrai en-tête — observé en
 * conditions réelles : la colonne "marché" auto-détectée était en réalité "Corriger manuellement
 * si l'information renseignée automatiquement est erronée ou manquante – Marché", un intitulé
 * d'aide, pas le vrai en-tête court "Marché" présent par ailleurs dans le même fichier. Le
 * candidat le plus court est presque toujours le vrai en-tête plutôt qu'un texte d'aide. */
function shortestMatch(colonnes: string[], predicate: (c: string) => boolean): string | null {
  const matches = colonnes.filter(predicate);
  if (matches.length === 0) return null;
  return matches.reduce((a, b) => (b.length < a.length ? b : a));
}

/** Ligne "Actif"/"Inactif" : sert à repérer la feuille et la première ligne de données. */
const STATUT_RE = /^(actif|inactif)\.?$/i;
/** Libellés d'instruction de saisie (pas de vrais en-têtes) à ignorer dans la reconstruction. */
const LIBELLE_GENERIQUE_RE = /^(auto|saisie|menu deroulant|menu déroulant)\.?$/i;

function norm(v: unknown): string {
  return String(v ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** A, B, ..., Z, AA, AB, ... pour nommer une colonne sans en-tête détecté. */
function colLetter(n: number): string {
  let s = '';
  n += 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Décode une plage Excel "B5:D7" en indices de lignes/colonnes 1-based. */
function decodeRange(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  const colToNum = (s: string): number => {
    let n = 0;
    for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  };
  return { r1: Number(m[2]), c1: colToNum(m[1]), r2: Number(m[4]), c2: colToNum(m[3]) };
}

/**
 * Convertit une feuille en matrice de chaînes, en propageant la valeur des
 * cellules fusionnées sur toute leur plage (sinon seule la cellule en haut à
 * gauche de la fusion porte une valeur, ce qui casse la reconstruction des
 * en-têtes et des lignes de données).
 */
function sheetToMatrix(sheet: ExcelJS.Worksheet): string[][] {
  const matrix: string[][] = [];
  for (let r = 1; r <= sheet.rowCount; r++) {
    const row: string[] = [];
    const rowObj = sheet.getRow(r);
    for (let c = 1; c <= sheet.columnCount; c++) row.push(cellToString(rowObj.getCell(c).value));
    matrix.push(row);
  }
  const merges = (sheet.model.merges ?? []) as string[];
  for (const ref of merges) {
    const range = decodeRange(ref);
    if (!range) continue;
    const v = matrix[range.r1 - 1]?.[range.c1 - 1];
    if (!norm(v)) continue;
    for (let r = range.r1; r <= range.r2; r++) {
      if (!matrix[r - 1]) matrix[r - 1] = [];
      for (let c = range.c1; c <= range.c2; c++) {
        if (!norm(matrix[r - 1][c - 1])) matrix[r - 1][c - 1] = v;
      }
    }
  }
  return matrix;
}

/**
 * Choisit la feuille la plus probable : celle qui contient une cellule
 * "Actif"/"Inactif" (signe qu'il s'agit du tableau de conditions plutôt que
 * d'un onglet annexe — mode opératoire, table de marques, listes
 * déroulantes...), sinon la plus volumineuse.
 */
function choisirFeuille(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  let best = workbook.worksheets[0];
  let bestScore = -1;
  for (const ws of workbook.worksheets) {
    let hasStatut = false;
    for (let r = 1; r <= ws.rowCount && !hasStatut; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        if (STATUT_RE.test(norm(row.getCell(c).value))) {
          hasStatut = true;
          break;
        }
      }
    }
    const score = (hasStatut ? 1_000_000_000 : 0) + ws.rowCount * ws.columnCount;
    if (score > bestScore) {
      bestScore = score;
      best = ws;
    }
  }
  return best;
}

function cellToString(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    // Formule, richtext, date, hyperlink...
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyVal = value as any;
    if (anyVal.result !== undefined) return cellToString(anyVal.result);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (anyVal.richText) return anyVal.richText.map((r: any) => r.text).join('');
    if (anyVal.text) return String(anyVal.text);
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value);
  }
  return String(value).trim();
}

/**
 * Lit un fichier Excel de conditions commerciales.
 *
 * Ces fichiers ont en général une mise en forme riche : plusieurs onglets
 * annexes (mode opératoire, table de marques, listes déroulantes...), des
 * en-têtes de colonne étalés sur plusieurs lignes (avec des cellules
 * fusionnées pour les regroupements), et une ligne d'instruction de saisie
 * ("AUTO", "SAISIE", "MENU DEROULANT") au-dessus des vrais libellés. On ne
 * peut donc pas se contenter de lire la ligne 1 du premier onglet :
 * - la bonne feuille est celle qui contient une colonne de statut
 *   "Actif"/"Inactif" (repère fiable du tableau de données, par opposition
 *   aux onglets annexes) ; à défaut, on prend la plus volumineuse ;
 * - la première ligne de données est celle qui porte ce statut ;
 * - les lignes qui précèdent (jusqu'à 8) sont concaténées pour reconstruire
 *   le libellé complet de chaque colonne.
 */
export async function parseConditionsFile(buffer: ArrayBuffer): Promise<ParsedConditions> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  if (workbook.worksheets.length === 0) throw new Error('Le fichier Excel ne contient aucune feuille.');

  const sheet = choisirFeuille(workbook);
  const matrix = sheetToMatrix(sheet);
  const maxCols = matrix.reduce((m, r) => Math.max(m, r?.length ?? 0), 0);
  const nbCellulesRemplies = (row: string[] | undefined): number =>
    (row || []).reduce((n, c) => n + (norm(c) ? 1 : 0), 0);

  // Première ligne de données : celle qui porte la valeur "Actif"/"Inactif".
  let dataStart = -1;
  for (let r = 0; r < matrix.length && dataStart < 0; r++) {
    for (let c = 0; c < maxCols; c++) {
      if (STATUT_RE.test(norm(matrix[r]?.[c]))) {
        dataStart = r;
        break;
      }
    }
  }
  if (dataStart < 0) {
    // Repli : première ligne "dense" suivie d'une ligne de densité comparable.
    const seuil = Math.max(2, Math.round(maxCols * 0.2));
    for (let r = 0; r < matrix.length - 1; r++) {
      if (nbCellulesRemplies(matrix[r]) >= seuil && nbCellulesRemplies(matrix[r + 1]) >= seuil) {
        dataStart = r + 1;
        break;
      }
    }
    if (dataStart < 0) dataStart = 0;
  }

  // Zone d'en-tête : les lignes non vides qui précèdent le début des données.
  const headerRows: number[] = [];
  for (let r = Math.max(0, dataStart - 8); r < dataStart; r++) {
    if (nbCellulesRemplies(matrix[r]) > 0) headerRows.push(r);
  }

  const used = new Set<string>();
  const colonnes: string[] = [];
  const colonneIndexSource: number[] = [];
  for (let c = 0; c < maxCols; c++) {
    const parts: string[] = [];
    for (const r of headerRows) {
      const v = norm(matrix[r]?.[c]).replace(/\s*\/\s*/g, ' / ');
      if (!v) continue;
      if (parts.some((p) => p.toLowerCase() === v.toLowerCase())) continue;
      if (LIBELLE_GENERIQUE_RE.test(v)) continue;
      parts.push(v);
    }
    let label = parts.join(' – ').trim();
    const hasData = matrix.slice(dataStart).some((row) => norm(row?.[c]));
    if (!label && !hasData) continue; // colonne vide de bout en bout : ignorée
    if (!label) label = `Colonne ${colLetter(c)}`;
    let nom = label;
    let i = 2;
    while (used.has(nom.toLowerCase())) nom = `${label} (${i++})`;
    used.add(nom.toLowerCase());
    colonnes.push(nom);
    colonneIndexSource.push(c);
  }
  if (colonnes.length === 0) throw new Error('Aucun en-tête de colonne détecté.');

  const rows: Record<string, string>[] = [];
  let nbLignesVides = 0;
  for (let r = dataStart; r < matrix.length; r++) {
    const record: Record<string, string> = {};
    let hasValue = false;
    colonnes.forEach((nom, idx) => {
      const val = norm(matrix[r]?.[colonneIndexSource[idx]]);
      if (val) hasValue = true;
      record[nom] = val;
    });
    if (!hasValue) {
      nbLignesVides++;
      continue;
    }
    rows.push(record);
  }

  const colonneCodeCandidate =
    shortestMatch(colonnes, (c) => CODE_HEADER_HINTS.includes(c.trim().toLowerCase())) ??
    shortestMatch(colonnes, (c) => /code\s*sous[\s-]*segment/i.test(c)) ??
    shortestMatch(colonnes, (c) => /sous[\s-]*segment/i.test(c)) ??
    null;

  const colonneMarcheCandidate =
    shortestMatch(colonnes, (c) => MARCHE_HEADER_HINTS.includes(c.trim().toLowerCase())) ??
    shortestMatch(colonnes, (c) => /march[ée]s?/i.test(c)) ??
    null;

  return { colonnes, rows, nbLignesVides, colonneCodeCandidate, colonneMarcheCandidate };
}

/** Détecte les codes sous-segment dupliqués (comparaison trim + casse insensible). */
export function findDuplicateCodes(rows: Record<string, string>[], colonneCode: string): number {
  const seen = new Map<string, number>();
  for (const row of rows) {
    const key = (row[colonneCode] ?? '').trim().toLowerCase();
    if (!key) continue;
    seen.set(key, (seen.get(key) ?? 0) + 1);
  }
  // Array.from plutôt que `for...of seen.values()` directement : la cible de compilation SPFx
  // est ES5, qui n'itère nativement que les tableaux (une itération directe sur un itérateur
  // de Map nécessiterait --downlevelIteration ou une cible ES2015+).
  let duplicates = 0;
  for (const count of Array.from(seen.values())) if (count > 1) duplicates++;
  return duplicates;
}

/** Écrit la liste des colonnes dans une feuille auxiliaire masquée et renvoie une formule de
 * référence de plage vers celle-ci — utilisé par buildBlankMappingWorkbook quand la formule de
 * liste "inline" dépasserait la limite Excel de 255 caractères. */
function buildAuxColumnListFormula(workbook: ExcelJS.Workbook, colonnes: string[]): string {
  const aux = workbook.addWorksheet('ListeColonnes');
  aux.state = 'veryHidden';
  colonnes.forEach((c, i) => {
    aux.getCell(`A${i + 1}`).value = c;
  });
  return `ListeColonnes!$A$1:$A$${colonnes.length}`;
}

/** Génère un fichier Excel de mapping vierge: Variable | Colonne correspondante | (aide) Colonnes disponibles. */
export async function buildBlankMappingWorkbook(
  variables: string[],
  colonnesDisponibles: string[],
): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Mapping');
  sheet.columns = [
    { header: 'Variable', key: 'variable', width: 35 },
    { header: 'Colonne correspondante', key: 'colonne', width: 40 },
    { header: 'Colonnes disponibles (aide)', key: 'aide', width: 45 },
  ];
  sheet.getRow(1).font = { bold: true };

  const maxRows = Math.max(variables.length, colonnesDisponibles.length, 1);
  for (let i = 0; i < maxRows; i++) {
    sheet.addRow({
      variable: variables[i] ?? '',
      colonne: '',
      aide: colonnesDisponibles[i] ?? '',
    });
  }

  // Liste déroulante de validation sur la colonne B (Colonne correspondante). Excel limite une
  // formule de validation de liste "inline" (valeurs séparées par virgules entre guillemets) à
  // 255 caractères — au-delà, Excel ignore silencieusement la validation à l'ouverture, sans
  // avertir ni l'application ni l'utilisateur (observé au-delà d'une trentaine de colonnes selon
  // la longueur de leurs noms). Si la liste dépasse cette limite, on la déporte dans une feuille
  // auxiliaire masquée et on référence cette plage à la place — pas de limite de longueur dans ce
  // cas, seulement une limite de nombre de lignes (largement suffisante ici).
  if (colonnesDisponibles.length > 0) {
    const inlineFormula = `"${colonnesDisponibles.join(',').replace(/"/g, '')}"`;
    const formula =
      inlineFormula.length <= 255 ? inlineFormula : buildAuxColumnListFormula(workbook, colonnesDisponibles);
    for (let i = 0; i < variables.length; i++) {
      const cell = sheet.getCell(`B${i + 2}`);
      cell.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [formula],
        showErrorMessage: true,
        error: 'Merci de choisir une colonne existante dans la liste.',
      };
    }
  }

  return workbook.xlsx.writeBuffer();
}

export interface ParsedMappingRow {
  variable: string;
  colonne: string;
}

/** Lit un fichier de mapping rempli (colonnes A=Variable, B=Colonne correspondante). */
export async function parseMappingFile(buffer: ArrayBuffer): Promise<ParsedMappingRow[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Le fichier de mapping ne contient aucune feuille.');

  const result: ParsedMappingRow[] = [];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const variable = cellToString(row.getCell(1).value);
    const colonne = cellToString(row.getCell(2).value);
    if (!variable) continue;
    result.push({ variable, colonne });
  }
  return result;
}
