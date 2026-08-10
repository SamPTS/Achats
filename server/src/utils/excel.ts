import ExcelJS from 'exceljs';

export interface ParsedConditions {
  colonnes: string[];
  rows: Record<string, string>[]; // valeurs en texte, cellules vides = ''
  nbLignesVides: number;
  colonneCodeCandidate: string | null;
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

function cellToString(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    // Formule, richtext, date, hyperlink...
    const anyVal = value as any;
    if (anyVal.result !== undefined) return cellToString(anyVal.result);
    if (anyVal.richText) return anyVal.richText.map((r: any) => r.text).join('');
    if (anyVal.text) return String(anyVal.text);
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value);
  }
  return String(value).trim();
}

/** Lit le premier onglet d'un fichier Excel: en-têtes (ligne 1) + lignes de données. */
export async function parseConditionsFile(buffer: Buffer): Promise<ParsedConditions> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Le fichier Excel ne contient aucune feuille.');

  const headerRow = sheet.getRow(1);
  const colonnes: string[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell) => {
    const label = cellToString(cell.value);
    if (label) colonnes.push(label);
  });
  if (colonnes.length === 0) throw new Error('Aucun en-tête de colonne détecté sur la première ligne.');

  const rows: Record<string, string>[] = [];
  let nbLignesVides = 0;
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    if (!row || row.cellCount === 0) continue;
    const record: Record<string, string> = {};
    let hasValue = false;
    colonnes.forEach((col, idx) => {
      const cell = row.getCell(idx + 1);
      const val = cellToString(cell.value);
      if (val) hasValue = true;
      record[col] = val;
    });
    if (!hasValue) {
      nbLignesVides++;
      continue;
    }
    rows.push(record);
  }

  const colonneCodeCandidate =
    colonnes.find((c) => CODE_HEADER_HINTS.includes(c.trim().toLowerCase())) ?? null;

  return { colonnes, rows, nbLignesVides, colonneCodeCandidate };
}

/** Détecte les codes sous-segment dupliqués (comparaison trim + casse insensible). */
export function findDuplicateCodes(rows: Record<string, string>[], colonneCode: string): number {
  const seen = new Map<string, number>();
  for (const row of rows) {
    const key = (row[colonneCode] ?? '').trim().toLowerCase();
    if (!key) continue;
    seen.set(key, (seen.get(key) ?? 0) + 1);
  }
  let duplicates = 0;
  for (const count of seen.values()) if (count > 1) duplicates++;
  return duplicates;
}

/** Génère un fichier Excel de mapping vierge: Variable | Colonne correspondante | (aide) Colonnes disponibles. */
export async function buildBlankMappingWorkbook(
  variables: string[],
  colonnesDisponibles: string[]
): Promise<Buffer> {
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

  // Liste déroulante de validation sur la colonne B (Colonne correspondante)
  if (colonnesDisponibles.length > 0) {
    const formula = `"${colonnesDisponibles.join(',').replace(/"/g, '')}"`;
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

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

export interface ParsedMappingRow {
  variable: string;
  colonne: string;
}

/** Lit un fichier de mapping rempli (colonnes A=Variable, B=Colonne correspondante). */
export async function parseMappingFile(buffer: Buffer): Promise<ParsedMappingRow[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);
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
