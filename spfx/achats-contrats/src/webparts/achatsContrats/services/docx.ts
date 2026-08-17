/**
 * Extraction et fusion des variables {{Variable}} dans un template Word —
 * porté depuis server/src/utils/docx.ts. La logique elle-même (travail
 * direct sur le XML de word/document.xml via JSZip) est inchangée : elle
 * n'utilisait déjà aucune API Node.js. Seul le type d'entrée/sortie change
 * (ArrayBuffer plutôt que Buffer, generateAsync en 'arraybuffer' plutôt que
 * 'nodebuffer').
 */
import JSZip from 'jszip';

const DOCUMENT_XML_PATH = 'word/document.xml';
const VARIABLE_PATTERN = /\{\{\s*([^{}]+?)\s*\}\}/g;
const RUN_TEXT_PATTERN = /<w:t((?:\s[^>]*)?)>([\s\S]*?)<\/w:t>/g;

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function encodeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

interface RunInfo {
  xmlStart: number;
  xmlEnd: number; // exclusif, position dans la chaîne document.xml
  attrs: string;
  decodedText: string;
  logicalStart: number;
  logicalEnd: number; // exclusif, position dans le texte logique concaténé
}

function collectRuns(xml: string): RunInfo[] {
  const runs: RunInfo[] = [];
  let logicalCursor = 0;
  let match: RegExpExecArray | null;
  RUN_TEXT_PATTERN.lastIndex = 0;
  while ((match = RUN_TEXT_PATTERN.exec(xml)) !== null) {
    const decodedText = decodeXmlEntities(match[2]);
    const run: RunInfo = {
      xmlStart: match.index,
      xmlEnd: match.index + match[0].length,
      attrs: match[1] ?? '',
      decodedText,
      logicalStart: logicalCursor,
      logicalEnd: logicalCursor + decodedText.length,
    };
    runs.push(run);
    logicalCursor = run.logicalEnd;
  }
  return runs;
}

/** Détecte les variables {{Variable}} présentes dans le corps du document (texte fusionné inter-runs). */
export async function extractVariablesFromDocx(buffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(DOCUMENT_XML_PATH);
  if (!file) throw new Error("Fichier Word invalide : word/document.xml introuvable.");
  const xml = await file.async('string');
  const runs = collectRuns(xml);
  const logicalText = runs.map((r) => r.decodedText).join('');

  const seen = new Set<string>();
  const ordered: string[] = [];
  let m: RegExpExecArray | null;
  VARIABLE_PATTERN.lastIndex = 0;
  while ((m = VARIABLE_PATTERN.exec(logicalText)) !== null) {
    const name = m[1].trim();
    if (name && !seen.has(name)) {
      seen.add(name);
      ordered.push(name);
    }
  }
  return ordered;
}

/**
 * Remplace chaque variable {{Variable}} par sa valeur dans le corps du document, en préservant
 * intégralement la mise en forme (styles, runs, tableaux, en-têtes/pieds de page) : seul le texte
 * contenu dans les balises <w:t> concernées est modifié.
 */
export async function fillDocxTemplate(
  buffer: ArrayBuffer,
  values: Record<string, string>,
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(DOCUMENT_XML_PATH);
  if (!file) throw new Error("Fichier Word invalide : word/document.xml introuvable.");
  const xml = await file.async('string');
  const runs = collectRuns(xml);
  const logicalText = runs.map((r) => r.decodedText).join('');

  interface MatchInfo {
    start: number;
    end: number;
    value: string;
  }
  const matches: MatchInfo[] = [];
  let m: RegExpExecArray | null;
  VARIABLE_PATTERN.lastIndex = 0;
  while ((m = VARIABLE_PATTERN.exec(logicalText)) !== null) {
    const name = m[1].trim();
    matches.push({
      start: m.index,
      end: m.index + m[0].length,
      value: values[name] ?? '',
    });
  }
  if (matches.length === 0) {
    // Aucune variable à remplacer : document inchangé.
    return buffer;
  }

  // Pour chaque run, collecter les segments à remplacer (coordonnées logiques locales).
  const overlapsByRun: Array<Array<{ localStart: number; localEnd: number; piece: string }>> = runs.map(
    () => [],
  );
  for (const match of matches) {
    let firstRunTouched = true;
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i];
      const overlapStart = Math.max(match.start, run.logicalStart);
      const overlapEnd = Math.min(match.end, run.logicalEnd);
      if (overlapEnd > overlapStart) {
        const localStart = overlapStart - run.logicalStart;
        const localEnd = overlapEnd - run.logicalStart;
        overlapsByRun[i].push({ localStart, localEnd, piece: firstRunTouched ? match.value : '' });
        firstRunTouched = false;
      }
    }
  }

  // Reconstruire le XML en remplaçant les runs affectés, du dernier au premier (offsets stables).
  let newXml = xml;
  for (let i = runs.length - 1; i >= 0; i--) {
    const overlaps = overlapsByRun[i];
    if (overlaps.length === 0) continue;
    const run = runs[i];
    overlaps.sort((a, b) => a.localStart - b.localStart);
    let rebuilt = '';
    let cursor = 0;
    for (const seg of overlaps) {
      rebuilt += run.decodedText.slice(cursor, seg.localStart) + seg.piece;
      cursor = seg.localEnd;
    }
    rebuilt += run.decodedText.slice(cursor);

    let attrs = run.attrs;
    if (!/xml:space=/.test(attrs)) attrs = `${attrs} xml:space="preserve"`;
    const newTag = `<w:t${attrs}>${encodeXmlText(rebuilt)}</w:t>`;
    newXml = newXml.slice(0, run.xmlStart) + newTag + newXml.slice(run.xmlEnd);
  }

  zip.file(DOCUMENT_XML_PATH, newXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}
