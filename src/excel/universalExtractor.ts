import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { DEFAULT_COUNTRY } from '../data/seedData';
import { Expert, Partner, Subcontractor } from '../types/database';

export type ExtractionTarget = 'experts' | 'subcontractors' | 'partners';

export type ExpertFieldKey =
  | 'ignore'
  | 'fullName'
  | 'lastName'
  | 'firstName'
  | 'experienceAndSpecialty'
  | 'domains'
  | 'yearsExperience'
  | 'email'
  | 'phone'
  | 'website'
  | 'country'
  | 'nationality'
  | 'scope'
  | 'reference'
  | 'notes';

export type OrgFieldKey =
  | 'ignore'
  | 'name'
  | 'specialtyAndNotes'
  | 'domains'
  | 'email'
  | 'phone'
  | 'website'
  | 'country'
  | 'contactName'
  | 'reference'
  | 'notes';

export type ColumnMappingKey = ExpertFieldKey | OrgFieldKey;

export interface ExtractedRawTable {
  id: string;
  sourceFileName: string;
  label: string;
  headers: string[];
  rows: string[][];
  hyperlinks: string[][];
  suggestedTarget: ExtractionTarget;
}

export const EXPERT_FIELD_OPTIONS: Array<{ key: ExpertFieldKey; label: string }> = [
  { key: 'fullName', label: 'Nom & Prénom' },
  { key: 'lastName', label: 'Nom' },
  { key: 'firstName', label: 'Prénom' },
  { key: 'experienceAndSpecialty', label: 'Expérience & Spécialité (texte libre)' },
  { key: 'domains', label: 'Domaines / Spécialité' },
  { key: 'yearsExperience', label: 'Années d’expérience' },
  { key: 'email', label: 'E-mails' },
  { key: 'phone', label: 'Numéro de téléphone' },
  { key: 'website', label: 'Site web / Lien / Hyperlink' },
  { key: 'country', label: 'Pays' },
  { key: 'nationality', label: 'Nationalité' },
  { key: 'scope', label: 'Portée (International / Nationale)' },
  { key: 'reference', label: 'Références (projets réalisés)' },
  { key: 'notes', label: 'Notes / Plateforme / Source' },
  { key: 'ignore', label: '— Ignorer cette colonne —' },
];

export const ORG_FIELD_OPTIONS: Array<{ key: OrgFieldKey; label: string }> = [
  { key: 'name', label: 'Nom du Bureau / Entreprise' },
  { key: 'specialtyAndNotes', label: 'Spécialité & Description' },
  { key: 'domains', label: 'Domaines' },
  { key: 'email', label: 'E-mails' },
  { key: 'phone', label: 'Téléphone' },
  { key: 'website', label: 'Site web / Hyperlink' },
  { key: 'country', label: 'Pays' },
  { key: 'contactName', label: 'Contact principal' },
  { key: 'reference', label: 'Référence / Code' },
  { key: 'notes', label: 'Notes' },
  { key: 'ignore', label: '— Ignorer cette colonne —' },
];

const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
const URL_REGEX = /https?:\/\/[^\s,;]+/gi;
const PHONE_REGEX = /(?:\+|00)?\d[\d\s().-]{6,}\d/g;

function normalizeText(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function isPureEmailCell(text: string): boolean {
  const t = text.trim();
  if (!t) return false;
  return /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(t);
}

function hasEmailOrUrlOrPhone(texts: string[], links: string[] = []): boolean {
  if (links.some((l) => /^https?:\/\//i.test(l) || /^mailto:/i.test(l))) {
    return true;
  }
  const joined = texts.join(' ');
  if (/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/.test(joined)) return true;
  if (/https?:\/\/[^\s,;]+/i.test(joined) || /www\.[^\s,;]+/i.test(joined)) return true;
  const phoneMatches = joined.match(/(?:\+|00)\d[\d\s().-]{5,}\d/g);
  if (phoneMatches && phoneMatches.length > 0) return true;
  return false;
}

function isLikelyHeaderRow(texts: string[], links: string[] = []): boolean {
  const nonEmpty = texts.map((t) => t.trim()).filter(Boolean);
  if (nonEmpty.length === 0) return false;
  // Data rows usually contain emails, URLs, or international phone numbers
  if (hasEmailOrUrlOrPhone(nonEmpty, links)) return false;
  // Long paragraphs (> 110 chars in one cell) are data descriptions, not table headers
  if (nonEmpty.some((t) => t.length > 90)) return false;

  const headerKeywordRe =
    /^(code|id|n°|num|#)$|(nom|prenom|expert|consultant|bureau|cabinet|laboratoire|entreprise|societe|raison sociale|organisme|structure|prestataire|partenaire|sous-traitant|specialite|domaine|secteur|activite|experience|annee|mail|courriel|e-mail|tel|telephone|phone|gsm|mobile|fax|site|web|url|lien|hyperlink|pays|country|ville|adresse|localisation|contact|responsable|portee|nationalite|plateform|source|reference|ref\b|note|observation|commentaire)/;

  let matches = 0;
  for (const cell of nonEmpty) {
    if (headerKeywordRe.test(normalizeText(cell))) {
      matches++;
    }
  }
  return matches >= 2 || (nonEmpty.length === 1 && matches === 1);
}

function cellValueToStringAndLink(
  value: ExcelJS.CellValue,
  cellHyperlink?: string
): {
  text: string;
  link: string;
} {
  let link = cellHyperlink?.trim() || '';
  if (value === null || value === undefined) return { text: '', link };
  if (typeof value === 'string') return { text: value.trim(), link };
  if (typeof value === 'number' || typeof value === 'boolean') {
    return { text: String(value), link };
  }
  if (value instanceof Date) {
    return { text: value.toISOString().slice(0, 10), link };
  }
  if (typeof value === 'object') {
    if ('hyperlink' in value && typeof value.hyperlink === 'string') {
      link = value.hyperlink.trim();
    }
    if ('text' in value && typeof value.text === 'string') {
      return { text: value.text.trim(), link };
    }
    if ('result' in value && value.result !== undefined) {
      return { text: String(value.result).trim(), link };
    }
    if ('richText' in value && Array.isArray(value.richText)) {
      return {
        text: value.richText.map((rt) => rt.text).join('').trim(),
        link,
      };
    }
  }
  return { text: '', link };
}

type GridRow = { texts: string[]; links: string[] };

function splitTablesByBlankRows(gridRows: GridRow[]): GridRow[][] {
  const blocks: GridRow[][] = [];
  let current: GridRow[] = [];

  for (const row of gridRows) {
    const hasContent =
      row.texts.some((t) => t.trim() !== '') || row.links.some((l) => l.trim() !== '');
    if (hasContent) {
      current.push(row);
    } else if (current.length > 0) {
      blocks.push(current);
      current = [];
    }
  }
  if (current.length > 0) {
    blocks.push(current);
  }
  return blocks;
}

/**
 * Merges multi-section blocks in a sheet when blank rows or section title rows
 * separate groups of records that share the same header structure without repeating headers.
 */
function parseStructuredBlocksIntoTables(blocks: GridRow[][]): Array<{
  headers: string[];
  rows: string[][];
  hyperlinks: string[][];
}> {
  const tables: Array<{
    headers: string[];
    rows: string[][];
    hyperlinks: string[][];
    rawColCount: number;
  }> = [];

  for (const rawBlock of blocks) {
    if (rawBlock.length === 0) continue;

    // Skip leading single-cell title banners if there are multi-cell rows below
    let startIdx = 0;
    while (startIdx < rawBlock.length - 1) {
      const row = rawBlock[startIdx];
      const nonEmpty = row.texts.filter((t) => t.trim() !== '');
      if (
        nonEmpty.length === 1 &&
        !hasEmailOrUrlOrPhone(row.texts, row.links) &&
        !isLikelyHeaderRow(row.texts, row.links)
      ) {
        const nextRowNonEmpty = rawBlock[startIdx + 1].texts.filter((t) => t.trim() !== '');
        if (nextRowNonEmpty.length >= 2) {
          startIdx++;
          continue;
        }
      }
      break;
    }

    const block = rawBlock.slice(startIdx);
    if (block.length === 0) continue;

    const firstRow = block[0];
    const firstIsHeader =
      block.length >= 2 &&
      (isLikelyHeaderRow(firstRow.texts, firstRow.links) ||
        !hasEmailOrUrlOrPhone(firstRow.texts, firstRow.links));

    const prevTable = tables[tables.length - 1];

    // If this block does NOT start with a header and we have a compatible previous table,
    // append all rows directly to the previous table instead of dropping or mis-parsing them!
    if (!firstIsHeader && prevTable) {
      const normalizedRows = block
        .filter((r) => {
          const nonEmpty = r.texts.filter((t) => t.trim() !== '');
          // Skip pure section banner rows (1 non-empty cell without contact info)
          if (nonEmpty.length === 1 && !hasEmailOrUrlOrPhone(r.texts, r.links)) {
            return false;
          }
          return nonEmpty.length > 0 || r.links.some((l) => l.trim() !== '');
        });

      for (const r of normalizedRows) {
        const paddedTexts = Array.from({ length: prevTable.rawColCount }, (_, c) =>
          (r.texts[c] ?? '').trim()
        );
        const paddedLinks = Array.from({ length: prevTable.rawColCount }, (_, c) =>
          (r.links[c] ?? '').trim()
        );
        prevTable.rows.push(paddedTexts);
        prevTable.hyperlinks.push(paddedLinks);
      }
      continue;
    }

    // Check if this block starts with the EXACT same header row as the previous table (e.g. page-break repeat)
    if (firstIsHeader && prevTable) {
      const prevHeaderNorm = prevTable.headers.map(normalizeText).join('|');
      const currActiveHeaders = firstRow.texts
        .map((t) => t.trim())
        .filter((_, idx) =>
          block.some((r) => (r.texts[idx] ?? '').trim() !== '')
        )
        .map(normalizeText)
        .join('|');
      if (currActiveHeaders && currActiveHeaders === prevHeaderNorm) {
        for (const r of block.slice(1)) {
          const paddedTexts = Array.from({ length: prevTable.rawColCount }, (_, c) =>
            (r.texts[c] ?? '').trim()
          );
          const paddedLinks = Array.from({ length: prevTable.rawColCount }, (_, c) =>
            (r.links[c] ?? '').trim()
          );
          prevTable.rows.push(paddedTexts);
          prevTable.hyperlinks.push(paddedLinks);
        }
        continue;
      }
    }

    const maxCols = Math.max(...block.map((r) => Math.max(r.texts.length, r.links.length)));
    const activeColIndices: number[] = [];
    for (let c = 0; c < maxCols; c++) {
      const anyVal = block.some(
        (r) => (r.texts[c] ?? '').trim() !== '' || (r.links[c] ?? '').trim() !== ''
      );
      if (anyVal) activeColIndices.push(c);
    }
    if (activeColIndices.length === 0) continue;

    let headers: string[];
    let dataRows: GridRow[];

    if (firstIsHeader) {
      headers = activeColIndices.map(
        (c, idx) => firstRow.texts[c]?.trim() || `Colonne ${idx + 1}`
      );
      dataRows = block.slice(1);
    } else {
      // No header row at all: row 0 is already data!
      headers = activeColIndices.map((_, idx) => `Colonne ${idx + 1}`);
      dataRows = block;
    }

    if (dataRows.length === 0) continue;

    const rows = dataRows.map((r) =>
      activeColIndices.map((c) => (r.texts[c] ?? '').trim())
    );
    const hyperlinks = dataRows.map((r) =>
      activeColIndices.map((c) => (r.links[c] ?? '').trim())
    );

    tables.push({
      headers,
      rows,
      hyperlinks,
      rawColCount: maxCols,
    });
  }

  return tables.map(({ headers, rows, hyperlinks }) => {
    // Trim trailing/unused columns cleanly while keeping alignment
    const active: number[] = [];
    for (let c = 0; c < headers.length; c++) {
      const hasData =
        rows.some((r) => (r[c] ?? '').trim() !== '') ||
        hyperlinks.some((r) => (r[c] ?? '').trim() !== '');
      if (hasData) active.push(c);
    }
    const cols = active.length > 0 ? active : headers.map((_, i) => i);
    return {
      headers: cols.map((c) => headers[c]),
      rows: rows.map((r) => cols.map((c) => r[c] ?? '')),
      hyperlinks: hyperlinks.map((r) => cols.map((c) => r[c] ?? '')),
    };
  });
}

export function detectSuggestedTarget(
  headers: string[],
  rows: string[][]
): ExtractionTarget {
  const headerText = normalizeText(headers.join(' '));
  const sampleText = normalizeText(
    rows
      .slice(0, 12)
      .flat()
      .join(' ')
  );

  let expertScore = 0;
  let bureauScore = 0;

  if (
    /prenom|nom\s*&\s*prenom|expert|experience|annees|nationalite|portee|diplome|cv|plateform|linkedin|signalhire|viadeo/.test(
      headerText
    )
  ) {
    expertScore += 4;
  }
  if (
    /bureau|cabinet|laboratoire|entreprise|societe|raison sociale|sous-traitant|telephone|tel\b|gsm|contact principal/.test(
      headerText
    )
  ) {
    bureauScore += 4;
  }
  if (/partenaire|partner/.test(headerText)) {
    return 'partners';
  }

  if (
    /consultant|ingenieur|enseignant|chercheur|expert|ans d'experience|ans d’experience|dr\s*-?\s*eng/.test(
      sampleText
    )
  ) {
    expertScore += 2;
  }
  if (/sarl|btp|laboratoire|geotechnique|topographie|sondages|essais|\+225|\+216|cabinet|geometre/.test(sampleText)) {
    bureauScore += 2;
  }

  return expertScore > bureauScore ? 'experts' : 'subcontractors';
}

export async function extractTablesFromXlsx(file: File): Promise<ExtractedRawTable[]> {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const results: ExtractedRawTable[] = [];

  workbook.eachSheet((sheet) => {
    const gridRows: GridRow[] = [];
    const colCount = Math.max(sheet.actualColumnCount || 0, sheet.columnCount || 0);

    sheet.eachRow({ includeEmpty: true }, (row) => {
      const texts: string[] = [];
      const links: string[] = [];
      const rowCols = Math.max(colCount, row.cellCount);
      for (let c = 1; c <= rowCols; c++) {
        const cell = row.getCell(c);
        // Only propagate vertical merges across rows (not horizontal title banner merges)
        const isVerticalSlave =
          cell.isMerged &&
          cell.master &&
          Number(cell.master.row) !== Number(row.number);
        const rawVal =
          cell.value !== null && cell.value !== undefined
            ? cell.value
            : isVerticalSlave
              ? cell.master.value
              : cell.value;
        const rawHyperlink =
          cell.hyperlink || (isVerticalSlave ? cell.master.hyperlink : undefined);

        const { text, link } = cellValueToStringAndLink(rawVal, rawHyperlink);
        texts.push(text);
        links.push(link);
      }
      gridRows.push({ texts, links });
    });

    const blocks = splitTablesByBlankRows(gridRows);
    const parsedTables = parseStructuredBlocksIntoTables(blocks);

    parsedTables.forEach((parsed, idx) => {
      if (parsed.rows.length === 0) return;
      results.push({
        id: `xlsx-${file.name}-${sheet.id}-${idx}`,
        sourceFileName: file.name,
        label:
          parsedTables.length > 1
            ? `${sheet.name} (Tableau ${idx + 1})`
            : sheet.name,
        headers: parsed.headers,
        rows: parsed.rows,
        hyperlinks: parsed.hyperlinks,
        suggestedTarget: detectSuggestedTarget(parsed.headers, parsed.rows),
      });
    });
  });

  return results;
}

function parseCsvRows(text: string): string[][] {
  const firstLine = text.split(/\r?\n/)[0] ?? '';
  const semiCount = (firstLine.match(/;/g) ?? []).length;
  const commaCount = (firstLine.match(/,/g) ?? []).length;
  const tabCount = (firstLine.match(/\t/g) ?? []).length;
  const delimiter =
    tabCount > semiCount && tabCount > commaCount
      ? '\t'
      : semiCount >= commaCount
        ? ';'
        : ',';

  const rows: string[][] = [];
  let row: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        current += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(current.trim());
      current = '';
    } else if (ch === '\r' && next === '\n') {
      row.push(current.trim());
      rows.push(row);
      row = [];
      current = '';
      i++;
    } else if (ch === '\n' || ch === '\r') {
      row.push(current.trim());
      rows.push(row);
      row = [];
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.length > 0 || row.length > 0) {
    row.push(current.trim());
    rows.push(row);
  }
  return rows;
}

export async function extractTablesFromCsv(file: File): Promise<ExtractedRawTable[]> {
  const text = await file.text();
  const parsedRows = parseCsvRows(text);
  const gridRows: GridRow[] = parsedRows.map((texts) => ({
    texts,
    links: texts.map(() => ''),
  }));
  const blocks = splitTablesByBlankRows(gridRows);
  const parsedTables = parseStructuredBlocksIntoTables(blocks);

  return parsedTables
    .filter((t) => t.rows.length > 0)
    .map((trimmed, idx) => ({
      id: `csv-${file.name}-${idx}`,
      sourceFileName: file.name,
      label: parsedTables.length > 1 ? `${file.name} (Tableau ${idx + 1})` : file.name,
      headers: trimmed.headers,
      rows: trimmed.rows,
      hyperlinks: trimmed.hyperlinks,
      suggestedTarget: detectSuggestedTarget(trimmed.headers, trimmed.rows),
    }));
}

function getElementsByLocalName(root: Element | Document, localName: string): Element[] {
  const out: Element[] = [];
  const all = root.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName === localName) {
      out.push(el);
    }
  }
  return out;
}

function getDirectChildrenByLocalName(parent: Element, localName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) {
      out.push(child);
    }
  }
  return out;
}

/**
 * Collects direct rows/cells/paragraphs while unwrapping Word <w:sdt> content controls
 * without descending into nested <w:tbl> sub-tables.
 */
function collectElementsStoppingAtTbl(
  root: Element,
  targetLocalName: string,
  stopAtTbl = true
): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    for (let i = 0; i < el.children.length; i++) {
      const child = el.children[i];
      if (stopAtTbl && child.localName === 'tbl') continue;
      if (child.localName === targetLocalName) {
        out.push(child);
        continue;
      }
      walk(child);
    }
  };
  walk(root);
  return out;
}

function getWordCellSpanAndMerge(tc: Element): {
  gridSpan: number;
  vMerge: 'none' | 'restart' | 'continue';
} {
  let gridSpan = 1;
  let vMerge: 'none' | 'restart' | 'continue' = 'none';
  const tcPr = getDirectChildrenByLocalName(tc, 'tcPr')[0];
  if (!tcPr) return { gridSpan, vMerge };

  const gridSpanEl = getDirectChildrenByLocalName(tcPr, 'gridSpan')[0];
  if (gridSpanEl) {
    const raw = gridSpanEl.getAttribute('w:val') || gridSpanEl.getAttribute('val');
    const parsed = Number(raw);
    if (!Number.isNaN(parsed) && parsed > 1) {
      gridSpan = parsed;
    }
  }

  const vMergeEl = getDirectChildrenByLocalName(tcPr, 'vMerge')[0];
  if (vMergeEl) {
    const raw = vMergeEl.getAttribute('w:val') || vMergeEl.getAttribute('val');
    vMerge = raw === 'restart' ? 'restart' : 'continue';
  }

  return { gridSpan, vMerge };
}

function getWordParagraphInfo(
  p: Element,
  relsMap: Map<string, string>
): {
  text: string;
  links: string[];
  isHeading: boolean;
  isListItem: boolean;
} {
  let text = '';
  const links: string[] = [];
  let isHeading = false;
  let isListItem = false;

  const pPr = getDirectChildrenByLocalName(p, 'pPr')[0];
  if (pPr) {
    const pStyle = getDirectChildrenByLocalName(pPr, 'pStyle')[0];
    const styleVal = (pStyle?.getAttribute('w:val') || pStyle?.getAttribute('val') || '').toLowerCase();
    if (/heading|titre|title|subtitle/.test(styleVal)) {
      isHeading = true;
    }
    if (getDirectChildrenByLocalName(pPr, 'numPr').length > 0) {
      isListItem = true;
    }
  }

  const walk = (node: Element) => {
    if (node.localName === 'hyperlink') {
      const rId =
        node.getAttribute('r:id') ||
        node.getAttributeNS(
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'id'
        );
      if (rId && relsMap.has(rId)) {
        links.push(relsMap.get(rId)!);
      }
    }
    if (node.localName === 't') {
      text += node.textContent ?? '';
      return;
    }
    if (node.localName === 'tab') {
      text += ' ';
      return;
    }
    if (node.localName === 'br' || node.localName === 'cr') {
      text += '\n';
      return;
    }
    for (let i = 0; i < node.children.length; i++) {
      walk(node.children[i]);
    }
  };

  walk(p);
  return { text: text.trim(), links, isHeading, isListItem };
}

function hasTblAncestor(el: Element): boolean {
  let curr: Element | null = el.parentElement;
  while (curr) {
    if (curr.localName === 'tbl') return true;
    curr = curr.parentElement;
  }
  return false;
}

/**
 * Fallback extraction for Word (.docx) documents where Bureaux d'études or Experts
 * are written outside tables (bullet lists, headings, cards, or key-value paragraphs).
 */
function extractParagraphRecordsFromDocx(
  doc: Document,
  relsMap: Map<string, string>,
  fileName: string
): ExtractedRawTable | null {
  const allParagraphs = getElementsByLocalName(doc, 'p').filter((p) => !hasTblAncestor(p));
  if (allParagraphs.length === 0) return null;

  type ParaItem = {
    text: string;
    links: string[];
    isHeading: boolean;
    isListItem: boolean;
  };

  const items: ParaItem[] = allParagraphs.map((p) => getWordParagraphInfo(p, relsMap));

  // Group paragraphs into logical entry blocks
  const blocks: ParaItem[][] = [];
  let currentBlock: ParaItem[] = [];

  const flushBlock = () => {
    if (currentBlock.length > 0) {
      blocks.push(currentBlock);
      currentBlock = [];
    }
  };

  const entryStartKeyRe =
    /^(nom|bureau|cabinet|laboratoire|entreprise|societe|raison sociale|organisme|expert|consultant)\s*[:\-–]/i;

  for (const item of items) {
    if (!item.text && item.links.length === 0) {
      flushBlock();
      continue;
    }

    const lineHasContact = hasEmailOrUrlOrPhone([item.text], item.links);

    if (item.isHeading) {
      flushBlock();
      currentBlock.push(item);
      continue;
    }

    if (entryStartKeyRe.test(item.text) && currentBlock.length > 0) {
      flushBlock();
    }

    if (item.isListItem && lineHasContact) {
      flushBlock();
      blocks.push([item]);
      continue;
    }

    currentBlock.push(item);
  }
  flushBlock();

  const rows: string[][] = [];
  const hyperlinks: string[][] = [];

  const parseBlockToRecord = (block: ParaItem[]) => {
    const rawLines = block.flatMap((b) =>
      b.text
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean)
    );
    const blockLinks = Array.from(new Set(block.flatMap((b) => b.links)));
    const fullText = rawLines.join('\n');
    const hasContact = hasEmailOrUrlOrPhone([fullText], blockLinks);

    // Must contain at least a contact signal (email / phone / URL) or explicit key-value lines
    const hasStructuredKeys = rawLines.some((l) =>
      /^(nom|bureau|cabinet|laboratoire|entreprise|societe|specialite|domaine|email|e-mail|tel|telephone|site|web)\s*:/i.test(
        l
      )
    );
    if (!hasContact && !hasStructuredKeys) return;

    // Single-line delimited entry (e.g. "1. AGTS CI — Géotechnique — contact@... — +225...")
    if (rawLines.length === 1 && !hasStructuredKeys) {
      const line = rawLines[0].replace(/^(?:\d+[.)]\s*|[-•*]\s*)/, '').trim();
      const parts = line
        .split(/\s*(?:\||—|–|;|\t)\s*/)
        .map((p) => p.trim())
        .filter(Boolean);

      const emails = extractEmails(line);
      const phones = extractPhones(line);
      const urls = extractUrls(line, blockLinks[0]);
      const nameCandidate =
        parts.find(
          (p) =>
            extractEmails(p).length === 0 &&
            extractUrls(p).length === 0 &&
            extractPhones(p).length === 0 &&
            p.length >= 2
        ) ?? line.split(/[,:]/)[0].trim();

      if (!nameCandidate || nameCandidate.length > 140) return;
      const desc = parts
        .filter(
          (p) =>
            p !== nameCandidate &&
            !emails.includes(p) &&
            !urls.includes(p) &&
            !phones.includes(p)
        )
        .join(' · ');

      rows.push([
        nameCandidate,
        desc,
        emails.join(', '),
        phones.join(' / '),
        urls[0] ?? '',
      ]);
      hyperlinks.push(['', '', '', '', blockLinks[0] ?? urls[0] ?? '']);
      return;
    }

    // Multi-line card block
    let name = '';
    const descParts: string[] = [];
    const emails = extractEmails(fullText);
    const phones = extractPhones(fullText);
    const urls = extractUrls(fullText, blockLinks[0]);

    for (const line of rawLines) {
      const cleanLine = line.replace(/^(?:\d+[.)]\s*|[-•*]\s*)/, '').trim();
      const kvMatch = cleanLine.match(/^([^:]{2,32}):\s*(.+)$/);
      if (kvMatch) {
        const k = normalizeText(kvMatch[1]);
        const v = kvMatch[2].trim();
        if (/^(nom|bureau|cabinet|laboratoire|entreprise|societe|raison sociale|organisme|expert)$/.test(k)) {
          name = v;
        } else if (!/^(mail|courriel|e-mail|tel|telephone|phone|gsm|site|web|url|lien)$/.test(k)) {
          descParts.push(cleanLine);
        }
      } else if (
        !name &&
        !isPureEmailCell(cleanLine) &&
        extractUrls(cleanLine).length === 0 &&
        extractPhones(cleanLine).join('') !== cleanLine &&
        cleanLine.length <= 130
      ) {
        name = cleanLine;
      } else if (!isPureEmailCell(cleanLine) && extractUrls(cleanLine).join('') !== cleanLine) {
        descParts.push(cleanLine);
      }
    }

    if (!name) return;

    rows.push([
      name,
      descParts.join(' — '),
      emails.join(', '),
      phones.join(' / '),
      urls[0] ?? '',
    ]);
    hyperlinks.push(['', '', '', '', blockLinks[0] ?? urls[0] ?? '']);
  };

  for (const block of blocks) {
    parseBlockToRecord(block);
  }

  if (rows.length === 0) return null;

  const headers = [
    'Nom / Bureau',
    'Spécialité / Description',
    'E-mails',
    'Téléphone',
    'Site web / Lien',
  ];

  return {
    id: `docx-paragraphs-${fileName}`,
    sourceFileName: fileName,
    label: 'Texte & listes hors tableaux (Word)',
    headers,
    rows,
    hyperlinks,
    suggestedTarget: detectSuggestedTarget(headers, rows),
  };
}

export async function extractTablesFromDocx(file: File): Promise<ExtractedRawTable[]> {
  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const docXmlEntry = zip.file('word/document.xml');
  if (!docXmlEntry) {
    throw new Error('Document Word (.docx) invalide : word/document.xml introuvable.');
  }

  // Load hyperlink relationships from document + any headers/footers
  const relsMap = new Map<string, string>();
  const relFiles = Object.keys(zip.files).filter((name) =>
    /^word\/_rels\/.*\.xml\.rels$/i.test(name)
  );
  for (const relPath of relFiles) {
    const relEntry = zip.file(relPath);
    if (!relEntry) continue;
    const relsXml = await relEntry.async('text');
    const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
    for (const rel of getElementsByLocalName(relsDoc, 'Relationship')) {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      const mode = rel.getAttribute('TargetMode');
      if (
        id &&
        target &&
        (mode === 'External' || /^https?:\/\//i.test(target) || target.startsWith('mailto:'))
      ) {
        relsMap.set(id, target);
      }
    }
  }

  const results: ExtractedRawTable[] = [];
  const xmlPaths = ['word/document.xml', ...Object.keys(zip.files).filter((n) => /^word\/(header|footer)\d*\.xml$/i.test(n))];

  let globalTblIdx = 0;

  for (const xmlPath of xmlPaths) {
    const entry = zip.file(xmlPath);
    if (!entry) continue;
    const xmlStr = await entry.async('text');
    const doc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    const tables = getElementsByLocalName(doc, 'tbl');

    tables.forEach((tbl) => {
      globalTblIdx += 1;
      const trList = collectElementsStoppingAtTbl(tbl, 'tr', true);
      const gridRows: GridRow[] = [];
      const vMergePrev: Array<{ text: string; link: string }> = [];

      for (const tr of trList) {
        const tcList = collectElementsStoppingAtTbl(tr, 'tc', true);
        const texts: string[] = [];
        const links: string[] = [];
        let colCursor = 0;

        for (const tc of tcList) {
          const { gridSpan, vMerge } = getWordCellSpanAndMerge(tc);
          const paragraphs = collectElementsStoppingAtTbl(tc, 'p', true);
          const pTexts: string[] = [];
          const pLinks: string[] = [];
          for (const p of paragraphs) {
            const parsed = getWordParagraphInfo(p, relsMap);
            if (parsed.text) pTexts.push(parsed.text);
            for (const l of parsed.links) {
              if (!pLinks.includes(l)) pLinks.push(l);
            }
          }

          let cellText = pTexts.join('\n').trim();
          let cellLink = pLinks.find((l) => /^https?:\/\//i.test(l)) ?? pLinks[0] ?? '';

          // Inherit vertically merged cell value when vMerge continues without new text
          if (vMerge === 'continue' && !cellText && !cellLink && vMergePrev[colCursor]) {
            cellText = vMergePrev[colCursor].text;
            cellLink = vMergePrev[colCursor].link;
          } else {
            vMergePrev[colCursor] = { text: cellText, link: cellLink };
          }

          texts.push(cellText);
          links.push(cellLink);
          for (let s = 1; s < gridSpan; s++) {
            texts.push('');
            links.push('');
          }
          colCursor += gridSpan;
        }

        if (texts.some((t) => t !== '') || links.some((l) => l !== '')) {
          gridRows.push({ texts, links });
        }
      }

      const blocks = splitTablesByBlankRows(gridRows);
      const parsedTables = parseStructuredBlocksIntoTables(blocks);

      parsedTables.forEach((parsed, partIdx) => {
        if (parsed.rows.length === 0) return;
        results.push({
          id: `docx-${file.name}-tbl-${globalTblIdx}-${partIdx}`,
          sourceFileName: file.name,
          label:
            parsedTables.length > 1
              ? `Tableau Word #${globalTblIdx}.${partIdx + 1}`
              : `Tableau Word #${globalTblIdx}`,
          headers: parsed.headers,
          rows: parsed.rows,
          hyperlinks: parsed.hyperlinks,
          suggestedTarget: detectSuggestedTarget(parsed.headers, parsed.rows),
        });
      });
    });

    if (xmlPath === 'word/document.xml') {
      const paraTable = extractParagraphRecordsFromDocx(doc, relsMap, file.name);
      if (paraTable) {
        results.push(paraTable);
      }
    }
  }

  return results;
}

export async function extractTablesFromFile(file: File): Promise<ExtractedRawTable[]> {
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.xlsx')) {
    return extractTablesFromXlsx(file);
  }
  if (lower.endsWith('.docx')) {
    return extractTablesFromDocx(file);
  }
  if (lower.endsWith('.csv') || lower.endsWith('.tsv')) {
    return extractTablesFromCsv(file);
  }
  throw new Error(
    `Format non pris en charge (${file.name}). Utilisez des fichiers Excel (.xlsx, .csv) ou Word (.docx).`
  );
}

function inferColumnContentKind(
  colIdx: number,
  rows: string[][],
  hyperlinks: string[][] = []
): 'email' | 'website' | 'phone' | 'text' {
  const sample = rows.slice(0, 30);
  let nonEmpty = 0;
  let emailHits = 0;
  let urlHits = 0;
  let phoneHits = 0;

  for (let r = 0; r < sample.length; r++) {
    const val = (sample[r][colIdx] ?? '').trim();
    const link = (hyperlinks[r]?.[colIdx] ?? '').trim();
    if (!val && !link) continue;
    nonEmpty++;
    if (/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/.test(val)) {
      emailHits++;
    }
    if (
      /^https?:\/\//i.test(link) ||
      /https?:\/\/[^\s,;]+/i.test(val) ||
      /www\.[^\s,;]+/i.test(val)
    ) {
      urlHits++;
    }
    if (
      !isPureEmailCell(val) &&
      !val.includes('@') &&
      !val.includes('http') &&
      /(?:\+|00)?\d[\d\s().-]{6,}\d/.test(val)
    ) {
      phoneHits++;
    }
  }

  if (nonEmpty === 0) return 'text';
  if (emailHits / nonEmpty >= 0.4) return 'email';
  if (urlHits / nonEmpty >= 0.4) return 'website';
  if (phoneHits / nonEmpty >= 0.4) return 'phone';
  return 'text';
}

export function autoMapExpertColumns(
  headers: string[],
  rows: string[][] = [],
  hyperlinks: string[][] = []
): ExpertFieldKey[] {
  const mapped: ExpertFieldKey[] = headers.map((header, colIdx) => {
    const h = normalizeText(header);
    if (/^(code|id|n°|num|#)$/.test(h)) return 'ignore';
    if (/(nom\s*[&/]\s*prenom|prenom\s*[&/]\s*nom|nom complet|full\s*name|expert|consultant|candidat)/.test(h)) {
      return 'fullName';
    }
    if (/^(nom|last\s*name|surname)$/.test(h)) return 'lastName';
    if (/^(prenom|first\s*name)$/.test(h)) return 'firstName';
    if (/(experience.*specialite|specialite.*experience|profil|biographie|resume|parcours)/.test(h)) {
      return 'experienceAndSpecialty';
    }
    if (/(annee|experience|\bexp\b|anciennete)/.test(h)) return 'yearsExperience';
    if (/(domaine|specialite|secteur|competence|expertise)/.test(h)) return 'domains';
    if (/(mail|courriel|e-mail)/.test(h)) return 'email';
    if (/(tel|telephone|phone|gsm|mobile)/.test(h)) return 'phone';
    if (/(hyperlink|site|web|url|lien|profil)/.test(h)) return 'website';
    if (/(pays|country|localisation)/.test(h)) return 'country';
    if (/(nationalite|nationality)/.test(h)) return 'nationality';
    if (/(portee|scope)/.test(h)) return 'scope';
    if (/(plateform|platform|source|reseau)/.test(h)) return 'notes';
    if (/(reference|ref\b|projets?\s*realise)/.test(h)) return 'reference';
    if (/(note|observation|commentaire|remarque)/.test(h)) return 'notes';

    if (rows.length > 0) {
      const kind = inferColumnContentKind(colIdx, rows, hyperlinks);
      if (kind !== 'text') return kind;
    }
    return 'notes';
  });

  if (rows.length > 0) {
    mapped.forEach((field, colIdx) => {
      if (field === 'notes') {
        const kind = inferColumnContentKind(colIdx, rows, hyperlinks);
        if (kind !== 'text') mapped[colIdx] = kind;
      }
    });

    if (!mapped.some((m) => m === 'fullName' || m === 'lastName' || m === 'firstName')) {
      const candidateIdx = mapped.findIndex((m) => m === 'notes');
      if (candidateIdx !== -1) mapped[candidateIdx] = 'fullName';
    }
  }

  return mapped;
}

export function autoMapOrgColumns(
  headers: string[],
  rows: string[][] = [],
  hyperlinks: string[][] = []
): OrgFieldKey[] {
  const mapped: OrgFieldKey[] = headers.map((header, colIdx) => {
    const h = normalizeText(header);
    if (/^(code|id|n°|num|#)$/.test(h)) return 'ignore';
    if (
      /(bureau|raison sociale|societe|entreprise|cabinet|laboratoire|organisme|partenaire|sous-traitant|prestataire|structure|entite|denomination|designation|intitule|libelle|\bnom\b)/.test(
        h
      )
    ) {
      return 'name';
    }
    if (/(specialite|prestation|activite|description|service|detail)/.test(h)) {
      return 'specialtyAndNotes';
    }
    if (/(domaine|secteur|metier|discipline)/.test(h)) return 'domains';
    if (/(mail|courriel|e-mail)/.test(h)) return 'email';
    if (/(tel|telephone|phone|gsm|mobile|fax)/.test(h)) return 'phone';
    if (/(hyperlink|site|web|url|lien)/.test(h)) return 'website';
    if (/(pays|country|ville|adresse|localisation|siege)/.test(h)) return 'country';
    if (/(contact|responsable|gerant|directeur|interlocuteur)/.test(h)) return 'contactName';
    if (/(reference|ref\b|source|plateform)/.test(h)) return 'reference';
    if (/(note|observation|commentaire|remarque)/.test(h)) return 'notes';

    if (rows.length > 0) {
      const kind = inferColumnContentKind(colIdx, rows, hyperlinks);
      if (kind !== 'text') return kind;
    }
    return 'notes';
  });

  if (rows.length > 0) {
    mapped.forEach((field, colIdx) => {
      if (field === 'notes') {
        const kind = inferColumnContentKind(colIdx, rows, hyperlinks);
        if (kind !== 'text') mapped[colIdx] = kind;
      }
    });

    if (!mapped.includes('name')) {
      const candidateIdx = mapped.findIndex((m) => m === 'notes' || m === 'specialtyAndNotes');
      if (candidateIdx !== -1) mapped[candidateIdx] = 'name';
    }
  }

  return mapped;
}

function extractEmails(text: string): string[] {
  const matches = text.match(EMAIL_REGEX) ?? [];
  return Array.from(new Set(matches.map((m) => m.trim())));
}

function extractUrls(text: string, link?: string): string[] {
  const candidates: string[] = [];
  if (link && /^https?:\/\//i.test(link)) {
    candidates.push(link.trim());
  }
  const matches = text.match(URL_REGEX) ?? [];
  for (const m of matches) {
    const clean = m.replace(/[),.;]+$/, '');
    if (!candidates.includes(clean)) candidates.push(clean);
  }
  return candidates;
}

function extractPhones(text: string): string[] {
  const lines = text
    .split(/[\n/;|]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  const phones: string[] = [];
  for (const line of lines) {
    const matches = line.match(PHONE_REGEX) ?? [];
    for (const m of matches) {
      const clean = m.replace(/\s+/g, ' ').trim();
      if (clean.replace(/\D/g, '').length >= 7 && !phones.includes(clean)) {
        phones.push(clean);
      }
    }
  }
  return phones;
}

function inferCountryFromTextOrPhone(text: string): string | null {
  const norm = normalizeText(text);
  if (/\+225|00225|cote d'ivoire|abidjan|yamoussoukro/.test(norm)) return "Côte d'Ivoire";
  if (/\+216|00216|tunisie|tunis|sfax|hammamet|gabes|enig|anged/.test(norm)) return 'Tunisie';
  if (/\+212|00212|maroc|rabat|casablanca|tanger/.test(norm)) return 'Maroc';
  if (/\+221|00221|senegal|dakar/.test(norm)) return 'Sénégal';
  if (/\+213|00213|algerie|alger/.test(norm)) return 'Algérie';
  if (/\+237|00237|cameroun|yaounde|douala/.test(norm)) return 'Cameroun';
  if (/\+223|00223|mali|bamako/.test(norm)) return 'Mali';
  if (/\+226|00226|burkina|ouagadougou/.test(norm)) return 'Burkina Faso';
  if (/\+222|00222|mauritanie|nouakchott/.test(norm)) return 'Mauritanie';
  if (/\+241|00241|gabon|libreville/.test(norm)) return 'Gabon';
  if (/\+33|0033|\bfrance\b|paris/.test(norm)) return 'France';
  return null;
}

function inferDomainsFromText(text: string, fallbackDomain: string): string[] {
  const norm = normalizeText(text);
  const detected: string[] = [];
  const rules: Array<{ pattern: RegExp; domain: string }> = [
    {
      pattern: /dechet|economie circulaire|biomethanisation|valorisation organique|sweep/,
      domain: 'Gestion des déchets & Économie circulaire',
    },
    {
      pattern: /geotechnique|sondage|carotte|sols|fondation|pressiometr|penetr|essai.*materiau|labosol|labogem|lbtp|leec|geobtp/,
      domain: 'Géotechnique & Essais de sols',
    },
    {
      pattern: /topograph|bathymetr|geodes|geomatique|photogrammetr|drone|\bsig\b|cartograph|geometre/,
      domain: 'Topographie, Géodésie & SIG',
    },
    {
      pattern: /geophys|tomographie|georadar|sismique|masw|hydrogeolog/,
      domain: 'Géophysique & Hydrogéologie',
    },
    {
      pattern: /environnement|eies|ecolog|climat/,
      domain: 'Environnement & EIES',
    },
    {
      pattern: /hydraulique|assainissement|hydrolog|hydrograph|eau\b/,
      domain: 'Hydraulique & Assainissement',
    },
    {
      pattern: /genie civil|infrastructure|\bbtp\b|ouvrages|route/,
      domain: 'Génie Civil & Infrastructures',
    },
    {
      pattern: /energie|solaire|photovoltaique|carbone|biogaz/,
      domain: 'Énergie & Transition Climatique',
    },
    {
      pattern: /amenagement|urbain|foncier/,
      domain: 'Aménagement Urbain',
    },
    {
      pattern: /audit|finance/,
      domain: 'Finance Publique & Audit',
    },
  ];

  for (const r of rules) {
    if (r.pattern.test(norm) && !detected.includes(r.domain)) {
      detected.push(r.domain);
    }
  }
  if (detected.length === 0 && fallbackDomain) {
    detected.push(fallbackDomain);
  }
  return detected;
}

function splitPersonFullName(raw: string): {
  firstName: string;
  lastName: string;
  titlePrefix: string;
} {
  let clean = raw.replace(/\s+/g, ' ').trim();
  const titleMatch = clean.match(
    /^(?:Dr\s*-?\s*Eng|Dr\.?|Pr\.?|Prof\.?|Ing\.?|M\.?|Mme\.?|Mr\.?)\s+/i
  );
  const titlePrefix = titleMatch ? titleMatch[0].trim() : '';
  if (titleMatch) {
    clean = clean.slice(titleMatch[0].length).trim();
  }
  const parts = clean.split(' ').filter(Boolean);
  if (parts.length <= 1) {
    return { firstName: '', lastName: clean || 'Inconnu', titlePrefix };
  }
  return {
    firstName: parts[0],
    lastName: parts.slice(1).join(' '),
    titlePrefix,
  };
}

function parseYearsExperience(text: string): number {
  const match = text.match(/(\d{1,2})\s*\+?\s*ans?/i);
  if (match) {
    const val = Number(match[1]);
    if (!Number.isNaN(val) && val > 0 && val <= 65) return val;
  }
  const sinceMatch = text.match(/depuis\s+(19\d{2}|20\d{2})/i);
  if (sinceMatch) {
    const yr = Number(sinceMatch[1]);
    const diff = new Date().getFullYear() - yr;
    if (diff > 0 && diff <= 65) return diff;
  }
  return 0;
}

export function buildExpertsFromTable(
  table: ExtractedRawTable,
  mapping: ExpertFieldKey[],
  startingNumber: number
): Expert[] {
  const now = new Date().toISOString();
  const experts: Expert[] = [];

  table.rows.forEach((row, idx) => {
    let firstName = '';
    let lastName = '';
    let titlePrefix = '';
    let email = '';
    const phones: string[] = [];
    let website = '';
    let country = '';
    let nationality = '';
    let scopeRaw = '';
    let reference = '';
    let yearsExperience = 0;
    const explicitDomains: string[] = [];
    const noteParts: string[] = [];
    const allRowText = row.join(' ');
    const allRowLinks = table.hyperlinks[idx] ?? [];

    row.forEach((cell, cIdx) => {
      const field = mapping[cIdx] ?? 'ignore';
      const link = allRowLinks[cIdx] ?? '';
      if (!cell && !link) return;

      switch (field) {
        case 'ignore':
          break;
        case 'fullName': {
          const parsed = splitPersonFullName(cell);
          firstName = parsed.firstName;
          lastName = parsed.lastName;
          titlePrefix = parsed.titlePrefix;
          break;
        }
        case 'firstName':
          firstName = cell;
          break;
        case 'lastName':
          lastName = cell;
          break;
        case 'email': {
          const emails = extractEmails(cell);
          if (emails.length > 0 && !email) {
            email = emails[0];
            if (emails.length > 1) noteParts.push(`Autres e-mails : ${emails.slice(1).join(', ')}`);
          } else if (cell && !email) {
            email = cell;
          }
          break;
        }
        case 'phone': {
          const found = extractPhones(cell);
          if (found.length > 0) {
            found.forEach((p) => {
              if (!phones.includes(p)) phones.push(p);
            });
          } else if (cell) {
            phones.push(cell.replace(/\n+/g, ' / ').trim());
          }
          break;
        }
        case 'website': {
          const urls = extractUrls(cell, link);
          if (urls.length > 0 && !website) website = urls[0];
          else if (cell && !website) website = cell;
          break;
        }
        case 'country':
          country = cell;
          break;
        case 'nationality':
          nationality = cell;
          break;
        case 'scope':
          scopeRaw = cell;
          break;
        case 'reference':
          reference = reference ? `${reference} · ${cell}` : cell;
          break;
        case 'yearsExperience': {
          const num = Number(cell.replace(/[^\d.]/g, ''));
          if (!Number.isNaN(num) && num > 0) yearsExperience = Math.round(num);
          else yearsExperience = Math.max(yearsExperience, parseYearsExperience(cell));
          break;
        }
        case 'domains': {
          const parts = cell
            .split(/[,;|/\n]+/)
            .map((s) => s.trim())
            .filter(Boolean);
          parts.forEach((p) => {
            if (!explicitDomains.includes(p)) explicitDomains.push(p);
          });
          break;
        }
        case 'experienceAndSpecialty':
          yearsExperience = Math.max(yearsExperience, parseYearsExperience(cell));
          noteParts.push(cell.replace(/\n+/g, ' '));
          break;
        case 'notes': {
          const headerLabel = normalizeText(table.headers[cIdx] ?? '');
          const cleaned = cell.replace(/\n+/g, ' ').trim();
          if (/(plateform|platform|source|reseau)/.test(headerLabel)) {
            noteParts.push(`Plateforme : ${cleaned}`);
          } else {
            noteParts.push(cleaned);
          }
          break;
        }
      }
    });

    // Fallback detection across row if not mapped or empty
    if (!lastName && !firstName) {
      const firstNonEmpty =
        row.find(
          (c) =>
            c.trim() !== '' &&
            !isPureEmailCell(c) &&
            extractUrls(c).length === 0
        ) ?? '';
      if (!firstNonEmpty) return;
      const parsed = splitPersonFullName(firstNonEmpty);
      firstName = parsed.firstName;
      lastName = parsed.lastName;
      titlePrefix = parsed.titlePrefix;
    }

    if (!email) {
      const fallbackEmails = extractEmails(allRowText);
      if (fallbackEmails.length > 0) email = fallbackEmails[0];
    }
    if (phones.length === 0) {
      phones.push(...extractPhones(allRowText));
    }
    if (!website) {
      const firstLink = allRowLinks.find((l) => /^https?:\/\//i.test(l));
      const fallbackUrls = extractUrls(allRowText, firstLink);
      if (fallbackUrls.length > 0) website = fallbackUrls[0];
    }
    if (!yearsExperience) {
      yearsExperience = parseYearsExperience(allRowText);
    }

    const inferredCountry = country || inferCountryFromTextOrPhone(allRowText) || DEFAULT_COUNTRY;
    const domains =
      explicitDomains.length > 0
        ? explicitDomains
        : inferDomainsFromText(allRowText, 'Environnement & EIES');

    const scope =
      /nationale/i.test(scopeRaw)
        ? 'Nationale'
        : /international/i.test(scopeRaw)
          ? 'International'
          : /international|sweep|europe|afrique|monde/i.test(allRowText)
            ? 'International'
            : 'Nationale';

    if (titlePrefix) {
      noteParts.unshift(`Titre : ${titlePrefix}`);
    }

    const code = `EXP-${String(startingNumber + experts.length).padStart(3, '0')}`;
    experts.push({
      id: `exp-ext-${Date.now()}-${Math.random().toString(36).slice(2, 6)}-${idx}`,
      code,
      lastName: lastName || firstName || 'Sans nom',
      firstName,
      scope,
      nationality,
      country: inferredCountry,
      domains,
      yearsExperience,
      reference,
      email,
      phone: phones.join(' / ') || undefined,
      website,
      attachments: [],
      projects: [],
      notes: noteParts.join(' — '),
      updatedAt: now,
    });
  });

  return experts;
}

export function buildSubcontractorsFromTable(
  table: ExtractedRawTable,
  mapping: OrgFieldKey[],
  startingNumber: number,
  mergeDuplicates = true
): Subcontractor[] {
  const now = new Date().toISOString();
  const list: Subcontractor[] = [];
  const byNameKey = new Map<string, Subcontractor>();

  table.rows.forEach((row, idx) => {
    let name = '';
    let contactName = '';
    let email = '';
    const extraEmails: string[] = [];
    const phones: string[] = [];
    let website = '';
    let country = '';
    let reference = '';
    const explicitDomains: string[] = [];
    const noteParts: string[] = [];
    const allRowText = row.join(' ');
    const allRowLinks = table.hyperlinks[idx] ?? [];

    row.forEach((cell, cIdx) => {
      const field = mapping[cIdx] ?? 'ignore';
      const link = allRowLinks[cIdx] ?? '';
      if (!cell && !link) return;

      switch (field) {
        case 'ignore':
          break;
        case 'name':
          name = cell.replace(/\n+/g, ' ').trim();
          break;
        case 'contactName':
          contactName = cell.replace(/\n+/g, ' ').trim();
          break;
        case 'email': {
          const emails = extractEmails(cell);
          emails.forEach((em) => {
            if (!email) email = em;
            else if (em !== email && !extraEmails.includes(em)) extraEmails.push(em);
          });
          if (emails.length === 0 && cell && !email) email = cell;
          break;
        }
        case 'phone': {
          const found = extractPhones(cell);
          if (found.length > 0) {
            found.forEach((p) => {
              if (!phones.includes(p)) phones.push(p);
            });
          } else if (cell) {
            phones.push(cell.replace(/\n+/g, ' / '));
          }
          break;
        }
        case 'website': {
          const urls = extractUrls(cell, link);
          if (urls.length > 0 && !website) website = urls[0];
          else if (cell && !website) website = cell;
          break;
        }
        case 'country':
          country = cell;
          break;
        case 'reference':
          reference = reference ? `${reference} · ${cell}` : cell;
          break;
        case 'domains': {
          const parts = cell
            .split(/[,;|/\n]+/)
            .map((s) => s.trim())
            .filter(Boolean);
          parts.forEach((p) => {
            if (!explicitDomains.includes(p)) explicitDomains.push(p);
          });
          break;
        }
        case 'specialtyAndNotes':
        case 'notes':
          noteParts.push(cell.replace(/\n+/g, ' ').trim());
          break;
      }
    });

    if (!name) {
      name = (
        row.find(
          (c) =>
            c.trim() !== '' &&
            !isPureEmailCell(c) &&
            extractUrls(c).length === 0
        ) ?? ''
      )
        .replace(/\n+/g, ' ')
        .trim();
      if (!name) return;
    }

    // Always harvest emails, phones & URLs across the whole row so nothing is missed
    const rowEmails = extractEmails(allRowText);
    for (const em of rowEmails) {
      if (!email) email = em;
      else if (em !== email && !extraEmails.includes(em)) extraEmails.push(em);
    }

    const rowPhones = extractPhones(allRowText);
    for (const p of rowPhones) {
      if (!phones.includes(p)) phones.push(p);
    }

    if (!website) {
      const firstLink = allRowLinks.find((l) => /^https?:\/\//i.test(l));
      const fallbackUrls = extractUrls(allRowText, firstLink);
      if (fallbackUrls.length > 0) website = fallbackUrls[0];
    }

    const inferredCountry =
      country || inferCountryFromTextOrPhone(allRowText) || DEFAULT_COUNTRY;
    const domains =
      explicitDomains.length > 0
        ? explicitDomains
        : inferDomainsFromText(allRowText, 'Génie Civil & Infrastructures');

    if (extraEmails.length > 0) {
      noteParts.push(`Autres e-mails : ${extraEmails.join(', ')}`);
    }

    const key = normalizeText(name);
    if (mergeDuplicates && byNameKey.has(key)) {
      const existing = byNameKey.get(key)!;
      const existingPhones = existing.phone
        .split('/')
        .map((p) => p.trim())
        .filter(Boolean);
      for (const p of phones) {
        if (!existingPhones.includes(p)) existingPhones.push(p);
      }
      existing.phone = existingPhones.join(' / ');
      for (const d of domains) {
        if (!existing.domains.includes(d)) existing.domains.push(d);
      }
      if (!existing.email && email) existing.email = email;
      else if (email && existing.email !== email) {
        const currentNotes = existing.notes || '';
        if (!currentNotes.includes(email)) {
          existing.notes = currentNotes ? `${currentNotes} — ${email}` : email;
        }
      }
      if (!existing.website && website) existing.website = website;
      if (!existing.contactName && contactName) existing.contactName = contactName;
      for (const np of noteParts) {
        if (np && !(existing.notes || '').includes(np)) {
          existing.notes = existing.notes ? `${existing.notes} — ${np}` : np;
        }
      }
      return;
    }

    const code = `BE-${String(startingNumber + list.length).padStart(3, '0')}`;
    const item: Subcontractor = {
      id: `sub-ext-${Date.now()}-${Math.random().toString(36).slice(2, 6)}-${idx}`,
      code,
      name,
      country: inferredCountry,
      domains,
      contactName,
      email,
      website,
      phone: phones.join(' / '),
      reference,
      notes: Array.from(new Set(noteParts)).join(' — '),
      updatedAt: now,
    };
    list.push(item);
    byNameKey.set(key, item);
  });

  return list;
}

export function buildPartnersFromTable(
  table: ExtractedRawTable,
  mapping: OrgFieldKey[],
  startingNumber: number,
  mergeDuplicates = true
): Partner[] {
  const subs = buildSubcontractorsFromTable(
    table,
    mapping,
    startingNumber,
    mergeDuplicates
  );
  return subs.map((s, idx) => ({
    id: `prt-ext-${Date.now()}-${Math.random().toString(36).slice(2, 6)}-${idx}`,
    code: `PRT-${String(startingNumber + idx).padStart(3, '0')}`,
    name: s.name,
    country: s.country,
    domains: s.domains,
    reference: s.reference,
    expertIds: [],
    email: s.email,
    website: s.website,
    notes: [
      s.contactName ? `Contact : ${s.contactName}` : '',
      s.phone ? `Tél : ${s.phone}` : '',
      s.notes ?? '',
    ]
      .filter(Boolean)
      .join(' — '),
    updatedAt: s.updatedAt,
  }));
}
