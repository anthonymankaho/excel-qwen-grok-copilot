import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import mammoth from 'mammoth';

const RAW_ROWS = 40;
const RAW_COLS = 30;
const PROFILE_COLS = 200;
const SHEET_CAP = 20;
const VISION_PAGE_CAP = 12;

function columnLetter(index) {
  let n = index;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function cellDisplay(cell) {
  const value = cell?.value;
  if (value == null) return '';
  if (value instanceof Date) return String(cell.text || value.toISOString()).replace(/\s+/g, ' ').trim();
  if (typeof value === 'object' && value.formula) {
    const result = value.result == null || value.result === '' ? (cell.text || '') : value.result;
    const formula = String(value.formula).replace(/^=/, '');
    return `${result} [=${formula}]`.replace(/\s+/g, ' ').trim();
  }
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return value.richText.map(part => part.text || '').join('').replace(/\s+/g, ' ').trim();
  }
  const shown = cell.text != null && String(cell.text) !== '' ? cell.text : value;
  return String(shown).replace(/\s+/g, ' ').trim();
}

function columnProfile(sheet, column, rowCount) {
  let filled = 0;
  let blank = 0;
  let numeric = 0;
  let formulas = 0;
  let sum = 0;
  const samples = [];
  for (let row = 2; row <= rowCount; row++) {
    const cell = sheet.getRow(row).getCell(column);
    if (cell.value && typeof cell.value === 'object' && cell.value.formula) formulas += 1;
    const text = cellDisplay(cell).replace(/\s*\[=.*$/, '');
    if (!text) { blank += 1; continue; }
    filled += 1;
    if (samples.length < 8 && !samples.includes(text)) samples.push(text);
    const n = Number(text.replace(/,/g, ''));
    if (Number.isFinite(n)) { numeric += 1; sum += n; }
  }
  const header = cellDisplay(sheet.getRow(1).getCell(column)) || columnLetter(column);
  return `${columnLetter(column)} ${header}: filled ${filled}, blank ${blank}, numeric ${numeric}, sum ${Math.round(sum * 100) / 100}, formulas ${formulas}, samples ${samples.join(', ')}`;
}

export async function extractWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sections = [];
  let complete = true;
  const sheets = workbook.worksheets.slice(0, SHEET_CAP);
  if (workbook.worksheets.length > SHEET_CAP) complete = false;
  for (const sheet of sheets) {
    const rowCount = sheet.rowCount || 0;
    const colCount = sheet.columnCount || 0;
    const rawRows = Math.min(rowCount, RAW_ROWS);
    const rawCols = Math.min(colCount, RAW_COLS);
    const profileCols = Math.min(colCount, PROFILE_COLS);
    const rowsFit = rowCount <= RAW_ROWS;
    const colsFit = colCount <= RAW_COLS && colCount <= PROFILE_COLS;
    const sheetComplete = rowsFit && colsFit;
    if (!sheetComplete) complete = false;
    const coverage = sheetComplete
      ? `Coverage: COMPLETE. rows 1-${rowCount} of ${rowCount}, columns ${colCount}.`
      : `Coverage: INCOMPLETE. raw rows 1-${rawRows} of ${rowCount}, raw columns ${rawCols} of ${colCount}. Next offset: ${rawRows}.`;
    const profile = [];
    for (let column = 1; column <= profileCols; column++) profile.push(columnProfile(sheet, column, rowCount));
    const lines = [];
    for (let row = 1; row <= rawRows; row++) {
      const cells = [];
      for (let column = 1; column <= Math.max(rawCols, 1); column++) cells.push(cellDisplay(sheet.getRow(row).getCell(column)));
      lines.push(cells.join(' | '));
    }
    sections.push(`[Sheet: ${sheet.name}]\n${coverage}\n${profile.join('\n')}\n${lines.join('\n')}`.trim());
  }
  if (workbook.worksheets.length > SHEET_CAP) sections.push(`Coverage: INCOMPLETE. Further sheets were not read. Next sheet index: ${SHEET_CAP + 1}.`);
  return { text: sections.join('\n\n'), complete };
}

function decodeBasic(value) {
  return String(value || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

export function docxHtmlToText(html) {
  let src = String(html || '');
  src = src.replace(/<table\b[^>]*>[\s\S]*?<\/table>/gi, table => {
    const rows = [...table.matchAll(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi)].map(row => {
      const cells = [...row[0].matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi)]
        .map(cell => decodeBasic(cell[1].replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim());
      return cells.join(' | ');
    }).filter(Boolean);
    return `\n${rows.join('\n')}\n`;
  });
  src = src.replace(/<h([1-6])\b[^>]*>([\s\S]*?)<\/h\1>/gi, (_, level, inner) => {
    const text = decodeBasic(inner.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
    return `\n${'#'.repeat(Number(level))} ${text}\n`;
  });
  src = src.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<[^>]+>/g, '');
  return decodeBasic(src).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

export async function extractDocx(buffer) {
  const result = await mammoth.convertToHtml({ buffer });
  const text = docxHtmlToText(result.value || '');
  if (!text.trim()) throw new Error('This Word file has no readable text.');
  return { text, complete: true };
}

export function formatPdfPages(rawText, { maxChars = 160000 } = {}) {
  let parts = String(rawText || '').split('\f');
  if (parts.length && !parts[parts.length - 1].trim()) parts = parts.slice(0, -1);
  const blocks = [];
  let used = 0;
  for (const page of parts) {
    const block = `Page ${blocks.length + 1}:\n${page.trim()}`;
    if (blocks.length && used + block.length > maxChars) {
      const omitted = parts.length - blocks.length;
      return { text: `${blocks.join('\n\n')}\n\nCoverage: INCOMPLETE. ${omitted} page(s) were not sent. Next page: ${blocks.length + 1}.`, complete: false };
    }
    blocks.push(block);
    used += block.length + 2;
  }
  if (!blocks.length) return { text: '', complete: false };
  return { text: `${blocks.join('\n\n')}\n\nCoverage: COMPLETE`, complete: true };
}

export function joinVisionPages(transcriptions, { total = 0, cap = VISION_PAGE_CAP } = {}) {
  const pages = Array.isArray(transcriptions) ? transcriptions : [];
  const blocks = pages.map((text, index) => `Page ${index + 1}:\n${String(text || '').trim()}`);
  const known = Number.isInteger(total) && total > 0 ? total : 0;
  let coverage = 'Coverage: COMPLETE';
  if (known > pages.length) coverage = `Coverage: INCOMPLETE. ${known - pages.length} page(s) were not sent. Next page: ${pages.length + 1}.`;
  else if (!known && pages.length >= cap) coverage = `Coverage: INCOMPLETE. Stopped at page ${cap}. Total page count was not available.`;
  return { text: `${blocks.join('\n\n')}\n\n${coverage}`, complete: coverage === 'Coverage: COMPLETE' };
}

function runTool(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { windowsHide: true });
    let out = '';
    let err = '';
    child.stdout.on('data', part => { out += part; });
    child.stderr.on('data', part => { err += part; });
    child.on('error', error => reject(new Error(`${command} is not available on this PC: ${error.message}`)));
    child.on('close', code => resolve({ code, out, err }));
  });
}

export async function extractPdf(buffer, { transcribePage } = {}) {
  const stamp = `eqgc-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-pdf-'));
  const tmp = path.join(dir, `${stamp}.pdf`);
  try {
    fs.writeFileSync(tmp, buffer);
    const layout = await runTool('pdftotext', ['-layout', '-enc', 'UTF-8', tmp, '-']);
    if (layout.code !== 0) throw new Error(layout.err.trim() || `pdftotext exited with ${layout.code}`);
    if (layout.out.trim()) return formatPdfPages(layout.out);
    const raw = await runTool('pdftotext', ['-raw', '-enc', 'UTF-8', tmp, '-']);
    if (raw.code !== 0) throw new Error(raw.err.trim() || `pdftotext exited with ${raw.code}`);
    if (raw.out.trim()) return formatPdfPages(raw.out);
    if (typeof transcribePage !== 'function') {
      throw new Error('This PDF has no extractable text layer (likely scanned images). Add your Qwen key in Settings to enable vision transcription, or attach a text-based PDF export.');
    }
    const info = await runTool('pdfinfo', [tmp]).catch(() => ({ code: 1, out: '' }));
    const total = Number(String(info.out || '').match(/^Pages:\s+(\d+)/m)?.[1] || 0);
    const limit = Math.min(VISION_PAGE_CAP, total || VISION_PAGE_CAP);
    const pages = [];
    for (let page = 1; page <= limit; page++) {
      const prefix = path.join(dir, `${stamp}-p${page}`);
      const rendered = await runTool('pdftoppm', ['-png', '-r', '150', '-f', String(page), '-l', String(page), tmp, prefix]);
      if (rendered.code !== 0) {
        if (!pages.length) throw new Error(rendered.err.trim() || `pdftoppm exited with ${rendered.code}`);
        break;
      }
      const pngName = fs.readdirSync(dir).find(name => name.startsWith(`${stamp}-p${page}-`) && name.endsWith('.png'));
      if (!pngName) break;
      pages.push(await transcribePage(fs.readFileSync(path.join(dir, pngName)), page, total || limit));
    }
    if (!pages.length) throw new Error('No pages could be rendered from this PDF.');
    return joinVisionPages(pages, { total, cap: VISION_PAGE_CAP });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
