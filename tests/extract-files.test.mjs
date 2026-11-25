import test from 'node:test';
import assert from 'node:assert/strict';
import { extractWorkbook, extractDocx, formatPdfPages, joinVisionPages } from '../extract-files.mjs';
import { buildSkillContext } from '../relay-core.mjs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

test('an Excel extract names the sheet, keeps formulas, and marks full coverage', async () => {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addRow(['Voucher', 'Debit']);
  sheet.addRow(['GJ001', { formula: '50+50', result: 100 }]);
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  const result = await extractWorkbook(buffer);
  assert.match(result.text, /\[Sheet: Data\]/);
  assert.match(result.text, /GJ001 \| 100 \[=\s*50\+50\]/);
  assert.match(result.text, /Coverage: COMPLETE/);
  assert.equal(result.complete, true);
});

test('a wide Excel extract stops the raw grid and names the next offset', async () => {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Wide');
  const header = Array.from({ length: 3 }, (_, i) => `C${i + 1}`);
  sheet.addRow(header);
  for (let row = 0; row < 45; row++) sheet.addRow([`R${row}`, row, row]);
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  const result = await extractWorkbook(buffer);
  assert.match(result.text, /Coverage: INCOMPLETE/);
  assert.match(result.text, /Next offset: 40/);
  assert.equal(result.complete, false);
});

test('a Word extract keeps headings and table rows', async () => {
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.folder('word').file('document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Close memo</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Account</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Amount</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Cash</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>`);
  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  const result = await extractDocx(buffer);
  assert.match(result.text, /Close memo/);
  assert.match(result.text, /Account \| Amount/);
  assert.match(result.text, /Cash \| 100/);
});

test('PDF text is split into numbered pages and a cut keeps the next page', () => {
  const full = formatPdfPages('Hello PDF attach test\fSecond page\f');
  assert.match(full.text, /Page 1:\nHello PDF attach test/);
  assert.match(full.text, /Page 2:\nSecond page/);
  assert.match(full.text, /Coverage: COMPLETE/);
  const cut = formatPdfPages('AAAA\fBBBB\fCCCC', { maxChars: 20 });
  assert.match(cut.text, /Page 1:/);
  assert.match(cut.text, /Coverage: INCOMPLETE/);
  assert.match(cut.text, /Next page: 2/);
});

test('scanned PDF pages stay separate and pages past the cap are named', () => {
  const joined = joinVisionPages(['alpha table', 'beta 100'], { total: 5, cap: 12 });
  assert.match(joined.text, /Page 1:\nalpha table/);
  assert.match(joined.text, /Page 2:\nbeta 100/);
  assert.match(joined.text, /Coverage: INCOMPLETE/);
  assert.match(joined.text, /3 page\(s\) were not sent/);
  assert.match(joined.text, /Next page: 3/);
  assert.equal(joined.complete, false);
});

test('skill triggers select a match and do not force the default assistant', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eqgc-skill-triggers-'));
  try {
    fs.mkdirSync(path.join(dir, 'qwen-chat-assistant'));
    fs.writeFileSync(path.join(dir, 'qwen-chat-assistant', 'SKILL.md'), '---\nname: qwen-chat-assistant\ndescription: General chat assistant for every question.\n---\n\n' + 'x'.repeat(8000));
    fs.mkdirSync(path.join(dir, 'month-end-close'));
    fs.writeFileSync(path.join(dir, 'month-end-close', 'SKILL.md'), '---\nname: month-end-close\ndescription: Steps for a period close.\ntriggers: 月結, aging report\n---\n\nRead the ledger.\n');
    const picked = buildSkillContext('請做月結', dir);
    assert.deepEqual(picked.selected, ['month-end-close']);
    assert.equal(picked.text.includes('qwen-chat-assistant'), false);
    const none = buildSkillContext('hello there', dir);
    assert.deepEqual(none.selected, []);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
