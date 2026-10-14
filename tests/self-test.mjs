import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Document, Packer, Paragraph, Table, TableRow, TableCell } from 'docx';

const require = createRequire(import.meta.url);
const { saveRecords } = require('../electron/wordStore.cjs');
const nativeFs = require('node:fs');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const source = await import('node:fs/promises').then(fs => fs.readFile(new URL('../src/main.js', import.meta.url), 'utf8'));
assert.match(source, /indexedDB/);
assert.match(source, /Packer\.toBlob/);
assert.match(source, /record\.id/);
assert.match(source, /localeCompare/);
assert.match(source, /data-edit/);
assert.match(source, /editingRecord\?\.id/);

const headers = ['日期', '领域', '发生了什么', '我的判断', '证据/原话', '下一步', '结果'];
const financeHeaders = ['日期', '主题/材料', '我认为的核心机制', '事实依据', '我的推论', '反例或不确定性', 'AI校正后的表述'];
const row = (values) => new TableRow({ children: values.map(value => new TableCell({ children: [new Paragraph(value)] })) });
const fixture = new Document({ sections: [{ children: [
  new Table({ rows: [row(headers), row(headers.map(() => '[填写]'))] }),
  new Table({ rows: [row(financeHeaders), row(financeHeaders.map(() => '[填写]'))] })
] }] });
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'kk-self-test-'));
const docPath = path.join(tempDir, 'KK个人成长系统.docx');
const fixtureZip = await JSZip.loadAsync(await Packer.toBuffer(fixture));
const fixtureDocument = new DOMParser().parseFromString(await fixtureZip.file('word/document.xml').async('string'), 'application/xml');
const financeTableFixture = Array.from(fixtureDocument.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'tbl')).find(table => {
  const firstRow = Array.from(table.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr');
  return firstRow && Array.from(firstRow.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')).map(node => node.textContent).join('|') === financeHeaders.join('|');
});
const financeGrid = Array.from(financeTableFixture.childNodes).find(node => node.nodeType === 1 && node.localName === 'tblGrid');
const financeHeader = Array.from(financeTableFixture.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr');
financeTableFixture.insertBefore(financeHeader, financeGrid);
fixtureZip.file('word/document.xml', new XMLSerializer().serializeToString(fixtureDocument));
await fs.writeFile(docPath, await fixtureZip.generateAsync({ type: 'nodebuffer' }));
const records = [
  { id: '1', type: '每日或事件记录', values: { 日期: '2099-10-03', 领域: '工作', 发生了什么: 'three', 我的判断: '', '证据/原话': '', 下一步: '', 结果: '' } },
  { id: '2', type: '每日或事件记录', values: { 日期: '2099-10-02', 领域: '工作', 发生了什么: 'two', 我的判断: '', '证据/原话': '', 下一步: '', 结果: '' } }
];
const originalRename = nativeFs.promises.rename;
let renameAttempts = 0;
nativeFs.promises.rename = async (...args) => {
  renameAttempts += 1;
  if (renameAttempts <= 2) {
    const error = new Error('temporary file lock');
    error.code = 'EPERM';
    throw error;
  }
  return originalRename(...args);
};
let result;
try { result = await saveRecords(docPath, records); }
finally { nativeFs.promises.rename = originalRename; }
assert.equal(result.added, 2);
assert.equal(result.updated, 0);
assert.equal(renameAttempts, 4);
assert.equal((await fs.readdir(path.join(tempDir, 'backups'))).length, 1);
assert.equal((await fs.readdir(tempDir)).some(name => name.endsWith('.tmp')), false);
assert.equal((await fs.readFile(`${docPath}.kk-index.json`, 'utf8')).includes('"1"'), true);
const savedZip = await JSZip.loadAsync(await fs.readFile(docPath));
const savedDocument = new DOMParser().parseFromString(await savedZip.file('word/document.xml').async('string'), 'application/xml');
const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const recordTable = Array.from(savedDocument.getElementsByTagNameNS(wordNs, 'tbl')).find(table => {
  const firstRow = Array.from(table.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr');
  return firstRow && Array.from(firstRow.getElementsByTagNameNS(wordNs, 't')).map(node => node.textContent).join('|') === headers.join('|');
});
const financeTable = Array.from(savedDocument.getElementsByTagNameNS(wordNs, 'tbl')).find(table => {
  const firstRow = Array.from(table.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr');
  return firstRow && Array.from(firstRow.getElementsByTagNameNS(wordNs, 't')).map(node => node.textContent).join('|') === financeHeaders.join('|');
});
const widthOf = node => Number(node.getAttributeNS(wordNs, 'w'));
for (const table of [recordTable, financeTable]) {
  const tableChildren = Array.from(table.childNodes).filter(node => node.nodeType === 1).map(node => node.localName);
  assert.ok(tableChildren.indexOf('tblPr') < tableChildren.indexOf('tblGrid'));
  assert.ok(tableChildren.indexOf('tblGrid') < tableChildren.indexOf('tr'));
}
assert.equal(recordTable.getElementsByTagNameNS(wordNs, 'tblLayout')[0].getAttributeNS(wordNs, 'type'), 'fixed');
const gridWidths = Array.from(recordTable.getElementsByTagNameNS(wordNs, 'gridCol')).map(widthOf);
assert.equal(widthOf(recordTable.getElementsByTagNameNS(wordNs, 'tblW')[0]), gridWidths.reduce((sum, width) => sum + width, 0));
for (const row of Array.from(recordTable.childNodes).filter(node => node.nodeType === 1 && node.localName === 'tr')) {
  assert.deepEqual(Array.from(row.getElementsByTagNameNS(wordNs, 'tcW')).map(widthOf), gridWidths);
  assert.equal(row.getElementsByTagNameNS(wordNs, 'cantSplit').length, 0);
}

const legacyRow = Array.from(recordTable.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr' && Array.from(node.getElementsByTagNameNS(wordNs, 't')).some(text => text.textContent === 'three'));
for (const start of Array.from(legacyRow.getElementsByTagNameNS(wordNs, 'bookmarkStart'))) {
  const id = start.getAttributeNS(wordNs, 'id');
  const end = Array.from(legacyRow.getElementsByTagNameNS(wordNs, 'bookmarkEnd')).find(node => node.getAttributeNS(wordNs, 'id') === id);
  start.parentNode.removeChild(start);
  if (end) end.parentNode.removeChild(end);
}
savedZip.file('word/document.xml', new XMLSerializer().serializeToString(savedDocument));
await fs.writeFile(docPath, await savedZip.generateAsync({ type: 'nodebuffer' }));
const editedLegacy = { ...records[0], values: { ...records[0].values, 发生了什么: 'three\ncontinued' } };
result = await saveRecords(docPath, [editedLegacy]);
assert.equal(result.added, 0);
assert.equal(result.updated, 1);
const editedMarked = { ...records[1], values: { ...records[1].values, 结果: 'follow-up' } };
result = await saveRecords(docPath, [editedMarked]);
assert.equal(result.added, 0);
assert.equal(result.updated, 1);
const editedZip = await JSZip.loadAsync(await fs.readFile(docPath));
const editedDocument = new DOMParser().parseFromString(await editedZip.file('word/document.xml').async('string'), 'application/xml');
const editedTable = Array.from(editedDocument.getElementsByTagNameNS(wordNs, 'tbl')).find(table => {
  const firstRow = Array.from(table.childNodes).find(node => node.nodeType === 1 && node.localName === 'tr');
  return firstRow && Array.from(firstRow.getElementsByTagNameNS(wordNs, 't')).map(node => node.textContent).join('|') === headers.join('|');
});
const editedRows = Array.from(editedTable.childNodes).filter(node => node.nodeType === 1 && node.localName === 'tr').slice(1).filter(row => !Array.from(row.getElementsByTagNameNS(wordNs, 't')).every(node => node.textContent === '[填写]'));
assert.equal(editedRows.length, 2);
assert.ok(editedRows.some(row => Array.from(row.getElementsByTagNameNS(wordNs, 't')).map(node => node.textContent).join('') === '2099-10-03工作threecontinued'));
assert.ok(editedRows.some(row => Array.from(row.getElementsByTagNameNS(wordNs, 't')).some(node => node.textContent === 'follow-up')));
assert.equal(editedRows.every(row => row.getElementsByTagNameNS(wordNs, 'bookmarkStart').length === 1), true);
const secondLegacyRow = editedRows.find(row => Array.from(row.getElementsByTagNameNS(wordNs, 't')).some(node => node.textContent === 'continued'));
for (const start of Array.from(secondLegacyRow.getElementsByTagNameNS(wordNs, 'bookmarkStart'))) {
  const id = start.getAttributeNS(wordNs, 'id');
  const end = Array.from(secondLegacyRow.getElementsByTagNameNS(wordNs, 'bookmarkEnd')).find(node => node.getAttributeNS(wordNs, 'id') === id);
  start.parentNode.removeChild(start);
  if (end) end.parentNode.removeChild(end);
}
editedZip.file('word/document.xml', new XMLSerializer().serializeToString(editedDocument));
await fs.writeFile(docPath, await editedZip.generateAsync({ type: 'nodebuffer' }));
const indexPath = `${docPath}.kk-index.json`;
const updatedIndex = JSON.parse(await fs.readFile(indexPath, 'utf8'));
updatedIndex.ids['1'].signature = 'stale-signature';
await fs.writeFile(indexPath, JSON.stringify(updatedIndex));
result = await saveRecords(docPath, [{ ...editedLegacy, values: { ...editedLegacy.values, 结果: 'reviewed' }, previousValues: editedLegacy.values }]);
assert.equal(result.added, 0);
assert.equal(result.updated, 1);
const finalZip = await JSZip.loadAsync(await fs.readFile(docPath));
const finalDocument = new DOMParser().parseFromString(await finalZip.file('word/document.xml').async('string'), 'application/xml');
assert.ok(Array.from(finalDocument.getElementsByTagNameNS(wordNs, 't')).some(node => node.textContent === 'reviewed'));
assert.equal((await fs.readdir(path.join(tempDir, 'backups'))).length, 4);

const beforeLockedWrite = await fs.readFile(docPath);
nativeFs.promises.rename = async (source, destination) => {
  if (String(destination).endsWith('.pending.docx')) return originalRename(source, destination);
  const error = new Error('document is locked');
  error.code = 'EPERM';
  throw error;
};
try {
  await assert.rejects(
    saveRecords(docPath, [{ ...records[0], id: '3', values: { ...records[0].values, 日期: '2099-10-04' } }]),
    /文件可能正被 Word.*原文档未被替换.*backups.*pending\.docx/
  );
} finally { nativeFs.promises.rename = originalRename; }
assert.deepEqual(await fs.readFile(docPath), beforeLockedWrite);
assert.equal((await fs.readdir(path.join(tempDir, 'backups'))).length, 5);
assert.equal((await fs.readdir(tempDir)).some(name => name.endsWith('.tmp')), false);
assert.equal((await fs.readdir(tempDir)).some(name => name.endsWith('.pending.docx')), true);
console.log('cross-platform source self-test passed');
