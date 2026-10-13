const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const TABLES = {
  '每日或事件记录': ['日期', '领域', '发生了什么', '我的判断', '证据/原话', '下一步', '结果'],
  '金融认知记录': ['日期', '主题/材料', '我认为的核心机制', '事实依据', '我的推论', '反例或不确定性', 'AI校正后的表述'],
  '外贸案例记录': ['日期', '客户阶段', '需求与证据', '我的动作', '客户反应', '复盘结论', '下次实验']
};

const q = (name) => `w:${name}`;
const textOf = (node) => Array.from(node.getElementsByTagName(q('t'))).map((item) => item.textContent || '').join('');
const rowCells = (row) => Array.from(row.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tc'));
const rowValues = (row) => rowCells(row).map(textOf);
const isPlaceholder = (values) => values.every((value) => !value.trim() || /^\[|^_+$/.test(value.trim()));
const canonical = (record) => JSON.stringify({ type: record.type, values: record.values });
const signature = (record) => crypto.createHash('sha256').update(canonical(record)).digest('hex');

function clone(node) {
  return node.cloneNode(true);
}

function firstChild(node, name) {
  return Array.from(node.childNodes || []).find((child) => child.nodeType === 1 && child.nodeName === q(name));
}

function setRowFlag(row, name) {
  let trPr = firstChild(row, 'trPr');
  if (!trPr) { trPr = row.ownerDocument.createElement(q('trPr')); row.insertBefore(trPr, row.firstChild); }
  let flag = firstChild(trPr, name);
  if (!flag) { flag = row.ownerDocument.createElement(q(name)); trPr.appendChild(flag); }
  flag.setAttribute(q('val'), '1');
}

function setCellText(cell, value) {
  const document = cell.ownerDocument;
  let paragraph = firstChild(cell, 'p');
  if (!paragraph) { paragraph = document.createElement(q('p')); cell.appendChild(paragraph); }
  const paragraphProperties = firstChild(paragraph, 'pPr');
  let styleRunProperties = null;
  const firstRun = Array.from(paragraph.childNodes || []).find((node) => node.nodeType === 1 && node.nodeName === q('r'));
  if (firstRun) styleRunProperties = firstChild(firstRun, 'rPr');
  Array.from(paragraph.childNodes || []).forEach((child) => {
    if (child !== paragraphProperties) paragraph.removeChild(child);
  });
  const run = document.createElement(q('r'));
  if (styleRunProperties) run.appendChild(clone(styleRunProperties));
  const lines = String(value || '').split(/\r?\n/);
  lines.forEach((line, index) => {
    if (index) run.appendChild(document.createElement(q('br')));
    const t = document.createElement(q('t'));
    if (/^\s|\s$/.test(line)) t.setAttribute('xml:space', 'preserve');
    t.appendChild(document.createTextNode(line));
    run.appendChild(t);
  });
  paragraph.appendChild(run);
  Array.from(cell.childNodes || []).forEach((child) => {
    if (child !== firstChild(cell, 'tcPr') && child !== paragraph) cell.removeChild(child);
  });
}

function findTable(document, headers) {
  return Array.from(document.getElementsByTagName(q('tbl'))).find((table) => {
    const firstRow = Array.from(table.childNodes || []).find((node) => node.nodeType === 1 && node.nodeName === q('tr'));
    return firstRow && JSON.stringify(rowValues(firstRow)) === JSON.stringify(headers);
  });
}

function prepareTable(table) {
  const rows = Array.from(table.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tr'));
  if (rows[0]) {
    let trPr = firstChild(rows[0], 'trPr');
    if (!trPr) { trPr = table.ownerDocument.createElement(q('trPr')); rows[0].insertBefore(trPr, rows[0].firstChild); }
    let header = firstChild(trPr, 'tblHeader');
    if (!header) { header = table.ownerDocument.createElement(q('tblHeader')); trPr.appendChild(header); }
    header.setAttribute(q('val'), '1');
  }
  rows.slice(1).forEach((row) => setRowFlag(row, 'cantSplit'));
}

function replaceRows(table, rows) {
  const document = table.ownerDocument;
  Array.from(table.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tr')).forEach((row) => table.removeChild(row));
  const properties = firstChild(table, 'tblPr');
  if (properties) table.insertBefore(rows[0], properties.nextSibling);
  else rows.forEach((row) => table.appendChild(row));
  rows.slice(1).forEach((row) => table.appendChild(row));
  return document;
}

async function readIndex(docPath) {
  try { return JSON.parse(await fsp.readFile(`${docPath}.kk-index.json`, 'utf8')); }
  catch { return { ids: {}, signatures: {} }; }
}

async function writeAtomic(filePath, data) {
  const temp = `${filePath}.${process.pid}.tmp`;
  await fsp.writeFile(temp, data);
  await fsp.rename(temp, filePath);
}

async function findDefaultDocument() {
  const candidates = [
    path.resolve(process.cwd(), '..', 'KK个人成长系统.docx'),
    path.resolve(process.cwd(), '..', '..', 'KK个人成长系统.docx'),
    path.resolve(__dirname, '..', '..', 'KK个人成长系统.docx'),
    path.resolve(process.execPath, '..', '..', '..', 'KK个人成长系统.docx'),
    path.resolve(process.execPath, '..', '..', '..', '..', 'KK个人成长系统.docx'),
    path.resolve(process.execPath, '..', '..', 'KK个人成长系统.docx')
  ];
  for (const candidate of candidates) if (fs.existsSync(candidate)) return candidate;
  return null;
}

async function saveRecords(docPath, records) {
  if (!docPath) throw new Error('请先选择 KK个人成长系统.docx');
  const resolved = path.resolve(docPath);
  await fsp.access(resolved, fs.constants.R_OK | fs.constants.W_OK);
  const index = await readIndex(resolved);
  const zip = await JSZip.loadAsync(await fsp.readFile(resolved));
  const xml = await zip.file('word/document.xml').async('string');
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const grouped = new Map();
  for (const record of records) {
    if (!TABLES[record.type] || !record.values || !record.values['日期']) continue;
    const id = record.id || crypto.randomUUID();
    const normalized = { id, type: record.type, values: record.values, date: record.values['日期'] };
    const sig = signature(normalized);
    if (index.ids[id] || index.signatures[sig]) continue;
    if (!grouped.has(record.type)) grouped.set(record.type, []);
    grouped.get(record.type).push(normalized);
    index.ids[id] = { type: record.type, signature: sig };
    index.signatures[sig] = id;
  }
  if (!grouped.size) return { added: 0, skipped: records.length, backup: null, path: resolved };
  let added = 0;
  for (const [type, incoming] of grouped) {
    const headers = TABLES[type];
    const table = findTable(document, headers);
    if (!table) throw new Error(`找不到 Word 表格：${type}`);
    prepareTable(table);
    const rows = Array.from(table.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tr'));
    const header = rows[0];
    const data = rows.slice(1);
    const placeholders = data.filter((row) => isPlaceholder(rowValues(row)));
    const actual = data.filter((row) => !isPlaceholder(rowValues(row)));
    for (const record of incoming) {
      let row = placeholders.shift();
      if (!row) row = clone(actual[actual.length - 1] || data[data.length - 1] || header);
      const cells = rowCells(row);
      headers.forEach((field, index) => setCellText(cells[index], record.values[field] || ''));
      setRowFlag(row, 'cantSplit');
      actual.push(row);
      added += 1;
    }
    actual.sort((a, b) => (rowValues(a)[0] || '').localeCompare(rowValues(b)[0] || ''));
    replaceRows(table, [header, ...actual, ...placeholders]);
  }
  const backupDir = path.join(path.dirname(resolved), 'backups');
  await fsp.mkdir(backupDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backup = path.join(backupDir, `${path.basename(resolved, '.docx')}_${stamp}.docx`);
  await fsp.copyFile(resolved, backup);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(document));
  await writeAtomic(resolved, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  await writeAtomic(`${resolved}.kk-index.json`, JSON.stringify(index, null, 2));
  return { added, skipped: records.length - added, backup, path: resolved };
}

async function saveRecord(docPath, record) { return saveRecords(docPath, [record]); }

module.exports = { TABLES, findDefaultDocument, saveRecord, saveRecords };
