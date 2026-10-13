const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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

function clearRowFlag(row, name) {
  const trPr = firstChild(row, 'trPr');
  const flag = trPr && firstChild(trPr, name);
  if (flag) trPr.removeChild(flag);
  if (trPr && !Array.from(trPr.childNodes || []).some((node) => node.nodeType === 1)) row.removeChild(trPr);
}

function wordWidth(node) {
  return Number(node?.getAttributeNS(WORD_NS, 'w') || node?.getAttribute(q('w')) || 0);
}

function setWidth(node, width, withType = true) {
  node.setAttribute(q('w'), String(width));
  if (withType) node.setAttribute(q('type'), 'dxa');
}

function normalizeTableOrder(table) {
  const properties = firstChild(table, 'tblPr');
  const grid = firstChild(table, 'tblGrid');
  if (properties && table.firstChild !== properties) table.insertBefore(properties, table.firstChild);
  if (grid) {
    const following = properties ? properties.nextSibling : table.firstChild;
    if (following !== grid) table.insertBefore(grid, following || null);
  }
}

function setFixedTableLayout(table) {
  const rows = Array.from(table.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tr'));
  const count = rowCells(rows[0] || {}).length;
  if (!count) return;

  let properties = firstChild(table, 'tblPr');
  if (!properties) { properties = table.ownerDocument.createElement(q('tblPr')); table.insertBefore(properties, table.firstChild); }
  let layout = firstChild(properties, 'tblLayout');
  if (!layout) {
    layout = table.ownerDocument.createElement(q('tblLayout'));
    const following = Array.from(properties.childNodes || []).find((node) => node.nodeType === 1 && ['tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'].includes(node.localName));
    if (following) properties.insertBefore(layout, following);
    else properties.appendChild(layout);
  }
  layout.setAttribute(q('type'), 'fixed');

  let grid = firstChild(table, 'tblGrid');
  if (!grid) { grid = table.ownerDocument.createElement(q('tblGrid')); table.insertBefore(grid, rows[0]); }
  let columns = Array.from(grid.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('gridCol'));
  let widths = columns.map(wordWidth);
  if (widths.length !== count || widths.some((width) => !width)) {
    widths = rowCells(rows[0]).map((cell) => wordWidth(firstChild(firstChild(cell, 'tcPr') || {}, 'tcW')));
  }
  if (widths.length !== count || widths.some((width) => !width)) {
    const base = Math.floor(10000 / count);
    widths = Array.from({ length: count }, (_, index) => index === count - 1 ? 10000 - base * (count - 1) : base);
  }
  if (columns.length !== count) {
    columns.forEach((column) => grid.removeChild(column));
    columns = widths.map(() => grid.appendChild(table.ownerDocument.createElement(q('gridCol'))));
  }
  columns.forEach((column, index) => setWidth(column, widths[index], false));

  let tableWidth = firstChild(properties, 'tblW');
  if (!tableWidth) {
    tableWidth = table.ownerDocument.createElement(q('tblW'));
    const following = Array.from(properties.childNodes || []).find((node) => node.nodeType === 1 && ['jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblPrChange'].includes(node.localName));
    if (following) properties.insertBefore(tableWidth, following);
    else properties.appendChild(tableWidth);
  }
  setWidth(tableWidth, widths.reduce((sum, width) => sum + width, 0));

  rows.forEach((row) => rowCells(row).forEach((cell, index) => {
    const width = widths[index];
    if (!width) return;
    let cellProperties = firstChild(cell, 'tcPr');
    if (!cellProperties) { cellProperties = table.ownerDocument.createElement(q('tcPr')); cell.insertBefore(cellProperties, cell.firstChild); }
    let cellWidth = firstChild(cellProperties, 'tcW');
    if (!cellWidth) { cellWidth = table.ownerDocument.createElement(q('tcW')); cellProperties.insertBefore(cellWidth, cellProperties.firstChild); }
    setWidth(cellWidth, width);
  }));
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
  rows.slice(1).forEach((row) => clearRowFlag(row, 'cantSplit'));
  setFixedTableLayout(table);
}

function replaceRows(table, rows) {
  Array.from(table.childNodes || []).filter((node) => node.nodeType === 1 && node.nodeName === q('tr')).forEach((row) => table.removeChild(row));
  const properties = firstChild(table, 'tblPr');
  const grid = firstChild(table, 'tblGrid');
  const anchor = grid || properties;
  if (anchor) table.insertBefore(rows[0], anchor.nextSibling);
  else rows.forEach((row) => table.appendChild(row));
  rows.slice(1).forEach((row) => table.appendChild(row));
  normalizeTableOrder(table);
}

async function readIndex(docPath) {
  try { return JSON.parse(await fsp.readFile(`${docPath}.kk-index.json`, 'utf8')); }
  catch { return { ids: {}, signatures: {} }; }
}

async function writeAtomic(filePath, data) {
  const temp = `${filePath}.${process.pid}.${crypto.randomUUID()}.tmp`;
  let dataWritten = false;
  try {
    await fsp.writeFile(temp, data, { flag: 'wx' });
    dataWritten = true;
    for (let attempt = 0; ; attempt += 1) {
      try {
        await fsp.rename(temp, filePath);
        return;
      } catch (error) {
        if (!['EPERM', 'EACCES', 'EBUSY'].includes(error.code) || attempt === 4) throw error;
        await new Promise((resolve) => setTimeout(resolve, 150 * (2 ** attempt)));
      }
    }
  } catch (error) {
    if (['EPERM', 'EACCES', 'EBUSY'].includes(error.code)) {
      let recoveryPath = temp;
      if (dataWritten) {
        const pending = `${filePath}.${process.pid}.${crypto.randomUUID()}.pending.docx`;
        try {
          await fsp.rename(temp, pending);
          recoveryPath = pending;
        } catch {}
      } else {
        await fsp.rm(temp, { force: true }).catch(() => {});
      }
      throw new Error(`无法替换文件：${filePath}。文件可能正被 Word、文件预览或云盘同步占用。请关闭它们后重试。原文档未被替换，备份保留在同目录的 backups 文件夹；本次待保存的文档保留在：${recoveryPath}`, { cause: error });
    }
    await fsp.rm(temp, { force: true }).catch(() => {});
    throw error;
  }
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
  Array.from(document.getElementsByTagName(q('tbl'))).forEach(normalizeTableOrder);
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
      clearRowFlag(row, 'cantSplit');
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
