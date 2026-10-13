import { Document, Packer, Paragraph, HeadingLevel, Table, TableCell, TableRow, WidthType, TableLayoutType, ShadingType, AlignmentType } from 'docx';
import './styles.css';

const TYPES = {
  '每日或事件记录': {
    headers: ['日期', '领域', '发生了什么', '我的判断', '证据/原话', '下一步', '结果'],
    columnWidths: [900, 1000, 1600, 1500, 1650, 1400, 1310],
    fields: [
      ['日期', 'date'], ['领域', 'choice', ['工作', '英语', '金融', '认知', '逻辑', '审美', '健身', '其他']],
      ['发生了什么', 'long'], ['我的判断', 'long'], ['证据/原话', 'long'], ['下一步', 'long'], ['结果', 'long']
    ]
  },
  '金融认知记录': {
    headers: ['日期', '主题/材料', '我认为的核心机制', '事实依据', '我的推论', '反例或不确定性', 'AI校正后的表述'],
    columnWidths: [900, 1300, 1500, 1250, 1350, 1550, 1510],
    fields: [['日期', 'date'], ['主题/材料', 'long'], ['我认为的核心机制', 'long'], ['事实依据', 'long'], ['我的推论', 'long'], ['反例或不确定性', 'long'], ['AI校正后的表述', 'long']]
  },
  '外贸案例记录': {
    headers: ['日期', '客户阶段', '需求与证据', '我的动作', '客户反应', '复盘结论', '下次实验'],
    columnWidths: [900, 1200, 1600, 1450, 1250, 1450, 1510],
    fields: [['日期', 'date'], ['客户阶段', 'choice', ['首次询盘', '已报价', '价格异议', '沉默', '样品', '谈判', '成交', '其他']], ['需求与证据', 'long'], ['我的动作', 'long'], ['客户反应', 'long'], ['复盘结论', 'long'], ['下次实验', 'long']]
  }
};

const DB_NAME = 'kk-growth-record';
const DB_VERSION = 1;
const STORE = 'records';
const app = document.querySelector('#app');
let currentType = '每日或事件记录';

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      const store = db.createObjectStore(STORE, { keyPath: 'id' });
      store.createIndex('type', 'type');
      store.createIndex('date', 'date');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function allRecords() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE).objectStore(STORE).getAll();
    request.onsuccess = () => resolve(request.result.sort((a, b) => `${a.date}${a.updatedAt}`.localeCompare(`${b.date}${b.updatedAt}`)));
    request.onerror = () => reject(request.error);
  });
}

async function putRecord(record) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE, 'readwrite').objectStore(STORE).put(record);
    request.onsuccess = resolve;
    request.onerror = () => reject(request.error);
  });
}

async function deleteRecord(id) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE, 'readwrite').objectStore(STORE).delete(id);
    request.onsuccess = resolve;
    request.onerror = () => reject(request.error);
  });
}

function uid() { return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`; }
function escapeHtml(value) { return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch])); }
function today() { return new Date().toISOString().slice(0, 10); }

function render() {
  const definition = TYPES[currentType];
  app.innerHTML = `
    <div class="shell">
      <aside class="sidebar">
        <div class="brand"><h1>KK 成长记录</h1><p>工作、学习与认知的长期数据库</p></div>
        <div class="side-label">记录类型</div>
        <nav class="type-nav">${Object.keys(TYPES).map(type => `<button class="type-button ${type === currentType ? 'active' : ''}" data-type="${escapeHtml(type)}">${escapeHtml(type)}</button>`).join('')}</nav>
        <div class="side-tools"><button class="button" id="choose-word" ${window.kkDesktop ? '' : 'hidden'}>选择 Word 文档</button><button class="button" id="export">${window.kkDesktop ? '打开 Word 文档' : '导出 Word 文档'}</button><button class="button" id="export-json">导出记录 JSON</button><button class="button" id="import">导入记录 JSON</button><input id="import-file" type="file" accept="application/json" hidden /></div>
        <div class="muted">电脑端保存会直接写入 KK个人成长系统.docx，并先自动备份。手机端保存到本地，导出 JSON 后在电脑端导入。</div>
      </aside>
      <main class="main">
        <div class="topbar"><strong id="storage-title">本地记录</strong><span class="path" id="storage-path">手机端保存在本地；电脑端可直接写入 Word</span></div>
        <section class="panel">
          <div class="panel-head"><h2>${escapeHtml(currentType)}</h2><span>保存后可随时导出为 Word</span></div>
          <form class="form" id="record-form">${definition.fields.map(([label, kind, options]) => fieldHtml(label, kind, options)).join('')}</form>
          <div class="actions"><div class="message" id="message"></div><div class="action-buttons"><button class="button" type="button" id="clear">清空</button><button class="button primary" type="button" id="save">保存记录</button></div></div>
          <div class="records"><h3>已保存记录</h3><div id="record-list"></div></div>
        </section>
      </main>
    </div>`;
  document.querySelectorAll('[data-type]').forEach(button => button.addEventListener('click', () => { currentType = button.dataset.type; render(); }));
  document.querySelector('#clear').addEventListener('click', clearForm);
  document.querySelector('#save').addEventListener('click', saveForm);
  document.querySelector('#export').addEventListener('click', exportWord);
  document.querySelector('#choose-word').addEventListener('click', chooseWord);
  document.querySelector('#export-json').addEventListener('click', exportJson);
  document.querySelector('#import').addEventListener('click', () => document.querySelector('#import-file').click());
  document.querySelector('#import-file').addEventListener('change', importJson);
  refreshDesktopStatus();
  loadList();
}

function fieldHtml(label, kind, options) {
  const id = `field-${Object.keys(TYPES[currentType].fields).length}-${label}`;
  if (kind === 'choice') return `<label for="${escapeHtml(id)}">${escapeHtml(label)}</label><select id="${escapeHtml(id)}" data-label="${escapeHtml(label)}">${options.map(option => `<option>${escapeHtml(option)}</option>`).join('')}</select>`;
  if (kind === 'date') return `<label for="${escapeHtml(id)}">${escapeHtml(label)}</label><input id="${escapeHtml(id)}" data-label="${escapeHtml(label)}" type="date" value="${today()}" />`;
  return `<label for="${escapeHtml(id)}">${escapeHtml(label)}</label><textarea id="${escapeHtml(id)}" data-label="${escapeHtml(label)}" placeholder="填写${escapeHtml(label)}"></textarea>`;
}

function formValues() { return TYPES[currentType].fields.map(([label]) => [label, document.querySelector(`[data-label="${CSS.escape(label)}"]`).value.trim()]); }
function clearForm() { document.querySelector('#record-form').reset(); document.querySelector('[type=date]').value = today(); }

async function saveForm() {
  const pairs = formValues();
  const values = Object.fromEntries(pairs);
  const message = document.querySelector('#message');
  if (!values['日期'] || !pairs.slice(1).some(([, value]) => value)) { message.textContent = '请填写日期，并至少填写一项内容。'; message.className = 'message error'; return; }
  const record = { id: uid(), type: currentType, date: values['日期'], values, updatedAt: new Date().toISOString() };
  try {
    const result = window.kkDesktop ? await window.kkDesktop.saveRecord(record) : null;
    await putRecord(record);
    message.textContent = result ? `已直接写入 Word（新增 ${result.added} 条，已备份原文件）。` : '已保存到本地数据库。';
    message.className = 'message';
  } catch (error) {
    message.textContent = `保存失败：${error.message || error}`;
    message.className = 'message error';
    return;
  }
  clearForm();
  await loadList();
}

async function loadList() {
  const list = document.querySelector('#record-list');
  if (!list) return;
  const records = (await allRecords()).filter(record => record.type === currentType).reverse();
  if (!records.length) { list.innerHTML = '<div class="empty">还没有记录。</div>'; return; }
  const headers = TYPES[currentType].headers;
  list.innerHTML = `<table><thead><tr>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join('')}<th>操作</th></tr></thead><tbody>${records.map(record => `<tr>${headers.map(header => `<td>${escapeHtml(record.values[header])}</td>`).join('')}<td><button class="button danger" data-delete="${record.id}">删除</button></td></tr>`).join('')}</tbody></table>`;
  list.querySelectorAll('[data-delete]').forEach(button => button.addEventListener('click', async () => { if (confirm('确定删除这条记录吗？')) { await deleteRecord(button.dataset.delete); loadList(); } }));
}

function paragraph(text, bold = false) { return new Paragraph({ children: [{ text: text || '', bold }], spacing: { after: 80 } }); }
function wordCell(text, width, header = false) { return new TableCell({ width: { size: width, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, shading: header ? { fill: '2F5597', type: ShadingType.CLEAR } : undefined, children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { before: 0, after: 0 }, children: [{ text: text || '', bold: header, color: header ? 'FFFFFF' : '000000', size: 18 }] })] }); }
function wordTable(headers, records, columnWidths) {
  return new Table({ layout: TableLayoutType.FIXED, width: { size: 9360, type: WidthType.DXA }, columnWidths, rows: [new TableRow({ tableHeader: true, children: headers.map((header, index) => wordCell(header, columnWidths[index], true)) }), ...records.map(record => new TableRow({ children: headers.map((header, index) => wordCell(record.values[header] || '', columnWidths[index])) }))] });
}

async function exportWord() {
  if (window.kkDesktop) {
    try { await window.kkDesktop.openWord(); } catch (error) { document.querySelector('#message').textContent = `打开 Word 失败：${error.message || error}`; }
    return;
  }
  const records = await allRecords();
  const children = [new Paragraph({ text: 'KK个人成长记录', heading: HeadingLevel.TITLE }), paragraph(`导出时间：${new Date().toLocaleString('zh-CN')}`), paragraph('记录保存在当前设备本地数据库中；本文件是当前数据的 Word 导出副本。')];
  for (const type of Object.keys(TYPES)) {
    const typeRecords = records.filter(record => record.type === type).sort((a, b) => a.date.localeCompare(b.date));
    children.push(new Paragraph({ text: type, heading: HeadingLevel.HEADING_1 }));
    children.push(typeRecords.length ? wordTable(TYPES[type].headers, typeRecords, TYPES[type].columnWidths) : paragraph('暂无记录。'));
  }
  const blob = await Packer.toBlob(new Document({ sections: [{ properties: {}, children }] }));
  downloadBlob(blob, `KK成长记录_${today()}.docx`);
}

function downloadBlob(blob, filename) { const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = filename; link.click(); setTimeout(() => URL.revokeObjectURL(link.href), 1000); }

async function exportJson() {
  const records = await allRecords();
  downloadBlob(new Blob([JSON.stringify(records, null, 2)], { type: 'application/json' }), `KK成长记录_${today()}.json`);
}

async function importJson(event) {
  const file = event.target.files[0];
  if (!file) return;
  const records = JSON.parse(await file.text());
  const valid = records.filter(record => record.id && record.type && record.values);
  let result = null;
  if (window.kkDesktop) result = await window.kkDesktop.mergeRecords(valid);
  for (const record of valid) await putRecord(record);
  await loadList();
  document.querySelector('#message').textContent = result ? `JSON 已合并到 Word：新增 ${result.added} 条，跳过 ${result.skipped} 条。` : 'JSON 记录已合并到本地数据库。';
  event.target.value = '';
}

async function chooseWord() {
  if (!window.kkDesktop) { document.querySelector('#message').textContent = '手机端不能选择电脑本地 Word，请在电脑 App 中操作。'; return; }
  const result = await window.kkDesktop.chooseWord();
  if (!result.canceled) refreshDesktopStatus();
}

async function refreshDesktopStatus() {
  if (!window.kkDesktop) return;
  const status = await window.kkDesktop.getWordStatus();
  document.querySelector('#storage-title').textContent = status.exists ? 'Word 主文档已连接' : '尚未连接 Word';
  document.querySelector('#storage-path').textContent = status.path || '请点击左侧“选择 Word 文档”';
}

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register(`${import.meta.env.BASE_URL}sw.js`).catch(() => {});
}
render();
