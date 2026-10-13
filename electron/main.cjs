const { app, BrowserWindow, shell, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { findDefaultDocument, saveRecord, saveRecords } = require('./wordStore.cjs');

const isDev = !app.isPackaged;
let mainWindow;
const configPath = path.join(app.getPath('userData'), 'kk-config.json');

function readConfig() {
  try { return JSON.parse(fs.readFileSync(configPath, 'utf8')); } catch { return {}; }
}
function writeConfig(config) {
  fs.mkdirSync(path.dirname(configPath), { recursive: true });
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
}
async function selectedDocument() {
  const configured = readConfig().documentPath;
  if (configured && fs.existsSync(configured)) return configured;
  return findDefaultDocument();
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 800,
    minWidth: 900,
    minHeight: 620,
    title: 'KK 成长记录',
    backgroundColor: '#eef2f6',
    webPreferences: { contextIsolation: true, sandbox: true, preload: path.join(__dirname, 'preload.cjs') }
  });
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  if (isDev) mainWindow.loadURL('http://localhost:5173');
  else mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
}

ipcMain.handle('word:get-status', async () => {
  const file = await selectedDocument();
  return { path: file, exists: Boolean(file && fs.existsSync(file)) };
});
ipcMain.handle('word:choose', async () => {
  const result = await dialog.showOpenDialog(mainWindow, { title: '选择 KK个人成长系统 Word 文档', properties: ['openFile'], filters: [{ name: 'Word 文档', extensions: ['docx'] }] });
  if (result.canceled || !result.filePaths[0]) return { canceled: true };
  const file = result.filePaths[0];
  writeConfig({ documentPath: file });
  return { path: file, exists: true };
});
ipcMain.handle('word:open', async () => {
  const file = await selectedDocument();
  if (!file) throw new Error('找不到 KK个人成长系统.docx，请先选择文件');
  const error = await shell.openPath(file);
  if (error) throw new Error(error);
  return { path: file };
});
ipcMain.handle('word:save-record', async (_event, record) => {
  const file = await selectedDocument();
  const result = await saveRecord(file, record);
  if (file) writeConfig({ documentPath: file });
  return result;
});
ipcMain.handle('word:merge-records', async (_event, records) => {
  const file = await selectedDocument();
  const result = await saveRecords(file, records);
  if (file) writeConfig({ documentPath: file });
  return result;
});

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
