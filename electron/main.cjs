const { app, BrowserWindow, shell } = require('electron');
const path = require('path');

const isDev = !app.isPackaged;
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 800,
    minWidth: 900,
    minHeight: 620,
    title: 'KK 成长记录',
    backgroundColor: '#eef2f6',
    webPreferences: { contextIsolation: true, sandbox: true }
  });
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  if (isDev) mainWindow.loadURL('http://localhost:5173');
  else mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
