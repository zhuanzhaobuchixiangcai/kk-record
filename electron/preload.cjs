const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('kkDesktop', {
  getWordStatus: () => ipcRenderer.invoke('word:get-status'),
  chooseWord: () => ipcRenderer.invoke('word:choose'),
  openWord: () => ipcRenderer.invoke('word:open'),
  saveRecord: (record) => ipcRenderer.invoke('word:save-record', record),
  mergeRecords: (records) => ipcRenderer.invoke('word:merge-records', records)
});
