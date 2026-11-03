const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('cellDesk', {
  platform: process.platform,
  openDialog: (filterKey) => ipcRenderer.invoke('docdeck:dialog:open', filterKey),
  saveDialog: (opts) => ipcRenderer.invoke('docdeck:dialog:save', opts),
  unsavedDialog: () => ipcRenderer.invoke('docdeck:dialog:unsaved'),
  readFile: (p) => ipcRenderer.invoke('docdeck:read-file', p),
  writeFile: (p, data) => ipcRenderer.invoke('docdeck:write-file', p, data),
  setTitle: (name, dirty) => ipcRenderer.send('docdeck:set-title', name, dirty),
  confirmClose: () => ipcRenderer.invoke('docdeck:confirm-close'),
  leaveDialog: () => ipcRenderer.invoke('docdeck:dialog:leave'),
  recents: {
    list: () => ipcRenderer.invoke('docdeck:recents:list'),
    add: (p) => ipcRenderer.invoke('docdeck:recents:add', p),
  },
  onRecentsChanged: (cb) => {
    const handler = () => cb();
    ipcRenderer.on('docdeck:recents-changed', handler);
    return () => ipcRenderer.removeListener('docdeck:recents-changed', handler);
  },
  onMenu: (cb) => {
    const handler = (_e, action) => cb(action);
    ipcRenderer.on('docdeck:menu', handler);
    return () => ipcRenderer.removeListener('docdeck:menu', handler);
  },
  onOpenPath: (cb) => {
    const handler = (_e, p) => cb(p);
    ipcRenderer.on('docdeck:open-path', handler);
    return () => ipcRenderer.removeListener('docdeck:open-path', handler);
  },
  queryDirty: (fn) => {
    ipcRenderer.on('docdeck:query-dirty', (e) => {
      e.returnValue = fn();
    });
  },
});
