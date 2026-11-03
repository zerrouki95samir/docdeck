const { app, BrowserWindow, Menu, dialog, ipcMain, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');

app.setName('DocDeck');

nativeTheme.themeSource = 'light';

if (process.env.DOCDECK_DEBUG_PORT) {
  app.commandLine.appendSwitch('remote-debugging-port', process.env.DOCDECK_DEBUG_PORT);
}

let win = null;
let allowClose = false;
let pendingOpenPath = null;

const OPENABLE = /\.(xlsx|xlsm|xls|csv|tsv|uni|pdf|docx)$/i;

/* ---------- recent files (main owns the list; menu + home read it) ---------- */

function recentsFile() {
  return path.join(app.getPath('userData'), 'recent-files.json');
}

/* one-time migration: carry recents over from the previous "CellDesk" install */
function migrateUserData() {
  try {
    const userData = app.getPath('userData');
    const oldFile = path.join(path.dirname(userData), 'CellDesk', 'recent-files.json');
    if (!fs.existsSync(recentsFile()) && fs.existsSync(oldFile)) {
      fs.mkdirSync(userData, { recursive: true });
      fs.copyFileSync(oldFile, recentsFile());
    }
  } catch {}
}

function loadRecents() {
  try {
    const raw = fs.readFileSync(recentsFile(), 'utf8');
    const list = JSON.parse(raw);
    return Array.isArray(list)
      ? list.filter((p) => typeof p === 'string' && OPENABLE.test(p))
      : [];
  } catch {
    return [];
  }
}

function saveRecents(list) {
  try {
    fs.writeFileSync(recentsFile(), JSON.stringify(list, null, 2), 'utf8');
  } catch {
    /* best effort */
  }
}

function addRecent(p) {
  if (!p || !OPENABLE.test(p)) return;
  const list = loadRecents().filter((x) => x !== p);
  list.unshift(p);
  saveRecents(list.slice(0, 10));
  buildMenu();
}

function clearRecents() {
  saveRecents([]);
  buildMenu();
}

function findOpenablePath(argv = []) {
  const candidates = argv.filter((a) => a && OPENABLE.test(a) && !a.startsWith('-'));
  return candidates.length ? candidates[candidates.length - 1] : null;
}

function forwardOrQueueOpenPath(p) {
  pendingOpenPath = p;
  if (win && win.webContents && !win.webContents.isLoading() && win.webContents.getURL()) {
    win.webContents.send('docdeck:open-path', p);
    pendingOpenPath = null;
  }
}

function sendMenu(action) {
  if (win) win.webContents.send('docdeck:menu', action);
}

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const recents = loadRecents();
  const recentItems = recents.length
    ? [
        ...recents.map((p, i) => ({
          label: `${i < 9 ? `${i + 1}.  ` : ''}${path.basename(p)}`,
          click: () => forwardOrQueueOpenPath(p),
        })),
        { type: 'separator' },
        {
          label: 'Clear Recent Files',
          click: () => {
            clearRecents();
            if (win) win.webContents.send('docdeck:recents-changed');
          },
        },
      ]
    : [{ label: 'No Recent Files', enabled: false }];

  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        { label: 'New Spreadsheet', accelerator: 'CmdOrCtrl+N', click: () => sendMenu('new-sheet') },
        { label: 'New PDF', click: () => sendMenu('new-pdf') },
        { label: 'New Word Document', click: () => sendMenu('new-docx') },
        { type: 'separator' },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => sendMenu('open') },
        { label: 'Open Recent', submenu: recentItems },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => sendMenu('save') },
        { label: 'Save As…', accelerator: 'Shift+CmdOrCtrl+S', click: () => sendMenu('save-as') },
        { type: 'separator' },
        { label: 'Export Current Sheet as CSV…', click: () => sendMenu('export-csv') },
        { label: 'Save as DocDeck Snapshot (.uni)…', click: () => sendMenu('save-uni') },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => sendMenu('print') },
        { type: 'separator' },
        { label: 'Workspace Home', accelerator: 'CmdOrCtrl+Shift+H', click: () => sendMenu('go-home') },
        ...(isMac ? [] : [{ type: 'separator' }, { role: 'quit' }]),
      ],
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  const iconPath = path.join(__dirname, '..', 'build', 'icon.png');
  win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 940,
    minHeight: 620,
    title: 'DocDeck',
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    backgroundColor: '#f6f7f9',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });

  if (process.env.VITE_DEV) {
    win.loadURL('http://localhost:5173');
  } else {
    win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  win.once('ready-to-show', () => win.show());

  win.webContents.on('did-finish-load', () => {
    if (pendingOpenPath) {
      win.webContents.send('docdeck:open-path', pendingOpenPath);
      pendingOpenPath = null;
    }
  });

  if (process.env.DOCDECK_DEBUG_PORT) {
    win.webContents.on('console-message', (event) => {
      const message = event && typeof event === 'object' && typeof event.message === 'string'
        ? event.message
        : arguments.length > 2
          ? arguments[2]
          : String(event);
      const src = event && event.sourceId
        ? ` (${String(event.sourceId).split('/').pop()}${event.lineNumber != null ? `:${event.lineNumber}` : ''})`
        : '';
      console.log(`[renderer] ${message}${src}`);
    });
  }

  win.on('close', (e) => {
    if (allowClose) return;
    e.preventDefault();
    let dirty = false;
    try {
      dirty = win.webContents.sendSync('docdeck:query-dirty') === true;
    } catch (_) {
      dirty = false;
    }
    if (!dirty) {
      allowClose = true;
      win.close();
      return;
    }
    dialog
      .showMessageBox(win, {
        type: 'warning',
        message: 'Save changes before closing?',
        detail: 'If you don\u2019t save, your changes will be lost.',
        buttons: ['Save', 'Discard Changes', 'Cancel'],
        defaultId: 0,
        cancelId: 2,
      })
      .then(({ response }) => {
        if (response === 1) {
          allowClose = true;
          win.close();
        } else if (response === 0) {
          sendMenu('save-then-exit');
        }
      });
  });

  win.on('closed', () => {
    win = null;
  });
}

// ---------- IPC ----------
ipcMain.handle('docdeck:dialog:open', async (_e, filterKey) => {
  const ALL = [
    { name: 'All Documents', extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'uni', 'pdf', 'docx'] },
    { name: 'PDF', extensions: ['pdf'] },
    { name: 'Word Documents', extensions: ['docx'] },
    { name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'uni'] },
    { name: 'All Files', extensions: ['*'] },
  ];
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    properties: ['openFile', 'multiSelections'],
    filters: filterKey === 'pdf' ? [{ name: 'PDF', extensions: ['pdf'] }] : filterKey === 'image' ? [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'] }] : ALL,
  });
  return canceled ? [] : filePaths;
});

ipcMain.handle('docdeck:dialog:save', async (_e, opts) => {
  const { canceled, filePath } = await dialog.showSaveDialog(win, {
    defaultPath: opts?.defaultPath || 'Workbook.xlsx',
    filters: opts?.filters || [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
  });
  return canceled || !filePath ? null : filePath;
});

ipcMain.handle('docdeck:dialog:unsaved', async () => {
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    message: 'Save changes?',
    detail: 'The current workbook has unsaved changes.',
    buttons: ['Save', 'Discard', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  });
  return response;
});

ipcMain.handle('docdeck:read-file', (_e, p) => fs.readFileSync(p).buffer);

/* ask about unsaved work when a document is about to be replaced or closed in place.
   0 = Save, 1 = Discard, 2 = Cancel */
ipcMain.handle('docdeck:dialog:leave', async () => {
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    message: 'Save changes?',
    detail: 'The document you are replacing has unsaved changes.',
    buttons: ['Save', 'Discard Changes', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  });
  return response;
});

ipcMain.handle('docdeck:recents:list', () => loadRecents());

ipcMain.handle('docdeck:recents:add', (_e, p) => {
  addRecent(p);
  return true;
});

ipcMain.on('docdeck:home', () => {
  if (win) win.webContents.send('docdeck:menu', 'go-home');
});

ipcMain.handle('docdeck:write-file', (_e, p, data) =>
  fs.promises.writeFile(p, Buffer.from(data))
);

ipcMain.handle('docdeck:confirm-close', () => {
  allowClose = true;
  if (win) win.close();
});

ipcMain.on('docdeck:set-title', (_e, name, dirty) => {
  if (win) win.setTitle(`${dirty ? '\u2022 ' : ''}${name} \u2014 DocDeck`);
});

// ---------- lifecycle ----------
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    if (win) {
      win.focus();
      const p = findOpenablePath(argv);
      if (p) forwardOrQueueOpenPath(p);
    }
  });

  app.on('open-file', (e, p) => {
    if (!OPENABLE.test(p)) return;
    e.preventDefault();
    forwardOrQueueOpenPath(p);
  });

  app.whenReady().then(() => {
    migrateUserData();
    buildMenu();
    createWindow();
    try {
      const dockIcon = path.join(__dirname, '..', 'build', 'icon.png');
      if (process.platform === 'darwin' && app.dock && fs.existsSync(dockIcon)) {
        app.dock.setIcon(dockIcon);
      }
    } catch {}
    const fromArgv = findOpenablePath([
      process.argv[0]?.endsWith('docdeck') || process.argv[0]?.endsWith('celldesk') ? '' : process.argv[1] || '',
      ...process.argv.slice(2),
    ]);
    if (fromArgv) forwardOrQueueOpenPath(fromArgv);

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  });

  app.on('window-all-closed', () => {
    app.quit();
  });
}
