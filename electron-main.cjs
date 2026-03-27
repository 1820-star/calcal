const path = require('path');
const fs = require('fs/promises');
const os = require('os');
const { app, BrowserWindow, shell, ipcMain } = require('electron');

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1180,
    minHeight: 760,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.loadFile(path.join(__dirname, 'index.html'));

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

ipcMain.handle('calcal:save-monthly-overview', async (_event, htmlContent) => {
  const targetDir = process.platform === 'win32'
    ? path.join('C:', 'CALCAL Portable')
    : path.join(os.homedir(), 'CALCAL Portable');

  const targetPath = path.join(targetDir, 'Monatsuebersicht-Aktuell.html');
  await fs.mkdir(targetDir, { recursive: true });
  await fs.writeFile(targetPath, htmlContent, 'utf8');
  return { ok: true, filePath: targetPath };
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
