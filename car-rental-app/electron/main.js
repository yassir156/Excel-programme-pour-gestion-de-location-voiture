const { app, BrowserWindow, shell } = require('electron');
const path = require('path');
const fs = require('fs');

const isDev = process.env.NODE_ENV === 'development';
const BACKEND_PORT = 4310;

function setupDataDirectories() {
  const userDataPath = app.getPath('userData');
  const dbDir = path.join(userDataPath, 'database');
  const uploadsDir = path.join(userDataPath, 'uploads');
  if (!fs.existsSync(dbDir)) fs.mkdirSync(dbDir, { recursive: true });
  if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
  process.env.CAR_RENTAL_DB_DIR = dbDir;
  process.env.CAR_RENTAL_UPLOADS_DIR = uploadsDir;
  process.env.CAR_RENTAL_PORT = String(BACKEND_PORT);
}

async function seedIfEmpty() {
  const { User } = require('../backend/src/models');
  const count = await User.count();
  if (count === 0) {
    const { seedDemoData } = require('../backend/seed/seedLogic');
    await seedDemoData();
  }
}

let mainWindow;

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    backgroundColor: '#0f172a',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.once('ready-to-show', () => mainWindow.show());

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  if (isDev) {
    await mainWindow.loadURL('http://127.0.0.1:5173');
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    await mainWindow.loadFile(path.join(__dirname, '..', 'frontend', 'dist', 'index.html'));
  }
}

app.whenReady().then(async () => {
  setupDataDirectories();
  const { startServer } = require('../backend/src/server');
  await startServer(BACKEND_PORT);
  await seedIfEmpty();
  await createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
