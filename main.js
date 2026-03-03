const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');

const {
  buildTimetableFromEntries,
  buildSuggestionForFile,
  ensureExt,
  findDefaultTimetablePath,
  isRegularFile,
  loadTimetableFromFile,
  sanitizeFilename,
} = require('./src/core/naming');

/** @type {Map<string, {path:string, year:number, entryCount:number, entries:Array<any>}>} */
const timetableCache = new Map();

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 820,
    minHeight: 620,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'renderer', 'index.html'));
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('get-initial-state', async () => {
  let baseTimetable = null;
  let baseTimetableSource = 'none';
  let customTimetable = null;
  let warning = '';
  const customPlanPath = getCustomPlanPath();

  const bundledPath = resolveBundledTimetablePath();
  const workspacePath = findDefaultTimetablePath(process.cwd());
  const baseCandidates = [];

  if (bundledPath) {
    baseCandidates.push({ source: 'bundled', filePath: bundledPath });
  }

  if (workspacePath && path.resolve(workspacePath) !== path.resolve(bundledPath || '')) {
    baseCandidates.push({ source: 'workspace', filePath: workspacePath });
  }

  for (const candidate of baseCandidates) {
    try {
      baseTimetable = loadTimetable(candidate.filePath);
      baseTimetableSource = candidate.source;
      break;
    } catch (error) {
      warning = appendWarning(warning, `${candidate.source} 시간표를 불러오지 못했습니다: ${error.message}`);
    }
  }

  try {
    customTimetable = loadCustomTimetable(customPlanPath);
  } catch (error) {
    warning = appendWarning(warning, `사용자 시간표를 불러오지 못했습니다: ${error.message}`);
  }

  const activeTimetable = customTimetable ?? baseTimetable ?? null;

  return {
    timetable: activeTimetable,
    customPlan: {
      path: customPlanPath,
      exists: Boolean(customTimetable),
      active: Boolean(customTimetable),
      basePath: baseTimetable?.path ?? null,
      baseSource: baseTimetableSource,
      bundledPath: bundledPath ?? null,
    },
    ...(warning ? { error: warning } : {}),
  };
});

ipcMain.handle('get-app-version', async () => {
  return app.getVersion();
});

ipcMain.handle('load-custom-plan', async () => {
  const customTimetable = loadCustomTimetable(getCustomPlanPath());
  return customTimetable;
});

ipcMain.handle('save-custom-plan', async (_, payload) => {
  const rawEntries = Array.isArray(payload?.entries) ? payload.entries : [];
  const year = Number(payload?.year) || new Date().getFullYear();
  const customPlanPath = getCustomPlanPath();
  const timetable = buildTimetableFromEntries(rawEntries, year, customPlanPath);

  if (timetable.entryCount === 0) {
    throw new Error('저장할 강의 항목이 없습니다.');
  }

  const serialized = {
    version: 1,
    year: timetable.year,
    entries: timetable.entries.map((entry) => ({
      subject: entry.subject,
      prefix: entry.prefix,
      orderNumber: entry.orderNumber,
      lectureTitle: entry.lectureTitle,
      professor: entry.professor,
      month: entry.month,
      day: entry.day,
      dateToken: entry.dateToken,
      period: entry.period,
      hours: entry.hours,
    })),
  };

  await fs.mkdir(path.dirname(customPlanPath), { recursive: true });
  await fs.writeFile(customPlanPath, `${JSON.stringify(serialized, null, 2)}\n`, 'utf8');
  timetableCache.set(path.resolve(customPlanPath), timetable);

  return {
    path: customPlanPath,
    timetable,
  };
});

ipcMain.handle('clear-custom-plan', async () => {
  const customPlanPath = getCustomPlanPath();
  try {
    await fs.unlink(customPlanPath);
  } catch (error) {
    if (error?.code !== 'ENOENT') {
      throw error;
    }
  }

  timetableCache.delete(path.resolve(customPlanPath));
  return { path: customPlanPath, cleared: true };
});

ipcMain.handle('select-timetable', async () => {
  const result = await dialog.showOpenDialog({
    title: '시간표 파일(.xlsx) 선택',
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle('load-timetable', async (_, filePath) => {
  return loadTimetable(filePath);
});

ipcMain.handle('build-suggestions', async (_, payload) => {
  const { timetablePath, filePaths } = payload || {};
  if (!timetablePath || !Array.isArray(filePaths) || filePaths.length === 0) {
    return [];
  }

  const timetable = loadTimetable(timetablePath);

  return filePaths
    .filter((filePath) => isRegularFile(filePath))
    .map((filePath) => buildSuggestionForFile(filePath, timetable.entries));
});

ipcMain.handle('select-output-dir', async () => {
  const result = await dialog.showOpenDialog({
    title: '출력 폴더 선택',
    properties: ['openDirectory', 'createDirectory'],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle('copy-renamed-files', async (_, payload) => {
  const { outputDir, tasks } = payload || {};

  if (!outputDir || !Array.isArray(tasks) || tasks.length === 0) {
    throw new Error('복사할 작업이 없습니다.');
  }

  await fs.mkdir(outputDir, { recursive: true });

  const results = [];

  for (const task of tasks) {
    const sourcePath = task?.sourcePath;
    const targetNameRaw = task?.targetName;

    if (!sourcePath || !targetNameRaw || !isRegularFile(sourcePath)) {
      results.push({
        sourcePath,
        status: 'failed',
        error: '원본 파일을 찾을 수 없거나 이름이 비어 있습니다.',
      });
      continue;
    }

    const originalExt = path.extname(sourcePath);
    const targetNameWithExt = ensureExt(targetNameRaw, originalExt);
    const safeTargetName = sanitizeFilename(targetNameWithExt);
    const finalPath = await findAvailablePath(path.join(outputDir, safeTargetName));

    try {
      await fs.copyFile(sourcePath, finalPath);
      results.push({
        sourcePath,
        status: 'copied',
        targetPath: finalPath,
      });
    } catch (error) {
      results.push({
        sourcePath,
        status: 'failed',
        error: error.message,
      });
    }
  }

  if (results.some((item) => item.status === 'copied')) {
    await shell.openPath(outputDir);
  }

  return results;
});

function getCustomPlanPath() {
  return path.join(app.getPath('userData'), 'lecture-plan.custom.json');
}

function resolveBundledTimetablePath() {
  const candidates = app.isPackaged
    ? [path.join(process.resourcesPath, 'assets', 'default-timetable.xlsx')]
    : [
        path.join(__dirname, 'assets', 'default-timetable.xlsx'),
        path.join(process.cwd(), 'assets', 'default-timetable.xlsx'),
      ];

  for (const candidate of candidates) {
    if (isRegularFile(candidate)) {
      return candidate;
    }
  }

  return null;
}

function appendWarning(origin, next) {
  if (!next) {
    return origin;
  }
  return origin ? `${origin}\n${next}` : next;
}

function loadTimetable(filePath) {
  const fullPath = path.resolve(filePath);
  const cached = timetableCache.get(fullPath);
  if (cached) {
    return cached;
  }

  const timetable = loadTimetableFromFile(fullPath);
  timetableCache.set(fullPath, timetable);
  return timetable;
}

function loadCustomTimetable(filePath) {
  const fullPath = path.resolve(filePath);
  if (!isRegularFile(fullPath)) {
    return null;
  }

  const cached = timetableCache.get(fullPath);
  if (cached) {
    return cached;
  }

  const raw = fsSync.readFileSync(fullPath, 'utf8');
  const parsed = JSON.parse(raw);
  const entries = Array.isArray(parsed?.entries) ? parsed.entries : [];
  const year = Number(parsed?.year) || new Date().getFullYear();
  const timetable = buildTimetableFromEntries(entries, year, fullPath);

  if (timetable.entryCount === 0) {
    throw new Error('사용자 시간표 파일에 유효한 항목이 없습니다.');
  }

  timetableCache.set(fullPath, timetable);
  return timetable;
}

async function findAvailablePath(targetPath) {
  const ext = path.extname(targetPath);
  const base = path.basename(targetPath, ext);
  const dir = path.dirname(targetPath);

  let candidate = targetPath;
  let index = 1;

  while (await pathExists(candidate)) {
    candidate = path.join(dir, `${base} (${index})${ext}`);
    index += 1;
  }

  return candidate;
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}
