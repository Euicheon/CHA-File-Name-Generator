const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('titleMaker', {
  getInitialState: () => ipcRenderer.invoke('get-initial-state'),
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  selectTimetable: () => ipcRenderer.invoke('select-timetable'),
  loadTimetable: (filePath) => ipcRenderer.invoke('load-timetable', filePath),
  loadCustomPlan: () => ipcRenderer.invoke('load-custom-plan'),
  saveCustomPlan: (payload) => ipcRenderer.invoke('save-custom-plan', payload),
  clearCustomPlan: () => ipcRenderer.invoke('clear-custom-plan'),
  buildSuggestions: (payload) => ipcRenderer.invoke('build-suggestions', payload),
  selectOutputDir: () => ipcRenderer.invoke('select-output-dir'),
  copyRenamedFiles: (payload) => ipcRenderer.invoke('copy-renamed-files', payload),
  getPathForFile: (file) => {
    try {
      if (file && typeof file.path === 'string' && file.path.length > 0) {
        return file.path;
      }
      return webUtils.getPathForFile(file);
    } catch {
      return '';
    }
  },
});
