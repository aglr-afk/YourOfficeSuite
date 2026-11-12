const { contextBridge, ipcRenderer, webUtils } = require('electron');

const CHANNELS = ['app:init', 'fs:list', 'fs:read', 'fs:write', 'fs:stat', 'fs:conflictCopyPath', 'fs:reveal', 'dlg:open',
  'dlg:saveAs', 'dlg:addFolder', 'dlg:message', 'recent:add', 'recent:remove', 'path:join', 'app:dirty', 'app:closeNow', 'app:active',
  'app:docState', 'dlg:pickFolder', 'fs:uniquePath', 'app:licenses', 'settings:get', 'settings:set', 'settings:pickUpdateDir',
  'session:save', 'fs:search', 'fs:rename', 'fs:duplicate', 'update:check', 'update:install', 'update:openPage'];

const isShell = window.top === window && location.pathname.startsWith('/__shell/');

if (isShell) {
  // Only the shell window gets file access. The main process checks the sender as well.
  contextBridge.exposeInMainWorld('office', {
    invoke: (channel, ...args) => {
      if (!CHANNELS.includes(channel)) return Promise.reject(new Error('Unknown channel'));
      return ipcRenderer.invoke(channel, ...args);
    },
    onCommand: (cb) => ipcRenderer.on('cmd', (_e, name) => cb(name)),
    onOpenFiles: (cb) => ipcRenderer.on('open-files', (_e, files) => cb(files)),
    onUpdateProgress: (cb) => ipcRenderer.on('update-progress', (_e, p) => cb(p)),
    pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch (_) { return ''; } },
  });
}
