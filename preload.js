const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('dialog:select-file'),
  selectAttachments: () => ipcRenderer.invoke('dialog:select-attachments'),
  parseFile: (filePath) => ipcRenderer.invoke('excel:parse-file', filePath),
  testSmtp: (config) => ipcRenderer.invoke('smtp:test-connection', config),
  sendEmail: (smtpConfig, mailData) => ipcRenderer.invoke('smtp:send-email', { smtpConfig, mailData }),
  exportReport: (reportData) => ipcRenderer.invoke('excel:export-report', reportData)
});

