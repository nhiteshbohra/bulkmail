const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('dialog:select-file'),
  parseFile: (filePath) => ipcRenderer.invoke('excel:parse-file', filePath),
  testSmtp: (config) => ipcRenderer.invoke('smtp:test-connection', config),
  sendEmail: (smtpConfig, mailData) => ipcRenderer.invoke('smtp:send-email', { smtpConfig, mailData }),
  exportReport: (reportData) => ipcRenderer.invoke('excel:export-report', reportData),
  
  // Gmail Web Automation IPC Bridge
  launchGoogleLogin: () => ipcRenderer.invoke('browser:launch-google-login'),
  checkGoogleAuth: () => ipcRenderer.invoke('browser:check-google-status'),
  sendEmailWeb: (mailData) => ipcRenderer.invoke('browser:send-email-via-web', { mailData }),
  closeGoogleBrowser: () => ipcRenderer.invoke('browser:close-browser')
});
