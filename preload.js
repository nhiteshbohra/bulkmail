const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('dialog:select-file'),
  selectAttachments: () => ipcRenderer.invoke('dialog:select-attachments'),
  parseFile: (filePath) => ipcRenderer.invoke('excel:parse-file', filePath),
  testSmtp: (config) => ipcRenderer.invoke('smtp:test-connection', config),
  sendEmail: (smtpConfig, mailData) => ipcRenderer.invoke('smtp:send-email', { smtpConfig, mailData }),
  exportReport: (reportData) => ipcRenderer.invoke('excel:export-report', reportData),
  // Schedule: data persistence
  saveSchedule:  (data) => ipcRenderer.invoke('schedule:save', data),
  loadSchedule:  ()     => ipcRenderer.invoke('schedule:load'),
  clearSchedule: ()     => ipcRenderer.invoke('schedule:clear'),
  // Schedule: Windows Task Scheduler (true background scheduling)
  createOsTask:  (opts) => ipcRenderer.invoke('schedule:create-os-task', opts),
  deleteOsTask:  ()     => ipcRenderer.invoke('schedule:delete-os-task'),
  checkOsTask:   ()     => ipcRenderer.invoke('schedule:check-os-task'),
  getLatestLog:  ()     => ipcRenderer.invoke('schedule:get-latest-log'),
  // Email verification & active domain checks
  validateEmail:      (email, options) => ipcRenderer.invoke('email:validate', email, options),
  validateEmailBatch: (emails, options) => ipcRenderer.invoke('email:validate-batch', emails, options),
  // Signature file synchronization
  loadSignatureFile:  () => ipcRenderer.invoke('signature:load'),
  saveSignatureFile:  (sig) => ipcRenderer.invoke('signature:save', sig),
});


