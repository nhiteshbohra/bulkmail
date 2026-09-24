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
  // Subject/Body template file (BODY.txt) synchronization
  loadBodyTemplateFile: () => ipcRenderer.invoke('bodyTemplate:load'),
  // Contact Registry: global dedupe + send/fail tracking across all campaigns
  loadRegistry:      () => ipcRenderer.invoke('registry:load'),
  recordSent:        (email, meta) => ipcRenderer.invoke('registry:record-sent', { email, meta }),
  recordFailed:      (email, error, meta) => ipcRenderer.invoke('registry:record-failed', { email, error, meta }),
  backfillRegistry:  () => ipcRenderer.invoke('registry:backfill'),
  checkBounces:      (smtpConfig, knownRecipients) => ipcRenderer.invoke('bounces:check', { smtpConfig, knownRecipients }),
  // Templates & Campaign History (Phase 1 & 2)
  sendSingleEmail:    (payload) => ipcRenderer.invoke('smtp:send-email', payload),
  loadTemplates:      () => ipcRenderer.invoke('templates:load'),
  saveTemplate:       (template) => ipcRenderer.invoke('templates:save', template),
  deleteTemplate:     (id) => ipcRenderer.invoke('templates:delete', id),
  listHistoryLogs:    () => ipcRenderer.invoke('campaigns:list-history'),
  getHistoryDetails:  (fileName) => ipcRenderer.invoke('campaigns:get-details', fileName),
});



