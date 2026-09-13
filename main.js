const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const { validateEmail, validateBatch } = require('./emailValidator');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 650,
    title: 'Excel Bulk Mailer',
    backgroundColor: '#0b0f19',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // Remove default menu for a clean app feel
  mainWindow.setMenuBarVisibility(false);
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

// IPC Handler: Open File Dialog for Excel / CSV selection
ipcMain.handle('dialog:select-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Excel or CSV File',
    properties: ['openFile'],
    filters: [
      { name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

// IPC Handler: Open File Dialog for Attachments / Images with Multiple Selection Support
ipcMain.handle('dialog:select-attachments', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Attachments (Images, Documents, Files)',
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Images & Documents', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'svg', 'bmp', 'pdf', 'docx', 'xlsx', 'csv', 'txt', 'zip'] },
      { name: 'Images (*.png, *.jpg, *.jpeg, *.gif, *.webp, *.svg)', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg'] },
      { name: 'Documents (*.pdf, *.docx, *.txt, *.xlsx)', extensions: ['pdf', 'docx', 'txt', 'xlsx', 'csv', 'zip'] },
      { name: 'All Files (*.*)', extensions: ['*'] }
    ]
  });

  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return [];
  }
  return result.filePaths;
});

// IPC Handler: Parse Excel/CSV File
ipcMain.handle('excel:parse-file', async (event, filePath) => {
  try {
    if (!fs.existsSync(filePath)) {
      throw new Error('Selected file does not exist.');
    }

    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Read headers and objects
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    if (!rawRows || rawRows.length === 0) {
      throw new Error('Spreadsheet is empty.');
    }

    const headers = rawRows[0].map(h => String(h).trim()).filter(h => h.length > 0);
    const dataObjects = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    return {
      success: true,
      fileName: path.basename(filePath),
      filePath: filePath,
      headers: headers,
      rows: dataObjects,
      totalRows: dataObjects.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// IPC Handler: Test SMTP Connection
ipcMain.handle('smtp:test-connection', async (event, config) => {
  try {
    const portNum = parseInt(config.port, 10) || 587;
    const transporter = nodemailer.createTransport({
      host: config.host,
      port: portNum,
      secure: config.secure === true || portNum === 465,
      auth: {
        user: config.user,
        pass: config.pass
      },
      connectionTimeout: 10000,
      greetingTimeout: 10000,
      socketTimeout: 10000
    });

    await transporter.verify();
    return { success: true, message: 'SMTP connection established successfully!' };
  } catch (error) {
    return { success: false, error: error.message || 'SMTP Connection failed.' };
  }
});

// IPC Handler: Send Single Email via SMTP
ipcMain.handle('smtp:send-email', async (event, { smtpConfig, mailData }) => {
  try {
    const portNum = parseInt(smtpConfig.port, 10) || 587;
    const transporter = nodemailer.createTransport({
      host: smtpConfig.host,
      port: portNum,
      secure: smtpConfig.secure === true || portNum === 465,
      auth: {
        user: smtpConfig.user,
        pass: smtpConfig.pass
      },
      connectionTimeout: 10000,
      greetingTimeout: 10000,
      socketTimeout: 15000
    });

    const mailOptions = {
      from: smtpConfig.fromName
        ? `"${smtpConfig.fromName}" <${smtpConfig.user}>`
        : smtpConfig.user,
      to: mailData.to,
      subject: mailData.subject,
      text: mailData.text || mailData.body,
      html: mailData.html || (mailData.isHtml ? mailData.body : (mailData.body || '').replace(/\n/g, '<br>'))
    };

    if (mailData.cc && String(mailData.cc).trim()) {
      mailOptions.cc = String(mailData.cc).trim();
    }

    if (mailData.bcc && String(mailData.bcc).trim()) {
      mailOptions.bcc = String(mailData.bcc).trim();
    }

    if (mailData.attachmentPath && String(mailData.attachmentPath).trim()) {
      const rawPaths = String(mailData.attachmentPath)
        .split(/[,;]+/)
        .map(p => p.trim().replace(/^["']|["']$/g, ''))
        .filter(Boolean);
      const attachments = [];
      for (const attachPath of rawPaths) {
        if (!fs.existsSync(attachPath)) {
          throw new Error(`Attachment file not found: ${attachPath}`);
        }
        const cleanFileName = path.basename(attachPath.replace(/\\/g, '/'));
        attachments.push({
          filename: cleanFileName,
          path: attachPath
        });
      }
      if (attachments.length > 0) {
        mailOptions.attachments = attachments;
      }
    }

    const info = await transporter.sendMail(mailOptions);
    return {
      success: true,
      messageId: info.messageId,
      response: info.response
    };
  } catch (error) {
    return {
      success: false,
      error: error.message || 'Failed to send email'
    };
  }
});



// IPC Handler: Export Execution Report to Excel File
ipcMain.handle('excel:export-report', async (event, reportData) => {
  try {
    const saveResult = await dialog.showSaveDialog(mainWindow, {
      title: 'Export Campaign Report',
      defaultPath: `Email_Campaign_Report_${Date.now()}.xlsx`,
      filters: [
        { name: 'Excel Workbook', extensions: ['xlsx'] },
        { name: 'CSV File', extensions: ['csv'] }
      ]
    });

    if (saveResult.canceled || !saveResult.filePath) {
      return { success: false, canceled: true };
    }

    const worksheet = XLSX.utils.json_to_sheet(reportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Campaign Report');

    XLSX.writeFile(workbook, saveResult.filePath);
    return { success: true, filePath: saveResult.filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// ===========================================================================
// IPC Handlers: Schedule Data (saved in project root for sender.js access)
// ===========================================================================
const SCHEDULE_FILE = path.join(__dirname, 'schedule_data.json');
const SENDER_SCRIPT = path.join(__dirname, 'sender.js');
const BAT_FILE = path.join(__dirname, 'run_schedule.bat');
const TASK_NAME = 'AutoMailExcelSchedule';

// Save schedule data to project root (where sender.js can find it)
ipcMain.handle('schedule:save', async (event, scheduleData) => {
  try {
    fs.writeFileSync(SCHEDULE_FILE, JSON.stringify(scheduleData, null, 2), 'utf8');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('schedule:load', async () => {
  try {
    if (!fs.existsSync(SCHEDULE_FILE)) return { success: true, data: null };
    const raw = fs.readFileSync(SCHEDULE_FILE, 'utf8');
    return { success: true, data: JSON.parse(raw) };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('schedule:clear', async () => {
  try {
    if (fs.existsSync(SCHEDULE_FILE)) fs.unlinkSync(SCHEDULE_FILE);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ===========================================================================
// IPC Handler: Register a Windows Task Scheduler job
// Finds node.exe, writes a .bat launcher, creates a schtasks entry
// so emails send even if the app is completely closed.
// ===========================================================================
ipcMain.handle('schedule:create-os-task', async (event, { isoString }) => {
  try {
    // Find the node.exe path
    const nodePath = await new Promise((resolve, reject) => {
      exec('where node', (err, stdout) => {
        if (err || !stdout.trim()) return reject(new Error('Node.js not found in PATH. Please ensure Node.js is installed.'));
        resolve(stdout.trim().split('\r\n')[0].split('\n')[0].trim());
      });
    });

    // Write a .bat launcher that sets the directory and invokes sender.js
    const batContent = `@echo off\r\ncd /d "${__dirname}"\r\n"${nodePath}" "${SENDER_SCRIPT}"\r\n`;
    fs.writeFileSync(BAT_FILE, batContent, 'utf8');

    // Parse date/time for schtasks
    const d = new Date(isoString);
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yyyy = d.getFullYear();
    const hh = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');

    // Try dd/mm/yyyy first (standard on this system), fallback to mm/dd/yyyy if needed
    let cmd = `schtasks /create /tn "${TASK_NAME}" /tr "\\"${BAT_FILE}\\"" /sc once /sd ${dd}/${mm}/${yyyy} /st ${hh}:${min} /f`;

    try {
      await new Promise((resolve, reject) => {
        exec(cmd, (err, stdout, stderr) => {
          if (err) return reject(new Error(stderr.trim() || err.message));
          resolve(stdout);
        });
      });
    } catch (firstErr) {
      cmd = `schtasks /create /tn "${TASK_NAME}" /tr "\\"${BAT_FILE}\\"" /sc once /sd ${mm}/${dd}/${yyyy} /st ${hh}:${min} /f`;
      await new Promise((resolve, reject) => {
        exec(cmd, (err, stdout, stderr) => {
          if (err) return reject(new Error(stderr.trim() || err.message || firstErr.message));
          resolve(stdout);
        });
      });
    }

    return { success: true, nodePath, taskName: TASK_NAME };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// Delete the Windows Task Scheduler job and clean up files
ipcMain.handle('schedule:delete-os-task', async () => {
  const errors = [];
  await new Promise(resolve => {
    exec(`schtasks /delete /tn "${TASK_NAME}" /f`, (err, stdout, stderr) => {
      if (err && !stderr.includes('cannot find')) errors.push(stderr.trim());
      resolve();
    });
  });
  try { if (fs.existsSync(SCHEDULE_FILE)) fs.unlinkSync(SCHEDULE_FILE); } catch (e) { }
  try { if (fs.existsSync(BAT_FILE)) fs.unlinkSync(BAT_FILE); } catch (e) { }
  return errors.length === 0
    ? { success: true }
    : { success: false, error: errors.join('; ') };
});

// Check if a Windows task is currently registered
ipcMain.handle('schedule:check-os-task', async () => {
  try {
    const output = await new Promise((resolve) => {
      exec(`schtasks /query /tn "${TASK_NAME}" /fo LIST`, (err, stdout) => {
        resolve(err ? '' : stdout);
      });
    });
    return { exists: output.includes(TASK_NAME), raw: output };
  } catch (e) {
    return { exists: false };
  }
});

// Get latest background campaign execution log (if sent while app was closed)
ipcMain.handle('schedule:get-latest-log', async () => {
  try {
    const files = fs.readdirSync(__dirname)
      .filter(f => f.startsWith('campaign_log_') && f.endsWith('.json'))
      .map(f => ({ name: f, time: fs.statSync(path.join(__dirname, f)).mtimeMs }))
      .sort((a, b) => b.time - a.time);

    if (files.length === 0) return { success: true, log: null };

    const latestPath = path.join(__dirname, files[0].name);
    const content = JSON.parse(fs.readFileSync(latestPath, 'utf8'));
    return {
      success: true,
      log: {
        fileName: files[0].name,
        timestamp: files[0].time,
        data: content
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ===========================================================================
// IPC Handlers: Email Verification (Syntax & Active Domain MX Validation)
// ===========================================================================
ipcMain.handle('email:validate', async (event, email, options) => {
  try {
    return await validateEmail(email, options);
  } catch (err) {
    return {
      isValid: false,
      isActive: false,
      reason: err.message || 'Validation error',
      details: { error: err.message }
    };
  }
});

ipcMain.handle('email:validate-batch', async (event, emails, options) => {
  try {
    const results = await validateBatch(emails, options);
    const validCount = results.filter(r => r.isValid && r.isActive).length;
    const invalidCount = results.length - validCount;
    return {
      success: true,
      total: results.length,
      validCount,
      invalidCount,
      results
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ===========================================================================
// IPC Handlers: Signature File Persistence
// ===========================================================================
ipcMain.handle('signature:load', async () => {
  try {
    const sigFile = path.join(__dirname, 'signature.txt');
    if (fs.existsSync(sigFile)) {
      return { success: true, signature: fs.readFileSync(sigFile, 'utf8') };
    }
    return { success: true, signature: null };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('signature:save', async (event, signatureText) => {
  try {
    const sigFile = path.join(__dirname, 'signature.txt');
    fs.writeFileSync(sigFile, signatureText || '', 'utf8');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});


