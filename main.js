const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

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
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));

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

// IPC Handler: Send Single Email
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
      text: mailData.body,
      // If body contains HTML tags, convert to html key
      html: mailData.isHtml ? mailData.body : mailData.body.replace(/\n/g, '<br>')
    };

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

// IPC Handler: Open External URLs in default OS Browser
ipcMain.handle('shell:open-external', async (event, url) => {
  if (url && (url.startsWith('http://') || url.startsWith('https://'))) {
    await shell.openExternal(url);
    return { success: true };
  }
  return { success: false, error: 'Invalid URL' };
});
