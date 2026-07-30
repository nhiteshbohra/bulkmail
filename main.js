const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const puppeteer = require('puppeteer-core');

let mainWindow;
let activeBrowser = null;

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

app.on('window-all-closed', async () => {
  if (activeBrowser) {
    try { await activeBrowser.close(); } catch (e) {}
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Helper: Locate installed Google Chrome or Microsoft Edge on Windows
function findChromeExecutable() {
  const possiblePaths = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    path.join(process.env.LOCALAPPDATA || '', 'Google', 'Chrome', 'Application', 'chrome.exe'),
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
  ];

  for (const p of possiblePaths) {
    if (p && fs.existsSync(p)) return p;
  }
  return null;
}

function getGmailProfilePath() {
  const profileDir = path.join(app.getPath('userData'), 'AutoMail_Gmail_Profile');
  if (!fs.existsSync(profileDir)) {
    fs.mkdirSync(profileDir, { recursive: true });
  }
  return profileDir;
}

async function getOrCreateBrowser(headless = false) {
  if (activeBrowser && activeBrowser.isConnected()) {
    return activeBrowser;
  }

  const chromePath = findChromeExecutable();
  if (!chromePath) {
    throw new Error('Google Chrome or Microsoft Edge was not found on your system. Please install Google Chrome.');
  }

  const userDataDir = getGmailProfilePath();

  activeBrowser = await puppeteer.launch({
    headless: headless,
    executablePath: chromePath,
    userDataDir: userDataDir,
    defaultViewport: null,
    args: [
      '--start-maximized',
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-blink-features=AutomationControlled'
    ]
  });

  return activeBrowser;
}

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
      text: mailData.body,
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

// IPC Handler: Launch Google Chrome Window for Sign-In
ipcMain.handle('browser:launch-google-login', async () => {
  try {
    const browser = await getOrCreateBrowser(false);
    const pages = await browser.pages();
    const page = pages[0] || await browser.newPage();

    await page.goto('https://mail.google.com/', { waitUntil: 'networkidle2', timeout: 60000 });
    return { 
      success: true, 
      message: 'Chrome launched with saved Google Profile. Please sign in to Gmail in the Chrome window!' 
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC Handler: Check Google Sign-In Status
ipcMain.handle('browser:check-google-status', async () => {
  try {
    const browser = await getOrCreateBrowser(false);
    const pages = await browser.pages();
    const page = pages[0] || await browser.newPage();

    const currentUrl = page.url();
    if (currentUrl.includes('mail.google.com') && !currentUrl.includes('signin') && !currentUrl.includes('ServiceLogin')) {
      return { signedIn: true, url: currentUrl };
    }

    await page.goto('https://mail.google.com/mail/u/0/#inbox', { waitUntil: 'domcontentloaded', timeout: 15000 });
    const finalUrl = page.url();
    const signedIn = finalUrl.includes('mail.google.com') && !finalUrl.includes('signin') && !finalUrl.includes('ServiceLogin');
    
    return { signedIn: signedIn, url: finalUrl };
  } catch (error) {
    return { signedIn: false, error: error.message };
  }
});

// IPC Handler: Send Email via Gmail Web Interface (Puppeteer Automation)
ipcMain.handle('browser:send-email-via-web', async (event, { mailData }) => {
  try {
    const browser = await getOrCreateBrowser(false);
    const pages = await browser.pages();
    const page = pages[0] || await browser.newPage();

    // Navigate to Gmail Compose Direct URL
    await page.goto('https://mail.google.com/mail/u/0/#inbox?compose=new', { 
      waitUntil: 'networkidle2', 
      timeout: 30000 
    });

    // 1. Wait for "To" Recipient Selector
    const toSelectors = [
      'textarea[name="to"]',
      'input[peoplekit-id]',
      'input[aria-label="To recipients"]',
      'div[aria-label="To recipients"] input',
      'input[aria-label="To"]',
      'div[aria-label="To"]'
    ];

    let toElement = null;
    for (const sel of toSelectors) {
      try {
        toElement = await page.waitForSelector(sel, { timeout: 4000, visible: true });
        if (toElement) break;
      } catch (e) {}
    }

    if (!toElement) {
      // Fallback: try pressing 'c' key to open compose dialog
      await page.keyboard.press('c');
      await new Promise(r => setTimeout(r, 1000));
      for (const sel of toSelectors) {
        try {
          toElement = await page.waitForSelector(sel, { timeout: 3000, visible: true });
          if (toElement) break;
        } catch (e) {}
      }
    }

    if (!toElement) {
      throw new Error('Could not locate Gmail Recipient (To) input box. Please make sure you are signed in to Gmail in Chrome.');
    }

    // Type recipient email
    await toElement.focus();
    await page.keyboard.type(mailData.to);
    await page.keyboard.press('Enter');
    await new Promise(r => setTimeout(r, 400));

    // 2. Type Subject
    const subjectSelector = 'input[name="subjectbox"], input[aria-label="Subject"]';
    const subjectElement = await page.waitForSelector(subjectSelector, { timeout: 8000, visible: true });
    if (subjectElement) {
      await subjectElement.focus();
      await page.keyboard.type(mailData.subject);
    }

    // 3. Type Body
    const bodySelector = 'div[aria-label="Message Body"], div[role="textbox"][aria-label*="Body"], div[aria-label*="Body"]';
    const bodyElement = await page.waitForSelector(bodySelector, { timeout: 8000, visible: true });
    if (bodyElement) {
      await bodyElement.focus();
      
      // Type body with newlines
      const lines = (mailData.body || '').split('\n');
      for (let i = 0; i < lines.length; i++) {
        await page.keyboard.type(lines[i]);
        if (i < lines.length - 1) {
          await page.keyboard.press('Enter');
        }
      }
    }

    await new Promise(r => setTimeout(r, 500));

    // 4. Trigger Send via Ctrl + Enter
    await page.keyboard.down('Control');
    await page.keyboard.press('Enter');
    await page.keyboard.up('Control');

    // Wait 2 seconds for dispatch confirmation
    await new Promise(r => setTimeout(r, 2000));

    return { success: true, message: 'Sent via Gmail Web!' };
  } catch (error) {
    return { success: false, error: error.message || 'Gmail Web Automation failed' };
  }
});

// IPC Handler: Close Browser
ipcMain.handle('browser:close-browser', async () => {
  try {
    if (activeBrowser) {
      await activeBrowser.close();
      activeBrowser = null;
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
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
