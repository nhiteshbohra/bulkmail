/**
 * sender.js - Standalone Bulk Email Sender with Multi-Day Daily Batch Support
 * Called by Windows Task Scheduler. Does NOT require Electron to be running.
 */
'use strict';
const path = require('path');
const fs = require('fs');
const nodemailer = require('nodemailer');
const { exec } = require('child_process');
const { validateEmail, formatEmailContent } = require('./emailValidator');

const DATA_FILE = path.join(__dirname, 'schedule_data.json');
const TASK_NAME = 'AutoMailExcelSchedule';
const LOCK_FILE = path.join(__dirname, '.sender.lock');
const SIG_FILE = path.join(__dirname, 'signature.txt');

function getActiveSignature(data) {
  if (data && data.signature && data.signature.trim()) {
    return data.signature.trim();
  }
  if (fs.existsSync(SIG_FILE)) {
    try {
      const content = fs.readFileSync(SIG_FILE, 'utf8').trim();
      if (content) return content;
    } catch (e) { }
  }
  return '';
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function writeLog(results, batchLabel) {
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const prefix = batchLabel ? `campaign_log_${batchLabel}_` : 'campaign_log_';
  const logPath = path.join(__dirname, prefix + ts + '.json');
  try { fs.writeFileSync(logPath, JSON.stringify(results, null, 2), 'utf8'); } catch (e) { }
  console.log('[AutoMail] Log written to', logPath);
  return logPath;
}

function cleanupSchedule() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      fs.copyFileSync(DATA_FILE, DATA_FILE + '.bak');
      const completedPath = path.join(__dirname, 'schedule_data_completed.json');
      fs.copyFileSync(DATA_FILE, completedPath);
      fs.unlinkSync(DATA_FILE);
    }
  } catch (e) { }
  exec('schtasks /delete /tn "' + TASK_NAME + '" /f', () => { });
  console.log('[AutoMail] Archived schedule to schedule_data_completed.json and cleaned up Windows task.');
}

function rescheduleNextTask(isoString) {
  const d = new Date(isoString);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  const bat = path.join(__dirname, 'run_schedule.bat');

  let cmd = `schtasks /create /tn "${TASK_NAME}" /tr "\\"${bat}\\"" /sc once /sd ${dd}/${mm}/${yyyy} /st ${hh}:${min} /f`;
  return new Promise((resolve) => {
    exec(cmd, (err, stdout) => {
      if (!err) {
        console.log('[AutoMail] Next task successfully registered:', stdout.trim());
        return resolve(true);
      }
      // Fallback for systems with mm/dd/yyyy locale
      cmd = `schtasks /create /tn "${TASK_NAME}" /tr "\\"${bat}\\"" /sc once /sd ${mm}/${dd}/${yyyy} /st ${hh}:${min} /f`;
      exec(cmd, (err2, stdout2) => {
        if (!err2) {
          console.log('[AutoMail] Next task registered (fallback mm/dd/yyyy):', stdout2.trim());
          return resolve(true);
        }
        console.error('[AutoMail] Failed to reschedule next task:', (err2 && err2.message) || err.message);
        resolve(false);
      });
    });
  });
}

async function run() {
  console.log('[AutoMail] ======================================================');
  console.log('[AutoMail] Sender process started at', new Date().toLocaleString());

  if (!fs.existsSync(DATA_FILE)) {
    console.log('[AutoMail] No schedule_data.json found. Exiting.');
    return;
  }

  // Prevent concurrent execution
  if (fs.existsSync(LOCK_FILE)) {
    try {
      const lockAge = Date.now() - fs.statSync(LOCK_FILE).mtimeMs;
      if (lockAge < 1000 * 60 * 30) { // 30 minutes
        console.warn('[AutoMail] Another sender process is currently running. Exiting.');
        return;
      }
    } catch (e) { }
  }
  try { fs.writeFileSync(LOCK_FILE, String(process.pid), 'utf8'); } catch (e) { }

  let data;
  try {
    data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch (e) {
    console.error('[AutoMail] Parse error reading schedule_data.json:', e.message);
    try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
    return;
  }

  const { smtpConfig, emailDelay = 0, useJitter = false, minDelay = 2, maxDelay = 6 } = data;
  if (!smtpConfig) {
    console.error('[AutoMail] Missing SMTP configuration. Aborting.');
    try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
    cleanupSchedule();
    return;
  }

  // Determine if this is a multi-batch schedule or a single legacy campaign
  const isMultiBatch = Array.isArray(data.batches) && data.batches.length > 0;
  let currentBatch = null;
  let itemsToSend = [];
  let batchLabel = '';

  if (isMultiBatch) {
    currentBatch = data.batches.find(b => b.status === 'pending');
    if (!currentBatch) {
      console.log('[AutoMail] All batches in schedule_data.json are already completed!');
      try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
      cleanupSchedule();
      return;
    }
    itemsToSend = currentBatch.items || [];
    batchLabel = `batch_${currentBatch.batchNumber}`;
    console.log(`[AutoMail] Processing Batch #${currentBatch.batchNumber} of ${data.batches.length} (${currentBatch.dayLabel || 'Day'})`);
    console.log(`[AutoMail] Batch item count: ${itemsToSend.length} emails`);
  } else {
    itemsToSend = data.campaignLogs || [];
    console.log(`[AutoMail] Processing single legacy campaign (${itemsToSend.length} emails)`);
  }

  if (itemsToSend.length === 0) {
    console.log('[AutoMail] No emails to send in current batch.');
    if (isMultiBatch && currentBatch) {
      currentBatch.status = 'completed';
      currentBatch.sentAt = new Date().toISOString();
      fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
    }
    try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
    return;
  }

  // Initialize SMTP transport
  const portNum = parseInt(smtpConfig.port, 10) || 587;
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: portNum,
    secure: smtpConfig.secure === true || portNum === 465,
    auth: { user: smtpConfig.user, pass: smtpConfig.pass },
    connectionTimeout: 15000,
    greetingTimeout: 15000,
    socketTimeout: 20000
  });

  try {
    await transporter.verify();
    console.log('[AutoMail] SMTP connection verified successfully');
  } catch (err) {
    console.error('[AutoMail] SMTP connection failed:', err.message);
    writeLog([{ error: 'SMTP connection failed: ' + err.message, timestamp: new Date().toISOString() }], 'error');
    try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
    return;
  }

  const results = [];
  let sent = 0, fail = 0;

  for (let i = 0; i < itemsToSend.length; i++) {
    const item = itemsToSend[i];
    const prefix = `[AutoMail] (${i + 1}/${itemsToSend.length})`;

    // Validate recipient email format and active domain to protect daily sending quota
    const validation = await validateEmail(item.recipientEmail);
    if (!validation.isValid || !validation.isActive) {
      console.warn(prefix, 'Skipping inactive/invalid email:', item.recipientEmail, `(${validation.reason})`);
      item.status = 'failed';
      item.error = `Skipped (Quota saved): ${validation.reason}`;
      item.sentTime = new Date().toLocaleString();
      results.push({ ...item });
      fail++;
      continue;
    }

    try {
      const fromAddr = smtpConfig.fromName
        ? '"' + smtpConfig.fromName + '" <' + smtpConfig.user + '>'
        : smtpConfig.user;

      // Format body and signature with clickable HTML links (e.g. LinkedIn, GitHub, email, tel)
      const signature = getActiveSignature(data);
      const formatted = formatEmailContent(item.body, signature);

      const opts = {
        from: fromAddr,
        to: item.recipientEmail,
        subject: item.topic,
        text: formatted.text,
        html: formatted.html
      };

      if (item.cc && String(item.cc).trim()) opts.cc = String(item.cc).trim();
      if (item.bcc && String(item.bcc).trim()) opts.bcc = String(item.bcc).trim();

      if (item.attachmentPath && String(item.attachmentPath).trim()) {
        const paths = String(item.attachmentPath)
          .split(/[,;]+/)
          .map(p => p.trim().replace(/^["']|["']$/g, ''))
          .filter(p => p && fs.existsSync(p));
        if (paths.length > 0) {
          opts.attachments = paths.map(p => ({
            filename: path.basename(p.replace(/\\/g, '/')),
            path: p
          }));
        }
      }

      await transporter.sendMail(opts);
      sent++;
      item.status = 'sent';
      item.sentTime = new Date().toLocaleString();
      item.error = '-';
      console.log(prefix, 'Sent ->', item.recipientEmail);
      results.push({ ...item });

      if (i < itemsToSend.length - 1) {
        let delaySec = parseInt(emailDelay, 10) || 0;
        if (useJitter) {
          const minSec = parseInt(minDelay, 10) || 1;
          const maxSec = parseInt(maxDelay, 10) || 5;
          if (maxSec >= minSec) {
            delaySec = Math.floor(Math.random() * (maxSec - minSec + 1)) + minSec;
          }
        }
        if (delaySec > 0) {
          await sleep(delaySec * 1000);
        }
      }

    } catch (err) {
      fail++;
      item.status = 'failed';
      item.error = err.message;
      item.sentTime = new Date().toLocaleString();
      console.error(prefix, 'Failed ->', item.recipientEmail, ':', err.message);
      results.push({ ...item });
    }

    // Periodically sync progress every 20 emails
    if (i > 0 && i % 20 === 0 && isMultiBatch) {
      try { fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8'); } catch (e) { }
    }
  }

  console.log(`[AutoMail] Batch complete! Sent: ${sent} | Failed: ${fail} | Total: ${itemsToSend.length}`);
  writeLog(results, batchLabel);

  if (isMultiBatch && currentBatch) {
    currentBatch.status = 'completed';
    currentBatch.sentAt = new Date().toISOString();
    currentBatch.sentCount = sent;
    currentBatch.failCount = fail;

    // Check if another pending batch exists
    const nextBatch = data.batches.find(b => b.status === 'pending');
    if (nextBatch) {
      // Save state so next batch is ready
      data.savedAt = new Date().toISOString();
      fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
      console.log(`[AutoMail] Rescheduling next batch #${nextBatch.batchNumber} (${nextBatch.dayLabel}) for ${nextBatch.isoString}...`);
      await rescheduleNextTask(nextBatch.isoString);
      console.log(`[AutoMail] Windows Task successfully chained to Batch #${nextBatch.batchNumber}!`);
    } else {
      console.log('[AutoMail] All batches completed successfully! Entire campaign finished.');
      cleanupSchedule();
    }
  } else {
    // Single legacy campaign finished
    cleanupSchedule();
  }

  try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
  console.log('[AutoMail] Sender finished at', new Date().toLocaleString());
  console.log('[AutoMail] ======================================================');
}

run().catch(err => {
  console.error('[AutoMail] Fatal error:', err);
  try { fs.unlinkSync(LOCK_FILE); } catch (e) { }
  process.exit(1);
});
