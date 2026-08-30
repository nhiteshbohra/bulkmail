/**
 * sender.js - Standalone Bulk Email Sender
 * Called by Windows Task Scheduler. Does NOT require Electron to be running.
 */
'use strict';
const path       = require('path');
const fs         = require('fs');
const nodemailer = require('nodemailer');
const { exec }   = require('child_process');

const DATA_FILE = path.join(__dirname, 'schedule_data.json');
const TASK_NAME = 'AutoMailExcelSchedule';

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function writeLog(results) {
  const ts      = new Date().toISOString().replace(/[:.]/g, '-');
  const logPath = path.join(__dirname, 'campaign_log_' + ts + '.json');
  try { fs.writeFileSync(logPath, JSON.stringify(results, null, 2), 'utf8'); } catch(e) {}
  console.log('[AutoMail] Log written to', logPath);
}

async function run() {
  console.log('[AutoMail] Sender started at', new Date().toLocaleString());
  if (!fs.existsSync(DATA_FILE)) {
    console.log('[AutoMail] No schedule_data.json found. Exiting.');
    return;
  }

  let data;
  try { data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch(e) { console.error('[AutoMail] Parse error:', e.message); return; }

  const { smtpConfig, campaignLogs, emailDelay = 0 } = data;
  if (!smtpConfig || !campaignLogs || campaignLogs.length === 0) {
    fs.unlinkSync(DATA_FILE);
    return;
  }

  const portNum = parseInt(smtpConfig.port, 10) || 587;
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: portNum,
    secure: smtpConfig.secure === true || portNum === 465,
    auth: { user: smtpConfig.user, pass: smtpConfig.pass },
    connectionTimeout: 15000,
    greetingTimeout:   15000,
    socketTimeout:     20000
  });

  try {
    await transporter.verify();
    console.log('[AutoMail] SMTP connection verified');
  } catch(err) {
    console.error('[AutoMail] SMTP failed:', err.message);
    writeLog([{ error: 'SMTP connection failed: ' + err.message, timestamp: new Date().toISOString() }]);
    return;
  }

  const results = [];
  let sent = 0, fail = 0;

  for (let i = 0; i < campaignLogs.length; i++) {
    const item = campaignLogs[i];
    const prefix = '[AutoMail] (' + (i+1) + '/' + campaignLogs.length + ')';

    if (!item.recipientEmail || !String(item.recipientEmail).includes('@')) {
      console.warn(prefix, 'Skipping invalid email:', item.recipientEmail);
      results.push({ ...item, status: 'failed', error: 'Invalid email', sentTime: new Date().toLocaleString() });
      fail++;
      continue;
    }

    try {
      const fromAddr = smtpConfig.fromName
        ? '"' + smtpConfig.fromName + '" <' + smtpConfig.user + '>'
        : smtpConfig.user;

      const opts = {
        from:    fromAddr,
        to:      item.recipientEmail,
        subject: item.topic,
        text:    item.body,
        html:    (item.body || '').replace(/\n/g, '<br>')
      };

      if (item.cc  && String(item.cc).trim())  opts.cc  = String(item.cc).trim();
      if (item.bcc && String(item.bcc).trim()) opts.bcc = String(item.bcc).trim();

      if (item.attachmentPath && String(item.attachmentPath).trim()) {
        const paths = String(item.attachmentPath).split(/[,;]+/)
          .map(p => p.trim()).filter(p => p && fs.existsSync(p));
        if (paths.length > 0) opts.attachments = paths.map(p => ({ path: p }));
      }

      await transporter.sendMail(opts);
      sent++;
      console.log(prefix, 'Sent ->', item.recipientEmail);
      results.push({ ...item, status: 'sent', sentTime: new Date().toLocaleString(), error: '-' });

      if (emailDelay > 0 && i < campaignLogs.length - 1) {
        await sleep(emailDelay * 1000);
      }
    } catch(err) {
      fail++;
      console.error(prefix, 'Failed ->', item.recipientEmail, ':', err.message);
      results.push({ ...item, status: 'failed', error: err.message, sentTime: new Date().toLocaleString() });
    }
  }

  console.log('[AutoMail] Done! Sent:', sent, '| Failed:', fail, '| Total:', campaignLogs.length);
  writeLog(results);

  try { fs.unlinkSync(DATA_FILE); } catch(e) {}

  exec('schtasks /delete /tn "' + TASK_NAME + '" /f', () => {});
  const bat = path.join(__dirname, 'run_schedule.bat');
  if (fs.existsSync(bat)) try { fs.unlinkSync(bat); } catch(e) {}
}

run().catch(err => { console.error('[AutoMail] Fatal error:', err); process.exit(1); });
