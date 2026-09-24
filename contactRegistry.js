/**
 * contactRegistry.js - Persistent cross-campaign contact memory.
 *
 * Tracks every recipient ever emailed (independent of any one Excel sheet) so the
 * app can dedupe future campaigns, retry transient failures, and drive follow-ups
 * and reply/bounce status without depending on the original spreadsheet still
 * being loaded. Plain CommonJS with no Electron dependency, so both the Electron
 * main process (via IPC) and the standalone background scripts (sender.js,
 * maintenance.js) can require() it directly.
 */
'use strict';
const fs = require('fs');
const path = require('path');

const REGISTRY_FILE = path.join(__dirname, 'contacts_registry.json');

function loadRegistry() {
  try {
    if (!fs.existsSync(REGISTRY_FILE)) return {};
    return JSON.parse(fs.readFileSync(REGISTRY_FILE, 'utf8'));
  } catch (e) {
    return {};
  }
}

function saveRegistry(registry) {
  fs.writeFileSync(REGISTRY_FILE, JSON.stringify(registry, null, 2), 'utf8');
}

function normEmail(email) {
  return String(email || '').trim().toLowerCase();
}

function getEntry(registry, email) {
  return registry[normEmail(email)] || null;
}

function isContacted(registry, email) {
  return !!registry[normEmail(email)];
}

/**
 * Records a successful send. `payload` snapshots the final rendered email so
 * retries/follow-ups don't need the original Excel row again.
 */
function recordSent(registry, email, { company, subject, body, cc, bcc, attachmentPath } = {}) {
  const key = normEmail(email);
  if (!key) return registry;
  const now = new Date().toISOString();
  const existing = registry[key];

  registry[key] = {
    company: company || (existing && existing.company) || '',
    firstSentAt: (existing && existing.firstSentAt) || now,
    lastSentAt: now,
    status: 'sent',
    sendCount: ((existing && existing.sendCount) || 0) + 1,
    followUpCount: (existing && existing.followUpCount) || 0,
    retryCount: 0,
    lastError: null,
    lastCheckedAt: (existing && existing.lastCheckedAt) || null,
    lastPayload: { subject: subject || '', body: body || '', cc: cc || '', bcc: bcc || '', attachmentPath: attachmentPath || '' }
  };
  return registry;
}

/**
 * Records a failed send attempt. Keeps `lastPayload` so maintenanceCore can retry it later.
 */
function recordFailed(registry, email, error, { company, subject, body, cc, bcc, attachmentPath } = {}) {
  const key = normEmail(email);
  if (!key) return registry;
  const now = new Date().toISOString();
  const existing = registry[key];

  registry[key] = {
    company: company || (existing && existing.company) || '',
    firstSentAt: (existing && existing.firstSentAt) || now,
    lastSentAt: (existing && existing.lastSentAt) || null,
    status: 'failed',
    sendCount: (existing && existing.sendCount) || 0,
    followUpCount: (existing && existing.followUpCount) || 0,
    retryCount: (existing && existing.retryCount) || 0,
    lastError: error || 'Unknown error',
    lastCheckedAt: (existing && existing.lastCheckedAt) || null,
    lastPayload: (existing && existing.lastPayload) || { subject: subject || '', body: body || '', cc: cc || '', bcc: bcc || '', attachmentPath: attachmentPath || '' }
  };
  return registry;
}

function recordBounced(registry, email, reason) {
  const key = normEmail(email);
  if (!key || !registry[key]) return registry;
  registry[key].status = 'bounced';
  registry[key].lastError = reason || 'Bounced';
  registry[key].lastCheckedAt = new Date().toISOString();
  return registry;
}

function recordReplied(registry, email) {
  const key = normEmail(email);
  if (!key || !registry[key]) return registry;
  registry[key].status = 'replied';
  registry[key].repliedAt = new Date().toISOString();
  registry[key].lastCheckedAt = new Date().toISOString();
  return registry;
}

/**
 * One-time seed from existing campaign_log_*.json files so dedupe/retry aren't
 * empty on the first run after this feature ships.
 */
function backfillFromLogs(appDir) {
  const registry = loadRegistry();
  let added = 0;

  const files = fs.readdirSync(appDir).filter(f => f.startsWith('campaign_log_') && f.endsWith('.json'));
  for (const f of files) {
    try {
      const content = JSON.parse(fs.readFileSync(path.join(appDir, f), 'utf8'));
      if (!Array.isArray(content)) continue;
      for (const item of content) {
        if (!item || !item.recipientEmail) continue;
        const key = normEmail(item.recipientEmail);
        if (registry[key]) continue; // don't clobber a newer, richer entry
        if (item.status === 'sent') {
          recordSent(registry, key, { company: item.company, subject: item.topic, body: item.body, cc: item.cc, bcc: item.bcc, attachmentPath: item.attachmentPath });
          added++;
        } else if (item.status === 'failed') {
          recordFailed(registry, key, item.error, { company: item.company, subject: item.topic, body: item.body, cc: item.cc, bcc: item.bcc, attachmentPath: item.attachmentPath });
          added++;
        }
      }
    } catch (e) { /* skip unreadable log file */ }
  }

  saveRegistry(registry);
  return { added, totalContacts: Object.keys(registry).length };
}

module.exports = {
  REGISTRY_FILE,
  loadRegistry,
  saveRegistry,
  normEmail,
  getEntry,
  isContacted,
  recordSent,
  recordFailed,
  recordBounced,
  recordReplied,
  backfillFromLogs
};
