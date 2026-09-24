/**
 * mailSync.js - On-demand bounce check via IMAP against the same Gmail
 * account/app-password already used for SMTP sending. No recurring task, no
 * reply tracking - just: "did any of these bounce?" when the user asks.
 */
'use strict';
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

// Gmail label applied to every bounce notification found, regardless of the
// specific rejection reason (address not found, mailbox full, blocked, ...).
// Gmail auto-creates the label on first use if it doesn't already exist.
const BOUNCE_LABEL = 'address-not-found';

function buildImapConfig(smtpConfig) {
  return {
    host: 'imap.gmail.com',
    port: 993,
    secure: true,
    auth: { user: smtpConfig.user, pass: smtpConfig.pass },
    logger: false
  };
}

// Bounce notifications either quote the original message headers, or say it
// in plain English ("wasn't delivered to X", "message to X has been blocked").
function extractBouncedAddress(text) {
  const patterns = [
    /Original-Recipient:\s*rfc822;\s*([^\s]+)/i,
    /Final-Recipient:\s*rfc822;\s*([^\s]+)/i,
    /(?:delivered to|to)\s*[:\-]?\s*<?([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)>?/i
  ];
  for (const re of patterns) {
    const m = text.match(re);
    if (m) return m[1].toLowerCase().replace(/[>,;.]+$/, '');
  }
  return null;
}

// Bounce phrasing varies endlessly by provider and rejection reason (mailbox
// full, address not found, blocked, quota exceeded, spam, ...). Rather than
// chase every wording, fall back to the one thing that's always true: the
// bounce notification always quotes the address we actually sent to
// somewhere in its text. If regex extraction misses (or extracts something
// that isn't one of ours), scan for any of our own recipients appearing in
// the message instead - catches "any reason", not just the phrasings above.
function findKnownRecipientInText(text, knownRecipientsLower) {
  const lower = text.toLowerCase();
  for (const email of knownRecipientsLower) {
    if (email && lower.includes(email)) return email;
  }
  return null;
}

/**
 * @param {object} smtpConfig - { user, pass } (same account as SMTP)
 * @param {Date} sinceDate - only look at bounce mail received after this
 * @param {string[]} knownRecipients - every address this campaign actually sent to;
 *   used as a fallback match when the regex can't parse a given bounce's phrasing
 * @returns {Promise<Array<{email: string, reason: string, date: string}>>}
 */
async function checkBounces(smtpConfig, sinceDate, knownRecipients = []) {
  const found = [];
  const knownLower = knownRecipients.map(e => String(e || '').trim().toLowerCase()).filter(Boolean);

  const client = new ImapFlow(buildImapConfig(smtpConfig));
  await client.connect();
  try {
    const lock = await client.getMailboxLock('INBOX');
    try {
      // mailer-daemon covers virtually all bounce reasons when sending via
      // Gmail's SMTP relay (Gmail itself generates the bounce back to the
      // sender regardless of why the remote server rejected it), but also
      // sweep common alternate senders/subjects some organizations use.
      const uids = await client.search({
        since: sinceDate,
        or: [
          { from: 'mailer-daemon' },
          { from: 'postmaster' },
          { subject: 'Delivery Status Notification' },
          { subject: 'Undelivered Mail' },
          { subject: 'Mail delivery failed' },
          { subject: 'Returned mail' },
          { subject: 'failure notice' }
        ]
      }, { uid: true });

      for (const uid of uids || []) {
        try {
          const msg = await client.fetchOne(uid, { source: true, envelope: true, internalDate: true });
          if (!msg || !msg.source) continue;
          const parsed = await simpleParser(msg.source);
          const text = (parsed.text || '') + ' ' + (parsed.html || '');

          let bounced = extractBouncedAddress(text);
          if (bounced && knownLower.length > 0 && !knownLower.includes(bounced)) {
            // Regex found *an* address, but it's not one we sent to (could be
            // a quoted CC, a different header) - fall back to a direct scan.
            bounced = findKnownRecipientInText(text, knownLower) || bounced;
          } else if (!bounced && knownLower.length > 0) {
            bounced = findKnownRecipientInText(text, knownLower);
          }
          if (!bounced) continue;

          found.push({
            email: bounced,
            reason: (parsed.subject || 'Delivery failed') + (parsed.text ? ' - ' + parsed.text.split('\n')[0].slice(0, 140) : ''),
            date: msg.internalDate ? msg.internalDate.toISOString() : new Date().toISOString()
          });

          // Tag the bounce notification itself in Gmail so it's easy to find
          // later, regardless of which specific reason caused it.
          try {
            await client.messageFlagsAdd([uid], [BOUNCE_LABEL], { uid: true, useLabels: true });
          } catch (e) { /* labeling is best-effort, don't fail the whole scan over it */ }
        } catch (e) { /* skip unparseable message */ }
      }
    } finally {
      lock.release();
    }
  } finally {
    try { await client.logout(); } catch (e) { /* ignore */ }
  }
  return found;
}

module.exports = { checkBounces, extractBouncedAddress };
