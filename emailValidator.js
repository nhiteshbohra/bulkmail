/**
 * emailValidator.js - Fast, Zero-Dependency Email & Active Domain Validator
 * 
 * Verifies email addresses before sending to avoid wasting daily email quotas
 * (e.g. Gmail 500/day limit) on invalid, inactive, dead-domain, or burner addresses.
 */
'use strict';

const dns = require('dns').promises;

// In-memory cache for domain lookup results: domain -> { active: boolean, reason: string, mx: Array }
const domainCache = new Map();

// DNS/network errors that mean "our resolver/connection is having trouble" rather
// than "this domain doesn't exist". Treating these as invalid was misclassifying
// entire dozens of good recipients as dead the moment the local network blipped -
// a DNS server refusing/timing out mid-campaign is a local outage, not proof a
// company's domain doesn't exist.
const CONNECTIVITY_ERROR_CODES = new Set([
  'ECONNREFUSED', 'ECONNRESET', 'ETIMEDOUT', 'EAI_AGAIN', 'ENETUNREACH', 'ENETDOWN', 'EHOSTUNREACH'
]);

function isConnectivityError(err) {
  return !!(err && CONNECTIVITY_ERROR_CODES.has(err.code));
}

// Pre-populate well-known, highly reliable active domains for 0ms instantaneous lookup
const KNOWN_ACTIVE_DOMAINS = [
  'gmail.com', 'googlemail.com',
  'outlook.com', 'hotmail.com', 'live.com', 'msn.com', 'office365.com',
  'yahoo.com', 'ymail.com', 'rocketmail.com',
  'icloud.com', 'me.com', 'mac.com',
  'proton.me', 'protonmail.com',
  'zoho.com', 'zohomail.com',
  'aol.com', 'comcast.net', 'gmx.com', 'mail.com'
];

KNOWN_ACTIVE_DOMAINS.forEach(d => {
  domainCache.set(d, { active: true, reason: 'Verified known mail provider', mx: ['pre-cached'] });
});

// Common typo domains to detect and reject immediately
const TYPO_DOMAINS = new Map([
  ['gmial.com', 'gmail.com'],
  ['gmai.com', 'gmail.com'],
  ['gamil.com', 'gmail.com'],
  ['gmaill.com', 'gmail.com'],
  ['gmaik.com', 'gmail.com'],
  ['hotmial.com', 'hotmail.com'],
  ['hotmai.com', 'hotmail.com'],
  ['yaho.com', 'yahoo.com'],
  ['yahooo.com', 'yahoo.com'],
  ['outlok.com', 'outlook.com'],
  ['outloo.com', 'outlook.com']
]);

// Top disposable / burner email domains
const DISPOSABLE_DOMAINS = new Set([
  'mailinator.com', 'tempmail.com', '10minutemail.com', 'guerrillamail.com',
  'sharklasers.com', 'yopmail.com', 'throwawaymail.com', 'trashmail.com',
  'dispostable.com', 'getairmail.com', 'mohmal.com', 'temp-mail.org',
  'crazymailing.com', 'burnermail.io', 'inboxkitten.com', 'fakemailgenerator.com'
]);

/**
 * Validates the basic syntax of an email string against RFC 5322 guidelines.
 * @param {string} email 
 * @returns {{ valid: boolean, reason?: string, localPart?: string, domain?: string }}
 */
function validateSyntax(email) {
  if (!email || typeof email !== 'string') {
    return { valid: false, reason: 'Email is empty or missing' };
  }

  const trimmed = email.trim();
  if (trimmed.length > 254) {
    return { valid: false, reason: 'Email exceeds maximum 254 characters' };
  }

  const atIndex = trimmed.lastIndexOf('@');
  if (atIndex === -1 || atIndex === 0 || atIndex === trimmed.length - 1) {
    return { valid: false, reason: 'Missing username or domain (malformed @)' };
  }

  const localPart = trimmed.slice(0, atIndex);
  const domain = trimmed.slice(atIndex + 1).toLowerCase();

  if (localPart.length > 64) {
    return { valid: false, reason: 'Local username exceeds maximum 64 characters' };
  }

  // Check for spaces or forbidden special characters
  if (/\s/.test(trimmed)) {
    return { valid: false, reason: 'Email contains illegal whitespace characters' };
  }

  // Check for consecutive dots in local part or domain
  if (localPart.includes('..') || domain.includes('..')) {
    return { valid: false, reason: 'Email contains consecutive dots (..)' };
  }

  // Check leading / trailing dots
  if (localPart.startsWith('.') || localPart.endsWith('.')) {
    return { valid: false, reason: 'Username cannot begin or end with a period' };
  }
  if (domain.startsWith('.') || domain.endsWith('.')) {
    return { valid: false, reason: 'Domain cannot begin or end with a period' };
  }

  // Validate local part pattern
  const localRegex = /^[a-zA-Z0-9!#$%&'*+/=?^_`{|}~.-]+$/;
  if (!localRegex.test(localPart)) {
    return { valid: false, reason: 'Username contains invalid characters' };
  }

  // Validate domain syntax (must have valid labels and at least a 2-char TLD)
  const domainRegex = /^[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/;
  if (!domainRegex.test(domain)) {
    return { valid: false, reason: 'Domain syntax is invalid (e.g. missing valid extension)' };
  }

  const tld = domain.split('.').pop();
  if (!tld || tld.length < 2 || /^\d+$/.test(tld)) {
    return { valid: false, reason: 'Domain has an invalid top-level domain (TLD)' };
  }

  return { valid: true, localPart, domain };
}

/**
 * Checks if a domain has active mail exchange (MX) or fallback (A) DNS records.
 * @param {string} domain 
 * @param {number} timeoutMs 
 * @returns {Promise<{ active: boolean, reason: string, mx?: Array }>}
 */
async function validateDomainMx(domain, timeoutMs = 4000) {
  const normDomain = domain.toLowerCase().trim();

  // Check cache first
  if (domainCache.has(normDomain)) {
    return domainCache.get(normDomain);
  }

  // Check known typo domains
  if (TYPO_DOMAINS.has(normDomain)) {
    const suggestion = TYPO_DOMAINS.get(normDomain);
    const res = { 
      active: false, 
      reason: `Likely typo domain "${normDomain}" (did you mean "${suggestion}"?)` 
    };
    domainCache.set(normDomain, res);
    return res;
  }

  // Check disposable domains
  if (DISPOSABLE_DOMAINS.has(normDomain)) {
    const res = { 
      active: false, 
      reason: `Disposable / temporary email service (${normDomain})` 
    };
    domainCache.set(normDomain, res);
    return res;
  }

  // DNS lookup with timeout
  const timeoutPromise = new Promise((_, reject) =>
    setTimeout(() => reject(new Error('DNS lookup timed out')), timeoutMs)
  );

  try {
    const mxLookup = dns.resolveMx(normDomain);
    const mxRecords = await Promise.race([mxLookup, timeoutPromise]);

    // Handle Null MX (RFC 7505: priority 0, exchange '.' explicitly means no mail service)
    if (mxRecords && mxRecords.length === 1 && (!mxRecords[0].exchange || mxRecords[0].exchange === '.')) {
      const res = { active: false, reason: `Domain explicitly rejects mail (Null MX record)` };
      domainCache.set(normDomain, res);
      return res;
    }

    if (Array.isArray(mxRecords) && mxRecords.length > 0) {
      const exchanges = mxRecords.map(r => r.exchange).filter(Boolean);
      if (exchanges.length > 0) {
        const res = { active: true, reason: 'Active mail server verified', mx: exchanges };
        domainCache.set(normDomain, res);
        return res;
      }
    }

    // If resolveMx succeeded but returned empty array, try A record fallback (RFC 5321)
    const aRecords = await Promise.race([dns.resolve4(normDomain), timeoutPromise]).catch(() => []);
    if (aRecords && aRecords.length > 0) {
      const res = { active: true, reason: 'Active host verified (fallback A record)', mx: [] };
      domainCache.set(normDomain, res);
      return res;
    }

    const res = { active: false, reason: `Domain "${normDomain}" has no active mail servers (no MX or A records)` };
    domainCache.set(normDomain, res);
    return res;
  } catch (err) {
    // Check if domain exists with fallback A record if MX lookup failed with NODATA
    if (err.code === 'ENODATA' || err.code === 'NODATA') {
      try {
        const aRecords = await Promise.race([dns.resolve4(normDomain), timeoutPromise]);
        if (aRecords && aRecords.length > 0) {
          const res = { active: true, reason: 'Active host verified (A record)', mx: [] };
          domainCache.set(normDomain, res);
          return res;
        }
      } catch (_) {}
    }

    let reason;
    if (err.code === 'ENOTFOUND' || err.code === 'NOTFOUND') {
      reason = `Domain "${normDomain}" does not exist`;
    } else if (err.code === 'ENODATA' || err.code === 'NODATA') {
      reason = `Domain "${normDomain}" has no mail server configured (no MX records)`;
    } else if ((err.message && err.message.includes('timed out')) || isConnectivityError(err)) {
      // Our own DNS resolver/network is unreachable right now - not evidence the
      // domain is bad. Allow cautiously and do NOT cache, so it gets a fair
      // re-check once connectivity recovers rather than being stuck "invalid".
      return { active: true, reason: `Local network/DNS resolver issue (${err.code || 'timeout'}) - allowed cautiously`, mx: [] };
    } else {
      reason = `Domain DNS lookup failed (${err.code || err.message})`;
    }

    const res = { active: false, reason };
    domainCache.set(normDomain, res);
    return res;
  }
}

/**
 * Validates a single email address completely: syntax, domain existence, and MX records.
 * 
 * @param {string} email - Email address to test.
 * @param {object} [options]
 * @param {boolean} [options.checkMx=true] - Perform DNS MX lookup.
 * @param {boolean} [options.rejectDisposable=true] - Reject temporary burner emails.
 * @returns {Promise<{ isValid: boolean, isActive: boolean, reason: string, details: object }>}
 */
async function validateEmail(email, options = {}) {
  const { checkMx = true, rejectDisposable = true } = options;

  const syntax = validateSyntax(email);
  if (!syntax.valid) {
    return {
      isValid: false,
      isActive: false,
      reason: syntax.reason,
      details: { email, syntaxValid: false, mxValid: false, domain: null }
    };
  }

  const { domain, localPart } = syntax;

  if (rejectDisposable && DISPOSABLE_DOMAINS.has(domain)) {
    return {
      isValid: false,
      isActive: false,
      reason: `Disposable / temporary email address (${domain})`,
      details: { email, syntaxValid: true, mxValid: false, domain, isDisposable: true }
    };
  }

  if (TYPO_DOMAINS.has(domain)) {
    const suggestion = TYPO_DOMAINS.get(domain);
    return {
      isValid: false,
      isActive: false,
      reason: `Suspicious typo domain "${domain}" (intended: "${suggestion}"?)`,
      details: { email, syntaxValid: true, mxValid: false, domain, typoSuggestion: suggestion }
    };
  }

  if (checkMx) {
    const mxResult = await validateDomainMx(domain);
    if (!mxResult.active) {
      return {
        isValid: false,
        isActive: false,
        reason: mxResult.reason,
        details: { email, syntaxValid: true, mxValid: false, domain, error: mxResult.reason }
      };
    }
  }

  return {
    isValid: true,
    isActive: true,
    reason: 'Active and valid',
    details: { email, syntaxValid: true, mxValid: true, domain, localPart }
  };
}

/**
 * Batch-validates an array of email addresses with concurrency control.
 * 
 * @param {Array<string|object>} emails - Array of email strings or objects with `recipientEmail`
 * @param {object} [options]
 * @param {number} [options.concurrency=15] - Maximum parallel DNS queries.
 * @param {function} [onProgress] - Optional progress callback: `(completed, total, itemResult)`
 * @returns {Promise<Array<object>>} Array of validation result objects.
 */
async function validateBatch(emails, options = {}, onProgress = null) {
  const { concurrency = 15, checkMx = true } = options;
  const results = new Array(emails.length);
  let completed = 0;
  let index = 0;

  async function worker() {
    while (index < emails.length) {
      const currentIndex = index++;
      const raw = emails[currentIndex];
      const emailStr = typeof raw === 'string' ? raw : (raw && raw.recipientEmail) || '';
      
      const validation = await validateEmail(emailStr, { checkMx });
      results[currentIndex] = {
        index: currentIndex,
        raw,
        email: emailStr,
        ...validation
      };

      completed++;
      if (typeof onProgress === 'function') {
        try { onProgress(completed, emails.length, results[currentIndex]); } catch (_) {}
      }
    }
  }

  const workers = [];
  const workerCount = Math.min(concurrency, emails.length);
  for (let w = 0; w < workerCount; w++) {
    workers.push(worker());
  }

  await Promise.all(workers);
  return results;
}

/**
 * Formats email body and signature into plain text and rich HTML with clickable links.
 * Automatically converts markdown links like [Text](url) or bare URLs into HTML anchors.
 * 
 * @param {string} bodyText 
 * @param {string} [signatureText] 
 * @returns {{ text: string, html: string }}
 */
function formatEmailContent(bodyText, signatureText) {
  let fullRaw = (bodyText || '').trim();
  if (signatureText && signatureText.trim()) {
    const sigTrimmed = signatureText.trim();
    const sigFirstLine = sigTrimmed.split('\n')[0].replace(/^[-–—\s]+/, '').trim();
    if (!sigFirstLine || !fullRaw.toLowerCase().includes(sigFirstLine.toLowerCase())) {
      if (sigTrimmed.startsWith('--')) {
        fullRaw += '\n\n' + sigTrimmed;
      } else {
        fullRaw += '\n\n--\n' + sigTrimmed;
      }
    }
  }

  const text = fullRaw;

  // Escape HTML entities
  let html = fullRaw
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Markdown links: [Title](https://...) or [Title](mailto:...)
  html = html.replace(/\[([^\]]+)\]\(((?:https?:\/\/|mailto:)[^\s<)]+)\)/gi, '<a href="$2" style="color: #2563eb; text-decoration: underline;" target="_blank">$1</a>');

  // Convert bare emails to mailto links: name@domain.com
  html = html.replace(/(^|[\s>(])([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)/g, '$1<a href="mailto:$2" style="color: #2563eb; text-decoration: underline;">$2</a>');

  // Convert bare URLs (not already inside href="...")
  html = html.replace(/(^|[\s>(])(https?:\/\/[^\s<)]+)/g, (match, p1, p2) => {
    return p1 + '<a href="' + p2 + '" style="color: #2563eb; text-decoration: underline;" target="_blank">' + p2 + '</a>';
  });

  // Convert newlines to HTML <br>
  html = html.replace(/\r\n/g, '<br>').replace(/\n/g, '<br>');

  return { text, html };
}

module.exports = {
  validateSyntax,
  validateDomainMx,
  validateEmail,
  validateBatch,
  formatEmailContent,
  isConnectivityError,
  domainCache,
  KNOWN_ACTIVE_DOMAINS,
  TYPO_DOMAINS,
  DISPOSABLE_DOMAINS
};

