// ==========================================================================
// AutoMail Excel - Renderer Process Application Logic
// ==========================================================================

document.addEventListener('DOMContentLoaded', () => {
  // Global State
  const state = {
    filePath: null,
    fileName: null,
    headers: [],
    rawRows: [],
    campaignLogs: [],
    status: 'idle', // 'idle' | 'running' | 'paused' | 'stopped'
    sendMode: 'web', // 'web' | 'smtp'
    currentFilter: 'all',
    smtpVerified: false,
    googleSignedIn: false,
    stopRequested: false,
    pauseRequested: false,
    globalAttachmentPaths: []
  };

  // Schedule state (declared early so updateStartButtonState can reference it)
  let scheduleTimeoutId = null;


  // DOM Elements
  const modeWebBtn = document.getElementById('modeWebBtn');
  const modeSmtpBtn = document.getElementById('modeSmtpBtn');
  const webEnginePanel = document.getElementById('webEnginePanel');
  const smtpEnginePanel = document.getElementById('smtpEnginePanel');
  const btnGoogleSignIn = document.getElementById('btnGoogleSignIn');
  const btnCheckGoogleAuth = document.getElementById('btnCheckGoogleAuth');
  const googleAuthBadge = document.getElementById('googleAuthBadge');

  const presetSelect = document.getElementById('presetSelect');
  const smtpHost = document.getElementById('smtpHost');
  const smtpPort = document.getElementById('smtpPort');
  const smtpSecure = document.getElementById('smtpSecure');
  const smtpUser = document.getElementById('smtpUser');
  const smtpPass = document.getElementById('smtpPass');
  const smtpFromName = document.getElementById('smtpFromName');
  const emailSignature = document.getElementById('emailSignature');
  const btnTestSmtp = document.getElementById('btnTestSmtp');
  const smtpFeedback = document.getElementById('smtpFeedback');
  const smtpStatusBadge = document.getElementById('smtpStatusBadge');

  // Initialize Email Signature from localStorage or signature.txt file (account-aware)
  function getSigStorageKey() {
    const user = (smtpUser && smtpUser.value) ? smtpUser.value.trim().toLowerCase() : '';
    return user ? `automail_sig_${user}` : 'automail_sig_default';
  }

  async function loadSignatureForCurrentAccount() {
    if (!emailSignature) return;
    const key = getSigStorageKey();
    let saved = localStorage.getItem(key) || localStorage.getItem('automail_signature') || '';
    if (!saved && window.electronAPI && window.electronAPI.loadSignatureFile) {
      try {
        const fileRes = await window.electronAPI.loadSignatureFile();
        if (fileRes && fileRes.success && fileRes.signature) {
          saved = fileRes.signature;
        }
      } catch (_) {}
    }
    emailSignature.value = saved;
  }

  if (emailSignature) {
    loadSignatureForCurrentAccount();
    emailSignature.addEventListener('input', () => {
      localStorage.setItem(getSigStorageKey(), emailSignature.value);
      localStorage.setItem('automail_signature', emailSignature.value);
      if (window.electronAPI && window.electronAPI.saveSignatureFile) {
        window.electronAPI.saveSignatureFile(emailSignature.value);
      }
    });
  }

  if (smtpUser) {
    smtpUser.addEventListener('input', loadSignatureForCurrentAccount);
    smtpUser.addEventListener('change', loadSignatureForCurrentAccount);
  }

  const dropZone = document.getElementById('dropZone');
  const btnSelectFile = document.getElementById('btnSelectFile');
  const fileInfoBox = document.getElementById('fileInfoBox');
  const fileNameText = document.getElementById('fileNameText');
  const fileRowsCount = document.getElementById('fileRowsCount');
  const btnRemoveFile = document.getElementById('btnRemoveFile');

  const mappingSection = document.getElementById('mappingSection');
  const mapEmail = document.getElementById('mapEmail');
  const mapCc = document.getElementById('mapCc');
  const mapBcc = document.getElementById('mapBcc');
  const mapTopic = document.getElementById('mapTopic');
  const mapBody = document.getElementById('mapBody');
  const mapAttachment = document.getElementById('mapAttachment');
  const btnSelectGlobalAttach = document.getElementById('btnSelectGlobalAttach');
  const globalAttachCountText = document.getElementById('globalAttachCountText');
  const btnClearAllGlobalAttach = document.getElementById('btnClearAllGlobalAttach');
  const globalAttachList = document.getElementById('globalAttachList');
  const splitMultipleEmails = document.getElementById('splitMultipleEmails');

  const emailDelay = document.getElementById('emailDelay');
  const btnStart = document.getElementById('btnStart');
  const btnPause = document.getElementById('btnPause');
  const btnResume = document.getElementById('btnResume');
  const btnStop = document.getElementById('btnStop');

  const progressBarFill = document.getElementById('progressBarFill');
  const progressText = document.getElementById('progressText');
  const progressPercent = document.getElementById('progressPercent');

  const metricTotal = document.getElementById('metricTotal');
  const metricSent = document.getElementById('metricSent');
  const metricFailed = document.getElementById('metricFailed');
  const metricPending = document.getElementById('metricPending');

  const logTableBody = document.getElementById('logTableBody');
  const btnExportReport = document.getElementById('btnExportReport');
  const btnResetLogs = document.getElementById('btnResetLogs');
  const logSearch = document.getElementById('logSearch');
  const filterTabs = document.querySelectorAll('.tab-btn');

  const filterCountAll = document.getElementById('filterCountAll');
  const filterCountSent = document.getElementById('filterCountSent');
  const filterCountFailed = document.getElementById('filterCountFailed');
  const filterCountInvalid = document.getElementById('filterCountInvalid');
  const filterCountPending = document.getElementById('filterCountPending');

  // Quota Shield & Email Verification Elements
  const chkValidateEmails = document.getElementById('chkValidateEmails');
  const btnValidateList = document.getElementById('btnValidateList');
  const validationSummaryBanner = document.getElementById('validationSummaryBanner');
  const valSummaryTitleText = document.getElementById('valSummaryTitleText');
  const btnCloseValSummary = document.getElementById('btnCloseValSummary');
  const valStatValid = document.getElementById('valStatValid');
  const valStatInvalid = document.getElementById('valStatInvalid');
  const btnFilterOutInvalid = document.getElementById('btnFilterOutInvalid');
  const metricSaved = document.getElementById('metricSaved');

  // Initialize Preset Listener
  if (presetSelect) {
    presetSelect.addEventListener('change', (e) => {
      const val = e.target.value;
      if (val === 'gmail') {
        smtpHost.value = 'smtp.gmail.com';
        smtpPort.value = '587';
        smtpSecure.checked = false;
      } else if (val === 'outlook') {
        smtpHost.value = 'smtp.office365.com';
        smtpPort.value = '587';
        smtpSecure.checked = false;
      } else if (val === 'yahoo') {
        smtpHost.value = 'smtp.mail.yahoo.com';
        smtpPort.value = '465';
        smtpSecure.checked = true;
      }
    });
  }

  // Test SMTP Connection Handler
  if (btnTestSmtp) {
    btnTestSmtp.addEventListener('click', async () => {
      const config = getSmtpConfig();
      if (!config.host || !config.user || !config.pass) {
        showSmtpFeedback(false, 'Please fill in Host, User, and Password fields.');
        return;
      }

      btnTestSmtp.disabled = true;
      btnTestSmtp.textContent = 'Verifying...';
      smtpFeedback.classList.add('hidden');

      const result = await window.electronAPI.testSmtp(config);

      btnTestSmtp.disabled = false;
      btnTestSmtp.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg> Test Connection`;

      if (result.success) {
        state.smtpVerified = true;
        showSmtpFeedback(true, result.message);
        updateSmtpStatus(true);
      } else {
        state.smtpVerified = false;
        showSmtpFeedback(false, result.error);
        updateSmtpStatus(false);
      }
      updateStartButtonState();
    });
  }

  function getSmtpConfig() {
    return {
      host: smtpHost.value.trim(),
      port: smtpPort.value.trim(),
      secure: smtpSecure.checked,
      user: smtpUser.value ? smtpUser.value.trim() : '',
      pass: smtpPass.value,
      fromName: smtpFromName.value ? smtpFromName.value.trim() : ''
    };
  }

  function showSmtpFeedback(isSuccess, message) {
    if (!smtpFeedback) return;
    smtpFeedback.className = `alert-box ${isSuccess ? 'success' : 'error'}`;
    smtpFeedback.textContent = message;
    smtpFeedback.classList.remove('hidden');
  }

  function updateSmtpStatus(isVerified) {
    if (!smtpStatusBadge) return;
    if (isVerified) {
      smtpStatusBadge.className = 'status-indicator connected';
      smtpStatusBadge.querySelector('.text').textContent = 'SMTP: Connected';
    } else {
      smtpStatusBadge.className = 'status-indicator disconnected';
      smtpStatusBadge.querySelector('.text').textContent = 'SMTP: Not Verified';
    }
  }

  // File Drag and Drop & Browse
  btnSelectFile.addEventListener('click', handleFileSelect);
  dropZone.addEventListener('click', (e) => {
    if (e.target !== btnSelectFile) handleFileSelect();
  });

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      processFile(file.path);
    }
  });

  btnRemoveFile.addEventListener('click', resetFileState);

  async function handleFileSelect() {
    const filePath = await window.electronAPI.selectFile();
    if (filePath) {
      processFile(filePath);
    }
  }

  async function processFile(filePath) {
    const result = await window.electronAPI.parseFile(filePath);
    if (!result.success) {
      alert(`Error loading file: ${result.error}`);
      return;
    }

    state.filePath = result.filePath;
    state.fileName = result.fileName;
    state.headers = result.headers;
    state.rawRows = result.rows;

    // Show File Info Box
    fileNameText.textContent = result.fileName;
    fileRowsCount.textContent = `${result.totalRows} recipients found`;
    fileInfoBox.classList.remove('hidden');
    dropZone.classList.add('hidden');

    // Populate Column Mappings
    populateMappingDropdowns(result.headers);
    mappingSection.classList.remove('hidden');

    // Initialize Campaign Logs
    buildCampaignLogs();
    renderLogTable();
    updateMetrics();
    updateStartButtonState();
  }

  function resetFileState() {
    state.filePath = null;
    state.fileName = null;
    state.headers = [];
    state.rawRows = [];
    state.campaignLogs = [];

    fileInfoBox.classList.add('hidden');
    mappingSection.classList.add('hidden');
    dropZone.classList.remove('hidden');

    logTableBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="6">No Excel file loaded yet. Load a file above to preview recipients.</td>
      </tr>
    `;

    updateMetrics();
    updateStartButtonState();
    btnExportReport.disabled = true;
    if (validationSummaryBanner) validationSummaryBanner.classList.add('hidden');
  }

  // Global Attachment Listeners (Supports Multiple Images & Files)
  if (btnSelectGlobalAttach) {
    btnSelectGlobalAttach.addEventListener('click', async () => {
      const selected = await window.electronAPI.selectAttachments();
      if (selected && selected.length > 0) {
        if (!Array.isArray(state.globalAttachmentPaths)) {
          state.globalAttachmentPaths = [];
        }
        selected.forEach(filePath => {
          if (!state.globalAttachmentPaths.includes(filePath)) {
            state.globalAttachmentPaths.push(filePath);
          }
        });
        renderGlobalAttachmentsUI();
        onMappingChanged();
      }
    });
  }

  if (btnClearAllGlobalAttach) {
    btnClearAllGlobalAttach.addEventListener('click', () => {
      state.globalAttachmentPaths = [];
      renderGlobalAttachmentsUI();
      onMappingChanged();
    });
  }

  function renderGlobalAttachmentsUI() {
    if (!globalAttachList || !globalAttachCountText) return;

    const count = (state.globalAttachmentPaths || []).length;
    if (count === 0) {
      globalAttachCountText.textContent = 'No global files selected';
      globalAttachList.innerHTML = '';
      if (btnClearAllGlobalAttach) btnClearAllGlobalAttach.classList.add('hidden');
      return;
    }

    globalAttachCountText.textContent = `${count} file${count > 1 ? 's' : ''} selected`;
    if (btnClearAllGlobalAttach) btnClearAllGlobalAttach.classList.remove('hidden');

    globalAttachList.innerHTML = '';
    state.globalAttachmentPaths.forEach((filePath, index) => {
      const fileName = filePath.split(/[\\/]/).pop();
      const ext = fileName.split('.').pop().toLowerCase();
      const isImg = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'svg', 'bmp', 'ico'].includes(ext);
      const icon = isImg ? '🖼️' : '📄';

      const chip = document.createElement('div');
      chip.className = 'global-attach-chip';
      chip.title = filePath;
      chip.innerHTML = `
        <span>${icon}</span>
        <span class="global-attach-chip-name">${escapeHtml(fileName)}</span>
        <button type="button" class="global-attach-chip-remove" title="Remove ${escapeHtml(fileName)}">&times;</button>
      `;

      chip.querySelector('.global-attach-chip-remove').addEventListener('click', (e) => {
        e.stopPropagation();
        state.globalAttachmentPaths.splice(index, 1);
        renderGlobalAttachmentsUI();
        onMappingChanged();
      });

      globalAttachList.appendChild(chip);
    });
  }

  function populateMappingDropdowns(headers) {
    const selectElements = [mapEmail, mapCc, mapBcc, mapTopic, mapBody, mapAttachment].filter(Boolean);

    selectElements.forEach(select => {
      select.innerHTML = '<option value="">-- Select Column --</option>';
      headers.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        select.appendChild(option);
      });
    });

    // Auto-detect matching headers
    headers.forEach(h => {
      const lower = h.toLowerCase();
      if (!mapEmail.value && (lower.includes('email') || lower.includes('mail') || lower.includes('address'))) {
        mapEmail.value = h;
      }
      if (mapCc && !mapCc.value && (lower === 'cc' || lower.includes('cc_email') || lower.includes('cc email'))) {
        mapCc.value = h;
      }
      if (mapBcc && !mapBcc.value && (lower === 'bcc' || lower.includes('bcc_email') || lower.includes('bcc email'))) {
        mapBcc.value = h;
      }
      if (!mapTopic.value && (lower.includes('topic') || lower.includes('subject') || lower.includes('title'))) {
        mapTopic.value = h;
      }
      if (!mapBody.value && (lower.includes('body') || lower.includes('content') || lower.includes('message') || lower.includes('text'))) {
        mapBody.value = h;
      }
      if (mapAttachment && !mapAttachment.value && (lower.includes('attach') || lower.includes('file') || lower.includes('document'))) {
        mapAttachment.value = h;
      }
    });

    // Add change listeners to rebuild log preview dynamically
    mapEmail.addEventListener('change', onMappingChanged);
    if (mapCc) mapCc.addEventListener('change', onMappingChanged);
    if (mapBcc) mapBcc.addEventListener('change', onMappingChanged);
    mapTopic.addEventListener('change', onMappingChanged);
    mapBody.addEventListener('change', onMappingChanged);
    if (mapAttachment) mapAttachment.addEventListener('change', onMappingChanged);
    if (splitMultipleEmails) {
      splitMultipleEmails.addEventListener('change', onMappingChanged);
    }
  }

  function onMappingChanged() {
    buildCampaignLogs();
    renderLogTable();
    updateMetrics();
    updateStartButtonState();
  }

  function extractEmails(rawStr) {
    if (!rawStr) return [];
    // Split by comma, semicolon, newline, carriage return, slash, or pipe
    const parts = String(rawStr).split(/[,;\n\r\/|]+/);
    const emails = parts
      .map(p => p.trim())
      .filter(p => p.length > 0 && p.includes('@'));
    return emails;
  }

  function buildCampaignLogs() {
    const emailCol = mapEmail.value;
    const ccCol = mapCc ? mapCc.value : '';
    const bccCol = mapBcc ? mapBcc.value : '';
    const topicCol = mapTopic.value;
    const bodyCol = mapBody.value;
    const attachmentCol = mapAttachment ? mapAttachment.value : '';
    const shouldSplit = splitMultipleEmails ? splitMultipleEmails.checked : true;

    state.campaignLogs = [];
    let currentId = 1;

    state.rawRows.forEach((row, idx) => {
      let recipientRaw = emailCol ? String(row[emailCol] || '').trim() : '';
      let ccRaw = ccCol ? String(row[ccCol] || '').trim() : '';
      let bccRaw = bccCol ? String(row[bccCol] || '').trim() : '';
      let topicRaw = topicCol ? String(row[topicCol] || '') : '';
      let bodyRaw = bodyCol ? String(row[bodyCol] || '') : '';
      let attachmentRaw = attachmentCol ? String(row[attachmentCol] || '').trim() : '';

      // Substitute row variables for placeholders e.g., {COLUMN_NAME}
      let finalTopic = interpolateTemplate(topicRaw, row);
      let finalBody = interpolateTemplate(bodyRaw, row);
      let finalAttachment = interpolateTemplate(attachmentRaw, row);

      // Combine row attachment(s) with global master attachments if set
      let attachmentList = [];
      if (finalAttachment) {
        attachmentList.push(...finalAttachment.split(/[,;]+/).map(a => a.trim()).filter(Boolean));
      }
      if (Array.isArray(state.globalAttachmentPaths) && state.globalAttachmentPaths.length > 0) {
        state.globalAttachmentPaths.forEach(gPath => {
          if (gPath && !attachmentList.includes(gPath)) {
            attachmentList.push(gPath);
          }
        });
      }

      const parsedEmails = extractEmails(recipientRaw);

      if (shouldSplit && parsedEmails.length > 1) {
        // Expand multi-email cell into separate recipient entries in campaign log table
        parsedEmails.forEach(singleEmail => {
          state.campaignLogs.push({
            id: currentId++,
            rowNumber: idx + 1,
            recipientEmail: singleEmail,
            cc: ccRaw,
            bcc: bccRaw,
            topic: finalTopic,
            body: finalBody,
            attachmentPath: attachmentList.join('; '),
            status: 'pending',
            sentTime: '-',
            error: '-'
          });
        });
      } else {
        state.campaignLogs.push({
          id: currentId++,
          rowNumber: idx + 1,
          recipientEmail: parsedEmails.length > 1 ? parsedEmails.join(', ') : recipientRaw,
          cc: ccRaw,
          bcc: bccRaw,
          topic: finalTopic,
          body: finalBody,
          attachmentPath: attachmentList.join('; '),
          status: 'pending',
          sentTime: '-',
          error: '-'
        });
      }
    });
  }

  function interpolateTemplate(templateStr, rowObject) {
    if (!templateStr) return '';
    return templateStr.replace(/\{([^}]+)\}/g, (match, key) => {
      const trimmedKey = key.trim();
      if (rowObject.hasOwnProperty(trimmedKey)) {
        return rowObject[trimmedKey];
      }
      return match;
    });
  }

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

  function _coreUpdateStart() {
    const hasData = state.campaignLogs.length > 0;
    const hasEmailMap = mapEmail.value !== '';
    const hasTopicMap = mapTopic.value !== '';
    const hasBodyMap = mapBody.value !== '';
    const hasHost = smtpHost.value.trim() !== '';

    btnStart.disabled = !(hasData && hasEmailMap && hasTopicMap && hasBodyMap && hasHost && state.status === 'idle');
    if (btnValidateList) {
      btnValidateList.disabled = !(hasData && state.status === 'idle');
    }
  }

  function updateStartButtonState() {
    _coreUpdateStart();
    // Sync schedule button enabled state
    const sdEl = document.getElementById('scheduleDateTime');
    const schBtn = document.getElementById('btnSchedule');
    if (sdEl && schBtn && sdEl.value && !scheduleTimeoutId) {
      const chosen = new Date(sdEl.value);
      schBtn.disabled = (chosen <= new Date()) || (state.campaignLogs.length === 0);
    }
    if (typeof updateBatchPreview === 'function') {
      updateBatchPreview();
    }
  }

  // Render Log Table with Filtering & Search
  function renderLogTable() {
    if (state.campaignLogs.length === 0) return;

    const searchTerm = logSearch.value.toLowerCase().trim();
    const filter = state.currentFilter;

    const filteredLogs = state.campaignLogs.filter(item => {
      const matchesFilter = filter === 'all' || item.status === filter;
      const matchesSearch = !searchTerm || 
        item.recipientEmail.toLowerCase().includes(searchTerm) || 
        item.topic.toLowerCase().includes(searchTerm);
      return matchesFilter && matchesSearch;
    });

    if (filteredLogs.length === 0) {
      logTableBody.innerHTML = `
        <tr class="empty-row">
          <td colspan="6">No entries match your search or filter criteria.</td>
        </tr>
      `;
      return;
    }

    logTableBody.innerHTML = filteredLogs.map(item => `
      <tr id="row-${item.id}">
        <td>${item.id}</td>
        <td><strong>${escapeHtml(item.recipientEmail || '(Missing Email)')}</strong></td>
        <td>${escapeHtml(item.topic || '-')}</td>
        <td>${getStatusBadge(item.status)}</td>
        <td style="font-size: 0.75rem; color: var(--text-secondary);">${escapeHtml(item.sentTime || '-')}</td>
        <td>${item.error !== '-' ? `<span class="error-text">${escapeHtml(item.error)}</span>` : '-'}</td>
      </tr>
    `).join('');
  }

  function updateSingleRowUI(item) {
    if (!item) return;
    const rowEl = document.getElementById(`row-${item.id}`);
    if (!rowEl) return;

    const cells = rowEl.querySelectorAll('td');
    if (cells.length >= 6) {
      cells[3].innerHTML = getStatusBadge(item.status);
      cells[4].textContent = item.sentTime || '-';
      cells[5].innerHTML = item.error && item.error !== '-' ? `<span class="error-text">${escapeHtml(item.error)}</span>` : '-';
    }
  }

  function getStatusBadge(status) {
    switch (status) {
      case 'sent':
        return '<span class="badge badge-sent">✓ Sent</span>';
      case 'failed':
        return '<span class="badge badge-failed">✕ Failed</span>';
      case 'invalid':
        return '<span class="badge badge-invalid">⚠️ Inactive / Invalid</span>';
      case 'sending':
        return '<span class="badge badge-sending">⌛ Sending...</span>';
      default:
        return '<span class="badge badge-pending">Pending</span>';
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function updateMetrics() {
    const total = state.campaignLogs.length;
    const sent = state.campaignLogs.filter(l => l.status === 'sent').length;
    const failed = state.campaignLogs.filter(l => l.status === 'failed').length;
    const invalid = state.campaignLogs.filter(l => l.status === 'invalid').length;
    const pending = state.campaignLogs.filter(l => l.status === 'pending').length;

    metricTotal.textContent = total;
    metricSent.textContent = sent;
    metricFailed.textContent = failed;
    if (metricSaved) metricSaved.textContent = invalid;
    metricPending.textContent = pending;

    filterCountAll.textContent = total;
    filterCountSent.textContent = sent;
    filterCountFailed.textContent = failed;
    if (filterCountInvalid) filterCountInvalid.textContent = invalid;
    filterCountPending.textContent = pending;

    const processed = sent + failed + invalid;
    const percent = total > 0 ? Math.round((processed / total) * 100) : 0;

    progressBarFill.style.width = `${percent}%`;
    progressPercent.textContent = `${percent}%`;

    if (state.status === 'running') {
      progressText.textContent = `Sending ${processed} of ${total}...`;
    } else if (state.status === 'paused') {
      progressText.textContent = `Campaign Paused (${processed} of ${total})`;
    } else if (state.status === 'stopped') {
      progressText.textContent = `Campaign Stopped (${processed} of ${total})`;
    } else if (processed === total && total > 0) {
      const savedNotice = invalid > 0 ? ` (${invalid} Quota Saved)` : '';
      progressText.textContent = `Campaign Complete! (${sent} Sent, ${failed} Failed${savedNotice})`;
    } else {
      progressText.textContent = 'Ready to start campaign';
    }

    if (total > 0) {
      btnExportReport.disabled = false;
      if (btnResetLogs) btnResetLogs.disabled = (state.status !== 'idle');
    } else {
      if (btnResetLogs) btnResetLogs.disabled = true;
    }
  }

  // Filter Tabs Event Listeners
  filterTabs.forEach(tab => {
    tab.addEventListener('click', (e) => {
      filterTabs.forEach(t => t.classList.remove('active'));
      e.target.classList.add('active');
      state.currentFilter = e.target.getAttribute('data-filter');
      renderLogTable();
    });
  });

  logSearch.addEventListener('input', renderLogTable);

  // Campaign Automation Control Buttons
  btnStart.addEventListener('click', startCampaign);
  btnPause.addEventListener('click', pauseCampaign);
  btnResume.addEventListener('click', resumeCampaign);
  btnStop.addEventListener('click', stopCampaign);

  async function startCampaign() {
    const smtpConfig = getSmtpConfig();
    if (!smtpConfig.host || !smtpConfig.user || !smtpConfig.pass) {
      alert('Please fill out your SMTP Server, User, and Password credentials first!');
      return;
    }

    state.status = 'running';
    state.stopRequested = false;
    state.pauseRequested = false;

    btnStart.classList.add('hidden');
    btnPause.classList.remove('hidden');
    btnStop.classList.remove('hidden');
    btnResume.classList.add('hidden');
    setInputsDisabled(true);

    const delaySec = parseInt(emailDelay.value, 10) || 0;

    for (let i = 0; i < state.campaignLogs.length; i++) {
      const item = state.campaignLogs[i];

      // Skip already sent or invalid rows if re-starting
      if (item.status === 'sent' || item.status === 'invalid') continue;

      // Handle Pause Request
      while (state.pauseRequested) {
        if (state.stopRequested) break;
        await sleep(200);
      }

      // Handle Stop Request
      if (state.stopRequested) {
        break;
      }

      // Pre-send validation to protect daily sending quota
      const shouldValidate = chkValidateEmails ? chkValidateEmails.checked : true;
      if (shouldValidate) {
        const validation = await window.electronAPI.validateEmail(item.recipientEmail);
        if (!validation.isValid || !validation.isActive) {
          item.status = 'invalid';
          item.error = `Skipped (Quota saved): ${validation.reason}`;
          item.sentTime = new Date().toLocaleTimeString();
          updateSingleRowUI(item);
          updateMetrics();
          continue; // Directly skip sending email, protecting daily quota!
        }
      } else if (!item.recipientEmail || !item.recipientEmail.includes('@')) {
        item.status = 'failed';
        item.error = 'Invalid email address format';
        item.sentTime = new Date().toLocaleTimeString();
        updateSingleRowUI(item);
        updateMetrics();
        continue;
      }

      // Mark row as sending
      item.status = 'sending';
      updateSingleRowUI(item);
      updateMetrics();

      // Format email body and signature with clickable HTML links
      const signatureStr = emailSignature ? emailSignature.value : '';
      const formatted = formatEmailContent(item.body, signatureStr);

      // Dispatch Email via SMTP
      const mailData = {
        to: item.recipientEmail,
        cc: item.cc || '',
        bcc: item.bcc || '',
        subject: item.topic,
        text: formatted.text,
        html: formatted.html,
        attachmentPath: item.attachmentPath,
        isHtml: true
      };

      const res = await window.electronAPI.sendEmail(smtpConfig, mailData);

      // If stop was requested while email was sending
      if (state.stopRequested) {
        item.status = 'pending';
        updateSingleRowUI(item);
        break;
      }

      if (res.success) {
        item.status = 'sent';
        item.sentTime = new Date().toLocaleTimeString();
        item.error = '-';
      } else {
        item.status = 'failed';
        item.sentTime = new Date().toLocaleTimeString();
        item.error = res.error;
      }

      updateSingleRowUI(item);
      updateMetrics();

      // Inter-email Delay (Interruptible)
      if (delaySec > 0 && i < state.campaignLogs.length - 1 && !state.stopRequested) {
        await interruptibleSleep(delaySec * 1000);
      }
    }

    // Always reset state status back to 'idle' when loop terminates or stops
    state.status = 'idle';
    state.stopRequested = false;
    state.pauseRequested = false;

    // Reset any hanging 'sending' status rows back to 'pending'
    state.campaignLogs.forEach(l => {
      if (l.status === 'sending') l.status = 'pending';
    });

    btnStart.classList.remove('hidden');
    btnPause.classList.add('hidden');
    btnResume.classList.add('hidden');
    btnStop.classList.add('hidden');
    setInputsDisabled(false);
    renderLogTable();
    updateMetrics();
    updateStartButtonState();
  }

  function pauseCampaign() {
    state.pauseRequested = true;
    state.status = 'paused';
    btnPause.classList.add('hidden');
    btnResume.classList.remove('hidden');
    updateMetrics();
  }

  function resumeCampaign() {
    state.pauseRequested = false;
    state.status = 'running';
    btnResume.classList.add('hidden');
    btnPause.classList.remove('hidden');
    updateMetrics();
  }

  function stopCampaign() {
    state.stopRequested = true;
    state.pauseRequested = false;
    state.status = 'idle'; // Reset status to idle immediately!

    // Reset any hanging 'sending' status rows back to 'pending'
    state.campaignLogs.forEach(l => {
      if (l.status === 'sending') l.status = 'pending';
    });

    btnPause.classList.add('hidden');
    btnResume.classList.add('hidden');
    btnStop.classList.add('hidden');
    btnStart.classList.remove('hidden');
    setInputsDisabled(false);
    renderLogTable();
    updateMetrics();
    updateStartButtonState();
  }

  async function interruptibleSleep(ms) {
    const step = 100;
    let elapsed = 0;
    while (elapsed < ms) {
      if (state.stopRequested) break;
      await new Promise(r => setTimeout(r, Math.min(step, ms - elapsed)));
      elapsed += step;
    }
  }

  function updateSingleRowUI(item) {
    const rowEl = document.getElementById(`row-${item.id}`);
    if (rowEl) {
      const statusCell = rowEl.children[3];
      const timeCell = rowEl.children[4];
      const errorCell = rowEl.children[5];

      if (statusCell) statusCell.innerHTML = getStatusBadge(item.status);
      if (timeCell) timeCell.textContent = item.sentTime;
      if (errorCell) {
        errorCell.innerHTML = item.error !== '-' 
          ? `<span class="error-text">${escapeHtml(item.error)}</span>` 
          : '-';
      }
    } else {
      renderLogTable();
    }
  }

  function setInputsDisabled(disabled) {
    btnSelectFile.disabled = disabled;
    mapEmail.disabled = disabled;
    mapTopic.disabled = disabled;
    mapBody.disabled = disabled;
    if (mapAttachment) mapAttachment.disabled = disabled;
    if (btnSelectGlobalAttach) btnSelectGlobalAttach.disabled = disabled;
    if (btnClearAllGlobalAttach) btnClearAllGlobalAttach.disabled = disabled;
    if (emailDelay) emailDelay.disabled = disabled;
    if (emailSignature) emailSignature.disabled = disabled;
    if (btnTestSmtp) btnTestSmtp.disabled = disabled;
    if (btnValidateList) btnValidateList.disabled = disabled;
    if (chkValidateEmails) chkValidateEmails.disabled = disabled;
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // Export Report to Excel
  btnExportReport.addEventListener('click', async () => {
    if (state.campaignLogs.length === 0) return;

    const exportData = state.campaignLogs.map(l => ({
      'ID': l.id,
      'Recipient Email': l.recipientEmail,
      'Topic / Subject': l.topic,
      'Attachment Path': l.attachmentPath || '-',
      'Status': l.status,
      'Time': l.sentTime,
      'Error Details': l.error
    }));

    const res = await window.electronAPI.exportReport(exportData);
    if (res.success) {
      alert(`Report exported successfully to:\n${res.filePath}`);
    } else if (!res.canceled) {
      alert(`Failed to export report: ${res.error}`);
    }
  });

  // Reset Statuses Button Event
  if (btnResetLogs) {
    btnResetLogs.addEventListener('click', () => {
      if (state.status !== 'idle') return;
      if (confirm('Reset all recipient statuses back to pending?')) {
        state.campaignLogs.forEach(item => {
          item.status = 'pending';
          item.sentTime = '-';
          item.error = '-';
        });
        renderLogTable();
        updateMetrics();
        updateStartButtonState();
      }
    });
  }

  // List Verification Feature (Pre-flight Scanner)
  if (btnValidateList) {
    btnValidateList.addEventListener('click', async () => {
      if (state.campaignLogs.length === 0 || state.status !== 'idle') return;

      btnValidateList.disabled = true;
      const originalHtml = btnValidateList.innerHTML;
      btnValidateList.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="animation: spin 1s linear infinite;"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg> Verifying...`;

      try {
        const res = await window.electronAPI.validateEmailBatch(state.campaignLogs);
        if (res.success && Array.isArray(res.results)) {
          let invalidCount = 0;
          res.results.forEach(r => {
            const item = state.campaignLogs[r.index];
            if (item) {
              if (!r.isValid || !r.isActive) {
                item.status = 'invalid';
                item.error = `Skipped (Quota saved): ${r.reason}`;
                invalidCount++;
              } else if (item.status === 'invalid') {
                item.status = 'pending';
                item.error = '-';
              }
            }
          });

          renderLogTable();
          updateMetrics();
          updateStartButtonState();

          if (validationSummaryBanner) {
            valStatValid.textContent = `${res.validCount} Active & Valid`;
            valStatInvalid.textContent = `${invalidCount} Inactive/Invalid`;
            validationSummaryBanner.classList.remove('hidden');
          }
        } else {
          alert('Email verification failed: ' + (res.error || 'Unknown error'));
        }
      } catch (err) {
        alert('Verification error: ' + err.message);
      } finally {
        btnValidateList.disabled = false;
        btnValidateList.innerHTML = originalHtml;
      }
    });
  }

  if (btnCloseValSummary) {
    btnCloseValSummary.addEventListener('click', () => {
      if (validationSummaryBanner) validationSummaryBanner.classList.add('hidden');
    });
  }

  if (btnFilterOutInvalid) {
    btnFilterOutInvalid.addEventListener('click', () => {
      const beforeCount = state.campaignLogs.length;
      state.campaignLogs = state.campaignLogs.filter(l => l.status !== 'invalid');
      const removedCount = beforeCount - state.campaignLogs.length;

      // Re-index remaining rows
      state.campaignLogs.forEach((item, i) => {
        item.id = i + 1;
      });

      renderLogTable();
      updateMetrics();
      updateStartButtonState();

      if (validationSummaryBanner) validationSummaryBanner.classList.add('hidden');
      alert(`Removed ${removedCount} invalid/inactive recipient(s) from campaign. ${state.campaignLogs.length} active recipients ready!`);
    });
  }

  // ===========================================================================
  // ===========================================================================
  // Schedule Campaign Feature (Multi-Day Batch Support + Windows Task Scheduler)
  // ===========================================================================
  const scheduleDateTime        = document.getElementById('scheduleDateTime');
  const btnSchedule             = document.getElementById('btnSchedule');
  const btnCancelSchedule       = document.getElementById('btnCancelSchedule');
  const scheduleCountdownBox    = document.getElementById('scheduleCountdownBox');
  const scheduleCountdownLabel  = document.getElementById('scheduleCountdownLabel');
  const scheduleCountdownTimer  = document.getElementById('scheduleCountdownTimer');
  const scheduleTargetTime      = document.getElementById('scheduleTargetTime');
  const scheduleBadge           = document.getElementById('scheduleBadge');
  const scheduleCard            = document.getElementById('scheduleCard');

  // Batch Configuration Elements
  const chkDailyBatches         = document.getElementById('chkDailyBatches');
  const inputDailyLimit         = document.getElementById('inputDailyLimit');
  const chkWeekdaysOnly         = document.getElementById('chkWeekdaysOnly');
  const batchOptionsSubRow      = document.getElementById('batchOptionsSubRow');
  const batchPreviewBox         = document.getElementById('batchPreviewBox');
  const batchPreviewTitle       = document.getElementById('batchPreviewTitle');
  const batchPreviewList        = document.getElementById('batchPreviewList');
  const activeBatchProgressBox  = document.getElementById('activeBatchProgressBox');
  const activeBatchProgressBadge= document.getElementById('activeBatchProgressBadge');
  const activeBatchTimeline     = document.getElementById('activeBatchTimeline');

  // scheduleTimeoutId declared at top of scope
  let countdownIntervalId = null;

  // ── helpers ────────────────────────────────────────────────────────────────

  function addWeekdays(startDate, numDays, weekdaysOnly = true) {
    let cur = new Date(startDate);
    let added = 0;
    while (added < numDays) {
      cur.setDate(cur.getDate() + 1);
      const day = cur.getDay(); // 0: Sun, 6: Sat
      if (!weekdaysOnly || (day !== 0 && day !== 6)) {
        added++;
      }
    }
    return cur;
  }

  function generateBatches(startDate, logs, dailyLimit, weekdaysOnly) {
    // Only schedule active/deliverable emails (exclude invalid addresses to protect quota)
    const activeLogs = logs.filter(l => l.status !== 'invalid');
    const total = activeLogs.length;
    const numBatches = Math.ceil(total / dailyLimit) || 1;
    const batches = [];
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    for (let i = 0; i < numBatches; i++) {
      const startIdx = i * dailyLimit;
      const endIdx = Math.min(startIdx + dailyLimit, total);
      const items = activeLogs.slice(startIdx, endIdx);
      if (items.length === 0 && i > 0) continue;
      const batchDate = addWeekdays(startDate, i, weekdaysOnly);
      const dayLabel = `${dayNames[batchDate.getDay()]}, ${monthNames[batchDate.getMonth()]} ${batchDate.getDate()}`;

      batches.push({
        batchNumber: i + 1,
        dayLabel: dayLabel,
        isoString: batchDate.toISOString(),
        status: 'pending',
        sentAt: null,
        sentCount: 0,
        failCount: 0,
        items: items
      });
    }
    return batches;
  }

  function updateBatchPreview() {
    if (!batchPreviewBox || !chkDailyBatches) return;

    if (!chkDailyBatches.checked) {
      batchPreviewBox.classList.add('hidden');
      if (batchOptionsSubRow) batchOptionsSubRow.style.opacity = '0.4';
      return;
    }

    if (batchOptionsSubRow) batchOptionsSubRow.style.opacity = '1';

    const val = scheduleDateTime.value;
    const total = state.campaignLogs.filter(l => l.status !== 'invalid').length;
    const limit = parseInt(inputDailyLimit.value, 10) || 500;
    const weekdays = chkWeekdaysOnly.checked;

    if (!val || total <= 0) {
      batchPreviewBox.classList.add('hidden');
      return;
    }

    const startDate = new Date(val);
    const batches = generateBatches(startDate, state.campaignLogs, limit, weekdays);

    if (batches.length <= 1) {
      batchPreviewBox.classList.add('hidden');
      return;
    }

    batchPreviewBox.classList.remove('hidden');
    batchPreviewTitle.textContent = `${total.toLocaleString()} active emails split across ${batches.length} daily batches (${limit}/day)`;

    batchPreviewList.innerHTML = '';
    batches.forEach(b => {
      const itemEl = document.createElement('div');
      itemEl.className = 'batch-preview-item';
      const timeStr = new Date(b.isoString).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      itemEl.innerHTML = `
        <div>
          <span class="batch-num">Day ${b.batchNumber}:</span>
          <span class="batch-day">${b.dayLabel} at ${timeStr}</span>
        </div>
        <span class="batch-cnt">${b.items.length} emails</span>
      `;
      batchPreviewList.appendChild(itemEl);
    });
  }

  function renderActiveBatchTimeline(batches) {
    if (!activeBatchTimeline || !batches || batches.length === 0) return;

    activeBatchProgressBox.classList.remove('hidden');
    const completedCount = batches.filter(b => b.status === 'completed').length;
    activeBatchProgressBadge.textContent = `${completedCount} of ${batches.length} Batches Completed`;

    activeBatchTimeline.innerHTML = '';
    let foundNext = false;

    batches.forEach(b => {
      const row = document.createElement('div');
      row.className = 'active-batch-row';
      const timeStr = new Date(b.isoString).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

      let tagClass = 'pending';
      let tagText = 'Scheduled';

      if (b.status === 'completed') {
        row.classList.add('completed');
        tagClass = 'completed';
        tagText = `Sent (${b.sentCount || b.items.length})`;
      } else if (!foundNext) {
        foundNext = true;
        row.classList.add('current');
        tagClass = 'next';
        tagText = 'Next Up';
      }

      row.innerHTML = `
        <div class="active-batch-left">
          <div class="active-batch-indicator"></div>
          <div>
            <div class="active-batch-name">Batch #${b.batchNumber} — ${b.dayLabel}</div>
            <div class="active-batch-subtext">${timeStr} • ${b.items.length} recipients</div>
          </div>
        </div>
        <div class="active-batch-right">
          <span class="batch-tag ${tagClass}">${tagText}</span>
        </div>
      `;
      activeBatchTimeline.appendChild(row);
    });
  }

  function refreshScheduleMin() {
    const now = new Date(Date.now() + 60000);
    const pad = n => String(n).padStart(2, '0');
    scheduleDateTime.min =
      `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
  }

  function startCountdownUI(isoString) {
    if (countdownIntervalId) clearInterval(countdownIntervalId);
    function tick() {
      const remaining = new Date(isoString) - new Date();
      if (remaining <= 0) {
        scheduleCountdownTimer.textContent = '00:00:00';
        return;
      }
      const s = Math.floor(remaining / 1000);
      scheduleCountdownTimer.textContent =
        `${String(Math.floor(s/3600)).padStart(2,'0')}:${String(Math.floor((s%3600)/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`;
    }
    tick();
    countdownIntervalId = setInterval(tick, 1000);
  }

  function applyScheduledUI(isoString, savedData) {
    const d = new Date(isoString);
    scheduleCountdownBox.classList.remove('hidden');
    scheduleBadge.classList.remove('hidden');
    btnSchedule.classList.add('hidden');
    btnCancelSchedule.classList.remove('hidden');
    scheduleDateTime.disabled = true;
    if (chkDailyBatches) chkDailyBatches.disabled = true;
    if (inputDailyLimit) inputDailyLimit.disabled = true;
    if (chkWeekdaysOnly) chkWeekdaysOnly.disabled = true;
    btnStart.disabled = true;

    if (batchPreviewBox) batchPreviewBox.classList.add('hidden');

    if (savedData && Array.isArray(savedData.batches) && savedData.batches.length > 0) {
      const pendingBatch = savedData.batches.find(b => b.status === 'pending');
      if (pendingBatch) {
        scheduleCountdownLabel.textContent = `Batch #${pendingBatch.batchNumber} of ${savedData.batches.length} fires in`;
        scheduleTargetTime.textContent = `Next Run: ${pendingBatch.dayLabel} at ${new Date(pendingBatch.isoString).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})}`;
        scheduleBadge.textContent = `Batch ${pendingBatch.batchNumber}/${savedData.batches.length} Active`;
      } else {
        scheduleCountdownLabel.textContent = `All ${savedData.batches.length} batches complete!`;
        scheduleBadge.textContent = `Campaign Completed`;
      }
      renderActiveBatchTimeline(savedData.batches);
    } else {
      scheduleTargetTime.textContent = `Scheduled for: ${d.toLocaleString()}`;
      scheduleBadge.textContent = `Scheduled`;
      if (activeBatchProgressBox) activeBatchProgressBox.classList.add('hidden');
    }

    if (scheduleDateTime.value !== isoString.slice(0,16)) {
      scheduleDateTime.value = isoString.slice(0,16);
    }
  }

  function armScheduleTimeout(isoString, msUntil) {
    if (scheduleTimeoutId) clearTimeout(scheduleTimeoutId);
    scheduleTimeoutId = setTimeout(async () => {
      clearInterval(countdownIntervalId);
      countdownIntervalId = null;
      scheduleTimeoutId   = null;

      // Allow background sender to execute and check updated schedule
      setTimeout(async () => {
        const res = await window.electronAPI.loadSchedule();
        if (res.success && res.data) {
          const nextPending = res.data.batches && res.data.batches.find(b => b.status === 'pending');
          if (nextPending) {
            const nextMs = new Date(nextPending.isoString) - Date.now();
            if (nextMs > 0) {
              armScheduleTimeout(nextPending.isoString, nextMs);
              startCountdownUI(nextPending.isoString);
              applyScheduledUI(nextPending.isoString, res.data);
              return;
            }
          }
        }
        resetScheduleUI();
      }, 5000);
    }, msUntil);
  }

  function resetScheduleUI() {
    if (countdownIntervalId) { clearInterval(countdownIntervalId); countdownIntervalId = null; }
    if (scheduleTimeoutId)   { clearTimeout(scheduleTimeoutId);   scheduleTimeoutId = null; }
    scheduleCountdownBox.classList.add('hidden');
    scheduleBadge.classList.add('hidden');
    btnCancelSchedule.classList.add('hidden');
    btnSchedule.classList.remove('hidden');
    scheduleDateTime.disabled = false;
    if (chkDailyBatches) chkDailyBatches.disabled = false;
    if (inputDailyLimit) inputDailyLimit.disabled = false;
    if (chkWeekdaysOnly) chkWeekdaysOnly.disabled = false;
    if (activeBatchProgressBox) activeBatchProgressBox.classList.add('hidden');
    scheduleDateTime.value    = '';
    btnSchedule.disabled      = true;
    refreshScheduleMin();
    updateBatchPreview();
  }

  // ── Input & Batch change listeners ─────────────────────────────────────────
  if (scheduleDateTime) {
    refreshScheduleMin();
    scheduleDateTime.addEventListener('input', () => {
      const val = scheduleDateTime.value;
      btnSchedule.disabled = !val || (new Date(val) <= new Date()) || state.campaignLogs.length === 0;
      updateBatchPreview();
    });
  }

  if (chkDailyBatches) {
    chkDailyBatches.addEventListener('change', updateBatchPreview);
  }
  if (inputDailyLimit) {
    inputDailyLimit.addEventListener('input', updateBatchPreview);
  }
  if (chkWeekdaysOnly) {
    chkWeekdaysOnly.addEventListener('change', updateBatchPreview);
  }

  // ── Schedule button ─────────────────────────────────────────────────────────
  if (btnSchedule) {
    btnSchedule.addEventListener('click', async () => {
      const val = scheduleDateTime.value;
      if (!val) { alert('Please pick a future date and time first.'); return; }
      const targetDate = new Date(val);
      const msUntil = targetDate - Date.now();
      if (msUntil <= 0) { alert('Please choose a time in the future!'); return; }
      if (state.campaignLogs.length === 0) { alert('Please load your Excel file and configure column mappings first.'); return; }

      const smtpConfig = getSmtpConfig();
      if (!smtpConfig.host || !smtpConfig.user || !smtpConfig.pass) {
        alert('Please fill out your SMTP Server, User, and Password credentials first!');
        return;
      }

      btnSchedule.disabled = true;

      const dailyLimit = parseInt(inputDailyLimit.value, 10) || 500;
      const weekdays = chkWeekdaysOnly.checked;
      const useBatches = chkDailyBatches.checked && state.campaignLogs.length > dailyLimit;

      let batches = null;
      let firstIso = targetDate.toISOString();

      if (useBatches) {
        batches = generateBatches(targetDate, state.campaignLogs, dailyLimit, weekdays);
        firstIso = batches[0].isoString;
      }

      // 1. Build payload to persist
      const payload = {
        isoString:    firstIso,
        smtpConfig:   smtpConfig,
        campaignLogs: JSON.parse(JSON.stringify(state.campaignLogs)),
        emailDelay:   parseInt(emailDelay.value, 10) || 0,
        dailyLimit:   dailyLimit,
        weekdaysOnly: weekdays,
        signature:    emailSignature ? emailSignature.value : '',
        savedAt:      new Date().toISOString(),
        batches:      batches
      };

      await window.electronAPI.saveSchedule(payload);

      // 2. Register Windows Task Scheduler background task for the first batch
      const osTaskRes = await window.electronAPI.createOsTask({ isoString: firstIso });

      // 3. Setup in-app timer (if app stays open)
      armScheduleTimeout(firstIso, msUntil);
      startCountdownUI(firstIso);
      applyScheduledUI(firstIso, payload);

      if (useBatches) {
        showScheduleToast(`⚡ Multi-day campaign scheduled! ${batches.length} batches created (${dailyLimit}/day max). Batch #1 set for ${targetDate.toLocaleString()}.`);
      } else if (osTaskRes && osTaskRes.success) {
        showScheduleToast(`⚡ Scheduled in Windows! Emails will send on ${targetDate.toLocaleString()} even if app is closed.`);
      } else {
        showScheduleToast(`Scheduled for ${targetDate.toLocaleString()} (In-app timer active).`);
      }
    });
  }

  // ── Cancel button ────────────────────────────────────────────────────────────
  if (btnCancelSchedule) {
    btnCancelSchedule.addEventListener('click', async () => {
      if (scheduleTimeoutId) { clearTimeout(scheduleTimeoutId); scheduleTimeoutId = null; }
      await window.electronAPI.deleteOsTask();
      await window.electronAPI.clearSchedule();
      resetScheduleUI();
      updateStartButtonState();
      showScheduleToast('Schedule cancelled and Windows background task removed.');
    });
  }

  // ── On startup: restore background logs or pending schedules ───────────────
  (async () => {
    // 1. Check if a background campaign completed while app was closed
    const logResult = await window.electronAPI.getLatestLog();
    if (logResult && logResult.success && logResult.log && Array.isArray(logResult.log.data)) {
      const logData = logResult.log.data;
      if (state.campaignLogs.length === 0) {
        state.campaignLogs = logData;
        renderLogTable();
        updateMetrics();

        const sent = logData.filter(l => l.status === 'sent').length;
        const failed = logData.filter(l => l.status === 'failed').length;

        const successBanner = document.createElement('div');
        successBanner.className = 'missed-schedule-banner';
        successBanner.style.background = 'linear-gradient(135deg, rgba(16, 185, 129, 0.15), rgba(5, 150, 105, 0.1))';
        successBanner.style.borderColor = 'rgba(16, 185, 129, 0.4)';
        successBanner.style.color = '#34d399';
        successBanner.innerHTML = `
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:#10b981;">
            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
          </svg>
          <span>
            <strong>Background Campaign Log Loaded:</strong> Activity results displayed (${sent} sent, ${failed} failed).
          </span>
          <button class="missed-banner-close" title="Dismiss" style="color:#34d399;">✕</button>
        `;
        successBanner.querySelector('.missed-banner-close').addEventListener('click', () => {
          successBanner.remove();
        });
        if (scheduleCard) scheduleCard.parentNode.insertBefore(successBanner, scheduleCard);
      }
    }

    // 2. Check pending schedule
    const result = await window.electronAPI.loadSchedule();
    if (!result.success || !result.data) return;

    const saved = result.data;
    if (scheduleCard) {
      scheduleCard.style.borderColor = 'rgba(139,92,246,0.4)';
    }

    // Check if multi-batch
    if (Array.isArray(saved.batches) && saved.batches.length > 0) {
      if (saved.campaignLogs && saved.campaignLogs.length > 0 && state.campaignLogs.length === 0) {
        state.campaignLogs = saved.campaignLogs;
        renderLogTable();
        updateMetrics();
      }

      const nextPending = saved.batches.find(b => b.status === 'pending');
      if (nextPending) {
        const targetDate = new Date(nextPending.isoString);
        const msUntil = targetDate - Date.now();

        armScheduleTimeout(nextPending.isoString, Math.max(0, msUntil));
        startCountdownUI(nextPending.isoString);
        applyScheduledUI(nextPending.isoString, saved);

        showScheduleToast(`⚡ Multi-day schedule active: Batch #${nextPending.batchNumber} (${nextPending.dayLabel}) fires next.`);
      } else {
        // All batches completed
        await window.electronAPI.deleteOsTask();
        await window.electronAPI.clearSchedule();
      }
      return;
    }

    // Legacy single schedule
    const targetDate = new Date(saved.isoString);
    const msUntil    = targetDate - Date.now();

    if (msUntil <= 0) {
      await window.electronAPI.deleteOsTask();
      await window.electronAPI.clearSchedule();
      return;
    }

    if (saved.campaignLogs && saved.campaignLogs.length > 0 && state.campaignLogs.length === 0) {
      state.campaignLogs = saved.campaignLogs;
      renderLogTable();
      updateMetrics();
    }

    armScheduleTimeout(saved.isoString, msUntil);
    startCountdownUI(saved.isoString);
    applyScheduledUI(saved.isoString, saved);

    showScheduleToast(`⚡ Scheduled in Windows Task Scheduler: Campaign fires at ${targetDate.toLocaleString()}`);
  })();

  // ── Screenshot Capture Mode Setup ─────────────────────────────────────────
  if (window.location.search.includes('screenshot=true')) {
    (async () => {
      if (smtpHost) smtpHost.value = 'smtp.gmail.com';
      if (smtpPort) smtpPort.value = '587';
      if (smtpUser) smtpUser.value = 'nhitesh.bohra@gmail.com';
      if (smtpFromName) smtpFromName.value = 'Hitesh Bohra';
      if (smtpPass) smtpPass.value = 'abcdefghijklmnop';
      if (emailSignature) emailSignature.value = '--\nBest regards,\nHitesh Bohra\n+91 9876543210 | [LinkedIn](https://linkedin.com/in/nhiteshbohra) | [GitHub](https://github.com/nhiteshbohra)';
      
      const badge = document.getElementById('smtpStatusBadge');
      if (badge) {
        badge.className = 'status-indicator connected';
        badge.querySelector('.text').textContent = 'SMTP: Verified';
      }

      if (chkJitterDelay && jitterRangeBox) {
        chkJitterDelay.checked = true;
        jitterRangeBox.classList.remove('hidden');
      }

      // Load sample file if available
      try {
        const samplePath = 'f:\\projects\\bulkmail\\sample_contacts.xlsx';
        const res = await window.electronAPI.parseFile(samplePath);
        if (res && res.success) {
          state.excelData = res;
          if (fileInfoBox && fileNameText && fileRowsCount) {
            fileNameText.textContent = res.fileName;
            fileRowsCount.textContent = `${res.totalRows} rows loaded`;
            fileInfoBox.classList.remove('hidden');
            if (dropZone) dropZone.classList.add('hidden');
          }
          if (mappingSection) mappingSection.classList.remove('hidden');
          populateColumnSelects(res.headers);
          
          // Auto map columns
          if (mapEmail) mapEmail.value = res.headers.find(h => /email/i.test(h)) || res.headers[0];
          if (mapTopic) mapTopic.value = res.headers.find(h => /topic|subject/i.test(h)) || res.headers[1] || res.headers[0];
          if (mapBody) mapBody.value = res.headers.find(h => /body|message/i.test(h)) || res.headers[2] || res.headers[0];
          generateLogsFromMapping();

          // Mark some logs as sent for beautiful UI counters
          if (state.campaignLogs.length >= 10) {
            for (let i = 0; i < 8; i++) {
              state.campaignLogs[i].status = 'sent';
              state.campaignLogs[i].sentTime = new Date().toLocaleTimeString();
              state.campaignLogs[i].error = '-';
            }
            state.campaignLogs[8].status = 'failed';
            state.campaignLogs[8].sentTime = new Date().toLocaleTimeString();
            state.campaignLogs[8].error = 'Recipient mailbox full (552)';
            renderLogTable();
            updateMetrics();
          }
        }
      } catch (e) { console.error('Screenshot pre-load error:', e); }
    })();
  }


  function showScheduleToast(message) {
    const toast = document.createElement('div');
    toast.className = 'schedule-toast';
    toast.textContent = message;
    document.body.appendChild(toast);
    requestAnimationFrame(() => { toast.classList.add('visible'); });
    setTimeout(() => {
      toast.classList.remove('visible');
      setTimeout(() => toast.remove(), 400);
    }, 5500);
  }

  // ===========================================================================
  // Phase 1 & Phase 2 Implementations: Jitter, Test Email, Preview, Templates, History
  // ===========================================================================

  // ── 1. Jitter Delay Controls ────────────────────────────────────────────────
  const chkJitterDelay = document.getElementById('chkJitterDelay');
  const jitterRangeBox = document.getElementById('jitterRangeBox');
  const jitterMin = document.getElementById('jitterMin');
  const jitterMax = document.getElementById('jitterMax');

  if (chkJitterDelay && jitterRangeBox) {
    chkJitterDelay.addEventListener('change', () => {
      if (chkJitterDelay.checked) {
        jitterRangeBox.classList.remove('hidden');
      } else {
        jitterRangeBox.classList.add('hidden');
      }
    });
  }

  // ── 2. Send Test Mail Modal ──────────────────────────────────────────────────
  const btnOpenTestMailModal = document.getElementById('btnOpenTestMailModal');
  const modalSendTest = document.getElementById('modalSendTest');
  const btnCloseTestMailModal = document.getElementById('btnCloseTestMailModal');
  const btnCancelTestMail = document.getElementById('btnCancelTestMail');
  const btnSubmitSendTest = document.getElementById('btnSubmitSendTest');
  const testMailRecipient = document.getElementById('testMailRecipient');
  const testMailSubject = document.getElementById('testMailSubject');
  const testMailFeedback = document.getElementById('testMailFeedback');

  function openTestMailModal() {
    if (modalSendTest) {
      testMailFeedback.classList.add('hidden');
      if (smtpUser && smtpUser.value) testMailRecipient.value = smtpUser.value;
      modalSendTest.classList.remove('hidden');
    }
  }

  function closeTestMailModal() {
    if (modalSendTest) modalSendTest.classList.add('hidden');
  }

  if (btnOpenTestMailModal) btnOpenTestMailModal.addEventListener('click', openTestMailModal);
  if (btnCloseTestMailModal) btnCloseTestMailModal.addEventListener('click', closeTestMailModal);
  if (btnCancelTestMail) btnCancelTestMail.addEventListener('click', closeTestMailModal);

  if (btnSubmitSendTest) {
    btnSubmitSendTest.addEventListener('click', async () => {
      const recipient = testMailRecipient.value.trim();
      if (!recipient) {
        testMailFeedback.textContent = 'Please enter a valid recipient email address.';
        testMailFeedback.className = 'alert-box alert-error';
        testMailFeedback.classList.remove('hidden');
        return;
      }

      const smtpConfig = getSmtpConfig();
      if (!smtpConfig.host || !smtpConfig.user || !smtpConfig.pass) {
        testMailFeedback.textContent = 'Please fill out your SMTP Server, User, and Password first.';
        testMailFeedback.className = 'alert-box alert-error';
        testMailFeedback.classList.remove('hidden');
        return;
      }

      btnSubmitSendTest.disabled = true;
      btnSubmitSendTest.textContent = 'Sending...';

      let sampleBody = 'This is a test email sent from AutoMail PRO to verify your SMTP connection and formatting.';
      if (state.campaignLogs && state.campaignLogs.length > 0) {
        sampleBody = state.campaignLogs[0].body || sampleBody;
      }
      const signatureText = emailSignature ? emailSignature.value : '';

      const testPayload = {
        smtpConfig,
        mailData: {
          to: recipient,
          subject: testMailSubject.value.trim() || '[TEST] AutoMail PRO Verification',
          body: sampleBody + (signatureText ? '\n\n' + signatureText : '')
        }
      };

      const res = await window.electronAPI.sendSingleEmail(testPayload);
      btnSubmitSendTest.disabled = false;
      btnSubmitSendTest.textContent = 'Send Test Now';

      if (res.success) {
        testMailFeedback.textContent = '✓ Test email sent successfully to ' + recipient + '!';
        testMailFeedback.className = 'alert-box alert-success';
        testMailFeedback.classList.remove('hidden');
      } else {
        testMailFeedback.textContent = '❌ Failed: ' + res.error;
        testMailFeedback.className = 'alert-box alert-error';
        testMailFeedback.classList.remove('hidden');
      }
    });
  }

  // ── 3. Live Email Preview Modal ──────────────────────────────────────────────
  const btnPreviewEmail = document.getElementById('btnPreviewEmail');
  const modalEmailPreview = document.getElementById('modalEmailPreview');
  const btnClosePreviewModal = document.getElementById('btnClosePreviewModal');
  const btnClosePreviewModal2 = document.getElementById('btnClosePreviewModal2');
  const btnPrevPreviewRow = document.getElementById('btnPrevPreviewRow');
  const btnNextPreviewRow = document.getElementById('btnNextPreviewRow');
  const previewRowText = document.getElementById('previewRowText');

  const previewFrom = document.getElementById('previewFrom');
  const previewTo = document.getElementById('previewTo');
  const previewSubject = document.getElementById('previewSubject');
  const previewBodyContainer = document.getElementById('previewBodyContainer');
  const previewCcRow = document.getElementById('previewCcRow');
  const previewCc = document.getElementById('previewCc');
  const previewAttachRow = document.getElementById('previewAttachRow');
  const previewAttachments = document.getElementById('previewAttachments');

  let currentPreviewRowIndex = 0;

  function renderPreviewRow(index) {
    if (!state.campaignLogs || state.campaignLogs.length === 0) return;
    if (index < 0) index = 0;
    if (index >= state.campaignLogs.length) index = state.campaignLogs.length - 1;
    currentPreviewRowIndex = index;

    const item = state.campaignLogs[index];
    previewRowText.textContent = `Row ${index + 1} of ${state.campaignLogs.length}`;

    const senderName = smtpFromName.value.trim();
    const senderEmail = smtpUser.value.trim() || 'user@domain.com';
    previewFrom.textContent = senderName ? `${senderName} <${senderEmail}>` : senderEmail;

    previewTo.textContent = item.recipientEmail || 'recipient@domain.com';
    previewSubject.textContent = item.topic || '(No Subject)';

    if (item.cc && String(item.cc).trim()) {
      previewCc.textContent = item.cc;
      previewCcRow.classList.remove('hidden');
    } else {
      previewCcRow.classList.add('hidden');
    }

    if (item.attachmentPath || state.globalAttachments.length > 0) {
      const globalNames = state.globalAttachments.map(p => p.split(/[\\/]/).pop());
      const itemAttach = item.attachmentPath ? [item.attachmentPath] : [];
      const allAttach = [...itemAttach, ...globalNames].join(', ');
      previewAttachments.textContent = allAttach;
      previewAttachRow.classList.remove('hidden');
    } else {
      previewAttachRow.classList.add('hidden');
    }

    const sigText = emailSignature ? emailSignature.value : '';
    let bodyText = item.body || '';
    if (sigText.trim()) {
      bodyText += '\n\n' + sigText;
    }
    previewBodyContainer.textContent = bodyText;
  }

  if (btnPreviewEmail) {
    btnPreviewEmail.addEventListener('click', () => {
      if (state.campaignLogs.length === 0) {
        alert('Please load an Excel file first.');
        return;
      }
      renderPreviewRow(0);
      modalEmailPreview.classList.remove('hidden');
    });
  }

  if (btnClosePreviewModal) btnClosePreviewModal.addEventListener('click', () => modalEmailPreview.classList.add('hidden'));
  if (btnClosePreviewModal2) btnClosePreviewModal2.addEventListener('click', () => modalEmailPreview.classList.add('hidden'));

  if (btnPrevPreviewRow) {
    btnPrevPreviewRow.addEventListener('click', () => {
      renderPreviewRow(currentPreviewRowIndex - 1);
    });
  }

  if (btnNextPreviewRow) {
    btnNextPreviewRow.addEventListener('click', () => {
      renderPreviewRow(currentPreviewRowIndex + 1);
    });
  }

  // Update preview button state whenever campaignLogs update
  const origUpdateStartButtonState = updateStartButtonState;
  updateStartButtonState = function() {
    origUpdateStartButtonState();
    if (btnPreviewEmail) {
      btnPreviewEmail.disabled = state.campaignLogs.length === 0;
    }
  };

  // ── 4. Email Templates System ────────────────────────────────────────────────
  const templateSelect = document.getElementById('templateSelect');
  const btnOpenSaveTemplateModal = document.getElementById('btnOpenSaveTemplateModal');
  const btnDeleteTemplate = document.getElementById('btnDeleteTemplate');
  const modalSaveTemplate = document.getElementById('modalSaveTemplate');
  const btnCloseSaveTemplateModal = document.getElementById('btnCloseSaveTemplateModal');
  const btnCancelSaveTemplate = document.getElementById('btnCancelSaveTemplate');
  const btnConfirmSaveTemplate = document.getElementById('btnConfirmSaveTemplate');
  const templateNameInput = document.getElementById('templateNameInput');

  let savedTemplatesList = [];

  async function loadTemplatesDropdown() {
    if (!templateSelect) return;
    const res = await window.electronAPI.loadTemplates();
    if (res.success && Array.isArray(res.templates)) {
      savedTemplatesList = res.templates;
      templateSelect.innerHTML = '<option value="">-- Custom / Default --</option>';
      res.templates.forEach(tpl => {
        const opt = document.createElement('option');
        opt.value = tpl.id;
        opt.textContent = tpl.name;
        templateSelect.appendChild(opt);
      });
    }
  }
  loadTemplatesDropdown();

  if (templateSelect) {
    templateSelect.addEventListener('change', () => {
      const selectedId = templateSelect.value;
      if (!selectedId) {
        if (btnDeleteTemplate) btnDeleteTemplate.classList.add('hidden');
        return;
      }
      if (btnDeleteTemplate) btnDeleteTemplate.classList.remove('hidden');
      const tpl = savedTemplatesList.find(t => t.id === selectedId);
      if (tpl && tpl.mappings) {
        if (tpl.mappings.email && mapEmail) mapEmail.value = tpl.mappings.email;
        if (tpl.mappings.topic && mapTopic) mapTopic.value = tpl.mappings.topic;
        if (tpl.mappings.body && mapBody) mapBody.value = tpl.mappings.body;
        if (tpl.mappings.cc && mapCc) mapCc.value = tpl.mappings.cc;
        if (tpl.mappings.bcc && mapBcc) mapBcc.value = tpl.mappings.bcc;
        if (tpl.mappings.attachment && mapAttachment) mapAttachment.value = tpl.mappings.attachment;
        if (tpl.signature && emailSignature) emailSignature.value = tpl.signature;
        generateLogsFromMapping();
      }
    });
  }

  if (btnOpenSaveTemplateModal) {
    btnOpenSaveTemplateModal.addEventListener('click', () => {
      if (templateNameInput) templateNameInput.value = '';
      if (modalSaveTemplate) modalSaveTemplate.classList.remove('hidden');
    });
  }

  if (btnCloseSaveTemplateModal) btnCloseSaveTemplateModal.addEventListener('click', () => modalSaveTemplate.classList.add('hidden'));
  if (btnCancelSaveTemplate) btnCancelSaveTemplate.addEventListener('click', () => modalSaveTemplate.classList.add('hidden'));

  if (btnConfirmSaveTemplate) {
    btnConfirmSaveTemplate.addEventListener('click', async () => {
      const name = templateNameInput.value.trim();
      if (!name) { alert('Please enter a template name.'); return; }

      const tplData = {
        name,
        mappings: {
          email: mapEmail ? mapEmail.value : '',
          topic: mapTopic ? mapTopic.value : '',
          body: mapBody ? mapBody.value : '',
          cc: mapCc ? mapCc.value : '',
          bcc: mapBcc ? mapBcc.value : '',
          attachment: mapAttachment ? mapAttachment.value : ''
        },
        signature: emailSignature ? emailSignature.value : ''
      };

      const res = await window.electronAPI.saveTemplate(tplData);
      if (res.success) {
        modalSaveTemplate.classList.add('hidden');
        await loadTemplatesDropdown();
        if (res.savedTemplate && templateSelect) {
          templateSelect.value = res.savedTemplate.id;
          if (btnDeleteTemplate) btnDeleteTemplate.classList.remove('hidden');
        }
        showScheduleToast('✓ Template saved successfully!');
      } else {
        alert('Failed to save template: ' + res.error);
      }
    });
  }

  if (btnDeleteTemplate) {
    btnDeleteTemplate.addEventListener('click', async () => {
      const selectedId = templateSelect.value;
      if (!selectedId) return;
      if (!confirm('Are you sure you want to delete this template?')) return;
      const res = await window.electronAPI.deleteTemplate(selectedId);
      if (res.success) {
        btnDeleteTemplate.classList.add('hidden');
        await loadTemplatesDropdown();
        showScheduleToast('Template deleted.');
      }
    });
  }

  // ── 5. Campaign History Drawer ─────────────────────────────────────────────
  const btnOpenHistoryDrawer = document.getElementById('btnOpenHistoryDrawer');
  const drawerHistory = document.getElementById('drawerHistory');
  const btnCloseHistoryDrawer = document.getElementById('btnCloseHistoryDrawer');
  const btnRefreshHistory = document.getElementById('btnRefreshHistory');
  const historyLogList = document.getElementById('historyLogList');
  const historyLogCountText = document.getElementById('historyLogCountText');
  const historyDetailsBox = document.getElementById('historyDetailsBox');
  const historyDetailsTitle = document.getElementById('historyDetailsTitle');
  const historyDetailsSummary = document.getElementById('historyDetailsSummary');
  const historyDetailsTableBody = document.getElementById('historyDetailsTableBody');
  const btnCloseHistoryDetails = document.getElementById('btnCloseHistoryDetails');

  async function fetchCampaignHistory() {
    if (!historyLogList) return;
    historyLogCountText.textContent = 'Loading logs...';
    historyLogList.innerHTML = '';
    const res = await window.electronAPI.listHistoryLogs();
    if (res.success && Array.isArray(res.logs)) {
      historyLogCountText.textContent = `${res.logs.length} Campaign logs found`;
      if (res.logs.length === 0) {
        historyLogList.innerHTML = '<div style="font-size:0.8rem; color:var(--text-secondary); text-align:center; padding:20px;">No past campaign logs available yet.</div>';
        return;
      }
      res.logs.forEach(log => {
        const item = document.createElement('div');
        item.className = 'history-log-item';
        item.innerHTML = `
          <div>
            <div class="history-log-title">${log.batchLabel} — <span style="font-weight:normal; font-size:0.78rem; color:var(--text-secondary);">${log.dateStr}</span></div>
            <div class="history-log-date" style="margin-top:2px;">File: <code>${log.fileName}</code></div>
          </div>
          <div style="display:flex; gap:6px; align-items:center;">
            <span class="history-stat-badge sent">${log.sent} Sent</span>
            ${log.failed > 0 ? `<span class="history-stat-badge failed">${log.failed} Failed</span>` : ''}
          </div>
        `;
        item.addEventListener('click', () => openHistoryLogDetails(log.fileName));
        historyLogList.appendChild(item);
      });
    } else {
      historyLogCountText.textContent = 'Error loading campaign history.';
    }
  }

  async function openHistoryLogDetails(fileName) {
    if (!historyDetailsBox) return;
    const res = await window.electronAPI.getHistoryDetails(fileName);
    if (res.success && Array.isArray(res.data)) {
      historyDetailsTitle.textContent = `Log Details: ${fileName}`;
      const sent = res.data.filter(i => i.status === 'sent').length;
      const failed = res.data.filter(i => i.status === 'failed').length;
      historyDetailsSummary.innerHTML = `
        <span>Total: <strong>${res.data.length}</strong></span>
        <span style="color:#34d399;">Sent: <strong>${sent}</strong></span>
        <span style="color:#f87171;">Failed: <strong>${failed}</strong></span>
      `;

      historyDetailsTableBody.innerHTML = '';
      res.data.forEach(row => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid rgba(255,255,255,0.05)';
        tr.innerHTML = `
          <td style="padding: 6px; color: var(--text-primary);">${row.recipientEmail || '-'}</td>
          <td style="padding: 6px;">
            <span class="badge ${row.status === 'sent' ? 'badge-sent' : 'badge-failed'}">${row.status || 'unknown'}</span>
          </td>
          <td style="padding: 6px; font-size: 0.7rem; color: var(--text-secondary);">${row.error || row.sentTime || '-'}</td>
        `;
        historyDetailsTableBody.appendChild(tr);
      });
      historyDetailsBox.classList.remove('hidden');
    }
  }

  if (btnOpenHistoryDrawer) {
    btnOpenHistoryDrawer.addEventListener('click', () => {
      fetchCampaignHistory();
      if (drawerHistory) drawerHistory.classList.remove('hidden');
    });
  }

  if (btnCloseHistoryDrawer) btnCloseHistoryDrawer.addEventListener('click', () => drawerHistory.classList.add('hidden'));
  if (btnRefreshHistory) btnRefreshHistory.addEventListener('click', fetchCampaignHistory);
  if (btnCloseHistoryDetails) btnCloseHistoryDetails.addEventListener('click', () => historyDetailsBox.classList.add('hidden'));
});



