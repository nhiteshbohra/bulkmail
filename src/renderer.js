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
    pauseRequested: false
  };

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
  const btnTestSmtp = document.getElementById('btnTestSmtp');
  const smtpFeedback = document.getElementById('smtpFeedback');
  const smtpStatusBadge = document.getElementById('smtpStatusBadge');

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
  const globalAttachNameText = document.getElementById('globalAttachNameText');
  const btnRemoveGlobalAttach = document.getElementById('btnRemoveGlobalAttach');
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
  const filterCountPending = document.getElementById('filterCountPending');

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
  }

  // Global Attachment Listeners
  if (btnSelectGlobalAttach) {
    btnSelectGlobalAttach.addEventListener('click', async () => {
      const selected = await window.electronAPI.selectFile();
      if (selected) {
        state.globalAttachmentPath = selected;
        globalAttachNameText.textContent = selected.split(/[\\/]/).pop();
        btnRemoveGlobalAttach.classList.remove('hidden');
        onMappingChanged();
      }
    });
  }

  if (btnRemoveGlobalAttach) {
    btnRemoveGlobalAttach.addEventListener('click', () => {
      state.globalAttachmentPath = null;
      globalAttachNameText.textContent = 'No global attachment set';
      btnRemoveGlobalAttach.classList.add('hidden');
      onMappingChanged();
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

      // Combine row attachment(s) with global master attachment if set
      let attachmentList = [];
      if (finalAttachment) {
        attachmentList.push(...finalAttachment.split(/[,;]+/).map(a => a.trim()).filter(Boolean));
      }
      if (state.globalAttachmentPath && !attachmentList.includes(state.globalAttachmentPath)) {
        attachmentList.push(state.globalAttachmentPath);
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

  function updateStartButtonState() {
    const hasData = state.campaignLogs.length > 0;
    const hasEmailMap = mapEmail.value !== '';
    const hasTopicMap = mapTopic.value !== '';
    const hasBodyMap = mapBody.value !== '';
    const hasHost = smtpHost.value.trim() !== '';

    btnStart.disabled = !(hasData && hasEmailMap && hasTopicMap && hasBodyMap && hasHost && state.status === 'idle');
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
    const pending = state.campaignLogs.filter(l => l.status === 'pending').length;

    metricTotal.textContent = total;
    metricSent.textContent = sent;
    metricFailed.textContent = failed;
    metricPending.textContent = pending;

    filterCountAll.textContent = total;
    filterCountSent.textContent = sent;
    filterCountFailed.textContent = failed;
    filterCountPending.textContent = pending;

    const processed = sent + failed;
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
      progressText.textContent = `Campaign Complete! (${sent} Sent, ${failed} Failed)`;
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

      // Skip already sent rows if re-starting
      if (item.status === 'sent') continue;

      // Handle Pause Request
      while (state.pauseRequested) {
        if (state.stopRequested) break;
        await sleep(200);
      }

      // Handle Stop Request
      if (state.stopRequested) {
        break;
      }

      // Skip invalid emails
      if (!item.recipientEmail || !item.recipientEmail.includes('@')) {
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

      // Dispatch Email via SMTP
      const mailData = {
        to: item.recipientEmail,
        cc: item.cc || '',
        bcc: item.bcc || '',
        subject: item.topic,
        body: item.body,
        attachmentPath: item.attachmentPath,
        isHtml: false
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
    emailDelay.disabled = disabled;
    if (btnTestSmtp) btnTestSmtp.disabled = disabled;
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

});
