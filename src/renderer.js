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
    currentFilter: 'all',
    smtpVerified: false,
    stopRequested: false,
    pauseRequested: false
  };

  // DOM Elements
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
  const mapTopic = document.getElementById('mapTopic');
  const mapBody = document.getElementById('mapBody');
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
  const logSearch = document.getElementById('logSearch');
  const filterTabs = document.querySelectorAll('.tab-btn');

  const filterCountAll = document.getElementById('filterCountAll');
  const filterCountSent = document.getElementById('filterCountSent');
  const filterCountFailed = document.getElementById('filterCountFailed');
  const filterCountPending = document.getElementById('filterCountPending');

  // Initialize Preset Listener
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

  // Test SMTP Connection Handler
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

  function getSmtpConfig() {
    return {
      host: smtpHost.value.trim(),
      port: smtpPort.value.trim(),
      secure: smtpSecure.checked,
      user: smtpUser.user ? smtpUser.user.trim() : smtpUser.value.trim(),
      pass: smtpPass.value,
      fromName: smtpFromName.value.trim()
    };
  }

  function showSmtpFeedback(isSuccess, message) {
    smtpFeedback.className = `alert-box ${isSuccess ? 'success' : 'error'}`;
    smtpFeedback.textContent = message;
    smtpFeedback.classList.remove('hidden');
  }

  function updateSmtpStatus(isVerified) {
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

  function populateMappingDropdowns(headers) {
    const selectElements = [mapEmail, mapTopic, mapBody];

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
      if (!mapTopic.value && (lower.includes('topic') || lower.includes('subject') || lower.includes('title'))) {
        mapTopic.value = h;
      }
      if (!mapBody.value && (lower.includes('body') || lower.includes('content') || lower.includes('message') || lower.includes('text'))) {
        mapBody.value = h;
      }
    });

    // Add change listeners to rebuild log preview dynamically
    mapEmail.addEventListener('change', onMappingChanged);
    mapTopic.addEventListener('change', onMappingChanged);
    mapBody.addEventListener('change', onMappingChanged);
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
    const topicCol = mapTopic.value;
    const bodyCol = mapBody.value;
    const shouldSplit = splitMultipleEmails ? splitMultipleEmails.checked : true;

    state.campaignLogs = [];
    let currentId = 1;

    state.rawRows.forEach((row, idx) => {
      let recipientRaw = emailCol ? String(row[emailCol] || '').trim() : '';
      let topicRaw = topicCol ? String(row[topicCol] || '') : '';
      let bodyRaw = bodyCol ? String(row[bodyCol] || '') : '';

      // Substitute row variables for placeholders e.g., {COLUMN_NAME}
      let finalTopic = interpolateTemplate(topicRaw, row);
      let finalBody = interpolateTemplate(bodyRaw, row);

      const parsedEmails = extractEmails(recipientRaw);

      if (shouldSplit && parsedEmails.length > 1) {
        // Expand multi-email cell into separate recipient entries in campaign log table
        parsedEmails.forEach(singleEmail => {
          state.campaignLogs.push({
            id: currentId++,
            rowNumber: idx + 1,
            recipientEmail: singleEmail,
            topic: finalTopic,
            body: finalBody,
            status: 'pending', // 'pending' | 'sending' | 'sent' | 'failed'
            sentTime: '-',
            error: '-'
          });
        });
      } else if (parsedEmails.length > 1) {
        // Keep as single row with comma-separated recipients
        state.campaignLogs.push({
          id: currentId++,
          rowNumber: idx + 1,
          recipientEmail: parsedEmails.join(', '),
          topic: finalTopic,
          body: finalBody,
          status: 'pending',
          sentTime: '-',
          error: '-'
        });
      } else {
        state.campaignLogs.push({
          id: currentId++,
          rowNumber: idx + 1,
          recipientEmail: recipientRaw,
          topic: finalTopic,
          body: finalBody,
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
        <td style="font-size: 0.75rem; color: var(--text-secondary);">${item.sentTime}</td>
        <td>${item.error !== '-' ? `<span class="error-text">${escapeHtml(item.error)}</span>` : '-'}</td>
      </tr>
    `).join('');
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
      .replace(/"/g, '&quot;');
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
        await sleep(500);
      }

      // Handle Stop Request
      if (state.stopRequested) {
        state.status = 'stopped';
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

      // Dispatch Email via Electron Main Process
      const mailData = {
        to: item.recipientEmail,
        subject: item.topic,
        body: item.body,
        isHtml: false
      };

      const res = await window.electronAPI.sendEmail(smtpConfig, mailData);

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

      // Inter-email Delay
      if (delaySec > 0 && i < state.campaignLogs.length - 1 && !state.stopRequested) {
        await sleep(delaySec * 1000);
      }
    }

    // Campaign Finished or Stopped
    if (!state.stopRequested && !state.pauseRequested) {
      state.status = 'idle';
    }

    btnStart.classList.remove('hidden');
    btnPause.classList.add('hidden');
    btnResume.classList.add('hidden');
    btnStop.classList.add('hidden');
    setInputsDisabled(false);
    updateStartButtonState();
    updateMetrics();
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
    state.status = 'stopped';
    btnPause.classList.add('hidden');
    btnResume.classList.add('hidden');
    btnStop.classList.add('hidden');
    btnStart.classList.remove('hidden');
    setInputsDisabled(false);
    updateStartButtonState();
    updateMetrics();
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
    emailDelay.disabled = disabled;
    btnTestSmtp.disabled = disabled;
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

});
