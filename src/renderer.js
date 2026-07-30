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
  const btnResetLogs = document.getElementById('btnResetLogs');
  const logSearch = document.getElementById('logSearch');
  const filterTabs = document.querySelectorAll('.tab-btn');

  const filterCountAll = document.getElementById('filterCountAll');
  const filterCountSent = document.getElementById('filterCountSent');
  const filterCountFailed = document.getElementById('filterCountFailed');
  const filterCountPending = document.getElementById('filterCountPending');

  // Mode Switcher Logic (Gmail Web vs SMTP)
  if (modeWebBtn && modeSmtpBtn) {
    modeWebBtn.addEventListener('click', () => {
      state.sendMode = 'web';
      modeWebBtn.classList.add('active');
      modeWebBtn.style.background = 'var(--accent-blue)';
      modeWebBtn.style.color = '#fff';
      modeSmtpBtn.classList.remove('active');
      modeSmtpBtn.style.background = 'transparent';
      modeSmtpBtn.style.color = 'var(--text-secondary)';

      if (webEnginePanel) webEnginePanel.classList.remove('hidden');
      if (smtpEnginePanel) smtpEnginePanel.classList.add('hidden');
      updateStartButtonState();
    });

    modeSmtpBtn.addEventListener('click', () => {
      state.sendMode = 'smtp';
      modeSmtpBtn.classList.add('active');
      modeSmtpBtn.style.background = 'var(--accent-blue)';
      modeSmtpBtn.style.color = '#fff';
      modeWebBtn.classList.remove('active');
      modeWebBtn.style.background = 'transparent';
      modeWebBtn.style.color = 'var(--text-secondary)';

      if (smtpEnginePanel) smtpEnginePanel.classList.remove('hidden');
      if (webEnginePanel) webEnginePanel.classList.add('hidden');
      updateStartButtonState();
    });
  }

  // Google Sign In Handler
  if (btnGoogleSignIn) {
    btnGoogleSignIn.addEventListener('click', async () => {
      btnGoogleSignIn.disabled = true;
      btnGoogleSignIn.textContent = 'Opening Chrome...';
      const res = await window.electronAPI.launchGoogleLogin();
      btnGoogleSignIn.disabled = false;
      btnGoogleSignIn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" style="margin-right: 6px;"><path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z"/><path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z"/></svg> Sign In with Google (Chrome Profile)`;
      
      if (res.success) {
        alert(res.message);
      } else {
        alert(`Error opening Chrome: ${res.error}`);
      }
    });
  }

  if (btnCheckGoogleAuth) {
    btnCheckGoogleAuth.addEventListener('click', checkGoogleAuthStatus);
  }

  async function checkGoogleAuthStatus() {
    if (!btnCheckGoogleAuth) return;
    btnCheckGoogleAuth.disabled = true;
    btnCheckGoogleAuth.textContent = 'Checking...';
    
    const res = await window.electronAPI.checkGoogleAuth();
    btnCheckGoogleAuth.disabled = false;
    btnCheckGoogleAuth.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M23 4v6h-6"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg> Verify Sign-In`;

    state.googleSignedIn = res.signedIn;
    if (googleAuthBadge) {
      if (res.signedIn) {
        googleAuthBadge.className = 'status-indicator connected';
        googleAuthBadge.querySelector('.text').textContent = 'Signed In';
      } else {
        googleAuthBadge.className = 'status-indicator disconnected';
        googleAuthBadge.querySelector('.text').textContent = 'Not Signed In';
      }
    }
    updateStartButtonState();
  }

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

    let isReady = false;
    if (state.sendMode === 'web') {
      isReady = hasData && hasEmailMap && hasTopicMap && hasBodyMap && state.status === 'idle';
    } else {
      const hasHost = smtpHost.value.trim() !== '';
      isReady = hasData && hasEmailMap && hasTopicMap && hasBodyMap && hasHost && state.status === 'idle';
    }

    btnStart.disabled = !isReady;
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
    if (state.sendMode === 'smtp') {
      const smtpConfig = getSmtpConfig();
      if (!smtpConfig.host || !smtpConfig.user || !smtpConfig.pass) {
        alert('Please fill out your SMTP Server, User, and Password credentials first!');
        return;
      }
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
    const smtpConfig = getSmtpConfig();

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

      // Dispatch Email depending on selected mode (Gmail Web vs SMTP)
      const mailData = {
        to: item.recipientEmail,
        subject: item.topic,
        body: item.body,
        isHtml: false
      };

      let res;
      if (state.sendMode === 'web') {
        res = await window.electronAPI.sendEmailWeb(mailData);
      } else {
        res = await window.electronAPI.sendEmail(smtpConfig, mailData);
      }

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
    emailDelay.disabled = disabled;
    if (btnTestSmtp) btnTestSmtp.disabled = disabled;
    if (btnGoogleSignIn) btnGoogleSignIn.disabled = disabled;
    if (btnCheckGoogleAuth) btnCheckGoogleAuth.disabled = disabled;
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
