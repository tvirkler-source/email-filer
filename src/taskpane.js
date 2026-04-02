// taskpane.js — main controller

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  await init();
});

// ── State ──────────────────────────────────────────────────────────────

let _prefs = { pinned: [], recent: [] };
let _searchDebounceTimer = null;
let _currentEmail = null;

// ── Init ───────────────────────────────────────────────────────────────

async function init() {
  try {
    await Auth.getToken();
    show('main-content');
    hide('auth-loading');

    // Load in parallel: email data, suggestions, preferences
    const [emailData, prefs] = await Promise.all([
      extractEmailData(),
      API.getPreferences().catch(() => ({ pinned: [], recent: [] })),
    ]);

    _currentEmail = emailData;
    _prefs = prefs;

    renderPinned();
    loadSuggestions(); // async, updates UI when ready
    seedFolderCache(); // async, pre-warms search cache

    bindEvents();
  } catch (err) {
    hide('auth-loading');
    show('auth-error');
    console.error('Init error:', err);
  }
}

// ── Email extraction ───────────────────────────────────────────────────

async function extractEmailData() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error('Failed to read email body'));
      }

      const bodyText = result.value || '';

      resolve({
        subject: item.subject || '(No subject)',
        senderEmail: item.from?.emailAddress || '',
        senderName: item.from?.displayName || '',
        dateReceived: item.dateTimeCreated?.toISOString() || new Date().toISOString(),
        bodyText: bodyText,
        bodySnippet: bodyText.slice(0, 500), // truncate for suggestion scoring
        // EML-compatible fields
        internetMessageId: item.internetMessageId || '',
      });
    });
  });
}

// ── Suggestions ───────────────────────────────────────────────────────

async function loadSuggestions() {
  const el = document.getElementById('suggestions-list');
  const loading = document.getElementById('suggestions-loading');

  try {
    const { suggestions } = await API.getSuggestions(
      _currentEmail.subject,
      _currentEmail.bodySnippet,
      _currentEmail.senderEmail
    );

    show('suggestions-section');
    hide(loading);

    if (!suggestions || suggestions.length === 0) {
      el.innerHTML = '<div class="empty-msg">No suggestions — use search below</div>';
      return;
    }

    el.innerHTML = '';
    suggestions.forEach(s => el.appendChild(buildFolderItem(s, 'suggestion')));

  } catch (err) {
    loading.textContent = 'Suggestions unavailable';
    console.error('Suggestions error:', err);
  }
}

// ── Folder search ──────────────────────────────────────────────────────

function bindEvents() {
  document.getElementById('folder-search').addEventListener('input', (e) => {
    clearTimeout(_searchDebounceTimer);
    const q = e.target.value.trim();

    if (q.length < 2) {
      document.getElementById('search-results').innerHTML = '';
      hide('search-spinner');
      return;
    }

    // Try cache first (instant)
    const cached = FolderCache.filterCached(q);
    if (cached !== null) {
      renderSearchResults(cached);
      return;
    }

    // Fall back to API with 200ms debounce
    _searchDebounceTimer = setTimeout(() => performSearch(q), 200);
  });

  document.getElementById('retry-auth-btn').addEventListener('click', async () => {
    Auth.clearToken();
    hide('auth-error');
    show('auth-loading');
    await init();
  });
}

async function performSearch(query) {
  show('search-spinner');
  try {
    const { folders } = await API.searchFolders(query);
    renderSearchResults(folders);
  } catch (err) {
    document.getElementById('search-results').innerHTML =
      '<div class="empty-msg">Search failed — try again</div>';
  } finally {
    hide('search-spinner');
  }
}

async function seedFolderCache() {
  try {
    const { folders } = await API.searchFolders('');
    FolderCache.populate(folders || []);
  } catch { /* non-critical */ }
}

function renderSearchResults(folders) {
  const el = document.getElementById('search-results');
  if (!folders || folders.length === 0) {
    el.innerHTML = '<div class="empty-msg">No folders found</div>';
    return;
  }
  el.innerHTML = '';
  folders.slice(0, 20).forEach(f => el.appendChild(buildFolderItem(f, 'search')));
}

// ── Pinned / recent ────────────────────────────────────────────────────

function renderPinned() {
  const el = document.getElementById('pinned-list');
  el.innerHTML = '';

  const items = [
    ..._prefs.pinned.map(f => ({ ...f, _type: 'pinned' })),
    ..._prefs.recent
      .filter(r => !_prefs.pinned.find(p => p.id === r.id))
      .slice(0, 5)
      .map(f => ({ ...f, _type: 'recent' })),
  ];

  if (items.length === 0) {
    el.innerHTML = '<div class="empty-msg">No recent folders yet</div>';
    return;
  }

  items.forEach(f => el.appendChild(buildFolderItem(f, f._type)));
}

// ── Folder item builder ────────────────────────────────────────────────

function buildFolderItem(folder, type) {
  const div = document.createElement('div');
  div.className = 'folder-item';
  div.dataset.folderId = folder.id;

  const isPinned = _prefs.pinned.some(p => p.id === folder.id);
  const confidenceClass = folder.confidence >= 0.75
    ? 'confidence-high'
    : folder.confidence >= 0.45
      ? 'confidence-mid'
      : 'confidence-low';

  div.innerHTML = `
    <div class="folder-icon">📁</div>
    <div class="folder-text">
      <div class="folder-name">${escapeHtml(folder.name)}</div>
      <div class="folder-path">${escapeHtml(folder.path || '')}</div>
    </div>
    ${type === 'suggestion' ? `
      <span class="confidence-badge ${confidenceClass}">${Math.round(folder.confidence * 100)}%</span>
    ` : ''}
    <button class="pin-btn ${isPinned ? 'pinned' : ''}"
            title="${isPinned ? 'Unpin' : 'Pin'}"
            data-pin-id="${folder.id}">
      ${isPinned ? '📌' : '🖇️'}
    </button>
  `;

  // Save click
  div.addEventListener('click', (e) => {
    if (e.target.closest('.pin-btn')) return;
    saveEmail(folder);
  });

  // Pin click
  div.querySelector('.pin-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    togglePin(folder, div.querySelector('.pin-btn'));
  });

  return div;
}

// ── Save flow ──────────────────────────────────────────────────────────

async function saveEmail(folder) {
  showOverlay('Saving to ' + folder.name + '…');
  try {
    await API.saveEmail(folder.driveId, folder.itemId, _currentEmail);

    // Update recent
    _prefs.recent = [folder, ..._prefs.recent.filter(r => r.id !== folder.id)].slice(0, 10);
    await API.savePreferences(_prefs).catch(() => {});
    renderPinned();

    hideOverlay();
    showStatus(`✓ Saved to "${folder.name}"`, 'success');

  } catch (err) {
    hideOverlay();
    showStatus(`✗ Failed to save: ${err.message}`, 'error');
  }
}

// ── Pin management ─────────────────────────────────────────────────────

async function togglePin(folder, btn) {
  const idx = _prefs.pinned.findIndex(p => p.id === folder.id);
  if (idx === -1) {
    _prefs.pinned.unshift(folder);
    btn.classList.add('pinned');
    btn.textContent = '📌';
    btn.title = 'Unpin';
  } else {
    _prefs.pinned.splice(idx, 1);
    btn.classList.remove('pinned');
    btn.textContent = '🖇️';
    btn.title = 'Pin';
  }
  renderPinned();
  await API.savePreferences(_prefs).catch(() => {});
}

// ── UI helpers ─────────────────────────────────────────────────────────

function show(elOrId) {
  const el = typeof elOrId === 'string' ? document.getElementById(elOrId) : elOrId;
  if (el) el.hidden = false;
}
function hide(elOrId) {
  const el = typeof elOrId === 'string' ? document.getElementById(elOrId) : elOrId;
  if (el) el.hidden = true;
}
function showOverlay(msg) {
  document.getElementById('saving-label').textContent = msg;
  show('saving-overlay');
}
function hideOverlay() {
  hide('saving-overlay');
}
function showStatus(msg, type) {
  const bar = document.getElementById('status-bar');
  bar.textContent = msg;
  bar.className = type;
  show(bar);
  setTimeout(() => hide(bar), 5000);
}
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
