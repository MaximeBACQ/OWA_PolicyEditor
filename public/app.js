'use strict';

// ============================================================
//  STATE
// ============================================================

const state = {
  // Connection
  connected: false,
  upn: '',
  tenant: '',

  // Policy list (array of { name, shortName, isDefault, userCount })
  policies: [],

  // Currently open policy in the editor
  currentPolicyName: null,
  policyParams: null,   // { params: {}, typeMap: {} }

  // Pending edits: { paramName: { original, current } }
  // "current" is the value the user has set but not yet saved.
  pendingChanges: {},

  // Users (array of { displayName, primarySmtpAddress, owaMailboxPolicy })
  users: [],

  // Which user the assign dialog is open for
  assignTargetUpn: null,
};

// ============================================================
//  API
// ============================================================

// Every API call goes through here so error handling is consistent.
async function apiFetch(method, path, body) {
  const opts = {
    method,
    headers: { 'Content-Type': 'application/json' },
  };
  if (body !== undefined) {
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(path, opts);
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(json.error || `HTTP ${res.status}`);
  }
  return json;
}

const api = {
  status:         ()           => apiFetch('GET',    '/api/status'),
  connect:        (upn)        => apiFetch('POST',   '/api/connect',   { upn }),
  disconnect:     ()           => apiFetch('POST',   '/api/disconnect'),

  getPolicies:    ()           => apiFetch('GET',    '/api/policies'),
  getPolicyParams:(name)       => apiFetch('GET',    `/api/policies/${encodeURIComponent(name)}/params`),
  patchPolicy:    (name, chgs) => apiFetch('PATCH',  `/api/policies/${encodeURIComponent(name)}`, { changes: chgs }),
  createPolicy:   (name)       => apiFetch('POST',   '/api/policies',  { name }),
  deletePolicy:   (name)       => apiFetch('DELETE', `/api/policies/${encodeURIComponent(name)}`),

  getUsers:       ()           => apiFetch('GET',    '/api/users'),
  assignPolicy:   (upn, p)     => apiFetch('POST',   `/api/users/${encodeURIComponent(upn)}/policy`, { policyName: p }),
  resetPolicy:    (upn)        => apiFetch('DELETE', `/api/users/${encodeURIComponent(upn)}/policy`),

  getLog:         ()           => apiFetch('GET',    '/api/log'),
};

// ============================================================
//  NAVIGATION
// ============================================================

const SCREENS = ['connect', 'policies', 'editor', 'users', 'log'];

function showScreen(name) {
  for (const id of SCREENS) {
    const el = document.getElementById(`screen-${id}`);
    el.hidden = (id !== name);
  }
}

// ============================================================
//  RENDERERS
// ============================================================

// ── Status badge ─────────────────────────────────────────────
function updateStatusBadge() {
  const badge = document.getElementById('status-badge');
  if (state.connected) {
    badge.textContent = `● ${state.upn}`;
    badge.className = 'status-badge status-connected';
    badge.title = state.tenant;
  } else {
    badge.textContent = '● Disconnected';
    badge.className = 'status-badge status-disconnected';
    badge.title = '';
  }
}

// ── Policy list table ─────────────────────────────────────────
function renderPolicies(policies) {
  const tbody = document.getElementById('policies-tbody');
  tbody.innerHTML = '';

  for (const p of policies) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escHtml(p.name)}</td>
      <td>${p.userCount}</td>
      <td>${p.isDefault ? '<span class="default-badge">Default</span>' : ''}</td>
      <td class="table-actions">
        <button class="btn-edit outline" data-name="${escAttr(p.name)}" style="font-size:0.8rem;padding:0.2rem 0.6rem;margin:0">Edit</button>
        <button class="btn-delete contrast outline" data-name="${escAttr(p.name)}" data-default="${p.isDefault}"
          ${p.isDefault ? 'disabled title="Cannot delete the default policy"' : ''}
          style="font-size:0.8rem;padding:0.2rem 0.6rem;margin:0">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('policies-loading').hidden = true;
  document.getElementById('policies-table').hidden = false;
}

// ── Policy editor ─────────────────────────────────────────────
// Builds the 4-column grid from params + typeMap + pending changes.
function renderEditor() {
  const { params, typeMap } = state.policyParams;

  // Classify each param into the four buckets.
  const cols = { true: [], false: [], notset: [], values: [] };

  for (const [name, value] of Object.entries(params)) {
    const typeName = typeMap[name] || '';
    const isBoolean = typeof value === 'boolean' ||
                      typeName.includes('Boolean');

    // Use pending value if there is one, otherwise the server value
    const current = state.pendingChanges[name] !== undefined
      ? state.pendingChanges[name].current
      : value;

    const entry = { name, value: current, originalServerValue: value, isBoolean, typeName };

    if (current === true)                          cols.true.push(entry);
    else if (current === false)                    cols.false.push(entry);
    else if (current === null || current === '')   cols.notset.push(entry);
    else                                           cols.values.push(entry);
  }

  // Sort each bucket alphabetically
  for (const key of Object.keys(cols)) {
    cols[key].sort((a, b) => a.name.localeCompare(b.name));
  }

  // Update column counts
  document.getElementById('count-true').textContent   = `(${cols.true.length})`;
  document.getElementById('count-false').textContent  = `(${cols.false.length})`;
  document.getElementById('count-notset').textContent = `(${cols.notset.length})`;
  document.getElementById('count-values').textContent = `(${cols.values.length})`;

  renderEditorCol('col-true-items',   cols.true,   'true');
  renderEditorCol('col-false-items',  cols.false,  'false');
  renderEditorCol('col-notset-items', cols.notset, 'notset');
  renderEditorCol('col-values-items', cols.values, 'values');

  updateSaveBadge();

  document.getElementById('editor-loading').hidden = true;
  document.getElementById('editor-grid').hidden = false;
}

// Render one column's param list
function renderEditorCol(containerId, entries, colType) {
  const container = document.getElementById(containerId);
  container.innerHTML = '';

  for (const entry of entries) {
    const row = document.createElement('div');
    const isModified = state.pendingChanges[entry.name] !== undefined;
    row.className = 'param-row' + (isModified ? ' modified' : '');
    row.dataset.param = entry.name;

    if (colType === 'true' || colType === 'false') {
      // Boolean params: render as a checkbox
      // Clicking toggles to the opposite value
      const checked = colType === 'true';
      row.innerHTML = `
        <input type="checkbox" ${checked ? 'checked' : ''} data-param="${escAttr(entry.name)}" />
        <span class="param-name">${escHtml(entry.name)}</span>
      `;
    } else if (colType === 'notset') {
      // Null params: show name + small buttons to set true/false/text
      // The param may be a known boolean (typeMap) or unknown type
      const boolButtons = entry.isBoolean
        ? `<button class="null-set-true" data-param="${escAttr(entry.name)}">T</button>
           <button class="null-set-false" data-param="${escAttr(entry.name)}">F</button>`
        : `<button class="null-set-true" data-param="${escAttr(entry.name)}">T</button>
           <button class="null-set-false" data-param="${escAttr(entry.name)}">F</button>
           <button class="null-set-text" data-param="${escAttr(entry.name)}">…</button>`;
      row.innerHTML = `
        <span class="param-name"><span class="val-null">(null)</span> ${escHtml(entry.name)}</span>
        <span class="null-actions">${boolButtons}</span>
      `;
    } else {
      // VALUES column: show inline text input
      const displayVal = Array.isArray(entry.value)
        ? entry.value.join(', ')
        : String(entry.value ?? '');
      row.innerHTML = `
        <input type="text" value="${escAttr(displayVal)}"
               data-param="${escAttr(entry.name)}"
               data-original="${escAttr(displayVal)}"
               title="${escAttr(entry.name)}" />
      `;
    }

    container.appendChild(row);
  }
}

// Keep the save button and change count in sync with pending changes
function updateSaveBadge() {
  const count = Object.keys(state.pendingChanges).length;
  const badge  = document.getElementById('editor-change-count');
  const saveBtn = document.getElementById('btn-save');

  if (count > 0) {
    badge.textContent = `${count} unsaved change${count !== 1 ? 's' : ''}`;
    badge.hidden = false;
    saveBtn.disabled = false;
  } else {
    badge.hidden = true;
    saveBtn.disabled = true;
  }
}

// ── Users table ───────────────────────────────────────────────
function renderUsers(users, filter) {
  const tbody = document.getElementById('users-tbody');
  tbody.innerHTML = '';

  const lc = (filter || '').toLowerCase();
  const filtered = lc
    ? users.filter(u =>
        u.displayName.toLowerCase().includes(lc) ||
        u.primarySmtpAddress.toLowerCase().includes(lc))
    : users;

  for (const u of filtered) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escHtml(u.displayName)}</td>
      <td>${escHtml(u.primarySmtpAddress)}</td>
      <td>${escHtml(u.owaMailboxPolicy || '(default)')}</td>
      <td class="table-actions">
        <button class="btn-assign-user outline" data-upn="${escAttr(u.primarySmtpAddress)}"
          style="font-size:0.8rem;padding:0.2rem 0.6rem;margin:0">Assign</button>
        <button class="btn-reset-user outline contrast" data-upn="${escAttr(u.primarySmtpAddress)}"
          style="font-size:0.8rem;padding:0.2rem 0.6rem;margin:0">Reset</button>
      </td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('users-loading').hidden = true;
  document.getElementById('users-table').hidden = false;
}

// ── Log viewer ────────────────────────────────────────────────
function renderLog(lines) {
  const pre = document.getElementById('log-content');
  pre.innerHTML = '';

  if (lines.length === 0) {
    pre.textContent = '(No changes logged yet)';
    pre.hidden = false;
    document.getElementById('log-loading').hidden = true;
    return;
  }

  for (const line of lines) {
    const span = document.createElement('span');
    if (/^\[/.test(line)) {
      // Timestamp header line — bright blue
      span.className = 'log-ts';
    } else if (line.includes('->')) {
      // Change diff line — yellow
      span.className = 'log-diff';
    } else {
      // Blank or separator — dimmed
      span.className = 'log-dim';
    }
    span.textContent = line + '\n';
    pre.appendChild(span);
  }

  document.getElementById('log-loading').hidden = true;
  pre.hidden = false;
}

// ── Save confirmation dialog ──────────────────────────────────
function openSaveDialog() {
  const tbody = document.getElementById('dialog-save-tbody');
  tbody.innerHTML = '';

  const changes = buildChangesList();

  document.getElementById('dialog-save-summary').textContent =
    `${changes.length} change${changes.length !== 1 ? 's' : ''} will be applied to ${state.currentPolicyName}`;

  for (const c of changes) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escHtml(c.name)}</td>
      <td class="diff-old">${escHtml(formatValue(c.oldValue))}</td>
      <td class="diff-arrow">→</td>
      <td class="diff-new">${escHtml(formatValue(c.newValue))}</td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('dialog-save').showModal();
}

// ============================================================
//  UTILS
// ============================================================

function escHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escAttr(s) {
  return String(s ?? '').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Human-readable representation of a param value for diffs and logs
function formatValue(v) {
  if (v === null || v === undefined || v === '') return '(null)';
  if (v === true)  return 'True';
  if (v === false) return 'False';
  if (Array.isArray(v)) return v.join(', ');
  return String(v);
}

// Build the array of changes from state.pendingChanges for PATCH and dialog
function buildChangesList() {
  return Object.entries(state.pendingChanges).map(([name, { original, current }]) => ({
    name,
    oldValue: original,
    newValue: current,
  }));
}

// Record a pending change; re-render the editor to update highlighting
function recordChange(paramName, newValue) {
  // Use the original server value as the baseline, not a previous pending value
  const originalServerValue = state.policyParams.params[paramName];

  // If the new value matches the original server value, remove the pending change
  if (valuesEqual(newValue, originalServerValue)) {
    delete state.pendingChanges[paramName];
  } else {
    state.pendingChanges[paramName] = {
      original: originalServerValue,
      current:  newValue,
    };
  }

  // Re-render the grid so the param moves columns and the border appears
  renderEditor();
}

// Equality check that handles booleans, null, strings, and arrays
function valuesEqual(a, b) {
  if (a === b) return true;
  if (Array.isArray(a) && Array.isArray(b)) {
    return a.length === b.length && a.every((v, i) => v === b[i]);
  }
  // Treat null/empty-string as equivalent (matches server Normalize-ParamValue)
  const isEmpty = v => v === null || v === undefined || v === '';
  if (isEmpty(a) && isEmpty(b)) return true;
  return false;
}

function showError(elementId, message) {
  const el = document.getElementById(elementId);
  el.textContent = message;
  el.hidden = false;
}

function clearError(elementId) {
  const el = document.getElementById(elementId);
  el.textContent = '';
  el.hidden = true;
}

// ============================================================
//  EVENT HANDLERS
// ============================================================

// ── Connect form ──────────────────────────────────────────────
document.getElementById('form-connect').addEventListener('submit', async (e) => {
  e.preventDefault();
  const upn = document.getElementById('input-upn').value.trim();
  const btn = document.getElementById('btn-connect');

  clearError('connect-error');
  btn.disabled = true;
  btn.textContent = 'Connecting… (check browser for OAuth popup)';

  try {
    const res = await api.connect(upn);
    state.connected = true;
    state.upn    = res.upn;
    state.tenant = res.tenant;
    updateStatusBadge();
    await loadPoliciesScreen();
  } catch (err) {
    showError('connect-error', err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Connect';
  }
});

// ── Disconnect ────────────────────────────────────────────────
document.getElementById('btn-disconnect').addEventListener('click', async () => {
  if (!confirm('Disconnect from Exchange Online?')) return;
  try {
    await api.disconnect();
    state.connected = false;
    state.upn = state.tenant = '';
    updateStatusBadge();
    showScreen('connect');
  } catch (err) {
    alert('Disconnect failed: ' + err.message);
  }
});

// ── Policy list actions ───────────────────────────────────────
document.getElementById('policies-tbody').addEventListener('click', async (e) => {
  const editBtn   = e.target.closest('.btn-edit');
  const deleteBtn = e.target.closest('.btn-delete');

  if (editBtn) {
    const name = editBtn.dataset.name;
    await loadEditorScreen(name);
  }

  if (deleteBtn && !deleteBtn.disabled) {
    const name = deleteBtn.dataset.name;
    if (!confirm(`Permanently delete policy "${name}"? This cannot be undone.`)) return;
    try {
      await api.deletePolicy(name);
      await loadPoliciesScreen();
    } catch (err) {
      showError('policies-error', 'Delete failed: ' + err.message);
    }
  }
});

// ── New policy dialog ─────────────────────────────────────────
document.getElementById('btn-new-policy').addEventListener('click', () => {
  document.getElementById('input-new-policy-name').value = '';
  clearError('new-policy-error');
  document.getElementById('dialog-new-policy').showModal();
});

document.getElementById('btn-new-policy-cancel').addEventListener('click', () => {
  document.getElementById('dialog-new-policy').close();
});

document.getElementById('btn-new-policy-confirm').addEventListener('click', async () => {
  const name = document.getElementById('input-new-policy-name').value.trim();
  if (!name) { showError('new-policy-error', 'Please enter a policy name.'); return; }

  clearError('new-policy-error');
  const btn = document.getElementById('btn-new-policy-confirm');
  btn.disabled = true;

  try {
    await api.createPolicy(name);
    document.getElementById('dialog-new-policy').close();
    // Open the new policy directly in the editor
    await loadEditorScreen(name);
  } catch (err) {
    showError('new-policy-error', err.message);
  } finally {
    btn.disabled = false;
  }
});

// ── Editor navigation ─────────────────────────────────────────
document.getElementById('btn-editor-back').addEventListener('click', async () => {
  const unsaved = Object.keys(state.pendingChanges).length;
  if (unsaved > 0) {
    if (!confirm(`You have ${unsaved} unsaved change${unsaved !== 1 ? 's' : ''}. Discard and go back?`)) return;
  }
  state.pendingChanges = {};
  await loadPoliciesScreen();
});

// ── Editor: checkbox toggles (booleans in TRUE/FALSE columns) ─
document.getElementById('editor-grid').addEventListener('change', (e) => {
  const cb = e.target.closest('input[type="checkbox"]');
  if (!cb) return;
  const paramName = cb.dataset.param;
  recordChange(paramName, cb.checked);
});

// ── Editor: text input changes (VALUES column) ────────────────
document.getElementById('editor-grid').addEventListener('change', (e) => {
  const input = e.target.closest('input[type="text"]');
  if (!input) return;
  const paramName  = input.dataset.param;
  const typeName   = state.policyParams.typeMap[paramName] || '';

  let newValue = input.value;

  // If the type is an array, split on commas
  if (typeName.includes('MultiValuedProperty') || typeName.includes('[]')) {
    newValue = newValue.split(',').map(s => s.trim()).filter(s => s !== '');
    if (newValue.length === 0) newValue = null;
  } else if (newValue.trim() === '') {
    newValue = null;
  }

  recordChange(paramName, newValue);
});

// ── Editor: NOT SET column buttons ────────────────────────────
document.getElementById('editor-grid').addEventListener('click', (e) => {
  const trueBtn  = e.target.closest('.null-set-true');
  const falseBtn = e.target.closest('.null-set-false');
  const textBtn  = e.target.closest('.null-set-text');

  if (trueBtn)  { recordChange(trueBtn.dataset.param, true);  return; }
  if (falseBtn) { recordChange(falseBtn.dataset.param, false); return; }
  if (textBtn) {
    const paramName = textBtn.dataset.param;
    const val = prompt(`Enter value for ${paramName}:`);
    if (val !== null) recordChange(paramName, val.trim() === '' ? null : val.trim());
  }
});

// ── Save button ───────────────────────────────────────────────
document.getElementById('btn-save').addEventListener('click', () => {
  if (Object.keys(state.pendingChanges).length === 0) return;
  openSaveDialog();
});

// ── Save dialog: cancel ───────────────────────────────────────
document.getElementById('btn-dialog-cancel').addEventListener('click', () => {
  document.getElementById('dialog-save').close();
});

// ── Save dialog: confirm & apply ──────────────────────────────
document.getElementById('btn-dialog-confirm').addEventListener('click', async () => {
  const btn = document.getElementById('btn-dialog-confirm');
  btn.disabled = true;
  btn.textContent = 'Applying…';
  clearError('editor-error');

  try {
    const changes = buildChangesList();
    await api.patchPolicy(state.currentPolicyName, changes);

    document.getElementById('dialog-save').close();
    state.pendingChanges = {};

    // Reload params so values reflect what Exchange now has
    await loadEditorParams(state.currentPolicyName);
  } catch (err) {
    document.getElementById('dialog-save').close();
    showError('editor-error', 'Failed to apply changes: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Apply to Exchange Online';
  }
});

// ── Navigation buttons ────────────────────────────────────────
document.getElementById('btn-goto-users').addEventListener('click', () => loadUsersScreen());
document.getElementById('btn-goto-log').addEventListener('click',   () => loadLogScreen());
document.getElementById('btn-users-back').addEventListener('click', () => loadPoliciesScreen());
document.getElementById('btn-log-back').addEventListener('click',   () => loadPoliciesScreen());

// ── User filter ───────────────────────────────────────────────
document.getElementById('user-search').addEventListener('input', (e) => {
  renderUsers(state.users, e.target.value);
});

// ── Users table actions ───────────────────────────────────────
document.getElementById('users-tbody').addEventListener('click', async (e) => {
  const assignBtn = e.target.closest('.btn-assign-user');
  const resetBtn  = e.target.closest('.btn-reset-user');

  if (assignBtn) {
    const upn = assignBtn.dataset.upn;
    state.assignTargetUpn = upn;

    // Populate the policy selector with current policy list
    const select = document.getElementById('dialog-policy-select');
    select.innerHTML = state.policies
      .map(p => `<option value="${escAttr(p.name)}">${escHtml(p.name)}</option>`)
      .join('');

    document.getElementById('dialog-assign-user').textContent = `Assign policy to: ${upn}`;
    document.getElementById('dialog-assign').showModal();
  }

  if (resetBtn) {
    const upn = resetBtn.dataset.upn;
    if (!confirm(`Reset ${upn} to the organization default OWA policy?`)) return;
    try {
      await api.resetPolicy(upn);
      await loadUsersScreen();
    } catch (err) {
      showError('users-error', 'Reset failed: ' + err.message);
    }
  }
});

document.getElementById('btn-assign-cancel').addEventListener('click', () => {
  document.getElementById('dialog-assign').close();
});

document.getElementById('btn-assign-confirm').addEventListener('click', async () => {
  const upn    = state.assignTargetUpn;
  const policy = document.getElementById('dialog-policy-select').value;
  const btn    = document.getElementById('btn-assign-confirm');

  btn.disabled = true;
  try {
    await api.assignPolicy(upn, policy);
    document.getElementById('dialog-assign').close();
    await loadUsersScreen();
  } catch (err) {
    alert('Assign failed: ' + err.message);
  } finally {
    btn.disabled = false;
  }
});

// ============================================================
//  SCREEN LOADERS
// ============================================================

async function loadPoliciesScreen() {
  showScreen('policies');
  clearError('policies-error');
  document.getElementById('policies-loading').hidden = false;
  document.getElementById('policies-table').hidden = true;

  try {
    state.policies = await api.getPolicies();
    renderPolicies(state.policies);
  } catch (err) {
    document.getElementById('policies-loading').hidden = true;
    showError('policies-error', 'Failed to load policies: ' + err.message);
  }
}

async function loadEditorScreen(policyName) {
  state.currentPolicyName = policyName;
  state.pendingChanges    = {};
  state.policyParams      = null;

  document.getElementById('editor-policy-name').textContent = policyName;
  showScreen('editor');
  clearError('editor-error');
  document.getElementById('editor-loading').hidden = false;
  document.getElementById('editor-grid').hidden = true;

  await loadEditorParams(policyName);
}

// Fetch params from the server and render — called on initial load and after save
async function loadEditorParams(policyName) {
  document.getElementById('editor-loading').hidden = false;
  document.getElementById('editor-grid').hidden = true;

  try {
    state.policyParams = await api.getPolicyParams(policyName);
    renderEditor();
  } catch (err) {
    document.getElementById('editor-loading').hidden = true;
    showError('editor-error', 'Failed to load params: ' + err.message);
  }
}

async function loadUsersScreen() {
  showScreen('users');
  clearError('users-error');
  document.getElementById('users-loading').hidden = false;
  document.getElementById('users-table').hidden = true;
  document.getElementById('user-search').value = '';

  try {
    state.users = await api.getUsers();
    renderUsers(state.users, '');
  } catch (err) {
    document.getElementById('users-loading').hidden = true;
    showError('users-error', 'Failed to load users: ' + err.message);
  }
}

async function loadLogScreen() {
  showScreen('log');
  clearError('log-error');
  document.getElementById('log-loading').hidden = false;
  document.getElementById('log-content').hidden = true;

  try {
    const { lines } = await api.getLog();
    renderLog(lines);
  } catch (err) {
    document.getElementById('log-loading').hidden = true;
    showError('log-error', 'Failed to load log: ' + err.message);
  }
}

// ============================================================
//  INIT
// ============================================================

async function init() {
  try {
    const status = await api.status();
    state.connected = status.connected;
    state.upn       = status.upn    || '';
    state.tenant    = status.tenant || '';
    updateStatusBadge();

    if (state.connected) {
      // Pre-fill UPN input in case they disconnect and reconnect
      document.getElementById('input-upn').value = state.upn;
      // await loadPoliciesScreen();
    } else {
      // Try to get the last used UPN from the status (server reads from config)
      // Server returns empty string when not connected — just show the connect screen
      showScreen('connect');
    }
  } catch {
    // Server not reachable yet — show connect screen as fallback
    showScreen('connect');
  }
}

init();
