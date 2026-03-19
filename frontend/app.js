const { invoke: rawInvoke } = window.__TAURI__.core;

// --- Auth state ---
let sessionToken = null;
let currentUser = null;
let currentPermissions = new Set();
let _graceTimerId = null;

/** Auth-aware invoke wrapper. Injects session_token into every call. */
async function invoke(command, args = {}) {
  // These commands don't need auth
  const noAuth = ['check_setup_needed', 'setup_admin', 'login'];
  if (!noAuth.includes(command) && sessionToken) {
    args.sessionToken = sessionToken;
  }
  const result = await rawInvoke(command, args);
  if (result && result.auth_error) {
    handleSessionExpired();
    throw new Error(result.error || 'Session expired');
  }
  return result;
}

function handleSessionExpired() {
  if (!currentUser) return;
  // Admin: immediate re-login
  if (currentUser.role === 'admin') {
    doLogout(true);
    return;
  }
  // Staff: grace prompt (30 seconds)
  showSessionExpiredPrompt();
}

function showLoginScreen() {
  document.getElementById('loginScreen').classList.remove('hidden');
  document.getElementById('mainApp').classList.add('hidden');
  document.getElementById('loginError').classList.add('hidden');
  document.getElementById('loginUsername').value = '';
  document.getElementById('loginPassword').value = '';
  document.getElementById('loginUsername').focus();
}

function hideLoginScreen() {
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('mainApp').classList.remove('hidden');
}

function showSessionExpiredPrompt() {
  const overlay = document.getElementById('sessionExpiredOverlay');
  overlay.classList.remove('hidden');
  document.getElementById('sessionExpiredPassword').value = '';
  document.getElementById('sessionExpiredError').classList.add('hidden');
  document.getElementById('sessionExpiredPassword').focus();
  let remaining = 30;
  const timerEl = document.getElementById('sessionExpiredTimer');
  timerEl.textContent = `${remaining} 秒後自動登出`;
  _graceTimerId = setInterval(() => {
    remaining--;
    timerEl.textContent = `${remaining} 秒後自動登出`;
    if (remaining <= 0) {
      clearInterval(_graceTimerId);
      doLogout(true);
    }
  }, 1000);
}

async function doLogout(skipApi) {
  clearInterval(_graceTimerId);
  if (!skipApi && sessionToken) {
    try { await rawInvoke('logout', { sessionToken: sessionToken }); } catch {}
  }
  sessionToken = null;
  currentUser = null;
  currentPermissions = new Set();
  document.getElementById('sessionExpiredOverlay').classList.add('hidden');
  showLoginScreen();
}

function applyPermissions(perms) {
  currentPermissions = new Set(perms);
  // Map tabs to required permissions
  const tabPerms = {
    'classes': 'classes.view',
    'reminders': 'classes.view',
    'settings': 'settings.modify',
    'fee-guide': 'classes.view',
    'docx-output': 'documents.view',
    'messages': 'documents.view',
    'makeup-plan': 'documents.view',
    'promote-notice': 'documents.view',
    'stock': 'textbooks.view',
    'eps-audit': 'eps.view',
  };
  document.querySelectorAll('.tab-button').forEach(btn => {
    const tab = btn.dataset.tab;
    if (tab === 'tasks') return; // always visible
    if (tab === 'admin') {
      btn.classList.toggle('hidden', !perms.some(p => p.startsWith('admin.')));
      return;
    }
    const req = tabPerms[tab];
    if (req && !currentPermissions.has(req)) {
      btn.classList.add('hidden');
    } else {
      btn.classList.remove('hidden');
    }
  });

  // User info bar
  if (currentUser) {
    document.getElementById('userDisplayName').textContent = currentUser.display_name || currentUser.username;
    document.getElementById('userRole').textContent = currentUser.role;
  }
}

// --- Helper: complete login after successful auth result ---
function completeLogin(result) {
  sessionToken = result.token;
  currentUser = result.user;
  applyPermissions(result.permissions);
  hideLoginScreen();
  loadState();
  loadFeeTemplate();
  loadDocxTemplates();
  loadMessageTemplates();
  loadMakeupTemplate();
  checkForUpdate();
}

// --- Auto-update check ---
async function checkForUpdate() {
  try {
    const updater = window.__TAURI__?.updater;
    const process = window.__TAURI__?.process;
    if (!updater || !process) return;
    const update = await updater.check();
    if (update) {
      const yes = confirm(`有新版本 v${update.version} 可用，是否更新？`);
      if (yes) {
        await update.downloadAndInstall();
        await process.relaunch();
      }
    }
  } catch (e) {
    console.log('Update check skipped:', e);
  }
}

// --- Login form handler ---
document.getElementById('loginForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const errEl = document.getElementById('loginError');
  errEl.classList.add('hidden');
  const username = document.getElementById('loginUsername').value.trim();
  const password = document.getElementById('loginPassword').value;
  if (!username || !password) return;

  const btn = document.getElementById('loginBtn');
  btn.disabled = true;
  btn.textContent = '處理中...';

  try {
    const result = await rawInvoke('login', { username, password });
    if (result.ok) {
      completeLogin(result);
    } else {
      errEl.textContent = result.error;
      errEl.classList.remove('hidden');
    }
  } catch (err) {
    errEl.textContent = '登入失敗: ' + err.message;
    errEl.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = '登入';
  }
});

// --- Setup form handler (first-time account creation) ---
document.getElementById('setupForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const errEl = document.getElementById('setupError');
  errEl.classList.add('hidden');
  const username = document.getElementById('setupUsername').value.trim();
  const displayName = document.getElementById('setupDisplayName').value.trim() || username;
  const password = document.getElementById('setupPassword').value;
  const confirm = document.getElementById('setupPasswordConfirm').value;
  if (!username || !password) return;
  if (password !== confirm) {
    errEl.textContent = '密碼不一致。';
    errEl.classList.remove('hidden');
    return;
  }
  try {
    const result = await rawInvoke('setup_admin', { username, password, displayName });
    if (result.ok) {
      completeLogin(result);
    } else {
      errEl.textContent = result.error;
      errEl.classList.remove('hidden');
    }
  } catch (err) {
    errEl.textContent = '建立失敗: ' + err.message;
    errEl.classList.remove('hidden');
  }
});

// --- Toggle between login and setup views ---
document.getElementById('showSetupBtn').addEventListener('click', (e) => {
  e.preventDefault();
  document.getElementById('loginView').style.display = 'none';
  document.getElementById('setupView').style.display = '';
});
document.getElementById('showLoginBtn').addEventListener('click', (e) => {
  e.preventDefault();
  document.getElementById('setupView').style.display = 'none';
  document.getElementById('loginView').style.display = '';
});

// --- Session expired re-auth ---
document.getElementById('sessionExpiredForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const password = document.getElementById('sessionExpiredPassword').value;
  const errEl = document.getElementById('sessionExpiredError');
  errEl.classList.add('hidden');
  if (!password || !currentUser) return;
  try {
    const result = await rawInvoke('login', { username: currentUser.username, password });
    if (result.ok) {
      clearInterval(_graceTimerId);
      sessionToken = result.token;
      currentUser = result.user;
      applyPermissions(result.permissions);
      document.getElementById('sessionExpiredOverlay').classList.add('hidden');
    } else {
      errEl.textContent = result.error;
      errEl.classList.remove('hidden');
    }
  } catch (err) {
    errEl.textContent = err.message;
    errEl.classList.remove('hidden');
  }
});
document.getElementById('sessionExpiredLogout').addEventListener('click', () => doLogout(true));

// --- Logout button ---
document.getElementById('logoutBtn').addEventListener('click', () => doLogout(false));

// --- Change password ---
document.getElementById('changePasswordBtn').addEventListener('click', async () => {
  const oldPw = prompt('請輸入目前密碼：');
  if (!oldPw) return;
  const newPw = prompt('請輸入新密碼（4 位數字）：');
  if (!newPw) return;
  const confirmPw = prompt('請再次輸入新密碼：');
  if (newPw !== confirmPw) { alert('密碼不一致。'); return; }
  try {
    const r = await invoke('change_password', { oldPassword: oldPw, newPassword: newPw });
    alert(r.ok ? '密碼已更新。' : (r.error || '更新失敗。'));
  } catch (err) { alert('錯誤: ' + err.message); }
});

// --- App init: check setup then show login ---
(async function authInit() {
  try {
    const r = await rawInvoke('check_setup_needed');
    if (r.setup_needed) {
      // Auto-create default admin account
      const setupResult = await rawInvoke('setup_admin', {
        username: 'Jeff',
        password: '9677',
        displayName: 'Jeff',
      });
      if (setupResult.ok) {
        completeLogin(setupResult);
        return;
      }
      // If auto-setup failed, show setup link so user can create manually
      document.getElementById('setupLink').style.display = '';
    }
  } catch (err) {
    console.error('Setup check failed:', err);
  }
  showLoginScreen();
})();

// ============================================
// ADMIN PANEL LOGIC
// ============================================
// Sub-tab switching
document.querySelectorAll('[data-admin-tab]').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('[data-admin-tab]').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('[data-admin-content]').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.querySelector(`[data-admin-content="${btn.dataset.adminTab}"]`).classList.add('active');
    if (btn.dataset.adminTab === 'users') loadAdminUsers();
    if (btn.dataset.adminTab === 'roles') loadAdminRoles();
    if (btn.dataset.adminTab === 'audit') loadAuditLog();
  });
});

async function loadAdminUsers() {
  try {
    const r = await invoke('list_users');
    if (!r.ok) return;
    const tbody = document.querySelector('#usersTable tbody');
    tbody.innerHTML = '';
    for (const u of r.users) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${esc(u.username)}</td>
        <td>${esc(u.display_name)}</td>
        <td>${esc(u.role)}</td>
        <td>${u.last_login || '—'}</td>
        <td class="${u.active ? 'status-active' : 'status-inactive'}">${u.active ? '啟用' : '停用'}</td>
        <td>
          <button class="action-btn" data-action="reset-pw" data-id="${u.id}">重設密碼</button>
          ${u.active
            ? `<button class="action-btn danger" data-action="deactivate" data-id="${u.id}">停用</button>`
            : `<button class="action-btn" data-action="reactivate" data-id="${u.id}">啟用</button>`}
          <button class="action-btn" data-action="edit-role" data-id="${u.id}" data-role="${u.role}">更改角色</button>
        </td>`;
      tbody.appendChild(tr);
    }
  } catch {}
}

function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

document.querySelector('#usersTable').addEventListener('click', async (e) => {
  const btn = e.target.closest('.action-btn');
  if (!btn) return;
  const id = btn.dataset.id;
  if (btn.dataset.action === 'reset-pw') {
    const pw = prompt('輸入新密碼（最少 8 字元）：');
    if (!pw) return;
    const r = await invoke('reset_password', { userId: id, newPassword: pw });
    alert(r.ok ? '密碼已重設。' : (r.error || '失敗'));
  } else if (btn.dataset.action === 'deactivate') {
    if (!confirm('確定停用此帳號？')) return;
    const r = await invoke('deactivate_user', { userId: id });
    alert(r.ok ? '已停用。' : (r.error || '失敗'));
    loadAdminUsers();
  } else if (btn.dataset.action === 'reactivate') {
    const r = await invoke('reactivate_user', { userId: id });
    alert(r.ok ? '已啟用。' : (r.error || '失敗'));
    loadAdminUsers();
  } else if (btn.dataset.action === 'edit-role') {
    const newRole = prompt('輸入新角色（如 admin, staff）：', btn.dataset.role);
    if (!newRole || newRole === btn.dataset.role) return;
    const r = await invoke('update_user', { userId: id, role: newRole });
    alert(r.ok ? '角色已更新。' : (r.error || '失敗'));
    loadAdminUsers();
  }
});

document.getElementById('createUserBtn').addEventListener('click', async () => {
  const username = prompt('用戶名稱：');
  if (!username) return;
  const displayName = prompt('顯示名稱：', username);
  const password = prompt('密碼（4 位數字）：');
  if (!password) return;
  const role = prompt('角色（admin / staff）：', 'staff');
  if (!role) return;
  const r = await invoke('create_user', { username, passwordVal: password, role, displayName: displayName || username });
  alert(r.ok ? '用戶已建立。' : (r.error || '失敗'));
  loadAdminUsers();
});

async function loadAdminRoles() {
  try {
    const rolesR = await invoke('list_roles');
    const permsR = await invoke('list_all_permissions');
    if (!rolesR.ok || !permsR.ok) return;

    const sel = document.getElementById('roleSelect');
    sel.innerHTML = '';
    for (const role of rolesR.roles) {
      const opt = document.createElement('option');
      opt.value = role;
      opt.textContent = role;
      sel.appendChild(opt);
    }

    const grid = document.getElementById('permissionsList');
    const categories = {};
    for (const p of permsR.permissions) {
      if (!categories[p.category]) categories[p.category] = [];
      categories[p.category].push(p);
    }
    grid.innerHTML = '';
    for (const [cat, perms] of Object.entries(categories)) {
      const div = document.createElement('div');
      div.className = 'perm-category';
      div.innerHTML = `<h4>${esc(cat)}</h4>`;
      for (const p of perms) {
        div.innerHTML += `<label class="perm-item"><input type="checkbox" value="${p.key}"> ${esc(p.name)}</label>`;
      }
      grid.appendChild(div);
    }

    sel.addEventListener('change', loadRolePerms);
    loadRolePerms();
  } catch {}
}

async function loadRolePerms() {
  const role = document.getElementById('roleSelect').value;
  const r = await invoke('list_role_permissions', { role });
  if (!r.ok) return;
  const permSet = new Set(r.permissions);
  const isAdmin = role === 'admin';
  document.querySelectorAll('#permissionsList input[type="checkbox"]').forEach(cb => {
    cb.checked = permSet.has(cb.value);
    cb.disabled = isAdmin;
  });
}

document.getElementById('saveRolePermsBtn').addEventListener('click', async () => {
  const role = document.getElementById('roleSelect').value;
  if (role === 'admin') { alert('Admin 角色不可修改。'); return; }
  const perms = [];
  document.querySelectorAll('#permissionsList input[type="checkbox"]:checked').forEach(cb => perms.push(cb.value));
  const r = await invoke('set_role_permissions', { role, permissionList: perms });
  alert(r.ok ? '權限已儲存。' : (r.error || '失敗'));
});

let auditOffset = 0;
const AUDIT_PAGE_SIZE = 50;
async function loadAuditLog() {
  try {
    const r = await invoke('get_audit_log', { limit: AUDIT_PAGE_SIZE, offset: auditOffset });
    if (!r.ok) return;
    const tbody = document.querySelector('#auditTable tbody');
    tbody.innerHTML = '';
    for (const e of r.entries) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${esc(e.timestamp)}</td>
        <td>${esc(e.username)}</td>
        <td>${esc(e.event_type)}</td>
        <td>${esc(e.details)}</td>
        <td class="${e.success ? 'status-active' : 'status-inactive'}">${e.success ? '成功' : '失敗'}</td>`;
      tbody.appendChild(tr);
    }
    document.getElementById('auditPageInfo').textContent =
      `${auditOffset + 1}–${Math.min(auditOffset + AUDIT_PAGE_SIZE, r.total)} / ${r.total}`;
    document.getElementById('auditPrevBtn').disabled = auditOffset === 0;
    document.getElementById('auditNextBtn').disabled = auditOffset + AUDIT_PAGE_SIZE >= r.total;
  } catch {}
}
document.getElementById('auditPrevBtn').addEventListener('click', () => { auditOffset = Math.max(0, auditOffset - AUDIT_PAGE_SIZE); loadAuditLog(); });
document.getElementById('auditNextBtn').addEventListener('click', () => { auditOffset += AUDIT_PAGE_SIZE; loadAuditLog(); });
document.getElementById('refreshAuditBtn').addEventListener('click', () => { auditOffset = 0; loadAuditLog(); });

window.addEventListener("unhandledrejection", (e) => {
  console.error("Unhandled promise rejection:", e.reason);
});

// --- Zoom persistence ---
let currentZoom = 1.0;
const ZOOM_STEP = 0.1;
const ZOOM_MIN = 0.5;
const ZOOM_MAX = 2.0;

function applyZoom(level) {
  currentZoom = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, level));
  document.documentElement.style.zoom = currentZoom;
}

let zoomSaveTimer = null;
function saveZoomDebounced() {
  clearTimeout(zoomSaveTimer);
  zoomSaveTimer = setTimeout(() => {
    invoke('set_zoom_level', { level: currentZoom }).catch(() => {});
  }, 500);
}

window.addEventListener("keydown", (e) => {
  if ((e.ctrlKey || e.metaKey) && (e.key === "=" || e.key === "+")) {
    e.preventDefault();
    applyZoom(currentZoom + ZOOM_STEP);
    saveZoomDebounced();
  } else if ((e.ctrlKey || e.metaKey) && e.key === "-") {
    e.preventDefault();
    applyZoom(currentZoom - ZOOM_STEP);
    saveZoomDebounced();
  } else if ((e.ctrlKey || e.metaKey) && e.key === "0") {
    e.preventDefault();
    applyZoom(1.0);
    saveZoomDebounced();
  }
});

window.addEventListener("wheel", (e) => {
  if (e.ctrlKey || e.metaKey) {
    e.preventDefault();
    const delta = e.deltaY > 0 ? -ZOOM_STEP : ZOOM_STEP;
    applyZoom(currentZoom + delta);
    saveZoomDebounced();
  }
}, { passive: false });

const weekdays = ["一", "二", "三", "四", "五", "六", "日"];

const statusLabels = {
  not_started: "未開始",
  active: "進行中",
  promoted: "已升級",
  merged: "已合併",
  terminated: "已結束",
  ended: "已完結",
};

const reminderConfig = {
  yellowWeeks: 3,
  redWeeks: 2,
};

const paymentConfig = {
  greenMax: 2,
  yellowMax: 3,
};

let appState = {
  classes: [],
  holidays: [],
  postpones: [],
  settings: { teacher: [], room: [], level: [], time: [] },
  stock_history: {},
};

const classBody = document.getElementById("classBody");
const holidayList = document.getElementById("holidayList");
const searchInput = document.getElementById("searchInput");
const statusFilter = document.getElementById("statusFilter");
const locationFilter = document.getElementById("locationFilter");
const levelFilter = document.getElementById("levelFilter");
const teacherSelect = document.getElementById("teacherSelect");
const roomSelect = document.getElementById("roomSelect");
const levelSelect = document.getElementById("levelSelect");
const timeSelect = document.getElementById("timeSelect");
const relayTeacherSelect = document.getElementById("relayTeacherSelect");
const relayFields = document.getElementById("relayFields");

const classForm = document.getElementById("classForm");
const startDateInput = classForm.querySelector('input[name="start_date"]');
const weekdaySelect = classForm.querySelector('select[name="weekday"]');
const weekdayFilter = document.getElementById("weekdayFilter");
const archiveFilter = document.getElementById("archiveFilter");
const issueFilter = document.getElementById("issueFilter");

const teacherList = document.getElementById("teacherList");
const roomList = document.getElementById("roomList");
const levelList = document.getElementById("levelList");
const timeList = document.getElementById("timeList");

const reminderBody = document.getElementById("reminderBody");
const reminderSort = document.getElementById("reminderSort");

const calendarView = document.getElementById("calendarView");
const calendarLabel = document.getElementById("calendarLabel");
const calendarPrevBtn = document.getElementById("calendarPrevBtn");
const calendarNextBtn = document.getElementById("calendarNextBtn");
const calendarMonthBtn = document.getElementById("calendarMonthBtn");
const calendarWeekBtn = document.getElementById("calendarWeekBtn");
const calendarTodayBtn = document.getElementById("calendarTodayBtn");
const calendarJumpMonth = document.getElementById("calendarJumpMonth");
const calendarJumpYear = document.getElementById("calendarJumpYear");
const locationSelect = document.getElementById("locationSelect");
const locationSegments = document.getElementById("locationSegments");
const locationBadge = document.getElementById("locationBadge");
const fabToggle = document.getElementById("fabToggle");
const fabActions = document.querySelector(".fab-actions");
const jumpTopBtn = document.getElementById("jumpTopBtn");
const timeCalcBtn = document.getElementById("timeCalcBtn");
const timeCalcPopover = document.getElementById("timeCalcPopover");
const timeCalcClose = document.getElementById("timeCalcClose");
const timeCalcCurrent = document.getElementById("timeCalcCurrent");
const timeCalcHours = document.getElementById("timeCalcHours");
const timeCalcMinutes = document.getElementById("timeCalcMinutes");
const timeCalcResult = document.getElementById("timeCalcResult");
const sidebarToggle = document.getElementById("sidebarToggle");
const layout = document.querySelector(".layout");
const locationBanner = document.getElementById("locationBanner");
const tabNav = document.querySelector(".sidebar-tabs");
const tabOrderToggle = document.getElementById("tabOrderToggle");

const scheduleModal = document.getElementById("scheduleModal");
const closeScheduleModalBtn = document.getElementById("closeScheduleModal");
const detailSku = document.getElementById("detailSku");
const detailRoom = document.getElementById("detailRoom");
const detailTeacher = document.getElementById("detailTeacher");
const detailTime = document.getElementById("detailTime");
const detailRelayTeacher = document.getElementById("detailRelayTeacher");
const detailRelayDate = document.getElementById("detailRelayDate");
const detailStudents = document.getElementById("detailStudents");
const saveClassDetailBtn = document.getElementById("saveClassDetail");
const deleteClassBtn = document.getElementById("deleteClassBtn");
const scheduleBody = document.getElementById("scheduleBody");
const addLessonDate = document.getElementById("addLessonDate");
const removeLessonDate = document.getElementById("removeLessonDate");
const addLessonBtn = document.getElementById("addLessonBtn");
const removeLessonBtn = document.getElementById("removeLessonBtn");
const postponeOriginalDate = document.getElementById("postponeOriginalDate");
const postponeMakeupDate = document.getElementById("postponeMakeupDate");
const postponeAutoWeek = document.getElementById("postponeAutoWeek");
const postponeReason = document.getElementById("postponeReason");
const addPostponeBtn = document.getElementById("addPostponeBtn");
const postponeList = document.getElementById("postponeList");
const overrideList = document.getElementById("overrideList");
const terminateClassBtn = document.getElementById("terminateClassBtn");

const exportClassesBtn = document.getElementById("exportClassesBtn");
const importClassesInput = document.getElementById("importClassesInput");

const priceForm = document.getElementById("priceForm");
const priceLevelSelect = document.getElementById("priceLevelSelect");
const priceValueInput = document.getElementById("priceValueInput");
const priceList = document.getElementById("priceList");
const priceAdjustPlus = document.getElementById("priceAdjustPlus");
const priceAdjustMinus = document.getElementById("priceAdjustMinus");

const feeLevelSelect = document.getElementById("feeLevelSelect");
const feeLocationSelect = document.getElementById("feeLocationSelect");
const feeMonthSelect = document.getElementById("feeMonthSelect");
const feeLetterSelect = document.getElementById("feeLetterSelect");
const feeYearSelect = document.getElementById("feeYearSelect");
const feeClassTimeInput = document.getElementById("feeClassTimeInput");
const feeStartDateInput = document.getElementById("feeStartDateInput");
const feeDeadlineInput = document.getElementById("feeDeadlineInput");
const feeTextbookInput = document.getElementById("feeTextbookInput");
const feeIdCardInput = document.getElementById("feeIdCardInput");
const feeTemplateOutput = document.getElementById("feeTemplateOutput");
const feeCopyBtn = document.getElementById("feeCopyBtn");
const feeClassPriceLabel = document.getElementById("feeClassPriceLabel");
const feeTotalLabel = document.getElementById("feeTotalLabel");
const feeResetBtn = document.getElementById("feeResetBtn");
const feeTutorPlus = document.getElementById("feeTutorPlus");
const feeTutorMinus = document.getElementById("feeTutorMinus");
const docxTemplateSelect = document.getElementById("docxTemplateSelect");
const docxClassSelect = document.getElementById("docxClassSelect");
const docxClassSelectSecondary = document.getElementById("docxClassSelectSecondary");
const docxRelayRowPrimary = document.getElementById("docxRelayRowPrimary");
const docxRelayTeacherPrimary = document.getElementById("docxRelayTeacherPrimary");
const docxRelayRowSecondary = document.getElementById("docxRelayRowSecondary");
const docxRelayTeacherSecondary = document.getElementById("docxRelayTeacherSecondary");
const docxSecondClassToggleRow = document.getElementById("docxSecondClassToggleRow");
const docxSecondClassToggle = document.getElementById("docxSecondClassToggle");
const docxSecondClassRow = document.getElementById("docxSecondClassRow");
const docxPreview = document.getElementById("docxPreview");
const docxPreviewSecondary = document.getElementById("docxPreviewSecondary");
const docxGenerateBtn = document.getElementById("docxGenerateBtn");
const docxOutput = document.getElementById("docxOutput");
const docxOpenFolderBtn = document.getElementById("docxOpenFolderBtn");

const promoteSourceClassSelect = document.getElementById("promoteSourceClassSelect");
const promoteTargetClassSelect = document.getElementById("promoteTargetClassSelect");
const promoteAddresseeInput = document.getElementById("promoteAddresseeInput");
const promoteMonthInput = document.getElementById("promoteMonthInput");
const promoteBodySuffix = document.getElementById("promoteBodySuffix");
const promoteFieldName = document.getElementById("promoteFieldName");
const promoteFieldStartDate = document.getElementById("promoteFieldStartDate");
const promoteFieldDuration = document.getElementById("promoteFieldDuration");
const promoteFieldTime = document.getElementById("promoteFieldTime");
const promoteFieldTeacher = document.getElementById("promoteFieldTeacher");
const promoteFieldLocation = document.getElementById("promoteFieldLocation");
const promoteFieldRemarks = document.getElementById("promoteFieldRemarks");
const promoteFieldSignatureDate = document.getElementById("promoteFieldSignatureDate");
const promoteTextbookInfo = document.getElementById("promoteTextbookInfo");
const promoteTextbookList = document.getElementById("promoteTextbookList");
const promoteIncludeTextbook = document.getElementById("promoteIncludeTextbook");
const promoteGenerateBtn = document.getElementById("promoteGenerateBtn");
const promoteOpenFolderBtn = document.getElementById("promoteOpenFolderBtn");
const promoteResetBtn = document.getElementById("promoteResetBtn");
const promoteOutput = document.getElementById("promoteOutput");
const promotePreviewCard = document.getElementById("promotePreviewCard");

const messageSearchInput = document.getElementById("messageSearchInput");
const messageCategorySelect = document.getElementById("messageCategorySelect");
const messageList = document.getElementById("messageList");
const messageOutput = document.getElementById("messageOutput");
const messageCopyBtn = document.getElementById("messageCopyBtn");
const messageTitle = document.getElementById("messageTitle");
const messageCategoryModal = document.getElementById("messageCategoryModal");
const messageCategoryInput = document.getElementById("messageCategoryInput");
const saveMessageCategoryBtn = document.getElementById("saveMessageCategoryBtn");
const closeMessageCategoryModal = document.getElementById("closeMessageCategoryModal");

const makeupRows = document.getElementById("makeupRows");
const makeupAddRowBtn = document.getElementById("makeupAddRowBtn");
const makeupOutput = document.getElementById("makeupOutput");
const makeupCopyBtn = document.getElementById("makeupCopyBtn");
const makeupResetBtn = document.getElementById("makeupResetBtn");

let activeClassId = "";
let calendarMode = "week";
let calendarAnchor = new Date();

let feeTemplateBase = "";
let feeClassAdjust = 0;
let feeTextbookAutoFilled = false;
let docxTemplates = [];
let messageTemplates = [];
let activeMessageName = "";
let tabOrderUnlocked = false;
let makeupTemplateBase = "";

const textbookForm = document.getElementById("textbookForm");
const textbookNameInput = document.getElementById("textbookNameInput");
const textbookPriceInput = document.getElementById("textbookPriceInput");
const textbookList = document.getElementById("textbookList");
const levelTextbookList = document.getElementById("levelTextbookList");

const inventoryTable = document.getElementById("inventoryTable");
const promotionTable = document.getElementById("promotionTable");
const promotionThreshold = document.getElementById("promotionThreshold");
const stockReviewBanner = document.getElementById("stockReviewBanner");
const stockQuickReview = document.getElementById("stockQuickReview");
const stockReviewBody = document.getElementById("stockReviewBody");
const stockReviewBtn = document.getElementById("stockReviewBtn");
const reviewWeekdayFilter = document.getElementById("reviewWeekdayFilter");
const reviewLevelFilter = document.getElementById("reviewLevelFilter");
const stockConfirmAllBtn = document.getElementById("stockConfirmAllBtn");
const stockSaveReviewBtn = document.getElementById("stockSaveReviewBtn");

// ---- Utility: Toast Notifications ----
function showToast(message, type = "error") {
  const container = document.getElementById("toastContainer");
  if (!container) return;
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  const timeout = type === "error" ? 4000 : 2500;
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(14px)";
    setTimeout(() => toast.remove(), 210);
  }, timeout);
}

// ---- Utility: Button Loading State ----
function setButtonLoading(btn, loading) {
  if (!btn) return;
  if (loading) {
    btn.dataset.originalText = btn.textContent;
    btn.textContent = "處理中...";
    btn.classList.add("loading");
  } else {
    if (btn.dataset.originalText !== undefined) {
      btn.textContent = btn.dataset.originalText;
    }
    btn.classList.remove("loading");
  }
}

// ---- Utility: Time Calculator ----
let timeCalcInterval = null;

function updateTimeCalcCurrent() {
  if (!timeCalcCurrent) return;
  const now = new Date();
  timeCalcCurrent.textContent = now.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  updateTimeCalcResult();
}

function updateTimeCalcResult() {
  if (!timeCalcResult || !timeCalcHours || !timeCalcMinutes) return;
  const now = new Date();
  const hours = parseInt(timeCalcHours.value, 10) || 0;
  const minutes = parseInt(timeCalcMinutes.value, 10) || 0;
  const result = new Date(now.getTime() + (hours * 60 + minutes) * 60000);
  timeCalcResult.textContent = result.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

function openTimeCalc() {
  if (!timeCalcPopover) return;
  timeCalcPopover.classList.remove("hidden");
  updateTimeCalcCurrent();
  timeCalcInterval = setInterval(updateTimeCalcCurrent, 1000);
}

function closeTimeCalc() {
  if (!timeCalcPopover) return;
  timeCalcPopover.classList.add("hidden");
  if (timeCalcInterval) {
    clearInterval(timeCalcInterval);
    timeCalcInterval = null;
  }
}

// ---- Task / To-Do (localStorage) ----
function loadTasks() {
  try {
    return JSON.parse(localStorage.getItem("dij_tasks") || "[]");
  } catch {
    return [];
  }
}

function saveTasks(tasks) {
  localStorage.setItem("dij_tasks", JSON.stringify(tasks));
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function addTask() {
  if (!taskInput) return;
  const text = taskInput.value.trim();
  if (!text) return;
  const tasks = loadTasks();
  tasks.push({
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    text,
    time: taskTimeInput?.value || "",
    done: false,
    created: new Date().toISOString(),
  });
  saveTasks(tasks);
  taskInput.value = "";
  if (taskTimeInput) taskTimeInput.value = "";
  taskInput.focus();
  renderTasks();
}

function toggleTask(id) {
  const tasks = loadTasks();
  const task = tasks.find((t) => t.id === id);
  if (task) {
    task.done = !task.done;
    saveTasks(tasks);
    renderTasks();
  }
}

function deleteTask(id) {
  const tasks = loadTasks().filter((t) => t.id !== id);
  saveTasks(tasks);
  renderTasks();
}

function clearDoneTasks() {
  const tasks = loadTasks().filter((t) => !t.done);
  saveTasks(tasks);
  renderTasks();
}

function renderTaskItem(task) {
  const item = document.createElement("div");
  item.className = `task-item${task.done ? " done" : ""}`;
  item.innerHTML = `
    <input type="checkbox" ${task.done ? "checked" : ""} data-task-id="${task.id}" />
    <span class="task-text">${escapeHtml(task.text)}</span>
    ${task.time ? `<span class="task-time">${task.time}</span>` : ""}
    <button class="task-delete" data-task-delete="${task.id}" type="button">&times;</button>
  `;
  return item;
}

function renderTasks() {
  if (!taskList || !taskDoneList) return;
  const tasks = loadTasks();
  const active = tasks.filter((t) => !t.done);
  const done = tasks.filter((t) => t.done);

  taskList.innerHTML = "";
  if (active.length) {
    active.forEach((task) => taskList.appendChild(renderTaskItem(task)));
  } else {
    const empty = document.createElement("div");
    empty.className = "task-empty";
    empty.textContent = "沒有待辦事項";
    taskList.appendChild(empty);
  }

  taskDoneList.innerHTML = "";
  done.forEach((task) => taskDoneList.appendChild(renderTaskItem(task)));
  if (taskDoneCount) {
    taskDoneCount.textContent = String(done.length);
  }
}

// ---- Utility: Copy to Clipboard ----
async function copyToClipboard(text, targetElement) {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      if (targetElement) {
        targetElement.select();
        document.execCommand("copy");
        targetElement.setSelectionRange(0, 0);
      }
    }
    showToast("已複製到剪貼板", "success");
  } catch (error) {
    showToast("複製失敗，請手動複製。", "error");
  }
}

// ---- Utility: Inline Field Validation ----
function showFieldError(input, message) {
  input.classList.add("input-error");
  let errEl = input.parentElement.querySelector(".field-error");
  if (!errEl) {
    errEl = document.createElement("span");
    errEl.className = "field-error";
    input.after(errEl);
  }
  errEl.textContent = message;
  const clear = () => {
    input.classList.remove("input-error");
    errEl.remove();
    input.removeEventListener("input", clear);
    input.removeEventListener("change", clear);
  };
  input.addEventListener("input", clear);
  input.addEventListener("change", clear);
}

function clearFieldErrors(form) {
  form.querySelectorAll(".input-error").forEach((el) => el.classList.remove("input-error"));
  form.querySelectorAll(".field-error").forEach((el) => el.remove());
}

function validateClassForm(form) {
  clearFieldErrors(form);
  let valid = true;
  const skuInput = form.querySelector('[name="sku"]');
  if (skuInput && !skuInput.value.trim()) {
    showFieldError(skuInput, "請填寫班別");
    valid = false;
  }
  const levelSel = form.querySelector('[name="level"]');
  if (levelSel && !levelSel.value) {
    showFieldError(levelSel, "請選擇等級");
    valid = false;
  }
  const startDateInput = form.querySelector('[name="start_date"]');
  if (startDateInput && !startDateInput.value) {
    showFieldError(startDateInput, "請填寫開課日期");
    valid = false;
  }
  return valid;
}

const exportSettingsBtn = document.getElementById("exportSettingsBtn");
const importSettingsInput = document.getElementById("importSettingsInput");
const teacherGenderFilter = document.getElementById("teacherGenderFilter");
const holidayStartDate = document.getElementById("holidayStartDate");
const holidayEndDate = document.getElementById("holidayEndDate");
const holidayOneDay = document.getElementById("holidayOneDay");

const taskInput = document.getElementById("taskInput");
const taskTimeInput = document.getElementById("taskTimeInput");
const taskAddBtn = document.getElementById("taskAddBtn");
const taskList = document.getElementById("taskList");
const taskDoneList = document.getElementById("taskDoneList");
const taskDoneCount = document.getElementById("taskDoneCount");
const taskDoneToggle = document.getElementById("taskDoneToggle");
const taskClearDone = document.getElementById("taskClearDone");

async function loadState() {
  const state = await invoke('load_state');
  appState = {
    app_config: { location: "" },
    ...state,
    classes: (state.classes || []).map((cls) => ({
      ...cls,
      doorplate_done: toBool(cls.doorplate_done),
      questionnaire_done: toBool(cls.questionnaire_done),
      intro_done: toBool(cls.intro_done),
    })),
  };
  if (locationSelect) {
    locationSelect.value = appState.app_config?.location || "";
  }
  updateLocationBanner();
  updateLocationUI(appState.app_config?.location || "");
  const savedZoom = parseFloat(appState.app_config?.zoom_level);
  if (savedZoom && savedZoom >= ZOOM_MIN && savedZoom <= ZOOM_MAX) {
    applyZoom(savedZoom);
  }
  applyTabOrder();
  renderFilters();
  renderClasses();
  renderHolidays();
  renderSettings();
  renderReminders();
  renderCalendar();
  renderFeeGuide();
  renderDocxClasses();
  updateDocxPreview();
  renderStockTab();
  renderPromoteClassSelects();
}

function renderFilters() {
  const locations = new Set(appState.classes.map((c) => c.location).filter(Boolean));
  const levels = new Set(appState.classes.map((c) => c.level).filter(Boolean));

  locationFilter.innerHTML = '<option value="">全部地點</option>';

  locations.forEach((loc) => {
    const option = document.createElement("option");
    option.value = loc;
    option.textContent = loc;
    locationFilter.appendChild(option);
  });

  levelFilter.innerHTML = '<option value="">全部等級</option>';
  levels.forEach((lvl) => {
    const option = document.createElement("option");
    option.value = lvl;
    option.textContent = lvl;
    levelFilter.appendChild(option);
  });

  if (reviewLevelFilter) {
    reviewLevelFilter.innerHTML = '<option value="">全部等級</option>';
    levels.forEach((lvl) => {
      const option = document.createElement("option");
      option.value = lvl;
      option.textContent = lvl;
      reviewLevelFilter.appendChild(option);
    });
  }
}

function renderSettings() {
  populateSelect(teacherSelect, appState.settings.teacher);
  populateSelect(relayTeacherSelect, appState.settings.teacher);
  populateSelect(roomSelect, appState.settings.room);
  populateSelect(levelSelect, appState.settings.level);
  populateSelect(timeSelect, appState.settings.time);
  setDefaultTime();
  relayFields.classList.toggle("hidden", levelSelect.value !== "初級");

  const genderFilter = teacherGenderFilter.value;
  const filteredTeachers = filterTeachers(appState.settings.teacher, genderFilter);

  renderSettingList(teacherList, "teacher", filteredTeachers);
  renderSettingList(roomList, "room", appState.settings.room);
  renderSettingList(levelList, "level", appState.settings.level);
  renderSettingList(timeList, "time", appState.settings.time);
  if (priceLevelSelect) {
    populateSelect(priceLevelSelect, appState.settings.level, priceLevelSelect.value || "");
  }
  if (feeLevelSelect) {
    populateSelect(feeLevelSelect, appState.settings.level, feeLevelSelect.value || "");
  }
  initFeeGuideSelectors();
  renderPriceSettings();
  renderTextbookSettings();
}

function renderDocxTemplates() {
  if (!docxTemplateSelect) return;
  const labelMap = {
    "class.docx": "列印門牌",
    "cs_sat.docx": "列印Cover (SAT)",
    "cs_weekday.docx": "列印Cover (WD)",
  };
  docxTemplateSelect.innerHTML = '<option value="">請選擇</option>';
  docxTemplates.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = labelMap[name] || name;
    docxTemplateSelect.appendChild(option);
  });
}

function renderDocxClasses() {
  if (!docxClassSelect) return;
  docxClassSelect.innerHTML = '<option value="">請選擇</option>';
  const sorted = [...appState.classes].sort((a, b) => (a.sku || "").localeCompare(b.sku || ""));
  sorted.forEach((cls) => {
    const option = document.createElement("option");
    option.value = cls.id;
    option.textContent = cls.sku;
    docxClassSelect.appendChild(option);
  });
  renderDocxSecondaryClasses();
}

function renderDocxSecondaryClasses() {
  if (!docxClassSelectSecondary) return;
  const primaryId = docxClassSelect?.value || "";
  docxClassSelectSecondary.innerHTML = '<option value="">請選擇</option>';
  const sorted = [...appState.classes].sort((a, b) => (a.sku || "").localeCompare(b.sku || ""));
  sorted.forEach((cls) => {
    if (primaryId && cls.id === primaryId) return;
    const option = document.createElement("option");
    option.value = cls.id;
    option.textContent = cls.sku;
    docxClassSelectSecondary.appendChild(option);
  });
}

function fillDocxPreview(container, cls, useRelay) {
  if (!container) return;
  const teacherName =
    useRelay && cls?.level === "初級" && cls?.relay_teacher ? cls.relay_teacher : cls?.teacher;
  const fields = {
    sku: formatDocxSku(cls?.sku),
    weekday: cls ? `星期${weekdays[cls.weekday] || "-"}` : "-",
    time: formatDocxTime(cls?.start_time),
    teacher: teacherName || "-",
    room: cls?.classroom || "-",
  };
  container.querySelectorAll("[data-field]").forEach((item) => {
    const key = item.dataset.field;
    item.textContent = fields[key] || "-";
  });
}

function updateDocxPreview() {
  if (!docxClassSelect) return;
  const classId = docxClassSelect.value;
  const cls = appState.classes.find((item) => item.id === classId);
  const useRelayPrimary = !!docxRelayTeacherPrimary?.checked;
  fillDocxPreview(docxPreview, cls, useRelayPrimary);
  if (docxSecondClassToggle?.checked) {
    const secondId = docxClassSelectSecondary?.value || "";
    const secondCls = appState.classes.find((item) => item.id === secondId);
    const useRelaySecondary = !!docxRelayTeacherSecondary?.checked;
    fillDocxPreview(docxPreviewSecondary, secondCls, useRelaySecondary);
  } else {
    fillDocxPreview(docxPreviewSecondary, null, false);
  }
  if (docxRelayRowPrimary) {
    const isBeginner = cls?.level === "初級";
    docxRelayRowPrimary.classList.toggle("hidden", !isBeginner);
    if (!isBeginner && docxRelayTeacherPrimary) {
      docxRelayTeacherPrimary.checked = false;
    }
    const docxAdvanced = document.getElementById("docxAdvancedOptions");
    if (docxAdvanced && isBeginner) docxAdvanced.open = true;
  }
  if (docxRelayRowSecondary) {
    const secondId = docxClassSelectSecondary?.value || "";
    const secondCls = appState.classes.find((item) => item.id === secondId);
    const isBeginnerSecondary = secondCls?.level === "初級";
    const showSecondary = !!docxSecondClassToggle?.checked && isBeginnerSecondary;
    docxRelayRowSecondary.classList.toggle("hidden", !showSecondary);
    if (!showSecondary && docxRelayTeacherSecondary) {
      docxRelayTeacherSecondary.checked = false;
    }
  }
}

function renderMessageFilters() {
  if (!messageCategorySelect) return;
  const categories = new Set();
  messageTemplates.forEach((item) => {
    if (item.category) categories.add(item.category);
  });
  messageCategorySelect.innerHTML = '<option value="">全部分類</option>';
  Array.from(categories)
    .sort((a, b) => a.localeCompare(b))
    .forEach((category) => {
      const option = document.createElement("option");
      option.value = category;
      option.textContent = category;
      messageCategorySelect.appendChild(option);
    });
}

function renderMessageList() {
  if (!messageList) return;
  const query = (messageSearchInput?.value || "").trim().toLowerCase();
  const category = messageCategorySelect?.value || "";
  const filtered = messageTemplates.filter((item) => {
    if (category && item.category !== category) return false;
    if (!query) return true;
    return item.label.toLowerCase().includes(query) || item.name.toLowerCase().includes(query);
  });
  messageList.innerHTML = "";
  if (!filtered.length) {
    const empty = document.createElement("div");
    empty.className = "list-item";
    empty.textContent = "沒有符合的訊息。";
    messageList.appendChild(empty);
    return;
  }
  filtered.forEach((item) => {
    const row = document.createElement("div");
    row.className = "list-item message-item";
    row.dataset.name = item.name;
    row.innerHTML = `
      <div class="message-meta">
        <strong>${item.label}</strong>
        <span class="pill">${item.category || "未分類"}</span>
      </div>
      <div class="message-actions">
        <button class="btn secondary" type="button" data-action="edit-category">分類</button>
        <button class="btn primary" type="button" data-action="select-message">選擇</button>
      </div>
    `;
    messageList.appendChild(row);
  });
}

function updateMakeupOutput() {
  if (!makeupOutput) return;
  const rows = Array.from(makeupRows?.querySelectorAll(".makeup-row") || []).map(getMakeupRowData);
  if (!rows.length) {
    makeupOutput.value = makeupTemplateBase || "";
    return;
  }

  const lineReplacer = (templateLine, data) => {
    const replacements = {
      SCHOOL: data.school,
      學校: data.school,
      DATE: data.dateText,
      日期: data.dateText,
      WEEKDAY: data.weekdayText,
      星期: data.weekdayText,
      TIME: data.timeText,
      時間: data.timeText,
      CLASS_NO: data.classCount,
      回數: data.classCount,
      第回: data.classCount,
      班次: data.classCount,
    };
    let line = templateLine;
    Object.entries(replacements).forEach(([key, value]) => {
      const pattern = new RegExp(`\\{\\{?${key}\\}?\\}`, "g");
      line = line.replace(pattern, value || "");
    });
    return line;
  };

  const content = (makeupTemplateBase || "").split(/\r?\n/);
  const placeholderLineIndex = content.findIndex((line) => /\{(?:學校|SCHOOL|日期|DATE|星期|WEEKDAY|時間|TIME|班次|CLASS_NO)\}/.test(line));
  const lines = [...content];
  if (placeholderLineIndex >= 0) {
    const templateLine = lines[placeholderLineIndex];
    const replacedLines = rows.map((row) => lineReplacer(templateLine, row));
    lines.splice(placeholderLineIndex, 1, ...replacedLines);
  } else {
    const replacedLines = rows.map((row) => lineReplacer("{學校} {日期} {星期} {時間} {班次}", row));
    lines.push(...replacedLines);
  }
  makeupOutput.value = lines.join("\n");
}

async function loadMakeupTemplate() {
  const response = await invoke('load_makeup_template');
  if (!response.ok) {
    if (makeupOutput) {
      makeupOutput.value = response.error || "無法載入補課安排模板。";
    }
    return;
  }
  makeupTemplateBase = response.content || "";
  if (makeupRows && !makeupRows.querySelector(".makeup-row")) {
    initMakeupRows();
  }
  updateMakeupOutput();
}

function resetMakeupForm() {
  initMakeupRows();
  updateMakeupOutput();
}

function updateLocationBanner() {
  if (!locationBanner) return;
  locationBanner.classList.remove("loc-k", "loc-l", "loc-h");
  const locationCode = (appState.app_config?.location || "").toUpperCase();
  if (locationCode === "K") {
    locationBanner.classList.add("loc-k");
  } else if (locationCode === "L") {
    locationBanner.classList.add("loc-l");
  } else if (locationCode === "H") {
    locationBanner.classList.add("loc-h");
  }
}

function updateLocationUI(value) {
  const code = (value || "").toUpperCase();
  // Sync segmented buttons
  if (locationSegments) {
    locationSegments.querySelectorAll(".loc-seg").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.value === code);
    });
  }
  // Sync location badge (collapsed sidebar)
  if (locationBadge) {
    locationBadge.classList.remove("loc-k", "loc-l", "loc-h");
    const labels = { K: "旺", L: "太", H: "港" };
    locationBadge.textContent = labels[code] || "";
    if (code) locationBadge.classList.add(`loc-${code.toLowerCase()}`);
  }
}

function applyTabOrder() {
  if (!tabNav) return;
  const orderRaw = appState.app_config?.tab_order || "";
  const order = orderRaw
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  if (!order.length) return;
  // Pinned tabs always stay first — move them to front before reordering the rest
  const pinnedButtons = Array.from(tabNav.querySelectorAll(".tab-button[data-pinned]"));
  pinnedButtons.forEach((btn) => tabNav.insertBefore(btn, tabNav.firstChild));
  // Ensure the divider after pinned tabs stays right after them
  const divider = tabNav.querySelector(".tab-divider");
  if (divider && pinnedButtons.length) {
    const lastPinned = pinnedButtons[pinnedButtons.length - 1];
    lastPinned.after(divider);
  }
  // Reorder non-pinned tabs according to saved order
  const buttons = Array.from(tabNav.querySelectorAll(".tab-button:not([data-pinned])"));
  const map = new Map(buttons.map((btn) => [btn.dataset.tab, btn]));
  order
    .filter((key) => map.has(key))
    .forEach((key) => {
      tabNav.appendChild(map.get(key));
      map.delete(key);
    });
  Array.from(map.values()).forEach((button) => tabNav.appendChild(button));
}

function setTabReorderState(enabled) {
  if (!tabNav) return;
  // Remove any existing reorder arrows
  tabNav.querySelectorAll(".tab-reorder-arrows").forEach((el) => el.remove());
  if (!enabled) return;
  const buttons = Array.from(tabNav.querySelectorAll(".tab-button:not([data-pinned])"));
  buttons.forEach((button, i) => {
    const arrows = document.createElement("span");
    arrows.className = "tab-reorder-arrows";
    const upBtn = document.createElement("button");
    upBtn.className = "tab-reorder-btn";
    upBtn.textContent = "▲";
    upBtn.disabled = i === 0;
    upBtn.addEventListener("click", (e) => { e.stopPropagation(); moveTab(button, -1); });
    const downBtn = document.createElement("button");
    downBtn.className = "tab-reorder-btn";
    downBtn.textContent = "▼";
    downBtn.disabled = i === buttons.length - 1;
    downBtn.addEventListener("click", (e) => { e.stopPropagation(); moveTab(button, 1); });
    arrows.appendChild(upBtn);
    arrows.appendChild(downBtn);
    button.appendChild(arrows);
  });
}

async function moveTab(button, direction) {
  const buttons = Array.from(tabNav.querySelectorAll(".tab-button:not([data-pinned])"));
  const idx = buttons.indexOf(button);
  const targetIdx = idx + direction;
  if (targetIdx < 0 || targetIdx >= buttons.length) return;
  const target = buttons[targetIdx];
  if (direction === -1) {
    tabNav.insertBefore(button, target);
  } else {
    tabNav.insertBefore(target, button);
  }
  // Refresh arrows to update disabled states
  setTabReorderState(true);
  // Save order
  const order = Array.from(tabNav.querySelectorAll(".tab-button:not([data-pinned])")).map((btn) => btn.dataset.tab);
  const response = await invoke('set_tab_order', { order });
  if (!response.ok) {
    showToast(response.error || "更新排序失敗。");
  }
}

function updateTabOrderToggleLabel() {
  if (!tabOrderToggle) return;
  tabOrderToggle.textContent = tabOrderUnlocked ? "🔓 解鎖排序" : "🔒 鎖定排序";
}

function populateSelect(select, values, selectedValue = "") {
  select.innerHTML = "";
  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = "請選擇";
  select.appendChild(emptyOption);
  const normalizedValues = Array.isArray(values) ? values : [];
  if (selectedValue && !normalizedValues.includes(selectedValue)) {
    const option = document.createElement("option");
    option.value = selectedValue;
    option.textContent = selectedValue;
    select.appendChild(option);
  }
  normalizedValues.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (selectedValue && normalizedValues.includes(selectedValue)) {
    select.value = selectedValue;
  }
}

function setDefaultTime() {
  if (!timeSelect.value) {
    const option = Array.from(timeSelect.options).find((item) => item.value === "1900-2100");
    if (option) {
      timeSelect.value = "1900-2100";
    }
  }
}

function renderSettingList(container, entryType, values) {
  container.innerHTML = "";
  values.forEach((value) => {
    const item = document.createElement("div");
    item.className = "list-item";
    item.innerHTML = `
      <div>${value}</div>
      <div class="item-actions">
        <button class="btn secondary" data-setting-type="${entryType}" data-setting-value="${value}" data-move="up">上移</button>
        <button class="btn secondary" data-setting-type="${entryType}" data-setting-value="${value}" data-move="down">下移</button>
        <button class="btn danger" data-setting-type="${entryType}" data-setting-value="${value}" data-delete="true">刪除</button>
      </div>
    `;
    container.appendChild(item);
  });
}

function renderPriceSettings() {
  if (!priceList) return;
  priceList.innerHTML = "";
  const levelPrices = appState.settings?.level_price || {};
  const levelOrder = Array.isArray(appState.settings.level) ? appState.settings.level : [];
  const seen = new Set();
  const rows = [];
  levelOrder.forEach((level) => {
    seen.add(level);
    rows.push({ level, price: levelPrices[level] });
  });
  Object.keys(levelPrices)
    .filter((level) => !seen.has(level))
    .sort((a, b) => a.localeCompare(b))
    .forEach((level) => rows.push({ level, price: levelPrices[level] }));
  if (!rows.length) {
    const empty = document.createElement("div");
    empty.className = "list-item";
    empty.innerHTML = "<div>尚未設定學費</div><div>—</div>";
    priceList.appendChild(empty);
    return;
  }
  rows.forEach((row) => {
    const item = document.createElement("div");
    item.className = "list-item";
    const priceLabel = Number.isFinite(row.price) ? `$${row.price}` : "未設定";
    item.innerHTML = `<div>${row.level}</div><div>${priceLabel}</div>`;
    priceList.appendChild(item);
  });
}

// ─── Textbook Settings ───────────────────────────────────────────────────────

function renderTextbookSettings() {
  renderTextbookList();
  renderLevelTextbookMapping();
}

function renderTextbookList() {
  if (!textbookList) return;
  textbookList.innerHTML = "";
  const textbooks = appState.settings?.textbook || {};
  const names = Object.keys(textbooks);
  if (!names.length) {
    const empty = document.createElement("div");
    empty.className = "list-item";
    empty.innerHTML = "<div>尚未新增教材</div><div>—</div>";
    textbookList.appendChild(empty);
    return;
  }
  names.forEach((name) => {
    const price = textbooks[name];
    const item = document.createElement("div");
    item.className = "list-item";
    item.innerHTML = `
      <div>${name}</div>
      <div>$${price}</div>
      <div class="action-group">
        <button class="btn secondary btn-sm" data-textbook-edit="${name}" data-price="${price}" type="button">編輯</button>
        <button class="btn danger btn-sm" data-textbook-delete="${name}" type="button">刪除</button>
      </div>`;
    textbookList.appendChild(item);
  });
}

function renderLevelTextbookMapping() {
  if (!levelTextbookList) return;
  levelTextbookList.innerHTML = "";
  const levels = appState.settings?.level || [];
  const levelTextbook = appState.settings?.level_textbook || {};
  const levelNext = appState.settings?.level_next || {};
  const textbookNames = Object.keys(appState.settings?.textbook || {});

  if (!levels.length) {
    levelTextbookList.innerHTML = '<p class="muted-hint">請先在「等級」設定中新增等級。</p>';
    return;
  }
  levels.forEach((level) => {
    const row = document.createElement("div");
    row.className = "level-textbook-row";

    const selectedBooks = Array.isArray(levelTextbook[level]) ? levelTextbook[level] : [];
    const checksHtml = textbookNames.length
      ? textbookNames.map((n) =>
          `<label class="ltb-check-label">
            <input type="checkbox" data-ltb-level="${level}" data-ltb-book="${n}" ${selectedBooks.includes(n) ? "checked" : ""}> ${n}
          </label>`
        ).join("")
      : '<span class="muted-hint" style="font-size:12px;">（尚未有教材）</span>';

    const nextOptions = levels
      .filter((l) => l !== level)
      .map((l) => `<option value="${l}" ${levelNext[level] === l ? "selected" : ""}>${l}</option>`)
      .join("");

    row.innerHTML = `
      <span class="level-label">${level}</span>
      <div class="ltb-checks-wrap">
        <span class="inline-label-text">教材</span>
        <div class="ltb-checks" data-ltb-level-group="${level}">${checksHtml}</div>
      </div>
      <label class="inline-label">下一等級
        <select data-lnext-level="${level}">
          <option value="">（無）</option>
          ${nextOptions}
        </select>
      </label>`;
    levelTextbookList.appendChild(row);
  });
}

// ─── Stock Tab ───────────────────────────────────────────────────────────────

function renderStockTab() {
  renderDataReviewBanner();
  renderInventoryTable();
  renderPromotionPlanningTable();
}

function renderDataReviewBanner() {
  if (!stockReviewBanner) return;
  const lastTs = appState.app_config?.last_review_ts || "";
  let daysSince = null;
  if (lastTs) {
    const last = new Date(lastTs);
    if (!isNaN(last)) {
      daysSince = Math.floor((Date.now() - last.getTime()) / 86400000);
    }
  }
  if (daysSince === null) {
    stockReviewBanner.innerHTML = `<div class="review-banner review-banner-warn">⚠ 尚未確認過課堂人數，建議進行核實以確保升班預測準確。</div>`;
  } else if (daysSince > 7) {
    stockReviewBanner.innerHTML = `<div class="review-banner review-banner-warn">⚠ 上次確認課堂人數：${daysSince} 天前，建議重新核實。</div>`;
  } else {
    stockReviewBanner.innerHTML = `<div class="review-banner review-banner-ok">✓ 課堂人數已於 ${daysSince} 天前確認。</div>`;
  }
}

function renderInventoryTable() {
  if (!inventoryTable) return;
  inventoryTable.innerHTML = "";
  const textbooks = appState.settings?.textbook || {};
  const liveStock = appState.settings?.textbook_stock || {};
  const history = appState.stock_history || {};
  const names = Object.keys(textbooks);
  if (!names.length) {
    inventoryTable.innerHTML = '<p class="muted-hint">尚未設定教材。請在「資料設定 → 教材管理」中新增。</p>';
    return;
  }

  // Month selector + snapshot controls
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const availableMonths = Object.keys(history).sort().reverse();

  const controls = document.createElement("div");
  controls.className = "stock-month-controls";
  controls.innerHTML = `
    <div class="stock-month-row">
      <label class="inline-label">
        檢視月份
        <select id="stockMonthSelect" class="stock-month-select">
          <option value="">目前存貨</option>
          ${availableMonths.map((m) => `<option value="${m}">${m}</option>`).join("")}
        </select>
      </label>
      <button id="stockSnapshotBtn" class="btn secondary btn-sm" type="button">
        儲存本月快照 (${currentMonth})
      </button>
    </div>`;
  inventoryTable.appendChild(controls);

  // Build the data table (editable for live stock, read-only for historical)
  function buildTable(displayStock, readOnly) {
    const table = document.createElement("table");
    table.className = "stock-table";
    table.id = "stockDataTable";
    if (readOnly) {
      table.innerHTML = `<thead><tr><th>教材</th><th>單價</th><th>該月存貨</th></tr></thead>`;
    } else {
      table.innerHTML = `<thead><tr><th>教材</th><th>單價</th><th>現存數量</th><th>操作</th></tr></thead>`;
    }
    const tbody = document.createElement("tbody");
    names.forEach((name) => {
      const price = textbooks[name];
      const count = displayStock[name] ?? (readOnly ? "—" : 0);
      const tr = document.createElement("tr");
      if (readOnly) {
        tr.innerHTML = `<td>${name}</td><td>$${price}</td><td>${count}</td>`;
      } else {
        tr.innerHTML = `
          <td>${name}</td>
          <td>$${price}</td>
          <td>
            <div class="count-cell">
              <button class="btn secondary btn-sm" data-stock-minus="${name}" type="button">－</button>
              <input type="number" class="stock-count-input" data-stock-name="${name}" value="${count}" min="0" />
              <button class="btn secondary btn-sm" data-stock-plus="${name}" type="button">＋</button>
            </div>
          </td>
          <td><button class="btn primary btn-sm" data-stock-save="${name}" type="button">儲存</button></td>`;
      }
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    return table;
  }

  inventoryTable.appendChild(buildTable(liveStock, false));

  // Month select handler
  controls.querySelector("#stockMonthSelect").addEventListener("change", (e) => {
    const selected = e.target.value;
    const existing = document.getElementById("stockDataTable");
    if (existing) existing.remove();
    if (!selected) {
      inventoryTable.appendChild(buildTable(liveStock, false));
    } else {
      inventoryTable.appendChild(buildTable(history[selected] || {}, true));
    }
  });

  // Snapshot button handler
  controls.querySelector("#stockSnapshotBtn").addEventListener("click", async () => {
    const response = await invoke('save_monthly_stock', { month: currentMonth, stockData: liveStock });
    if (response.ok) {
      appState.stock_history[currentMonth] = { ...liveStock };
      showToast(`已儲存 ${currentMonth} 月份快照`, "success");
      renderInventoryTable();
    } else {
      showToast(response.error || "儲存失敗。", "error");
    }
  });
}

function renderPromotionPlanningTable() {
  if (!promotionTable) return;
  promotionTable.innerHTML = "";
  const threshold = parseInt(promotionThreshold?.value || "3", 10);
  const textbooks = appState.settings?.textbook || {};
  const stock = appState.settings?.textbook_stock || {};
  const levelTextbook = appState.settings?.level_textbook || {};
  const levelNext = appState.settings?.level_next || {};

  const archivedStatuses = new Set(["promoted", "merged", "terminated", "ended"]);
  const endingSoon = (appState.classes || []).filter(
    (cls) =>
      !archivedStatuses.has(cls.status) &&
      cls.status !== "not_started" &&
      typeof cls.lessons_remaining === "number" &&
      cls.lessons_remaining > 0 &&
      cls.lessons_remaining <= threshold
  );

  if (!endingSoon.length) {
    promotionTable.innerHTML = `<p class="muted-hint">目前沒有剩餘堂數 ≤ ${threshold} 的進行中班別。</p>`;
    return;
  }

  // Build demand map: textbookName -> total students needing it
  const demandMap = {}; // textbookName -> needed count
  endingSoon.forEach((cls) => {
    const nextLevel = levelNext[cls.level] || "";
    const nextBooks = nextLevel ? (Array.isArray(levelTextbook[nextLevel]) ? levelTextbook[nextLevel] : []) : [];
    nextBooks.forEach((bookName) => {
      demandMap[bookName] = (demandMap[bookName] || 0) + (parseInt(cls.student_count, 10) || 0);
    });
  });

  const table = document.createElement("table");
  table.className = "stock-table";
  table.innerHTML = `<thead><tr><th>班別</th><th>等級</th><th>剩餘堂</th><th>學生人數</th><th>升班等級</th><th>所需教材</th></tr></thead>`;
  const tbody = document.createElement("tbody");

  endingSoon.forEach((cls) => {
    const nextLevel = levelNext[cls.level] || "—";
    const nextBooks = nextLevel !== "—" ? (Array.isArray(levelTextbook[nextLevel]) ? levelTextbook[nextLevel] : []) : [];
    const booksCell = nextBooks.length ? nextBooks.join("、") : "—";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${cls.sku || "—"}</td>
      <td>${cls.level || "—"}</td>
      <td>${cls.lessons_remaining}</td>
      <td>${cls.student_count || 0}</td>
      <td>${nextLevel}</td>
      <td>${booksCell}</td>`;
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  promotionTable.appendChild(table);

  // Formatted summary table per textbook: demand vs. stock
  const summaryEntries = Object.entries(demandMap).map(([bookName, needed]) => {
    const avail = stock[bookName] ?? 0;
    const diff = avail - needed;
    return { bookName, needed, avail, diff };
  });

  if (summaryEntries.length) {
    const summaryDiv = document.createElement("div");
    summaryDiv.className = "promotion-summary";

    const headerRow = document.createElement("div");
    headerRow.className = "summary-header-row";
    headerRow.innerHTML = `<strong>匯總</strong><button class="btn secondary btn-sm" id="copySummaryBtn" type="button">複製匯總</button>`;
    summaryDiv.appendChild(headerRow);

    const summaryTable = document.createElement("table");
    summaryTable.className = "stock-table summary-table";
    summaryTable.innerHTML = `<thead><tr><th>教材</th><th>需求量</th><th>現存量</th><th>狀態</th></tr></thead>`;
    const sTbody = document.createElement("tbody");
    summaryEntries.forEach(({ bookName, needed, avail, diff }) => {
      const statusCell = diff >= 0
        ? `<span class="stock-ok">充足（+${diff}）</span>`
        : `<span class="stock-danger">缺 ${Math.abs(diff)} 本</span>`;
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${bookName}</td><td>${needed} 本</td><td>${avail} 本</td><td>${statusCell}</td>`;
      sTbody.appendChild(tr);
    });
    summaryTable.appendChild(sTbody);
    summaryDiv.appendChild(summaryTable);
    promotionTable.appendChild(summaryDiv);

    summaryDiv.querySelector("#copySummaryBtn").addEventListener("click", () => {
      const lines = summaryEntries.map(({ bookName, needed, avail, diff }) => {
        const tag = diff >= 0 ? `充足（+${diff}）` : `缺 ${Math.abs(diff)} 本`;
        return `${bookName}：需 ${needed} 本，現存 ${avail} 本 → ${tag}`;
      });
      const text = "【升班教材匯總】\n" + lines.join("\n");
      navigator.clipboard.writeText(text).then(() => {
        showToast("匯總已複製到剪貼簿", "success");
      }).catch(() => {
        showToast("複製失敗，請手動複製。", "error");
      });
    });
  }
}

function renderQuickReviewPanel(show) {
  if (!stockQuickReview) return;
  if (!show) {
    stockQuickReview.classList.add("hidden");
    return;
  }
  stockQuickReview.classList.remove("hidden");
  if (!stockReviewBody) return;
  stockReviewBody.innerHTML = "";

  const weekdayVal = reviewWeekdayFilter ? reviewWeekdayFilter.value : "";
  const levelVal = reviewLevelFilter ? reviewLevelFilter.value : "";
  const levelOrder = Array.isArray(appState.settings?.level) ? appState.settings.level : [];

  let activeClasses = (appState.classes || []).filter((c) => c.status === "active");
  if (weekdayVal) activeClasses = activeClasses.filter((c) => String(c.weekday) === weekdayVal);
  if (levelVal) activeClasses = activeClasses.filter((c) => c.level === levelVal);

  activeClasses.sort((a, b) => {
    if (a.weekday !== b.weekday) return a.weekday - b.weekday;
    const aIdx = levelOrder.indexOf(a.level);
    const bIdx = levelOrder.indexOf(b.level);
    return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
  });

  activeClasses.forEach((cls) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${cls.sku || "—"}</td>
      <td>${cls.level || "—"}</td>
      <td><input type="number" class="review-count-input" data-class-id="${cls.id}" value="${cls.student_count || 0}" min="0" /></td>`;
    stockReviewBody.appendChild(tr);
  });
}

async function saveReviewTimestamp() {
  const ts = new Date().toISOString();
  const resp = await invoke('set_last_review_ts', { ts });
  if (resp?.ok) {
    appState.app_config = appState.app_config || {};
    appState.app_config.last_review_ts = ts;
    renderDataReviewBanner();
    renderQuickReviewPanel(false);
    showToast("課堂人數已確認。", "success");
  }
}

function filterTeachers(values, gender) {
  if (!gender) return values;
  if (gender === "male") {
    return values.filter((value) => value.endsWith("先生"));
  }
  if (gender === "female") {
    return values.filter((value) => value.endsWith("小姐"));
  }
  return values;
}

function renderClasses() {
  const query = searchInput.value.trim().toLowerCase();
  const status = statusFilter.value;
  const location = locationFilter.value;
  const level = levelFilter.value;
  const weekday = weekdayFilter.value;
  const archive = archiveFilter ? archiveFilter.value : "active";
  const issue = issueFilter ? issueFilter.value : "";
  const today = new Date();
  const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  const skuCounts = new Map();
  appState.classes.forEach((cls) => {
    const key = (cls.sku || "").trim();
    if (!key) return;
    skuCounts.set(key, (skuCounts.get(key) || 0) + 1);
  });

  const conflictMap = new Map();
  appState.classes.forEach((cls) => {
    if (!cls.classroom || !cls.start_time) return;
    const key = `${cls.classroom}|${cls.weekday}|${cls.start_time}`;
    if (!conflictMap.has(key)) {
      conflictMap.set(key, []);
    }
    const start = parseDate(cls.start_date);
    const end = parseDate(cls.end_date) || new Date(9999, 11, 31);
    if (!start) return;
    conflictMap.get(key).push({
      id: cls.id,
      start,
      end,
    });
  });

  const filtered = appState.classes.filter((cls) => {
    const derivedStatus = getDerivedStatus(cls);
    const isArchived = ["ended", "terminated", "merged", "promoted"].includes(derivedStatus);
    if (archive === "active" && isArchived) return false;
    if (archive === "archived" && !isArchived) return false;
    if (status && derivedStatus !== status) return false;
    if (location && cls.location !== location) return false;
    if (level && cls.level !== level) return false;
    if (weekday && String(cls.weekday) !== weekday) return false;
    if (query) {
      const haystack = `${cls.sku} ${cls.teacher} ${cls.classroom}`.toLowerCase();
      if (!haystack.includes(query)) return false;
    }
    if (issue) {
      const issues = getClassIssues(cls, {
        todayDate,
        skuCounts,
        conflictMap,
      });
      if (!issues.some((item) => item.key === issue)) return false;
    }
    return true;
  });

  classBody.innerHTML = "";
  filtered.forEach((cls) => {
    const row = document.createElement("tr");
    row.dataset.id = cls.id;
    const derivedStatus = getDerivedStatus(cls);
    const statusClass = derivedStatus === "ended" ? "ended" : derivedStatus === "active" ? "active" : "other";
    const statusLabel = derivedStatus;
    const endButton = cls.lessons_remaining === 0
      ? `<button class="btn secondary" data-action="end" data-id="${cls.id}">結束</button>`
      : "";
    const issues = getClassIssues(cls, {
      todayDate,
      skuCounts,
      conflictMap,
    });
    const issueHtml = issues.length
      ? issues.map((item) => `<span class="issue-badge ${item.key}">${item.label}</span>`).join("")
      : "-";

    row.innerHTML = `
      <td>${cls.sku}</td>
      <td>${cls.classroom || "-"}</td>
      <td>${cls.start_date || "-"}</td>
      <td>星期${weekdays[cls.weekday] || "-"}</td>
      <td>${formatDisplayTime(cls.start_time)}</td>
      <td>${cls.teacher || "-"}</td>
      <td>${cls.student_count}</td>
      <td>${cls.lesson_total}</td>
      <td>${cls.lessons_elapsed}</td>
      <td>${cls.lessons_remaining}</td>
      <td>${cls.end_date || "-"}</td>
      <td><span class="status ${statusClass}">${statusLabels[statusLabel] || statusLabel}</span></td>
      <td>${issueHtml}</td>
      <td>
        <div class="action-group">
          <button class="btn secondary" data-action="postpone" data-id="${cls.id}">改期</button>
          <button class="btn secondary" data-action="schedule" data-id="${cls.id}">日程</button>
          ${endButton}
        </div>
      </td>
    `;
    classBody.appendChild(row);
  });
}

function getClassIssues(cls, context) {
  const issues = [];
  const startDate = parseDate(cls.start_date);
  const hasStarted = startDate && startDate <= context.todayDate;
  const missingInfo = !cls.teacher || !cls.start_time || !cls.classroom || !cls.location;
  if (missingInfo) {
    issues.push({ key: "incomplete", label: "缺資料" });
  }
  const skuKey = (cls.sku || "").trim();
  if (skuKey && context.skuCounts.get(skuKey) > 1) {
    issues.push({ key: "duplicate", label: "重複" });
  }
  if (cls.classroom && cls.start_time) {
    const conflictKey = `${cls.classroom}|${cls.weekday}|${cls.start_time}`;
    const group = context.conflictMap.get(conflictKey) || [];
    const start = parseDate(cls.start_date);
    const end = parseDate(cls.end_date) || new Date(9999, 11, 31);
    if (start) {
      const hasOverlap = group.some((item) => {
        if (item.id === cls.id) return false;
        return item.start <= end && start <= item.end;
      });
      if (hasOverlap) {
        issues.push({ key: "conflict", label: "課室衝突" });
      }
    }
  }
  return issues;
}

function renderHolidays() {
  holidayList.innerHTML = "";
  const sorted = [...appState.holidays].sort((a, b) => (a.start_date || "").localeCompare(b.start_date || ""));
  const groups = new Map();
  sorted.forEach((holiday) => {
    const year = (holiday.start_date || holiday.end_date || "").split("-")[0] || "未設定";
    if (!groups.has(year)) {
      groups.set(year, []);
    }
    groups.get(year).push(holiday);
  });
  const years = Array.from(groups.keys()).sort((a, b) => b.localeCompare(a));
  years.forEach((year) => {
    const section = document.createElement("div");
    section.className = "holiday-year";
    section.innerHTML = `<div class="pill">${year}</div>`;
    groups.get(year).forEach((holiday) => {
      const item = document.createElement("div");
      item.className = "list-item";
      item.innerHTML = `
        <div>
          <div>${holiday.start_date} → ${holiday.end_date}</div>
          <div class="pill">${holiday.name || "假期"}</div>
        </div>
        <button class="btn danger" data-holiday-id="${holiday.id}">刪除</button>
      `;
      section.appendChild(item);
    });
    holidayList.appendChild(section);
  });
}

function renderReminders() {
  reminderBody.innerHTML = "";
  const sortKey = reminderSort ? reminderSort.value : "start_date";
  const today = new Date();
  const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const beginnerCutoff = new Date(todayDate);
  beginnerCutoff.setDate(beginnerCutoff.getDate() - 70);
  const nonBeginnerCutoff = new Date(todayDate);
  nonBeginnerCutoff.setDate(nonBeginnerCutoff.getDate() - 35);
  const filteredClasses = appState.classes.filter((cls) => {
    if (cls.level === "初級") {
      const startDate = parseDate(cls.start_date);
      if (!startDate) return true;
      const startDateOnly = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      if (startDateOnly > todayDate) return true;
      return startDateOnly >= beginnerCutoff;
    }
    const startDate = parseDate(cls.start_date);
    if (!startDate) return true;
    const startDateOnly = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    if (startDateOnly > todayDate) return true;
    return startDateOnly >= nonBeginnerCutoff;
  });
  const sortedClasses = [...filteredClasses].sort((a, b) => {
    if (sortKey === "sku") {
      return (a.sku || "").localeCompare(b.sku || "");
    }
    if (sortKey === "classroom") {
      return (a.classroom || "").localeCompare(b.classroom || "");
    }
    if (sortKey === "level") {
      const levelOrder = Array.isArray(appState.settings.level) ? appState.settings.level : [];
      const orderMap = new Map(levelOrder.map((level, index) => [level, index]));
      const rankA = orderMap.has(a.level) ? orderMap.get(a.level) : Number.MAX_SAFE_INTEGER;
      const rankB = orderMap.has(b.level) ? orderMap.get(b.level) : Number.MAX_SAFE_INTEGER;
      if (rankA !== rankB) {
        return rankA - rankB;
      }
      return (a.level || "").localeCompare(b.level || "");
    }
    const dateA = parseDate(a.start_date);
    const dateB = parseDate(b.start_date);
    const timeA = dateA ? dateA.getTime() : Number.MAX_SAFE_INTEGER;
    const timeB = dateB ? dateB.getTime() : Number.MAX_SAFE_INTEGER;
    return timeA - timeB;
  });
  sortedClasses.forEach((cls) => {
    const startDate = parseDate(cls.start_date);
    const relayDate = parseDate(cls.relay_date);
    const doorplate = getReminderBadge(startDate, "門牌", cls.doorplate_done);
    const isBeginner = cls.level === "初級";
    const questionnaire = isBeginner
      ? getReminderBadge(startDate ? addDays(startDate, 21) : null, "問卷", cls.questionnaire_done)
      : "-";
    const introTarget = relayDate ? addDays(relayDate, -7) : null;
    const intro = isBeginner ? getReminderBadge(introTarget, "介紹", cls.intro_done) : "-";
    const questionnaireCell = isBeginner
      ? renderReminderCell(cls.id, "questionnaire", questionnaire, cls.questionnaire_done)
      : questionnaire;
    const introCell = isBeginner ? renderReminderCell(cls.id, "intro", intro, cls.intro_done) : intro;

    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${cls.sku}</td>
      <td>${cls.classroom || "-"}</td>
      <td>${cls.relay_teacher || "-"}</td>
      <td>${cls.start_date || "-"}</td>
      <td>星期${weekdays[cls.weekday] || "-"}</td>
      <td>${formatDisplayTime(cls.start_time)}</td>
      <td>${cls.relay_date || "-"}</td>
      <td>${renderReminderCell(cls.id, "doorplate", doorplate, cls.doorplate_done)}</td>
      <td>${questionnaireCell}</td>
      <td>${introCell}</td>
    `;
    reminderBody.appendChild(row);
  });
}

async function renderCalendar() {
  const { rangeStart, rangeEnd, label } = getCalendarRange(calendarMode, calendarAnchor);
  calendarLabel.textContent = label;
  syncCalendarJumpSelects();
  let response;
  try {
    response = await invoke('get_calendar_data', { startDate: rangeStart, endDate: rangeEnd });
  } catch (err) {
    console.error("renderCalendar invoke error:", err);
    calendarView.innerHTML = `<div>無法載入行事曆: ${err}</div>`;
    return;
  }
  if (!response.ok) {
    calendarView.innerHTML = "<div>無法載入行事曆</div>";
    return;
  }
  const sessionsByDate = new Map();
  response.sessions.forEach((session) => {
    if (!sessionsByDate.has(session.date)) {
      sessionsByDate.set(session.date, []);
    }
    sessionsByDate.get(session.date).push(session);
  });

  const holidays = response.holidays || [];
  if (calendarMode === "month") {
    renderCalendarMonth(rangeStart, sessionsByDate, holidays);
  } else {
    renderCalendarWeek(rangeStart, sessionsByDate, holidays);
  }
}

function syncCalendarJumpSelects() {
  if (!calendarJumpMonth || !calendarJumpYear) return;
  const isMonth = calendarMode === "month";
  calendarLabel.classList.toggle("hidden", isMonth);
  calendarJumpMonth.classList.toggle("hidden", !isMonth);
  calendarJumpYear.classList.toggle("hidden", !isMonth);
  if (!isMonth) return;
  const year = calendarAnchor.getFullYear();
  const month = calendarAnchor.getMonth() + 1;
  if (calendarJumpYear.options.length === 0) {
    for (let y = year - 3; y <= year + 3; y += 1) {
      const opt = document.createElement("option");
      opt.value = String(y);
      opt.textContent = `${y}年`;
      calendarJumpYear.appendChild(opt);
    }
  }
  calendarJumpMonth.value = String(month);
  calendarJumpYear.value = String(year);
}

function renderCalendarMonth(rangeStart, sessionsByDate, holidays) {
  calendarView.className = "calendar grid-month";
  calendarView.innerHTML = "";
  weekdays.forEach((day) => {
    const header = document.createElement("div");
    header.className = "calendar-weekday-header";
    header.textContent = `週${day}`;
    calendarView.appendChild(header);
  });
  const startDate = parseDate(rangeStart);
  const todayStr = toIsoDate(new Date());
  for (let i = 0; i < 42; i += 1) {
    const current = addDays(startDate, i);
    const dateStr = toIsoDate(current);
    const sessions = sessionsByDate.get(dateStr) || [];
    const paymentCount = sessions.filter((session) => session.payment_due).length;
    const barClass = paymentCount === 0 ? "" : paymentCount <= paymentConfig.greenMax ? "green" : paymentCount <= paymentConfig.yellowMax ? "yellow" : "red";
    const holidayLabel = getHolidayLabel(dateStr, holidays);
    const dayCard = document.createElement("div");
    dayCard.className = "calendar-day";
    dayCard.dataset.date = dateStr;
    dayCard.dataset.action = "drill-down";
    if (holidayLabel) {
      dayCard.classList.add("holiday-day");
    }
    if (dateStr === todayStr) {
      dayCard.classList.add("today");
    }
    if (current.getMonth() !== calendarAnchor.getMonth()) {
      dayCard.classList.add("dim");
    }
    dayCard.innerHTML = `
      <div class="day-number">${current.getDate()}</div>
      <div class="day-bar ${barClass}"></div>
      <div class="day-count">課堂 ${sessions.length} 堂</div>
      ${holidayLabel ? `<div class="holiday">${holidayLabel}</div>` : ""}
    `;
    calendarView.appendChild(dayCard);
  }
}

function renderCalendarWeek(rangeStart, sessionsByDate, holidays) {
  calendarView.className = "calendar grid-week";
  calendarView.innerHTML = "";
  const startDate = parseDate(rangeStart);
  const todayStr = toIsoDate(new Date());
  const configuredRooms = Array.isArray(appState.settings.room) ? appState.settings.room.filter(Boolean) : [];
  const weekSessions = [];
  sessionsByDate.forEach((daySessions) => weekSessions.push(...daySessions));
  const fallbackRooms = Array.from(new Set(weekSessions.map((session) => session.room).filter(Boolean))).sort((a, b) => a.localeCompare(b));
  const baseRooms = configuredRooms.length ? configuredRooms : fallbackRooms;
  const hasUnassigned = weekSessions.some((session) => !session.room);
  const roomList = hasUnassigned ? [...baseRooms, "未設定課室"] : baseRooms;
  for (let i = 0; i < 7; i += 1) {
    const current = addDays(startDate, i);
    const dateStr = toIsoDate(current);
    const sessions = sessionsByDate.get(dateStr) || [];
    const paymentCount = sessions.filter((session) => session.payment_due).length;
    const barClass = paymentCount === 0 ? "" : paymentCount <= paymentConfig.greenMax ? "green" : paymentCount <= paymentConfig.yellowMax ? "yellow" : "red";
    const holidayLabel = getHolidayLabel(dateStr, holidays);
    const isSaturday = current.getDay() === 6;
    const dayRoomList = roomList.length ? roomList : Array.from(new Set(sessions.map((session) => session.room).filter(Boolean)));
    const roomBlocks = dayRoomList
      .map((roomName) => {
        const roomLabel = roomName || "未設定課室";
        const roomSessions = sessions.filter((session) => (session.room || "未設定課室") === roomLabel);
        const grouped = groupSessionsByTime(roomSessions);
        const fullItems = renderTimeBlocks(grouped);
        const hasMore = grouped.length > 1;
        const previewItems = hasMore ? renderTimeBlocks(grouped.slice(0, 1)) : fullItems;
        const showToggle = isSaturday && hasMore;
        return `
          <div class="room-block${showToggle ? " is-collapsible" : ""}">
            <div class="room-title">
              ${showToggle ? '<button class="room-toggle" type="button" data-action="toggle-room">展開</button>' : ""}
              ${roomLabel} (${roomSessions.length})
            </div>
            <div class="room-items room-items-preview">${previewItems || "—"}</div>
            ${showToggle ? `<div class="room-items room-items-full">${fullItems || "—"}</div>` : ""}
          </div>
        `;
      })
      .join("");
    const dayCard = document.createElement("div");
    dayCard.className = "calendar-day";
    if (holidayLabel) {
      dayCard.classList.add("holiday-day");
    }
    if (dateStr === todayStr) {
      dayCard.classList.add("today");
    }
    dayCard.innerHTML = `
      <div class="day-number">${current.getMonth() + 1}/${current.getDate()} (星期${weekdays[(current.getDay() + 6) % 7]})</div>
      <div class="day-bar ${barClass}"></div>
      ${holidayLabel ? `<div class="holiday">${holidayLabel}</div>` : ""}
      <div class="day-items">${roomBlocks || "—"}</div>
    `;
    calendarView.appendChild(dayCard);
  }
}

function getCalendarRange(mode, anchorDate) {
  const anchor = new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate());
  if (mode === "week") {
    const start = startOfWeek(anchor);
    const end = addDays(start, 6);
    return {
      rangeStart: toIsoDate(start),
      rangeEnd: toIsoDate(end),
      label: `${start.getMonth() + 1}/${start.getDate()} - ${end.getMonth() + 1}/${end.getDate()}`,
    };
  }
  const start = startOfMonth(anchor);
  const gridStart = startOfWeek(start);
  const gridEnd = addDays(gridStart, 41);
  return {
    rangeStart: toIsoDate(gridStart),
    rangeEnd: toIsoDate(gridEnd),
    label: `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, "0")}`,
  };
}

function startOfWeek(dateValue) {
  const jsDay = dateValue.getDay();
  const diff = (jsDay + 6) % 7;
  return addDays(dateValue, -diff);
}

function startOfMonth(dateValue) {
  return new Date(dateValue.getFullYear(), dateValue.getMonth(), 1);
}

function toIsoDate(dateValue) {
  const year = dateValue.getFullYear();
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const day = String(dateValue.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getHolidayLabel(dateStr, holidays) {
  const dateValue = parseDate(dateStr);
  if (!dateValue) return "";
  for (const holiday of holidays) {
    const start = parseDate(holiday.start_date);
    const end = parseDate(holiday.end_date);
    if (!start || !end) continue;
    if (dateValue >= start && dateValue <= end) {
      return holiday.name || "假期";
    }
  }
  return "";
}

function groupSessionsByTime(sessions) {
  const buckets = new Map();
  sessions.forEach((session) => {
    const key = session.time || "";
    if (!buckets.has(key)) {
      buckets.set(key, []);
    }
    buckets.get(key).push(session);
  });
  const entries = Array.from(buckets.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  return entries.map(([time, groupSessions]) => ({ time, sessions: groupSessions }));
}

function isEndingSoon(session) {
  const total = Number(session.lesson_total || 0);
  const index = Number(session.lesson_index || 0);
  if (!total || !index) return false;
  const remaining = total - index;
  return remaining >= 0 && remaining <= 2;
}

function isNewClass(session) {
  return Number(session.lesson_index || 0) === 1;
}

function renderTimeBlocks(grouped) {
  return grouped
    .map((group) => {
      const details = group.sessions
        .map((session) => {
          const displaySku = mergeLocationPrefix(
            session.sku,
            session.location || appState.app_config?.location || ""
          );
          const line = `(${session.lesson_index}/${session.lesson_total}) ${displaySku}`;
          let label = line;
          if (isEndingSoon(session)) {
            label = `<span class="ending-soon">${line}</span>`;
          } else if (isNewClass(session)) {
            label = `<span class="new-class">${line}</span>`;
          }
          return `${label}<br>${session.teacher || ""}`;
        })
        .join("<br>");
      return `
        <div class="time-block">
          <div class="time-label">${group.time || "未設定時間"} (${group.sessions.length})</div>
          <div>${details}</div>
        </div>
      `;
    })
    .join("");
}

function mergeLocationPrefix(sku, locationCode) {
  const value = (sku || "").trim();
  const location = (locationCode || "").trim().toUpperCase();
  if (!value || !location) return value;
  const match = /^([^0-9]+?)?([KLH])?(\d{1,2}[A-Z]\d{2})$/.exec(value);
  if (!match) return value;
  if (match[2]) return value;
  return `${match[1] || ""}${location}${match[3]}`;
}

function formatDocxSku(value) {
  const text = (value || "").trim();
  if (!text) return "-";
  const match = /^(.*?)([KLH])?(\d{1,2})([A-Z])(\d{2})$/.exec(text);
  if (!match) return text;
  const level = match[1];
  const location = match[2] || "";
  const month = match[3];
  const letter = match[4];
  const year = match[5];
  const prefix = level.startsWith("N") ? "" : "N";
  return `${prefix}${level}${location}${month}${letter}/${year}`;
}

function formatDocxTime(value) {
  const text = (value || "").trim();
  if (!text) return "-";
  const match = /(\d{2}):?(\d{2})\s*[-~]\s*(\d{2}):?(\d{2})/.exec(text);
  if (!match) return text;
  const toLabel = (hourStr, minStr) => {
    const hour = Number(hourStr);
    const minute = Number(minStr);
    const isPm = hour >= 12;
    const hour12 = ((hour + 11) % 12) + 1;
    return `${hour12}:${String(minute).padStart(2, "0")}${isPm ? "pm" : "am"}`;
  };
  const start = toLabel(match[1], match[2]);
  const end = toLabel(match[3], match[4]);
  return `${start} ~ ${end}`;
}

function formatDisplayTime(value) {
  const text = (value || "").trim();
  if (!text) return "-";
  const rangeMatch = /(\d{1,2}):?(\d{2})\s*[-~]\s*(\d{1,2}):?(\d{2})/.exec(text);
  if (rangeMatch) {
    const h1 = String(rangeMatch[1]).padStart(2, "0");
    const m1 = rangeMatch[2];
    const h2 = String(rangeMatch[3]).padStart(2, "0");
    const m2 = rangeMatch[4];
    return `${h1}:${m1} - ${h2}:${m2}`;
  }
  const singleMatch = /^(\d{1,2}):?(\d{2})$/.exec(text);
  if (singleMatch) {
    return `${String(singleMatch[1]).padStart(2, "0")}:${singleMatch[2]}`;
  }
  return text;
}

function parseDate(value) {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function toBool(value) {
  if (typeof value === "boolean") return value;
  const text = String(value ?? "").trim().toLowerCase();
  return text === "1" || text === "true" || text === "yes";
}

function getDerivedStatus(cls) {
  if (cls.start_date) {
    const startDate = parseDate(cls.start_date);
    if (startDate) {
      const today = new Date();
      const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      if (startDate > todayDate) {
        return "not_started";
      }
    }
  }
  if (cls.lessons_remaining === 0) {
    return "ended";
  }
  return cls.status || "active";
}

function addDays(dateValue, days) {
  const next = new Date(dateValue.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function buildFeeClassSku() {
  const level = feeLevelSelect?.value || "";
  const location = feeLocationSelect?.value || "";
  const month = feeMonthSelect?.value || "";
  const letter = feeLetterSelect?.value || "";
  const year = feeYearSelect?.value || "";
  if (!level || !month || !letter || !year) return "";
  return `${level}${location}${Number(month)}${letter}${year}`;
}

function initFeeGuideSelectors() {
  if (feeMonthSelect && feeMonthSelect.options.length <= 1) {
    for (let i = 1; i <= 12; i += 1) {
      const option = document.createElement("option");
      option.value = String(i);
      option.textContent = `${i}月`;
      feeMonthSelect.appendChild(option);
    }
  }
  if (feeLetterSelect && feeLetterSelect.options.length <= 1) {
    for (let code = 65; code <= 90; code += 1) {
      const letter = String.fromCharCode(code);
      const option = document.createElement("option");
      option.value = letter;
      option.textContent = letter;
      feeLetterSelect.appendChild(option);
    }
  }
  if (feeYearSelect && feeYearSelect.options.length <= 1) {
    const currentYear = new Date().getFullYear();
    for (let year = currentYear - 2; year <= currentYear + 2; year += 1) {
      const option = document.createElement("option");
      option.value = String(year).slice(-2);
      option.textContent = `${year}`;
      feeYearSelect.appendChild(option);
    }
  }
  if (feeLocationSelect && !feeLocationSelect.value) {
    const defaultLocation = appState.app_config?.location || "";
    if (defaultLocation) {
      feeLocationSelect.value = defaultLocation;
    }
  }
  if (feeYearSelect && !feeYearSelect.value) {
    feeYearSelect.value = String(new Date().getFullYear()).slice(-2);
  }
}

function resetFeeGuideForm() {
  if (feeLevelSelect) feeLevelSelect.value = "";
  if (feeLocationSelect) feeLocationSelect.value = "";
  if (feeMonthSelect) feeMonthSelect.value = "";
  if (feeLetterSelect) feeLetterSelect.value = "";
  if (feeYearSelect) feeYearSelect.value = "";
  if (feeClassTimeInput) feeClassTimeInput.value = "";
  if (feeStartDateInput) feeStartDateInput.value = "";
  if (feeDeadlineInput) feeDeadlineInput.value = "";
  if (feeTextbookInput) feeTextbookInput.value = "0";
  if (feeIdCardInput) feeIdCardInput.checked = false;
  feeClassAdjust = 0;
  feeTextbookAutoFilled = false;
  const feeTextbookLabel = document.getElementById("feeTextbookLabel");
  if (feeTextbookLabel) feeTextbookLabel.textContent = "書簿費";
  initFeeGuideSelectors();
  updateFeeGuideOutput();
}

function findClassBySkuInput(value) {
  const normalized = (value || "").trim().toUpperCase();
  if (!normalized) return null;
  const exact = appState.classes.find(
    (cls) => (cls.sku || "").trim().toUpperCase() === normalized
  );
  if (exact) return exact;
  return appState.classes.find((cls) => (cls.sku || "").trim().toUpperCase().endsWith(normalized)) || null;
}

function getLevelFromSkuInput(value) {
  const text = (value || "").trim();
  if (!text) return "";
  const levels = Array.isArray(appState.settings.level) ? appState.settings.level : [];
  const ordered = [...levels].sort((a, b) => b.length - a.length);
  return ordered.find((level) => text.startsWith(level)) || "";
}

function normalizeTimeRange(value) {
  if (!value) return "";
  return value.replace(/\s*[-–]\s*/g, " - ");
}

function formatChineseDate(dateStr) {
  const dateValue = parseDate(dateStr);
  if (!dateValue) return "";
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const day = String(dateValue.getDate()).padStart(2, "0");
  return `${month}月${day}日`;
}

function formatChineseDateParts(dateValue) {
  if (!dateValue) return "";
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const day = String(dateValue.getDate()).padStart(2, "0");
  return `${month}月${day}日`;
}

function createMakeupRow(index) {
  if (!makeupRows) return null;
  const row = document.createElement("div");
  row.className = "makeup-row";
  row.dataset.index = String(index);
  row.innerHTML = `
    <div class="field-inline">
      <label>
        學校
        <select data-field="school">
          <option value="">請選擇</option>
          <option value="太子校">太子校</option>
          <option value="旺角校">旺角校</option>
          <option value="香港校">香港校</option>
        </select>
      </label>
      <label>
        日期
        <input type="date" data-field="date" />
      </label>
      <label>
        星期
        <select data-field="weekday">
          <option value="">請選擇</option>
          <option value="1">星期一</option>
          <option value="2">星期二</option>
          <option value="3">星期三</option>
          <option value="4">星期四</option>
          <option value="5">星期五</option>
          <option value="6">星期六</option>
          <option value="0">星期日</option>
        </select>
      </label>
      <label>
        時間
        <input type="text" data-field="time" list="makeupTimeOptions" placeholder="19:00-21:30" />
      </label>
      <label>
        班次
        <input type="number" data-field="count" min="1" placeholder="例如 3" />
      </label>
    </div>
    <div class="field-inline">
      <div class="pill">第 ${index + 1} 行</div>
      ${index === 0 ? "" : '<button class="btn danger" type="button" data-action="remove-row">刪除</button>'}
    </div>
  `;
  return row;
}

function getMakeupRowData(row) {
  const school = row.querySelector('[data-field="school"]')?.value || "";
  const dateStr = row.querySelector('[data-field="date"]')?.value || "";
  const dateValue = parseDate(dateStr);
  const weekdaySelect = row.querySelector('[data-field="weekday"]');
  const weekdayText = weekdaySelect?.options[weekdaySelect.selectedIndex]?.text || "";
  const timeText = row.querySelector('[data-field="time"]')?.value?.trim() || "";
  const countValue = row.querySelector('[data-field="count"]')?.value || "";
  const classCount = countValue ? `第${countValue}回` : "";
  return {
    school,
    dateText: formatChineseDateParts(dateValue),
    weekdayText,
    timeText,
    classCount,
  };
}

function updateMakeupRowCount() {
  const counter = document.getElementById("makeupRowCount");
  if (counter) {
    const count = makeupRows?.querySelectorAll(".makeup-row").length || 0;
    counter.textContent = `${count} 行`;
  }
}

function reindexMakeupRows() {
  const rows = makeupRows?.querySelectorAll(".makeup-row") || [];
  rows.forEach((row, i) => {
    row.dataset.index = String(i);
    const pill = row.querySelector(".pill");
    if (pill) pill.textContent = `第 ${i + 1} 行`;
  });
  updateMakeupRowCount();
}

function initMakeupRows() {
  if (!makeupRows) return;
  makeupRows.innerHTML = "";
  const row = createMakeupRow(0);
  if (row) {
    makeupRows.appendChild(row);
  }
  updateMakeupRowCount();
}

function formatClassTimeDisplay(startDateStr, timeRange) {
  const dateValue = parseDate(startDateStr);
  if (!dateValue) return "";
  const jsDay = dateValue.getDay();
  const weekdayIndex = (jsDay + 6) % 7;
  const weekdayLabel = weekdays[weekdayIndex] ? `星期${weekdays[weekdayIndex]}` : "";
  const dateLabel = formatChineseDate(startDateStr);
  const timeLabel = normalizeTimeRange(timeRange);
  const timeBlock = timeLabel ? ` ${timeLabel}` : "";
  return `${dateLabel} (${weekdayLabel}${timeBlock})`;
}

function updateFeeDeadlineFromStart() {
  if (!feeDeadlineInput) return;
  const today = new Date();
  const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const deadline = addDays(todayDate, 4);
  feeDeadlineInput.value = toIsoDate(deadline);
}

function updateFeeGuideOutput() {
  if (!feeTemplateOutput) return;
  const className = buildFeeClassSku();
  const classTime = (feeClassTimeInput?.value || "").trim();
  const startDate = feeStartDateInput?.value || "";
  const deadlineDate = feeDeadlineInput?.value || "";
  const textbookPrice = Number(feeTextbookInput?.value || 0);
  const idCardFee = feeIdCardInput?.checked ? 50 : 0;
  const matchedClass = findClassBySkuInput(className);
  const level = matchedClass?.level || feeLevelSelect?.value || getLevelFromSkuInput(className);
  const levelPrices = appState.settings?.level_price || {};
  const baseClassPrice = Number(levelPrices[level] || 0);
  const classPrice = baseClassPrice + feeClassAdjust;
  const totalPrice = classPrice + textbookPrice + idCardFee;

  if (feeClassPriceLabel) {
    feeClassPriceLabel.textContent = `$${classPrice}`;
  }
  if (feeTotalLabel) {
    feeTotalLabel.textContent = `$${totalPrice}`;
  }

  const timeDisplay = startDate ? formatClassTimeDisplay(startDate, classTime) : "";
  const deadlineDisplay = formatChineseDate(deadlineDate);
  let template = feeTemplateBase;
  if (!template) {
    template = "此訊息由第一日語暨文化學校送出\n\n已給同學留位，報讀課程資料如下:\n(非持續進修基金課程)\n班別：\n時間：\n學費：\n課本：\n學生證：\n*費用共：*\n\n**課程如有變動，會另行以WhatsApp 通知。\n\n請同學於 * * 或之前透過以下方法繳交費用。";
  }
  const lines = template.split(/\r?\n/).map((line) => {
    const trimmed = line.trim();
    if (trimmed.startsWith("班別")) {
      return `班別：\t${className || "-"}`;
    }
    if (trimmed.startsWith("時間")) {
      return `時間：\t${timeDisplay || "-"}`;
    }
    if (/^學費\s*[：:]/.test(trimmed)) {
      return `學費：\t$${classPrice}`;
    }
    if (trimmed.startsWith("課本") || trimmed.startsWith("書簿")) {
      return `課本：\t$${textbookPrice}`;
    }
    if (trimmed.startsWith("學生證") || trimmed.startsWith("ID證") || trimmed.startsWith("ID卡")) {
      return `學生證：\t$${idCardFee}`;
    }
    if (trimmed.startsWith("*費用共") || trimmed.startsWith("費用共")) {
      return `*費用共：$${totalPrice}*`;
    }
    if (trimmed.startsWith("請同學於")) {
      const deadlineText = deadlineDisplay ? `*${deadlineDisplay}*` : "* *";
      return `請同學於 ${deadlineText} 或之前透過以下方法繳交費用。`;
    }
    return line;
  });
  feeTemplateOutput.value = lines.join("\n");
}

function renderFeeGuide() {
  updateFeeGuideOutput();
}

function autoFillTextbookPrice(level) {
  if (!feeTextbookInput || !level) return;
  if (!feeTextbookAutoFilled && feeTextbookInput.value && Number(feeTextbookInput.value) !== 0) return;
  const levelTextbook = appState.settings?.level_textbook || {};
  const textbooks = appState.settings?.textbook || {};
  const tbNames = Array.isArray(levelTextbook[level]) ? levelTextbook[level] : [];
  const label = document.getElementById("feeTextbookLabel");
  if (tbNames.length) {
    const total = tbNames.reduce((sum, n) => sum + (textbooks[n] || 0), 0);
    feeTextbookInput.value = total;
    feeTextbookAutoFilled = true;
    const nameStr = tbNames.join(" + ");
    if (label) label.textContent = `書簿費（${nameStr}，自動）`;
  } else {
    feeTextbookAutoFilled = false;
    if (label) label.textContent = "書簿費";
  }
}

function syncFeeGuideFromClass() {
  const className = buildFeeClassSku();
  const matched = findClassBySkuInput(className);
  const level = feeLevelSelect?.value || "";
  if (matched) {
    if (feeClassTimeInput) {
      feeClassTimeInput.value = matched.start_time || "";
    }
    if (feeStartDateInput) {
      feeStartDateInput.value = matched.start_date || "";
    }
    updateFeeDeadlineFromStart();
  }
  autoFillTextbookPrice(level);
  updateFeeGuideOutput();
}

async function loadFeeTemplate() {
  const response = await invoke('load_payment_template');
  if (response.ok) {
    feeTemplateBase = response.content || "";
  }
  updateFeeGuideOutput();
}

async function loadDocxTemplates() {
  const response = await invoke('list_docx_templates');
  if (!response.ok) {
    if (docxOutput) {
      docxOutput.textContent = response.error || "無法載入模板。";
    }
    return;
  }
  docxTemplates = Array.isArray(response.templates) ? response.templates : [];
  renderDocxTemplates();
}

async function loadMessageTemplates() {
  const response = await invoke('list_message_templates');
  if (!response.ok) {
    if (messageOutput) {
      messageOutput.value = response.error || "無法載入訊息。";
    }
    return;
  }
  messageTemplates = Array.isArray(response.templates) ? response.templates : [];
  renderMessageFilters();
  renderMessageList();
}

async function selectMessageTemplate(name) {
  const response = await invoke('load_message_content', { templateName: name });
  if (!response.ok) {
    if (messageOutput) {
      messageOutput.value = response.error || "無法載入訊息內容。";
    }
    return;
  }
  activeMessageName = name;
  if (messageTitle) {
    const target = messageTemplates.find((item) => item.name === name);
    messageTitle.textContent = target ? target.label : "訊息內容";
  }
  if (messageOutput) {
    messageOutput.value = response.content || "";
  }
}

function getReminderBadge(targetDate, label, isDone) {
  if (!targetDate) return "-";
  if (isDone) {
    return `<span class="reminder gray">已完成</span>`;
  }
  const today = new Date();
  const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const targetDateOnly = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  const diffMs = targetDateOnly - todayDate;
  const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
  const diffWeeks = Math.ceil(diffDays / 7);

  if (diffWeeks > reminderConfig.yellowWeeks) {
    return "-";
  }

  if (diffDays < 0) {
    const overdueWeeks = Math.ceil(Math.abs(diffDays) / 7);
    return `<span class="reminder red">已過期 ${overdueWeeks} 週</span>`;
  }

  if (diffWeeks <= reminderConfig.redWeeks) {
    return `<span class="reminder red">${label} ${diffWeeks} 週</span>`;
  }

  return `<span class="reminder yellow">${label} ${diffWeeks} 週</span>`;
}

function renderReminderCell(classId, key, badgeHtml, isDone) {
  const checked = isDone ? "checked" : "";
  return `
    <label class="reminder-check">
      <input type="checkbox" data-reminder="${key}" data-id="${classId}" ${checked} />
      ${badgeHtml}
    </label>
  `;
}

async function addClass(event) {
  event.preventDefault();
  const form = event.target;
  const submitBtn = form.querySelector('button[type="submit"]');
  if (!validateClassForm(form)) return;
  const data = Object.fromEntries(new FormData(form));
  if (data.level !== "初級") {
    data.relay_teacher = "";
    data.relay_date = "";
  }
  setButtonLoading(submitBtn, true);
  const response = await invoke('create_class', { data: {
    ...data,
    weekday: Number(data.weekday),
    student_count: Number(data.student_count || 0),
    lesson_total: Number(data.lesson_total || 0),
  } });
  setButtonLoading(submitBtn, false);
  if (!response.ok) {
    showToast(response.error || "新增班別失敗。");
    return;
  }
  showToast("班別已新增", "success");
  form.reset();
  await loadState();
}

async function addHoliday(event) {
  event.preventDefault();
  const form = event.target;
  const submitBtn = form.querySelector('button[type="submit"]');
  const data = Object.fromEntries(new FormData(form));
  setButtonLoading(submitBtn, true);
  await invoke('add_holiday', { data });
  setButtonLoading(submitBtn, false);
  form.reset();
  if (holidayEndDate) {
    holidayEndDate.disabled = false;
  }
  await loadState();
}

async function deleteHoliday(id) {
  await invoke('delete_holiday', { holidayId: id });
  await loadState();
}

async function postponeClass(id) {
  await openScheduleModal(id);
  const postponeSection = postponeList?.closest(".modal-section");
  if (postponeSection) {
    postponeSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

async function endClass(id) {
  const action = prompt("輸入操作：升級、合併、結束");
  if (!action) return;

  if (action === "結束") {
    await invoke('end_class_action', { classId: id, action: "terminate", targetId: "", newSku: "" });
  } else if (action === "合併") {
    const targetSku = prompt("輸入要合併的班別:");
    const target = appState.classes.find((cls) => cls.sku === targetSku);
    if (!target) {
      showToast("找不到目標班別。");
      return;
    }
    await invoke('end_class_action', { classId: id, action: "merge", targetId: target.id, newSku: "" });
  } else if (action === "升級") {
    const newSku = prompt("輸入升級後的新班別:");
    if (!newSku) return;
    await invoke('end_class_action', { classId: id, action: "promote", targetId: "", newSku: newSku });
  } else {
    showToast("未能識別操作。");
  }

  await loadState();
}

async function openScheduleModal(classId) {
  const response = await invoke('get_class_schedule', { classId: classId });
  if (!response.ok) {
    showToast(response.error || "無法載入日程。");
    return;
  }
  activeClassId = classId;
  const cls = response.class;
  detailSku.value = mergeLocationPrefix(
    cls.sku,
    cls.location || appState.app_config?.location || ""
  );
  detailRoom.value = cls.classroom || "";
  populateSelect(detailTeacher, appState.settings.teacher, cls.teacher || "");
  populateSelect(detailTime, appState.settings.time, cls.start_time || "");
  populateSelect(detailRelayTeacher, appState.settings.teacher, cls.relay_teacher || "");
  detailRelayDate.value = cls.relay_date || "";
  detailRelayTeacher.disabled = cls.level !== "初級";
  detailRelayDate.disabled = cls.level !== "初級";
  detailStudents.value = cls.student_count || 0;

  scheduleBody.innerHTML = "";
  const today = new Date();
  const todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  response.schedule.forEach((item) => {
    const row = document.createElement("tr");
    row.className = "schedule-row";
    const lessonDate = parseDate(item.date);
    if (lessonDate && lessonDate < todayDate) {
      row.classList.add("dim");
    }
    const typeLabel = item.type === "makeup" ? "補課" : "正常";
    const progressLabel = item.index ? `<span class="pill">${item.index}/${item.total || ""}</span>` : "";
    row.innerHTML = `
      <td>${item.date} ${progressLabel}</td>
      <td>${typeLabel}</td>
    `;
    scheduleBody.appendChild(row);
  });

  postponeList.innerHTML = "";
  response.postpones.forEach((postpone) => {
    const item = document.createElement("div");
    item.className = "list-item";
    item.innerHTML = `
      <div>${postpone.original_date} → ${postpone.make_up_date}</div>
      <div class="item-actions">
        <button class="btn danger" data-postpone-id="${postpone.id}">刪除</button>
      </div>
    `;
    postponeList.appendChild(item);
  });

  overrideList.innerHTML = "";
  response.overrides.forEach((override) => {
    const item = document.createElement("div");
    item.className = "list-item";
    const actionLabel = override.action === "add" ? "新增" : "取消";
    item.innerHTML = `
      <div>${override.date} (${actionLabel})</div>
      <div class="item-actions">
        <button class="btn danger" data-override-id="${override.id}">刪除</button>
      </div>
    `;
    overrideList.appendChild(item);
  });

  scheduleModal.classList.remove("hidden");
  trapFocus(scheduleModal);
}

function closeScheduleModal() {
  scheduleModal.classList.add("hidden");
  activeClassId = "";
  addLessonDate.value = "";
  removeLessonDate.value = "";
  postponeOriginalDate.value = "";
  postponeMakeupDate.value = "";
  postponeReason.value = "";
  if (postponeAutoWeek) {
    postponeAutoWeek.checked = false;
  }
  if (postponeMakeupDate) {
    postponeMakeupDate.disabled = false;
  }
}

// ---- Modal: Backdrop click + Escape key + Focus trap ----
function trapFocus(modal) {
  const focusable = modal.querySelectorAll(
    "button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex='-1'])"
  );
  if (!focusable.length) return;
  const first = focusable[0];
  const last = focusable[focusable.length - 1];
  first.focus();
  modal.addEventListener("keydown", function handleTab(e) {
    if (e.key !== "Tab") return;
    if (e.shiftKey) {
      if (document.activeElement === first) {
        e.preventDefault();
        last.focus();
      }
    } else {
      if (document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }
    if (!modal.contains(document.activeElement)) {
      e.preventDefault();
      first.focus();
    }
  });
}

scheduleModal.addEventListener("click", (e) => {
  if (e.target === scheduleModal) closeScheduleModal();
});

messageCategoryModal.addEventListener("click", (e) => {
  if (e.target === messageCategoryModal) {
    messageCategoryModal.classList.add("hidden");
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  const fab = document.querySelector(".fab");
  if (fab && fab.classList.contains("is-open")) {
    fab.classList.remove("is-open");
    closeTimeCalc();
    return;
  }
  if (!timeCalcPopover?.classList.contains("hidden")) {
    closeTimeCalc();
  } else if (!scheduleModal.classList.contains("hidden")) {
    closeScheduleModal();
  } else if (!messageCategoryModal.classList.contains("hidden")) {
    messageCategoryModal.classList.add("hidden");
  }
});

document.getElementById("classForm").addEventListener("submit", addClass);
document.getElementById("holidayForm").addEventListener("submit", addHoliday);
document.getElementById("refreshBtn").addEventListener("click", loadState);
if (sidebarToggle && layout) {
  sidebarToggle.addEventListener("click", () => {
    const isCollapsed = layout.classList.toggle("sidebar-collapsed");
    sidebarToggle.textContent = isCollapsed ? "▸" : "◂";
  });
}
if (fabToggle && fabActions) {
  fabToggle.addEventListener("click", () => {
    const fab = fabToggle.closest(".fab");
    if (fab) {
      fab.classList.toggle("is-open");
    }
  });
}
document.addEventListener("click", (e) => {
  const fab = document.querySelector(".fab");
  if (!fab || !fab.classList.contains("is-open")) return;
  if (!fab.contains(e.target)) {
    fab.classList.remove("is-open");
    closeTimeCalc();
  }
});
if (jumpTopBtn) {
  jumpTopBtn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}
if (timeCalcBtn) {
  timeCalcBtn.addEventListener("click", () => {
    if (!timeCalcPopover) return;
    const isVisible = !timeCalcPopover.classList.contains("hidden");
    if (isVisible) {
      closeTimeCalc();
    } else {
      openTimeCalc();
    }
  });
}
if (timeCalcClose) {
  timeCalcClose.addEventListener("click", closeTimeCalc);
}
if (timeCalcHours) {
  timeCalcHours.addEventListener("input", updateTimeCalcResult);
}
if (timeCalcMinutes) {
  timeCalcMinutes.addEventListener("input", updateTimeCalcResult);
}
if (reminderSort) {
  reminderSort.addEventListener("change", renderReminders);
}
if (holidayOneDay && holidayStartDate && holidayEndDate) {
  const syncHolidayEnd = () => {
    if (!holidayOneDay.checked) {
      holidayEndDate.disabled = false;
      return;
    }
    if (holidayStartDate.value) {
      holidayEndDate.value = holidayStartDate.value;
    }
    holidayEndDate.disabled = true;
  };
  holidayOneDay.addEventListener("change", syncHolidayEnd);
  holidayStartDate.addEventListener("change", syncHolidayEnd);
}
if (locationSelect) {
  locationSelect.addEventListener("change", async () => {
    updateLocationUI(locationSelect.value);
    const response = await invoke('set_app_location', { location: locationSelect.value });
    if (!response.ok) {
      showToast(response.error || "設定地點失敗。");
      return;
    }
    await loadState();
  });
}

if (locationSegments) {
  locationSegments.addEventListener("click", (event) => {
    const btn = event.target.closest(".loc-seg");
    if (!btn) return;
    locationSelect.value = btn.dataset.value;
    locationSelect.dispatchEvent(new Event("change"));
  });
}
searchInput.addEventListener("input", renderClasses);
statusFilter.addEventListener("change", renderClasses);
weekdayFilter.addEventListener("change", renderClasses);
locationFilter.addEventListener("change", renderClasses);
levelFilter.addEventListener("change", renderClasses);
if (archiveFilter) {
  archiveFilter.addEventListener("change", renderClasses);
}
if (issueFilter) {
  issueFilter.addEventListener("change", renderClasses);
}

levelSelect.addEventListener("change", () => {
  const isBeginner = levelSelect.value === "初級";
  relayFields.classList.toggle("hidden", !isBeginner);
  if (!isBeginner) {
    relayTeacherSelect.value = "";
    relayFields.querySelector('input[name="relay_date"]').value = "";
  }
});

startDateInput.addEventListener("change", () => {
  const selectedDate = parseDate(startDateInput.value);
  if (!selectedDate) return;
  const jsDay = selectedDate.getDay();
  const weekday = (jsDay + 6) % 7;
  weekdaySelect.value = String(weekday);
  setDefaultTime();
});

document.querySelectorAll(".settings-form").forEach((form) => {
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const entryType = event.target.dataset.settingType;
    const data = Object.fromEntries(new FormData(event.target));
    const response = await invoke('add_setting', { entryType: entryType, value: data.value || "" });
    if (!response.ok) {
      showToast(response.error || "新增設定失敗。");
      return;
    }
    event.target.reset();
    await loadState();
  });
});

if (priceForm) {
  priceForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const level = priceLevelSelect?.value || "";
    const price = Number(priceValueInput?.value || 0);
    const response = await invoke('set_level_price', { level, price });
    if (!response.ok) {
      showToast(response.error || "設定學費失敗。");
      return;
    }
    showToast("學費已更新", "success");
    priceValueInput.value = "";
    await loadState();
  });
}

if (priceAdjustPlus) {
  priceAdjustPlus.addEventListener("click", async () => {
    const response = await invoke('adjust_level_prices', { delta: 20 });
    if (!response.ok) {
      showToast(response.error || "調整學費失敗。");
      return;
    }
    showToast("全部學費 +$20", "success");
    await loadState();
  });
}

if (priceAdjustMinus) {
  priceAdjustMinus.addEventListener("click", async () => {
    const response = await invoke('adjust_level_prices', { delta: -20 });
    if (!response.ok) {
      showToast(response.error || "調整學費失敗。");
      return;
    }
    showToast("全部學費 -$20", "success");
    await loadState();
  });
}

// ─── Textbook form handlers ───────────────────────────────────────────────────

if (textbookForm) {
  textbookForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const name = textbookNameInput?.value.trim() || "";
    const price = Number(textbookPriceInput?.value || 0);
    if (!name) { showToast("請輸入教材名稱。"); return; }
    const response = await invoke('set_textbook', { name, price });
    if (!response.ok) { showToast(response.error || "新增失敗。"); return; }
    showToast("教材已儲存", "success");
    textbookNameInput.value = "";
    textbookPriceInput.value = "";
    await loadState();
  });
}

if (textbookList) {
  textbookList.addEventListener("click", async (event) => {
    const editBtn = event.target.closest("[data-textbook-edit]");
    if (editBtn) {
      textbookNameInput.value = editBtn.dataset.textbookEdit;
      textbookPriceInput.value = editBtn.dataset.price;
      textbookNameInput.focus();
      return;
    }
    const delBtn = event.target.closest("[data-textbook-delete]");
    if (delBtn) {
      const name = delBtn.dataset.textbookDelete;
      if (!confirm(`確定刪除教材「${name}」？相關等級對應及存貨記錄亦會一併移除。`)) return;
      const response = await invoke('delete_textbook', { name });
      if (!response.ok) { showToast(response.error || "刪除失敗。"); return; }
      showToast("教材已刪除", "success");
      await loadState();
    }
  });
}

if (levelTextbookList) {
  levelTextbookList.addEventListener("change", async (event) => {
    const ltbCheckbox = event.target.closest("[data-ltb-level][data-ltb-book]");
    if (ltbCheckbox) {
      const level = ltbCheckbox.dataset.ltbLevel;
      // Collect all checked books for this level
      const allChecks = levelTextbookList.querySelectorAll(`[data-ltb-level="${level}"][data-ltb-book]`);
      const selectedBooks = Array.from(allChecks).filter((c) => c.checked).map((c) => c.dataset.ltbBook);
      const response = await invoke('set_level_textbook', { level, textbookNames: selectedBooks });
      if (!response.ok) { showToast(response.error || "設定失敗。"); return; }
      appState.settings.level_textbook = appState.settings.level_textbook || {};
      appState.settings.level_textbook[level] = selectedBooks;
      showToast(`${level} 教材已更新`, "success");
      return;
    }
    const lnextSelect = event.target.closest("[data-lnext-level]");
    if (lnextSelect) {
      const level = lnextSelect.dataset.lnextLevel;
      const nextLevel = lnextSelect.value;
      const response = await invoke('set_level_next', { level, nextLevel: nextLevel });
      if (!response.ok) { showToast(response.error || "設定失敗。"); return; }
      appState.settings.level_next = appState.settings.level_next || {};
      appState.settings.level_next[level] = nextLevel;
      showToast(`${level} 下一等級已更新`, "success");
    }
  });
}

// ─── Stock tab handlers ───────────────────────────────────────────────────────

if (inventoryTable) {
  inventoryTable.addEventListener("click", async (event) => {
    const saveBtn = event.target.closest("[data-stock-save]");
    if (saveBtn) {
      const name = saveBtn.dataset.stockSave;
      const input = inventoryTable.querySelector(`[data-stock-name="${name}"]`);
      const count = parseInt(input?.value || "0", 10);
      const response = await invoke('set_textbook_stock', { name, count });
      if (!response.ok) { showToast(response.error || "儲存失敗。"); return; }
      appState.settings.textbook_stock = appState.settings.textbook_stock || {};
      appState.settings.textbook_stock[name] = count;
      showToast(`${name} 存貨已更新`, "success");
      renderPromotionPlanningTable();
      return;
    }
    const minusBtn = event.target.closest("[data-stock-minus]");
    if (minusBtn) {
      const name = minusBtn.dataset.stockMinus;
      const input = inventoryTable.querySelector(`[data-stock-name="${name}"]`);
      if (input) input.value = Math.max(0, (parseInt(input.value || "0", 10) - 1));
      return;
    }
    const plusBtn = event.target.closest("[data-stock-plus]");
    if (plusBtn) {
      const name = plusBtn.dataset.stockPlus;
      const input = inventoryTable.querySelector(`[data-stock-name="${name}"]`);
      if (input) input.value = (parseInt(input.value || "0", 10) + 1);
    }
  });
}

if (promotionThreshold) {
  promotionThreshold.addEventListener("change", () => renderPromotionPlanningTable());
}

if (stockReviewBtn) {
  stockReviewBtn.addEventListener("click", () => {
    const isVisible = !stockQuickReview?.classList.contains("hidden");
    renderQuickReviewPanel(!isVisible);
    stockReviewBtn.textContent = isVisible ? "核實課堂人數" : "收起";
  });
}

if (reviewWeekdayFilter) {
  reviewWeekdayFilter.addEventListener("change", () => {
    if (!stockQuickReview?.classList.contains("hidden")) renderQuickReviewPanel(true);
  });
}

if (reviewLevelFilter) {
  reviewLevelFilter.addEventListener("change", () => {
    if (!stockQuickReview?.classList.contains("hidden")) renderQuickReviewPanel(true);
  });
}

if (stockConfirmAllBtn) {
  stockConfirmAllBtn.addEventListener("click", async () => {
    await saveReviewTimestamp();
    stockReviewBtn.textContent = "核實課堂人數";
  });
}

if (stockSaveReviewBtn) {
  stockSaveReviewBtn.addEventListener("click", async () => {
    const inputs = stockReviewBody?.querySelectorAll(".review-count-input") || [];
    const updates = [];
    inputs.forEach((input) => {
      updates.push({ id: input.dataset.classId, student_count: parseInt(input.value || "0", 10) });
    });
    if (updates.length) {
      const response = await invoke('save_student_counts', { updates });
      if (!response.ok) { showToast(response.error || "儲存失敗。"); return; }
      // Update local state
      updates.forEach((u) => {
        const cls = appState.classes.find((c) => c.id === u.id);
        if (cls) cls.student_count = u.student_count;
      });
      renderPromotionPlanningTable();
    }
    await saveReviewTimestamp();
    stockReviewBtn.textContent = "核實課堂人數";
  });
}

const tasksDropdown = document.getElementById("tasksDropdown");
const tasksBackdrop = document.getElementById("tasksBackdrop");

function openTasksDropdown() {
  tasksDropdown.classList.add("open");
  tasksBackdrop.classList.add("visible");
}

function closeTasksDropdown() {
  tasksDropdown.classList.remove("open");
  tasksBackdrop.classList.remove("visible");
  // Remove active state from tasks button
  document.querySelectorAll('.tab-button[data-tab="tasks"]').forEach((b) => b.classList.remove("active"));
}

if (tasksBackdrop) {
  tasksBackdrop.addEventListener("click", closeTasksDropdown);
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && tasksDropdown && tasksDropdown.classList.contains("open")) {
    closeTasksDropdown();
  }
});

document.querySelectorAll(".tab-button").forEach((button) => {
  button.addEventListener("click", () => {
    // Tasks button toggles a dropdown instead of switching tabs
    if (button.dataset.tab === "tasks") {
      if (tasksDropdown.classList.contains("open")) {
        closeTasksDropdown();
      } else {
        openTasksDropdown();
        button.classList.add("active");
      }
      return;
    }
    // Close tasks dropdown if open when switching to another tab
    closeTasksDropdown();
    document.querySelectorAll(".tab-button").forEach((btn) => btn.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach((content) => content.classList.remove("active"));
    button.classList.add("active");
    const target = button.dataset.tab;
    document.querySelector(`[data-tab-content="${target}"]`).classList.add("active");
    try { localStorage.setItem("dij_active_tab", target); } catch {}
  });
});

if (tabOrderToggle) {
  tabOrderToggle.addEventListener("click", () => {
    tabOrderUnlocked = !tabOrderUnlocked;
    updateTabOrderToggleLabel();
    setTabReorderState(tabOrderUnlocked);
    if (tabOrderUnlocked) {
      showToast("使用箭頭按鈕排序標籤", "info");
    }
  });
  updateTabOrderToggleLabel();
  setTabReorderState(false);
}


if (feeLevelSelect) {
  feeLevelSelect.addEventListener("change", syncFeeGuideFromClass);
}
if (feeLocationSelect) {
  feeLocationSelect.addEventListener("change", syncFeeGuideFromClass);
}
if (feeMonthSelect) {
  feeMonthSelect.addEventListener("change", syncFeeGuideFromClass);
}
if (feeLetterSelect) {
  feeLetterSelect.addEventListener("change", syncFeeGuideFromClass);
}
if (feeYearSelect) {
  feeYearSelect.addEventListener("change", syncFeeGuideFromClass);
}

if (feeClassTimeInput) {
  feeClassTimeInput.addEventListener("input", updateFeeGuideOutput);
}

if (feeStartDateInput) {
  feeStartDateInput.addEventListener("change", () => {
    updateFeeDeadlineFromStart();
    updateFeeGuideOutput();
  });
}

if (feeDeadlineInput) {
  feeDeadlineInput.addEventListener("change", updateFeeGuideOutput);
}

if (feeTextbookInput) {
  feeTextbookInput.addEventListener("input", () => {
    // User manually edited — clear auto-fill flag and restore label
    if (feeTextbookAutoFilled) {
      feeTextbookAutoFilled = false;
      const label = document.getElementById("feeTextbookLabel");
      if (label) label.textContent = "書簿費";
    }
    updateFeeGuideOutput();
  });
}

if (feeIdCardInput) {
  feeIdCardInput.addEventListener("change", updateFeeGuideOutput);
}

if (feeCopyBtn) {
  feeCopyBtn.addEventListener("click", async () => {
    if (!feeTemplateOutput) return;
    await copyToClipboard(feeTemplateOutput.value, feeTemplateOutput);
  });
}

if (feeResetBtn) {
  feeResetBtn.addEventListener("click", resetFeeGuideForm);
}

if (feeTutorPlus) {
  feeTutorPlus.addEventListener("click", () => {
    feeClassAdjust += 20;
    updateFeeGuideOutput();
  });
}

if (feeTutorMinus) {
  feeTutorMinus.addEventListener("click", () => {
    feeClassAdjust -= 20;
    updateFeeGuideOutput();
  });
}

if (docxClassSelect) {
  docxClassSelect.addEventListener("change", () => {
    renderDocxSecondaryClasses();
    updateDocxPreview();
    if (docxOutput) {
      docxOutput.textContent = "";
    }
  });
}

if (docxTemplateSelect) {
  docxTemplateSelect.addEventListener("change", () => {
    if (docxSecondClassRow) {
      const needsSecond = docxTemplateSelect.value === "class.docx";
      if (docxSecondClassToggleRow) {
        docxSecondClassToggleRow.classList.toggle("hidden", !needsSecond);
      }
      const showSecondary = needsSecond && !!docxSecondClassToggle?.checked;
      docxSecondClassRow.classList.toggle("hidden", !showSecondary);
      if (!needsSecond && docxClassSelectSecondary) {
        docxClassSelectSecondary.value = "";
      }
      if (!needsSecond && docxSecondClassToggle) {
        docxSecondClassToggle.checked = false;
      }
      if (!needsSecond && docxRelayTeacherSecondary) {
        docxRelayTeacherSecondary.checked = false;
      }
    }
    if (docxPreviewSecondary) {
      const showPreview = docxTemplateSelect.value === "class.docx" && !!docxSecondClassToggle?.checked;
      docxPreviewSecondary.classList.toggle("hidden", !showPreview);
    }
    updateDocxPreview();
    if (docxOutput) {
      docxOutput.textContent = "";
    }
  });
}

if (docxClassSelectSecondary) {
  docxClassSelectSecondary.addEventListener("change", () => {
    updateDocxPreview();
    if (docxOutput) {
      docxOutput.textContent = "";
    }
  });
}

if (docxSecondClassToggle) {
  docxSecondClassToggle.addEventListener("change", () => {
    const showSecondary = !!docxSecondClassToggle.checked;
    if (docxSecondClassRow) {
      docxSecondClassRow.classList.toggle("hidden", !showSecondary);
    }
    if (docxPreviewSecondary) {
      docxPreviewSecondary.classList.toggle("hidden", !showSecondary);
    }
    if (!showSecondary && docxClassSelectSecondary) {
      docxClassSelectSecondary.value = "";
    }
    if (!showSecondary && docxRelayTeacherSecondary) {
      docxRelayTeacherSecondary.checked = false;
    }
    updateDocxPreview();
  });
}

if (docxRelayTeacherPrimary) {
  docxRelayTeacherPrimary.addEventListener("change", updateDocxPreview);
}
if (docxRelayTeacherSecondary) {
  docxRelayTeacherSecondary.addEventListener("change", updateDocxPreview);
}

if (docxGenerateBtn) {
  docxGenerateBtn.addEventListener("click", async () => {
    if (!docxTemplateSelect || !docxClassSelect) return;
    const templateName = docxTemplateSelect.value;
    const classId = docxClassSelect.value;
    const classIdSecondary = docxSecondClassToggle?.checked ? docxClassSelectSecondary?.value || "" : "";
    const useRelayPrimary = !!docxRelayTeacherPrimary?.checked;
    const useRelaySecondary = !!docxRelayTeacherSecondary?.checked;
    if (!templateName) {
      showToast("請選擇模板。");
      return;
    }
    if (!classId) {
      showToast("請選擇班別。");
      return;
    }
    setButtonLoading(docxGenerateBtn, true);
    const response = await invoke('generate_docx', {
      templateName: templateName,
      classId: classId,
      useRelayTeacher: useRelayPrimary,
      classIdSecondary: classIdSecondary,
      useRelayTeacherSecondary: useRelaySecondary,
    });
    setButtonLoading(docxGenerateBtn, false);
    if (!response.ok) {
      showToast(response.error || "生成文件失敗。");
      return;
    }
    if (docxOutput) {
      docxOutput.textContent = response.path ? `已輸出：${response.path}` : "已輸出";
    }
    showToast("文件已生成", "success");
  });
}

if (messageSearchInput) {
  messageSearchInput.addEventListener("input", renderMessageList);
}

if (messageCategorySelect) {
  messageCategorySelect.addEventListener("change", renderMessageList);
}

if (messageList) {
  messageList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    const row = event.target.closest(".message-item");
    if (!button || !row) return;
    const name = row.dataset.name || "";
    if (button.dataset.action === "select-message") {
      selectMessageTemplate(name);
    } else if (button.dataset.action === "edit-category") {
      if (messageCategoryModal && messageCategoryInput) {
        messageCategoryModal.dataset.name = name;
        const target = messageTemplates.find((item) => item.name === name);
        messageCategoryInput.value = target?.category || "";
        messageCategoryModal.classList.remove("hidden");
      }
    }
  });
}

if (messageCopyBtn) {
  messageCopyBtn.addEventListener("click", async () => {
    if (!messageOutput) return;
    await copyToClipboard(messageOutput.value, messageOutput);
  });
}

if (makeupAddRowBtn) {
  makeupAddRowBtn.addEventListener("click", () => {
    if (!makeupRows) return;
    const current = makeupRows.querySelectorAll(".makeup-row").length;
    const row = createMakeupRow(current);
    if (row) {
      const firstCount = makeupRows.querySelector('.makeup-row [data-field="count"]')?.value || "";
      const countInput = row.querySelector('[data-field="count"]');
      if (countInput && firstCount) {
        countInput.value = firstCount;
      }
      makeupRows.appendChild(row);
      updateMakeupRowCount();
      updateMakeupOutput();
    }
  });
}
if (makeupRows) {
  makeupRows.addEventListener("input", updateMakeupOutput);
  makeupRows.addEventListener("change", (event) => {
    const target = event.target;
    if (!target) return;
    if (target.matches('.makeup-row[data-index="0"] [data-field="count"]')) {
      const value = target.value || "";
      makeupRows.querySelectorAll('.makeup-row:not([data-index="0"]) [data-field="count"]').forEach((input) => {
        input.value = value;
      });
      updateMakeupOutput();
      return;
    }
    if (target.matches('[data-field="date"]')) {
      const row = target.closest(".makeup-row");
      const weekdaySelect = row?.querySelector('[data-field="weekday"]');
      const dateValue = parseDate(target.value);
      if (weekdaySelect && dateValue) {
        const jsDay = dateValue.getDay();
        const weekdayIndex = jsDay === 0 ? 0 : jsDay;
        weekdaySelect.value = String(weekdayIndex);
      }
      updateMakeupOutput();
    }
  });
  makeupRows.addEventListener("click", (event) => {
    const button = event.target.closest('button[data-action="remove-row"]');
    if (!button) return;
    const row = button.closest(".makeup-row");
    if (row) {
      row.remove();
      reindexMakeupRows();
      updateMakeupOutput();
    }
  });
}
if (makeupCopyBtn) {
  makeupCopyBtn.addEventListener("click", async () => {
    if (!makeupOutput) return;
    await copyToClipboard(makeupOutput.value, makeupOutput);
  });
}
if (makeupResetBtn) {
  makeupResetBtn.addEventListener("click", resetMakeupForm);
}

if (closeMessageCategoryModal) {
  closeMessageCategoryModal.addEventListener("click", () => {
    if (messageCategoryModal) {
      messageCategoryModal.classList.add("hidden");
    }
  });
}

if (saveMessageCategoryBtn) {
  saveMessageCategoryBtn.addEventListener("click", async () => {
    if (!messageCategoryModal) return;
    const name = messageCategoryModal.dataset.name || "";
    const category = messageCategoryInput?.value || "";
    if (!name) return;
    const response = await invoke('set_message_category', { templateName: name, category });
    if (!response.ok) {
      showToast(response.error || "分類儲存失敗。");
      return;
    }
    showToast("分類已儲存", "success");
    messageCategoryModal.classList.add("hidden");
    await loadMessageTemplates();
    renderMessageList();
  });
}

document.querySelectorAll(".collapse-toggle").forEach((button) => {
  button.addEventListener("click", () => {
    const card = button.closest(".card-collapsible");
    if (!card) return;
    card.classList.toggle("collapsed");
    button.textContent = card.classList.contains("collapsed") ? "展開" : "收起";
  });
});

teacherGenderFilter.addEventListener("change", renderSettings);

calendarMonthBtn.addEventListener("click", () => {
  calendarMode = "month";
  renderCalendar();
});

calendarWeekBtn.addEventListener("click", () => {
  calendarMode = "week";
  renderCalendar();
});

if (calendarTodayBtn) {
  calendarTodayBtn.addEventListener("click", () => {
    calendarAnchor = new Date();
    renderCalendar();
  });
}

calendarPrevBtn.addEventListener("click", () => {
  calendarAnchor = calendarMode === "week" ? addDays(calendarAnchor, -7) : new Date(calendarAnchor.getFullYear(), calendarAnchor.getMonth() - 1, 1);
  renderCalendar();
});

calendarNextBtn.addEventListener("click", () => {
  calendarAnchor = calendarMode === "week" ? addDays(calendarAnchor, 7) : new Date(calendarAnchor.getFullYear(), calendarAnchor.getMonth() + 1, 1);
  renderCalendar();
});

if (calendarJumpMonth) {
  calendarJumpMonth.addEventListener("change", () => {
    const month = Number(calendarJumpMonth.value) - 1;
    calendarAnchor = new Date(calendarAnchor.getFullYear(), month, 1);
    renderCalendar();
  });
}

if (calendarJumpYear) {
  calendarJumpYear.addEventListener("change", () => {
    const year = Number(calendarJumpYear.value);
    calendarAnchor = new Date(year, calendarAnchor.getMonth(), 1);
    renderCalendar();
  });
}

holidayList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-holiday-id]");
  if (button) {
    deleteHoliday(button.dataset.holidayId);
  }
});

calendarView.addEventListener("click", (event) => {
  const toggleBtn = event.target.closest("button[data-action='toggle-room']");
  if (toggleBtn) {
    const roomBlock = toggleBtn.closest(".room-block");
    if (!roomBlock) return;
    const isExpanded = roomBlock.classList.contains("is-expanded");
    roomBlock.classList.toggle("is-expanded", !isExpanded);
    toggleBtn.textContent = isExpanded ? "展開" : "收起";
    return;
  }

  if (calendarMode === "month") {
    const dayCard = event.target.closest("[data-action='drill-down']");
    if (!dayCard || !dayCard.dataset.date) return;
    calendarAnchor = parseDate(dayCard.dataset.date);
    calendarMode = "week";
    renderCalendar();
  }
});

classBody.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (button) {
    const action = button.dataset.action;
    const id = button.dataset.id;
    if (action === "postpone") {
      postponeClass(id);
    } else if (action === "schedule") {
      openScheduleModal(id);
    } else if (action === "end") {
      endClass(id);
    }
    return;
  }

  const row = event.target.closest("tr[data-id]");
  if (row) {
    openScheduleModal(row.dataset.id);
  }
});

function registerSettingActions(container) {
  container.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-setting-type]");
    if (!button) return;
    if (button.dataset.delete) {
      invoke('delete_setting', { entryType: button.dataset.settingType, value: button.dataset.settingValue }).then(loadState);
      return;
    }
    if (button.dataset.move) {
      invoke('move_setting', { entryType: button.dataset.settingType, value: button.dataset.settingValue, direction: button.dataset.move }).then(loadState);
    }
  });
}

registerSettingActions(teacherList);
registerSettingActions(roomList);
registerSettingActions(levelList);
registerSettingActions(timeList);

exportSettingsBtn.addEventListener("click", async () => {
  const response = await invoke('export_settings_csv');
  if (!response.ok) {
    showToast(response.error || "匯出失敗。");
    return;
  }
  const blob = new Blob([response.content], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "settings.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

importSettingsInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  const content = await file.text();
  const response = await invoke('import_settings_csv', { content });
  if (!response.ok) {
    showToast(response.error || "匯入失敗。");
  } else {
    showToast("設定已匯入", "success");
    await loadState();
  }
  event.target.value = "";
});

reminderBody.addEventListener("change", (event) => {
  const checkbox = event.target.closest('input[data-reminder]');
  if (!checkbox) return;
  const fieldMap = {
    doorplate: "doorplate_done",
    questionnaire: "questionnaire_done",
    intro: "intro_done",
  };
  const field = fieldMap[checkbox.dataset.reminder];
  if (!field) return;
  invoke('update_class', { classId: checkbox.dataset.id, updates: { [field]: checkbox.checked } }).then(loadState);
});

closeScheduleModalBtn.addEventListener("click", closeScheduleModal);

if (postponeAutoWeek && postponeOriginalDate && postponeMakeupDate) {
  const syncPostponeMakeup = async () => {
    if (!postponeAutoWeek.checked) {
      postponeMakeupDate.disabled = false;
      return;
    }
    if (!activeClassId || !postponeOriginalDate.value) return;
    const response = await invoke('get_make_up_date', { classId: activeClassId, originalDate: postponeOriginalDate.value });
    if (!response.ok) {
      showToast(response.error || "補課日期計算失敗。");
      return;
    }
    postponeMakeupDate.value = response.make_up_date || "";
    postponeMakeupDate.disabled = true;
  };
  postponeAutoWeek.addEventListener("change", syncPostponeMakeup);
  postponeOriginalDate.addEventListener("change", syncPostponeMakeup);
}

terminateClassBtn.addEventListener("click", async () => {
  if (!activeClassId) return;
  const lastDate = prompt("請輸入最後一堂日期 (YYYY-MM-DD):");
  if (!lastDate) return;
  const firstConfirm = confirm("確定要結束班別嗎？此動作會調整總課節。");
  if (!firstConfirm) return;
  const secondConfirm = confirm(`最後一堂日期：${lastDate}\n確認結束？`);
  if (!secondConfirm) return;
  const response = await invoke('terminate_class_with_last_date', { classId: activeClassId, lastDate: lastDate });
  if (!response.ok) {
    showToast(response.error || "結束班別失敗。");
    return;
  }
  showToast("班別已結束", "success");
  await loadState();
  closeScheduleModal();
});

saveClassDetailBtn.addEventListener("click", async () => {
  if (!activeClassId) return;
  const skuValue = detailSku.value.trim();
  if (!skuValue) {
    showToast("班別不能為空。");
    return;
  }
  setButtonLoading(saveClassDetailBtn, true);
  const response = await invoke('update_class', { classId: activeClassId, updates: {
    sku: skuValue,
    classroom: detailRoom.value.trim(),
    teacher: detailTeacher.value.trim(),
    start_time: detailTime.value.trim(),
    relay_teacher: detailRelayTeacher.value.trim(),
    relay_date: detailRelayDate.value,
    student_count: Number(detailStudents.value || 0),
  } });
  setButtonLoading(saveClassDetailBtn, false);
  if (response && response.ok === false) {
    showToast(response.error || "更新失敗。");
    return;
  }
  showToast("資料已儲存", "success");
  await loadState();
  await openScheduleModal(activeClassId);
});

if (deleteClassBtn) {
  deleteClassBtn.addEventListener("click", async () => {
    if (!activeClassId) return;
    const confirmed = confirm("確定要刪除班別？此動作無法復原。");
    if (!confirmed) return;
    const response = await invoke('delete_class', { classId: activeClassId });
    if (!response.ok) {
      showToast(response.error || "刪除班別失敗。");
      return;
    }
    showToast("班別已刪除", "success");
    closeScheduleModal();
    await loadState();
  });
}

addLessonBtn.addEventListener("click", async () => {
  if (!activeClassId || !addLessonDate.value) return;
  const response = await invoke('add_schedule_override', { classId: activeClassId, dateStr: addLessonDate.value, action: "add" });
  if (!response.ok) {
    showToast(response.error || "新增失敗。");
    return;
  }
  await openScheduleModal(activeClassId);
});

removeLessonBtn.addEventListener("click", async () => {
  if (!activeClassId || !removeLessonDate.value) return;
  const response = await invoke('add_schedule_override', { classId: activeClassId, dateStr: removeLessonDate.value, action: "remove" });
  if (!response.ok) {
    showToast(response.error || "取消失敗。");
    return;
  }
  await openScheduleModal(activeClassId);
});

addPostponeBtn.addEventListener("click", async () => {
  if (!activeClassId) return;
  let response;
  if (postponeAutoWeek && postponeAutoWeek.checked) {
    response = await invoke('add_postpone', { classId: activeClassId, originalDate: postponeOriginalDate.value, reason: postponeReason.value || "" });
  } else {
    response = await invoke('add_postpone_manual', {
      classId: activeClassId,
      originalDate: postponeOriginalDate.value,
      makeUpDate: postponeMakeupDate.value,
      reason: postponeReason.value || "",
    });
  }
  if (!response.ok) {
    showToast(response.error || "新增改期失敗。");
    return;
  }
  showToast("改期已新增", "success");
  if (response.reactivated) {
    showToast("班別已由已結束重新激活", "info");
  }
  await openScheduleModal(activeClassId);
});

postponeList.addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-postpone-id]");
  if (!button) return;
  await invoke('delete_postpone', { postponeId: button.dataset.postponeId });
  await openScheduleModal(activeClassId);
});

overrideList.addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-override-id]");
  if (!button) return;
  await invoke('delete_schedule_override', { overrideId: button.dataset.overrideId });
  await openScheduleModal(activeClassId);
});

exportClassesBtn.addEventListener("click", async () => {
  const response = await invoke('export_classes_csv');
  if (!response.ok) {
    showToast(response.error || "匯出失敗。");
    return;
  }
  const blob = new Blob([response.content], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "classes.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

importClassesInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  const content = await file.text();
  const response = await invoke('import_classes_csv', { content });
  if (!response.ok) {
    showToast(response.error || "匯入失敗。");
  } else {
    showToast("班別資料已匯入", "success");
    await loadState();
  }
  event.target.value = "";
});

// ---- Task Event Listeners ----
if (taskAddBtn) {
  taskAddBtn.addEventListener("click", addTask);
}
if (taskInput) {
  taskInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      addTask();
    }
  });
}
if (taskList) {
  taskList.addEventListener("change", (e) => {
    const checkbox = e.target.closest("input[data-task-id]");
    if (checkbox) toggleTask(checkbox.dataset.taskId);
  });
  taskList.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-task-delete]");
    if (btn) deleteTask(btn.dataset.taskDelete);
  });
}
if (taskDoneList) {
  taskDoneList.addEventListener("change", (e) => {
    const checkbox = e.target.closest("input[data-task-id]");
    if (checkbox) toggleTask(checkbox.dataset.taskId);
  });
  taskDoneList.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-task-delete]");
    if (btn) deleteTask(btn.dataset.taskDelete);
  });
}
if (taskDoneToggle && taskDoneList) {
  taskDoneToggle.addEventListener("click", () => {
    taskDoneList.classList.toggle("collapsed");
  });
}
if (taskClearDone) {
  taskClearDone.addEventListener("click", clearDoneTasks);
}

// ── Promote Notice Tab ──────────────────────────────────────────────────────

let promoteLevelLabels = { sourceShort: "", targetShort: "" };

function getPromoteBodyText() {
  const month = promoteMonthInput?.value || "";
  const { sourceShort, targetShort } = promoteLevelLabels;
  return `本班同學將於${month}月份${sourceShort}班畢業，並升上${targetShort}班。新課程詳情如下:`;
}

function updatePromoteBodySuffix() {
  if (promoteBodySuffix) {
    const { sourceShort, targetShort } = promoteLevelLabels;
    promoteBodySuffix.textContent = `月份${sourceShort}班畢業，並升上${targetShort}班。新課程詳情如下:`;
  }
}

function _buildLevelOrder() {
  const levelNext = appState.settings?.level_next || {};
  const allSources = new Set(Object.keys(levelNext));
  const allTargets = new Set(Object.values(levelNext));
  const start = [...allSources].find(l => !allTargets.has(l));
  const order = [];
  let cur = start;
  while (cur && !order.includes(cur)) {
    order.push(cur);
    cur = levelNext[cur];
  }
  return order;
}

function renderPromoteClassSelects() {
  if (!promoteSourceClassSelect) return;
  const classes = appState.classes || [];
  const archivedStatuses = new Set(["promoted", "merged", "terminated", "ended"]);
  const today = new Date();
  const oneMonthLater = new Date(today);
  oneMonthLater.setMonth(oneMonthLater.getMonth() + 1);

  // Only classes ending within 1 month
  const ending = classes.filter(c => {
    if (archivedStatuses.has(c.status) || c.status === "not_started") return false;
    if (!c.end_date) return false;
    const endDate = new Date(c.end_date);
    return endDate <= oneMonthLater;
  });

  // Sort by level using the level_next chain order
  const levelOrder = _buildLevelOrder();
  ending.sort((a, b) => {
    const ai = levelOrder.indexOf(a.level);
    const bi = levelOrder.indexOf(b.level);
    return (ai >= 0 ? ai : 999) - (bi >= 0 ? bi : 999);
  });

  promoteSourceClassSelect.innerHTML = `<option value="">請選擇</option>` +
    ending.map(c => `<option value="${c.id}">${c.sku}</option>`).join("");
  promoteTargetClassSelect.innerHTML = `<option value="">請選擇（先選來源班別）</option>`;

  // Default signature date to today's MM/YYYY
  if (promoteFieldSignatureDate && !promoteFieldSignatureDate.value) {
    const m = String(today.getMonth() + 1).padStart(2, "0");
    promoteFieldSignatureDate.value = `${m}/${today.getFullYear()}`;
  }
}

function renderPromotePreview() {
  if (!promotePreviewCard) return;
  const addressee = promoteAddresseeInput?.value || "";
  const body = getPromoteBodyText();
  const name = promoteFieldName?.value || "";
  const startDate = promoteFieldStartDate?.value || "";
  const duration = promoteFieldDuration?.value || "";
  const time = promoteFieldTime?.value || "";
  const teacher = promoteFieldTeacher?.value || "";
  const location = promoteFieldLocation?.value || "";
  const remarks = promoteFieldRemarks?.value || "";
  const sigDate = promoteFieldSignatureDate?.value || "";

  // Compute textbook fee for preview
  const targetId = promoteTargetClassSelect?.value;
  const classes = appState.classes || [];
  const targetCls = classes.find(c => c.id === targetId);
  let tbFeeStr = "";
  if (promoteIncludeTextbook?.checked !== false && targetCls) {
    const levelTextbook = appState.settings?.level_textbook || {};
    const textbookPrices = appState.settings?.textbook || {};
    const books = Array.isArray(levelTextbook[targetCls.level]) ? levelTextbook[targetCls.level] : [];
    if (books.length) {
      const parts = books.map(n => `${n} $${textbookPrices[n] || 0}`);
      const total = books.reduce((s, n) => s + (textbookPrices[n] || 0), 0);
      tbFeeStr = parts.join("、") + (books.length > 1 ? `（合計 $${total}）` : "");
    }
  }

  promotePreviewCard.innerHTML = `
    <p>致${addressee}，</p>
    <p>${body}</p>
    <table>
      <thead>
        <tr>
          <th>名稱</th><th>開課日期</th><th>全課程修讀期</th>
          <th>時間</th><th>導師</th><th>上課地點</th><th>備註</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>${name}</td><td>${startDate}</td><td>${duration}</td>
          <td>${time}</td><td>${teacher}</td><td>${location}</td><td>${remarks}</td>
        </tr>
      </tbody>
    </table>
    ${tbFeeStr ? `<p><strong>書本費：</strong>${tbFeeStr}</p>` : ""}
    <p class="promote-footer-text">同學收到成績後，如欲報讀升級課程，請即日辦理報名手續，未有即日報名的同學，將列入候補名單內，如需留位，必須於當日到校務處登記，另作安排。開課日不辦理新登記手續，由於學位所限，請同學合作。</p>
    <p style="text-align:right">校務處<br>${sigDate}</p>
  `;
}

async function onPromoteSourceChange() {
  const sourceId = promoteSourceClassSelect?.value;
  if (!promoteTargetClassSelect) return;
  promoteTargetClassSelect.innerHTML = `<option value="">請選擇</option>`;

  if (!sourceId) return;

  const classes = appState.classes || [];
  const source = classes.find(c => c.id === sourceId);
  if (!source) return;

  // Auto-set addressee from source level
  if (promoteAddresseeInput && !promoteAddresseeInput.value) {
    promoteAddresseeInput.value = `${source.level}班同學`;
  }

  // If source has a promoted_to_id, pick that class as target
  const levelNext = appState.settings?.level_next || {};
  const nextLevel = levelNext[source.level] || "";

  let targetCandidates = [];
  if (source.promoted_to_id) {
    targetCandidates = classes.filter(c => c.id === source.promoted_to_id);
  }
  if (targetCandidates.length === 0 && nextLevel) {
    targetCandidates = classes.filter(c =>
      c.level === nextLevel && !["terminated", "merged"].includes(c.status)
    );
  }

  promoteTargetClassSelect.innerHTML = `<option value="">請選擇</option>` +
    targetCandidates.map(c => `<option value="${c.id}">${c.sku}</option>`).join("") +
    `<option value="__custom__">自訂班別</option>`;

  // If there's exactly one candidate, auto-select it
  if (targetCandidates.length === 1) {
    promoteTargetClassSelect.value = targetCandidates[0].id;
    await onPromoteTargetChange();
  }
}

async function onPromoteTargetChange() {
  const targetId = promoteTargetClassSelect?.value;
  if (!targetId) return;

  const classes = appState.classes || [];
  const source = classes.find(c => c.id === promoteSourceClassSelect?.value);

  // Handle custom class: clear fields and let user fill from scratch
  if (targetId === "__custom__") {
    if (promoteFieldName) promoteFieldName.value = "";
    if (promoteFieldStartDate) promoteFieldStartDate.value = "";
    if (promoteFieldDuration) promoteFieldDuration.value = "";
    if (promoteFieldTime) promoteFieldTime.value = "";
    if (promoteFieldTeacher) promoteFieldTeacher.value = "";
    if (promoteFieldLocation) promoteFieldLocation.value = "";
    if (promoteFieldRemarks) promoteFieldRemarks.value = "";
    if (promoteFieldSignatureDate) {
      const now = new Date();
      promoteFieldSignatureDate.value = `${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
    }
    // Auto-derive level labels from source via level_next
    const levelNext = appState.settings?.level_next || {};
    const sourceShort = source ? source.level.replace("級", "") : "";
    const targetShort = levelNext[source?.level]?.replace("級", "") || "";
    promoteLevelLabels = { sourceShort, targetShort };
    updatePromoteBodySuffix();
    if (promoteMonthInput) promoteMonthInput.value = "";
    if (promoteTextbookInfo) promoteTextbookInfo.classList.add("hidden");
    renderPromotePreview();
    return;
  }

  const target = classes.find(c => c.id === targetId);

  const response = await invoke('get_promote_notice_data', { classId: targetId });
  if (!response.ok) {
    showToast(response.error || "無法載入班別資料。");
    return;
  }

  if (promoteFieldName) promoteFieldName.value = response.name || "";
  if (promoteFieldStartDate) promoteFieldStartDate.value = response.start_date_formatted || "";
  if (promoteFieldDuration) promoteFieldDuration.value = response.duration || "";
  if (promoteFieldTime) promoteFieldTime.value = response.time || "";
  if (promoteFieldTeacher) promoteFieldTeacher.value = response.teacher || "";
  if (promoteFieldLocation) promoteFieldLocation.value = response.location || "";
  if (promoteFieldRemarks) promoteFieldRemarks.value = response.remarks || "";

  // Signature date: always default to today's MM/YYYY
  if (promoteFieldSignatureDate) {
    const now = new Date();
    promoteFieldSignatureDate.value = `${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
  }

  // Update level labels for body text suffix
  const sourceShort = response.source_level_short || (source ? source.level.replace("級", "") : "");
  const targetShort = response.target_level_short || "";
  promoteLevelLabels = { sourceShort, targetShort };
  updatePromoteBodySuffix();

  // Pre-fill month from start_month
  if (promoteMonthInput && response.start_month) {
    promoteMonthInput.value = response.start_month;
  }

  // Show textbook fee for the target class level
  if (target && promoteTextbookInfo && promoteTextbookList) {
    const levelTextbook = appState.settings?.level_textbook || {};
    const textbookPrices = appState.settings?.textbook || {};
    const books = Array.isArray(levelTextbook[target.level]) ? levelTextbook[target.level] : [];
    if (books.length) {
      promoteTextbookInfo.classList.remove("hidden");
      let total = 0;
      promoteTextbookList.innerHTML = books.map(name => {
        const price = textbookPrices[name] || 0;
        total += price;
        return `<span class="promote-textbook-item">${name} <strong>$${price}</strong></span>`;
      }).join("") + (books.length > 1 ? `<span class="promote-textbook-item promote-textbook-total">合計 <strong>$${total}</strong></span>` : "");
    } else {
      promoteTextbookInfo.classList.add("hidden");
    }
  }

  renderPromotePreview();
}

const promoteNoticeForm = document.getElementById("promoteNoticeForm");
if (promoteNoticeForm) {
  promoteNoticeForm.addEventListener("submit", (e) => e.preventDefault());
}

if (promoteSourceClassSelect) {
  promoteSourceClassSelect.addEventListener("change", onPromoteSourceChange);
}
if (promoteTargetClassSelect) {
  promoteTargetClassSelect.addEventListener("change", onPromoteTargetChange);
}

const promoteInputFields = [
  promoteAddresseeInput, promoteMonthInput, promoteFieldName, promoteFieldStartDate,
  promoteFieldDuration, promoteFieldTime, promoteFieldTeacher, promoteFieldLocation,
  promoteFieldRemarks, promoteFieldSignatureDate,
];
promoteInputFields.forEach(el => {
  if (el) el.addEventListener("input", renderPromotePreview);
});
if (promoteIncludeTextbook) {
  promoteIncludeTextbook.addEventListener("change", renderPromotePreview);
}

if (promoteGenerateBtn) {
  promoteGenerateBtn.addEventListener("click", async () => {
    const targetId = promoteTargetClassSelect?.value;
    if (!targetId) { showToast("請先選擇目標班別。"); return; }
    const classes = appState.classes || [];
    const source = classes.find(c => c.id === promoteSourceClassSelect?.value);
    const target = classes.find(c => c.id === targetId);
    const isCustom = targetId === "__custom__";

    // Build textbook fee string for the DOCX
    let textbookFeeStr = "";
    if (!isCustom && promoteIncludeTextbook?.checked !== false && target) {
      const levelTextbook = appState.settings?.level_textbook || {};
      const textbookPrices = appState.settings?.textbook || {};
      const books = Array.isArray(levelTextbook[target.level]) ? levelTextbook[target.level] : [];
      if (books.length) {
        const parts = books.map(name => {
          const price = textbookPrices[name] || 0;
          return `${name} $${price}`;
        });
        const total = books.reduce((s, name) => s + (textbookPrices[name] || 0), 0);
        textbookFeeStr = parts.join("、") + (books.length > 1 ? `（合計 $${total}）` : "");
      }
    }

    const data = {
      sku: isCustom ? (source?.sku || "custom_promote") : (target?.sku || "promote"),
      addressee: promoteAddresseeInput?.value || "",
      body_text: getPromoteBodyText(),
      name: promoteFieldName?.value || "",
      start_date_formatted: promoteFieldStartDate?.value || "",
      duration: promoteFieldDuration?.value || "",
      time: promoteFieldTime?.value || "",
      teacher: promoteFieldTeacher?.value || "",
      location: promoteFieldLocation?.value || "",
      remarks: promoteFieldRemarks?.value || "",
      textbook_fee: textbookFeeStr,
      signature_date: promoteFieldSignatureDate?.value || "",
    };

    setButtonLoading(promoteGenerateBtn, true);
    const response = await invoke('generate_promote_notice', { data });
    setButtonLoading(promoteGenerateBtn, false);

    if (!response.ok) {
      showToast(response.error || "生成升班通知失敗。");
      return;
    }
    if (promoteOutput) {
      promoteOutput.textContent = response.path ? `已輸出：${response.path}` : "已輸出";
    }
    showToast("升班通知已生成", "success");
  });
}

if (promoteOpenFolderBtn) {
  promoteOpenFolderBtn.addEventListener("click", async () => {
    await invoke('open_output_folder');
  });
}

if (promoteResetBtn) {
  promoteResetBtn.addEventListener("click", () => {
    if (promoteSourceClassSelect) promoteSourceClassSelect.value = "";
    if (promoteTargetClassSelect) promoteTargetClassSelect.innerHTML = `<option value="">請選擇（先選來源班別）</option>`;
    if (promoteAddresseeInput) promoteAddresseeInput.value = "";
    if (promoteMonthInput) promoteMonthInput.value = "";
    if (promoteFieldName) promoteFieldName.value = "";
    if (promoteFieldStartDate) promoteFieldStartDate.value = "";
    if (promoteFieldDuration) promoteFieldDuration.value = "";
    if (promoteFieldTime) promoteFieldTime.value = "";
    if (promoteFieldTeacher) promoteFieldTeacher.value = "";
    if (promoteFieldLocation) promoteFieldLocation.value = "";
    if (promoteFieldRemarks) promoteFieldRemarks.value = "";
    if (promoteFieldSignatureDate) promoteFieldSignatureDate.value = "";
    if (promoteIncludeTextbook) promoteIncludeTextbook.checked = true;
    if (promoteTextbookInfo) promoteTextbookInfo.classList.add("hidden");
    if (promoteOutput) promoteOutput.textContent = "";
    promoteLevelLabels = { sourceShort: "", targetShort: "" };
    updatePromoteBodySuffix();
    if (promotePreviewCard) promotePreviewCard.innerHTML = `<p class="muted-hint">請選擇目標班別以預覽通知。</p>`;
  });
}

if (docxOpenFolderBtn) {
  docxOpenFolderBtn.addEventListener("click", async () => {
    await invoke('open_output_folder');
  });
}

// ---------------------------------------------------------------------------
// EPS Audit Tab
// ---------------------------------------------------------------------------
const epsDatePicker = document.getElementById("epsDatePicker");
const epsPrevDay = document.getElementById("epsPrevDay");
const epsNextDay = document.getElementById("epsNextDay");
const epsToday = document.getElementById("epsToday");
const epsDayOfWeek = document.getElementById("epsDayOfWeek");
const epsPeriodBefore = document.getElementById("epsPeriodBefore");
const epsPeriodAfter = document.getElementById("epsPeriodAfter");
const epsItemsTable = document.getElementById("epsItemsTable");
const epsClassSubtotal = document.getElementById("epsClassSubtotal");
const epsBookSubtotal = document.getElementById("epsBookSubtotal");
const epsOtherSubtotal = document.getElementById("epsOtherSubtotal");
const epsBeforeTotal = document.getElementById("epsBeforeTotal");
const epsAfterTotal = document.getElementById("epsAfterTotal");
const epsPastDay = document.getElementById("epsPastDay");
const epsSheetTotal = document.getElementById("epsSheetTotal");
const epsOp1 = document.getElementById("epsOp1");
const epsOp2 = document.getElementById("epsOp2");
const epsOp3 = document.getElementById("epsOp3");
const epsOpsSum = document.getElementById("epsOpsSum");
const epsCheckResult = document.getElementById("epsCheckResult");
const epsSaveBtn = document.getElementById("epsSaveBtn");
const epsExportBtn = document.getElementById("epsExportBtn");
const epsOutputPath = document.getElementById("epsOutputPath");
const epsSetPathBtn = document.getElementById("epsSetPathBtn");
const epsHistoryList = document.getElementById("epsHistoryList");

const epsDayNames = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];

const epsState = {
  items: [],
  records: { before: [], after: [] },
  audit: {
    operator_1_before: 0, operator_2_before: 0, operator_3_before: 0,
    operator_1_after: 0, operator_2_after: 0, operator_3_after: 0,
  },
  pastDayCarry: 0,
  currentDate: "",
  currentPeriod: "before",
  loaded: false,
};

function epsFormatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function epsSetDate(dateStr) {
  epsState.currentDate = dateStr;
  if (epsDatePicker) epsDatePicker.value = dateStr;
  const d = new Date(dateStr + "T00:00:00");
  if (epsDayOfWeek && !isNaN(d.getTime())) {
    const dow = (d.getDay() + 6) % 7; // Monday=0
    epsDayOfWeek.textContent = epsDayNames[dow] || "";
  }
  // Always reset to before-1900 when changing date
  epsState.currentPeriod = "before";
  if (epsPeriodBefore) epsPeriodBefore.classList.add("active");
  if (epsPeriodAfter) epsPeriodAfter.classList.remove("active");
  epsLoadRecord(dateStr);
}

async function epsInit() {
  if (epsState.loaded) return;
  try {
    const res = await invoke('load_eps_items');
    if (res && res.ok) {
      epsState.items = res.items;
    }
  } catch (e) {
    console.error("Failed to load EPS items", e);
  }
  epsState.loaded = true;
  const savedPath = appState.app_config?.eps_output_path || "";
  if (epsOutputPath) epsOutputPath.value = savedPath;
  const today = epsFormatDate(new Date());
  epsSetDate(today);
}

async function epsLoadRecord(dateStr) {
  try {
    const res = await invoke('load_eps_record', { dateStr: dateStr });
    if (res && res.ok) {
      epsState.records = res.records;
      epsState.audit = res.audit;
      epsState.pastDayCarry = res.past_day_carry || 0;
      epsLoadOperatorInputs();
    }
  } catch (e) {
    console.error("Failed to load EPS record", e);
    epsState.records = { before: [], after: [] };
    epsState.audit = {
      operator_1_before: 0, operator_2_before: 0, operator_3_before: 0,
      operator_1_after: 0, operator_2_after: 0, operator_3_after: 0,
    };
    epsState.pastDayCarry = 0;
  }
  epsRenderItems();
  epsComputeTotals();
}

function epsSaveOperatorInputs() {
  const p = epsState.currentPeriod;
  const suffix = p === "before" ? "_before" : "_after";
  epsState.audit["operator_1" + suffix] = Math.max(0, parseInt(epsOp1?.value) || 0);
  epsState.audit["operator_2" + suffix] = Math.max(0, parseInt(epsOp2?.value) || 0);
  epsState.audit["operator_3" + suffix] = Math.max(0, parseInt(epsOp3?.value) || 0);
}

function epsLoadOperatorInputs() {
  const p = epsState.currentPeriod;
  const suffix = p === "before" ? "_before" : "_after";
  if (epsOp1) epsOp1.value = epsState.audit["operator_1" + suffix] || 0;
  if (epsOp2) epsOp2.value = epsState.audit["operator_2" + suffix] || 0;
  if (epsOp3) epsOp3.value = epsState.audit["operator_3" + suffix] || 0;
  const label = document.getElementById("epsOpsLabel");
  if (label) label.textContent = p === "before" ? "操作員 (1900前)" : "操作員 (1900後)";
  const combinedRow = document.getElementById("epsOpsCombinedRow");
  if (combinedRow) combinedRow.classList.toggle("hidden", p === "before");
}

function epsRenderItems() {
  if (!epsItemsTable) return;
  const items = epsState.items;
  const period = epsState.currentPeriod;
  const periodData = epsState.records[period] || [];

  let html = `<table class="eps-table"><thead><tr>
    <th>項目</th><th>單價</th>
    <th>K</th><th>L</th><th>HK</th><th>小計</th>
  </tr></thead><tbody>`;

  let currentSection = null;
  const sectionLabels = { class: "班別", book: "書", other: "其他" };

  items.forEach((item, idx) => {
    if (item.section !== currentSection) {
      currentSection = item.section;
      html += `<tr class="eps-section-header"><td colspan="6">${sectionLabels[currentSection] || currentSection}</td></tr>`;
    }
    const rec = periodData[idx] || { qty_K: 0, qty_L: 0, qty_HK: 0 };
    const qk = rec.qty_K || 0;
    const ql = rec.qty_L || 0;
    const qh = rec.qty_HK || 0;
    const sub = item.price * (qk + ql + qh);

    html += `<tr>
      <td class="eps-item-name">${item.name}</td>
      <td class="eps-price">$${item.price.toLocaleString()}</td>
      <td><input type="number" class="eps-qty" data-idx="${idx}" data-loc="K" min="0" value="${qk}" /></td>
      <td><input type="number" class="eps-qty" data-idx="${idx}" data-loc="L" min="0" value="${ql}" /></td>
      <td><input type="number" class="eps-qty" data-idx="${idx}" data-loc="HK" min="0" value="${qh}" /></td>
      <td class="eps-subtotal" data-idx="${idx}">$${sub.toLocaleString()}</td>
    </tr>`;
  });

  html += `</tbody></table>`;
  epsItemsTable.innerHTML = html;

  // Attach input listeners
  epsItemsTable.querySelectorAll(".eps-qty").forEach((input) => {
    input.addEventListener("input", epsOnQtyChange);
  });
}

function epsOnQtyChange(e) {
  const idx = parseInt(e.target.dataset.idx);
  const loc = e.target.dataset.loc;
  const val = Math.max(0, parseInt(e.target.value) || 0);
  const period = epsState.currentPeriod;

  if (!epsState.records[period][idx]) {
    epsState.records[period][idx] = { item_name: epsState.items[idx]?.name || "", qty_K: 0, qty_L: 0, qty_HK: 0 };
  }
  epsState.records[period][idx][`qty_${loc}`] = val;

  // Update subtotal cell
  const item = epsState.items[idx];
  if (item) {
    const rec = epsState.records[period][idx];
    const sub = item.price * ((rec.qty_K || 0) + (rec.qty_L || 0) + (rec.qty_HK || 0));
    const cell = epsItemsTable.querySelector(`.eps-subtotal[data-idx="${idx}"]`);
    if (cell) cell.textContent = `$${sub.toLocaleString()}`;
  }

  epsComputeTotals();
}

function epsComputeTotals() {
  const items = epsState.items;
  let beforeTotal = 0, afterTotal = 0;
  const sectionBefore = { class: 0, book: 0, other: 0 };
  const sectionAfter = { class: 0, book: 0, other: 0 };

  for (const period of ["before", "after"]) {
    const data = epsState.records[period] || [];
    items.forEach((item, idx) => {
      const rec = data[idx] || { qty_K: 0, qty_L: 0, qty_HK: 0 };
      const sub = item.price * ((rec.qty_K || 0) + (rec.qty_L || 0) + (rec.qty_HK || 0));
      if (period === "before") {
        beforeTotal += sub;
        sectionBefore[item.section] = (sectionBefore[item.section] || 0) + sub;
      } else {
        afterTotal += sub;
        sectionAfter[item.section] = (sectionAfter[item.section] || 0) + sub;
      }
    });
  }

  // Combined section totals (both periods)
  if (epsClassSubtotal) epsClassSubtotal.textContent = `$${(sectionBefore.class + sectionAfter.class).toLocaleString()}`;
  if (epsBookSubtotal) epsBookSubtotal.textContent = `$${(sectionBefore.book + sectionAfter.book).toLocaleString()}`;
  if (epsOtherSubtotal) epsOtherSubtotal.textContent = `$${(sectionBefore.other + sectionAfter.other).toLocaleString()}`;

  if (epsBeforeTotal) epsBeforeTotal.textContent = `$${beforeTotal.toLocaleString()}`;
  if (epsAfterTotal) epsAfterTotal.textContent = `$${afterTotal.toLocaleString()}`;
  if (epsPastDay) epsPastDay.textContent = `$${epsState.pastDayCarry.toLocaleString()}`;

  const sheetTotal = beforeTotal;
  if (epsSheetTotal) epsSheetTotal.textContent = `$${sheetTotal.toLocaleString()}`;

  // Operators — save current inputs into state, then read both periods from state
  epsSaveOperatorInputs();

  const opsSumBefore = (epsState.audit.operator_1_before || 0)
    + (epsState.audit.operator_2_before || 0)
    + (epsState.audit.operator_3_before || 0);
  const opsSumAfter = (epsState.audit.operator_1_after || 0)
    + (epsState.audit.operator_2_after || 0)
    + (epsState.audit.operator_3_after || 0);

  // Current period's sum for display
  const currentOpsSum = epsState.currentPeriod === "before" ? opsSumBefore : opsSumAfter;
  if (epsOpsSum) epsOpsSum.textContent = `$${currentOpsSum.toLocaleString()}`;

  // Combined per-operator totals (before + after) — shown only in after-1900 view
  const c1 = (epsState.audit.operator_1_before || 0) + (epsState.audit.operator_1_after || 0);
  const c2 = (epsState.audit.operator_2_before || 0) + (epsState.audit.operator_2_after || 0);
  const c3 = (epsState.audit.operator_3_before || 0) + (epsState.audit.operator_3_after || 0);
  const opCombined1 = document.getElementById("epsOpCombined1");
  const opCombined2 = document.getElementById("epsOpCombined2");
  const opCombined3 = document.getElementById("epsOpCombined3");
  const opsCombined = document.getElementById("epsOpsCombined");
  if (opCombined1) opCombined1.textContent = `$${c1.toLocaleString()}`;
  if (opCombined2) opCombined2.textContent = `$${c2.toLocaleString()}`;
  if (opCombined3) opCombined3.textContent = `$${c3.toLocaleString()}`;
  if (opsCombined) opsCombined.textContent = `$${(c1 + c2 + c3).toLocaleString()}`;

  // Audit check
  const epsCheckResult2 = document.getElementById("epsCheckResult2");
  let statusText = "";
  let statusClass = "";
  let statusText2 = "";
  let statusClass2 = "";

  if (epsState.currentPeriod === "before") {
    // Before-1900: single check — opsSumBefore + pastDayCarry == beforeTotal
    const auditOps = opsSumBefore + epsState.pastDayCarry;
    if (auditOps === 0 && sheetTotal === 0) {
      statusText = "";
      statusClass = "";
    } else if (auditOps === sheetTotal) {
      statusText = "OK";
      statusClass = "eps-ok";
    } else {
      const diff = auditOps - sheetTotal;
      const sign = diff > 0 ? "+" : "";
      statusText = `MISMATCH (差額: ${sign}$${diff.toLocaleString()})`;
      statusClass = "eps-mismatch";
    }
    if (epsCheckResult2) epsCheckResult2.classList.add("hidden");
  } else {
    // After-1900: two indicators
    // Indicator 1 — current day: opsSumAfter == afterTotal
    if (opsSumAfter === 0 && afterTotal === 0) {
      statusText = "";
      statusClass = "";
    } else if (opsSumAfter === afterTotal) {
      statusText = "當日核數: OK";
      statusClass = "eps-ok";
    } else {
      const diff = opsSumAfter - afterTotal;
      const sign = diff > 0 ? "+" : "";
      statusText = `當日核數: MISMATCH (差額: ${sign}$${diff.toLocaleString()})`;
      statusClass = "eps-mismatch";
    }

    // Indicator 2 — full day audit: (beforeTotal + afterTotal) == (opsSumBefore + opsSumAfter)
    const sheetCombined = beforeTotal + afterTotal;
    const opsCombinedTotal = opsSumBefore + opsSumAfter;
    if (opsCombinedTotal === 0 && sheetCombined === 0) {
      statusText2 = "";
      statusClass2 = "";
    } else if (sheetCombined === opsCombinedTotal) {
      statusText2 = "總審計: OK";
      statusClass2 = "eps-ok";
    } else {
      const diff2 = opsCombinedTotal - sheetCombined;
      const sign2 = diff2 > 0 ? "+" : "";
      statusText2 = `總審計: MISMATCH (差額: ${sign2}$${diff2.toLocaleString()})`;
      statusClass2 = "eps-mismatch";
    }
    if (epsCheckResult2) {
      epsCheckResult2.textContent = statusText2;
      epsCheckResult2.className = "eps-check-result" + (statusClass2 ? " " + statusClass2 : "");
      epsCheckResult2.classList.remove("hidden");
    }
  }

  if (epsCheckResult) {
    epsCheckResult.textContent = statusText;
    epsCheckResult.className = "eps-check-result" + (statusClass ? " " + statusClass : "");
  }

  // Sticky bar
  const stickySheet = document.getElementById("epsStickySheet");
  const stickyOps = document.getElementById("epsStickyOps");
  const stickyPastDay = document.getElementById("epsStickyPastDay");
  const stickyStatus = document.getElementById("epsStickyStatus");
  if (epsState.currentPeriod === "before") {
    if (stickySheet) stickySheet.textContent = `$${sheetTotal.toLocaleString()}`;
    if (stickyOps) stickyOps.textContent = `$${opsSumBefore.toLocaleString()}`;
    if (stickyPastDay) stickyPastDay.textContent = `$${epsState.pastDayCarry.toLocaleString()}`;
  } else {
    if (stickySheet) stickySheet.textContent = `$${afterTotal.toLocaleString()}`;
    if (stickyOps) stickyOps.textContent = `$${opsSumAfter.toLocaleString()}`;
    if (stickyPastDay) stickyPastDay.textContent = `$${epsState.pastDayCarry.toLocaleString()}`;
  }
  if (stickyStatus) {
    stickyStatus.textContent = statusText;
    stickyStatus.className = "eps-sticky-status" + (statusClass ? " " + statusClass : "");
  }
}

// Period toggle
if (epsPeriodBefore) {
  epsPeriodBefore.addEventListener("click", () => {
    epsSaveOperatorInputs();
    epsState.currentPeriod = "before";
    epsPeriodBefore.classList.add("active");
    if (epsPeriodAfter) epsPeriodAfter.classList.remove("active");
    epsLoadOperatorInputs();
    epsRenderItems();
  });
}
if (epsPeriodAfter) {
  epsPeriodAfter.addEventListener("click", () => {
    epsSaveOperatorInputs();
    epsState.currentPeriod = "after";
    epsPeriodAfter.classList.add("active");
    if (epsPeriodBefore) epsPeriodBefore.classList.remove("active");
    epsLoadOperatorInputs();
    epsRenderItems();
  });
}

// Date navigation
if (epsDatePicker) {
  epsDatePicker.addEventListener("change", () => {
    if (epsDatePicker.value) epsSetDate(epsDatePicker.value);
  });
}
if (epsPrevDay) {
  epsPrevDay.addEventListener("click", () => {
    const d = new Date(epsState.currentDate + "T00:00:00");
    d.setDate(d.getDate() - 1);
    epsSetDate(epsFormatDate(d));
  });
}
if (epsNextDay) {
  epsNextDay.addEventListener("click", () => {
    const d = new Date(epsState.currentDate + "T00:00:00");
    d.setDate(d.getDate() + 1);
    epsSetDate(epsFormatDate(d));
  });
}
if (epsToday) {
  epsToday.addEventListener("click", () => {
    epsSetDate(epsFormatDate(new Date()));
  });
}

// Operator inputs trigger recompute
[epsOp1, epsOp2, epsOp3].forEach((el) => {
  if (el) el.addEventListener("input", epsComputeTotals);
});

// Save
if (epsSaveBtn) {
  epsSaveBtn.addEventListener("click", async () => {
    const dateStr = epsState.currentDate;
    if (!dateStr) return;
    epsSaveOperatorInputs();
    const audit = {
      operator_1_before: epsState.audit.operator_1_before || 0,
      operator_2_before: epsState.audit.operator_2_before || 0,
      operator_3_before: epsState.audit.operator_3_before || 0,
      operator_1_after: epsState.audit.operator_1_after || 0,
      operator_2_after: epsState.audit.operator_2_after || 0,
      operator_3_after: epsState.audit.operator_3_after || 0,
    };
    try {
      const res = await invoke('save_eps_record', { dateStr: dateStr, records: epsState.records, audit });
      if (res && res.ok) {
        const msg = res.status === "OK"
          ? `已儲存 (${dateStr}) — OK`
          : `已儲存 (${dateStr}) — MISMATCH (差額: $${(res.operators_sum_before - res.calculated_total).toLocaleString()})`;
        showToast(msg, res.status === "OK" ? "success" : "error");
        epsLoadHistory();
      } else {
        showToast("儲存失敗", "error");
      }
    } catch (e) {
      console.error("EPS save error", e);
      showToast("儲存失敗", "error");
    }
  });
}

// Export CSV
if (epsExportBtn) {
  epsExportBtn.addEventListener("click", async () => {
    const dateStr = epsState.currentDate;
    if (!dateStr) return;
    try {
      const res = await invoke('export_eps_csv', { dateStr: dateStr });
      if (res && res.ok) {
        const blob = new Blob([res.content], { type: "text/html;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = res.filename || `EPS_${dateStr}.htm`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        if (res.saved_path) {
          showToast("已儲存至 " + res.saved_path, "success");
        } else {
          showToast("報表已匯出", "success");
        }
      }
    } catch (e) {
      console.error("EPS export error", e);
      showToast("匯出失敗", "error");
    }
  });
}

if (epsSetPathBtn) {
  epsSetPathBtn.addEventListener("click", async () => {
    const pathVal = (epsOutputPath?.value || "").trim();
    try {
      const res = await invoke('set_eps_output_path', { path: pathVal });
      if (res && res.ok) {
        appState.app_config = appState.app_config || {};
        appState.app_config.eps_output_path = res.eps_output_path;
        showToast("匯出路徑已儲存", "success");
      } else {
        showToast(res?.error || "儲存失敗", "error");
      }
    } catch (e) {
      showToast("儲存失敗", "error");
    }
  });
}

// History
async function epsLoadHistory() {
  if (!epsHistoryList) return;
  try {
    const res = await invoke('list_eps_dates_endpoint');
    if (res && res.ok && res.dates) {
      if (res.dates.length === 0) {
        epsHistoryList.innerHTML = `<p class="muted-hint">尚無歷史記錄</p>`;
        return;
      }
      epsHistoryList.innerHTML = res.dates
        .slice()
        .reverse()
        .map((d) => `<div class="list-item eps-history-item" data-date="${d}">${d}</div>`)
        .join("");
      epsHistoryList.querySelectorAll(".eps-history-item").forEach((el) => {
        el.addEventListener("click", () => {
          epsSetDate(el.dataset.date);
        });
      });
    }
  } catch (e) {
    console.error("Failed to load EPS history", e);
  }
}

// Init EPS tab when it becomes active
const epsTabObserver = new MutationObserver(() => {
  const epsTab = document.querySelector('[data-tab-content="eps-audit"]');
  if (epsTab && epsTab.classList.contains("active")) {
    epsInit();
    epsLoadHistory();
  }
});
const contentEl = document.querySelector(".content");
if (contentEl) {
  epsTabObserver.observe(contentEl, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });
}

// Startup is now handled by authInit() -> login -> loadState()
// Only render tasks (local-only, no auth needed)
renderTasks();