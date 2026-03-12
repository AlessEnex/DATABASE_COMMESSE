(() => {
  if (window.__commesseAppLoaded) return;
  window.__commesseAppLoaded = true;

const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";
const AUTH_STORAGE_KEY = "sb-bsceqirconhqmxwipbyl-auth-token";
const FETCH_TIMEOUT_MS = 1500;
const NAV_ENTRY = performance.getEntriesByType("navigation")[0];
const PREFER_REST_ON_RELOAD = NAV_ENTRY && NAV_ENTRY.type === "reload";

const DEFAULT_RESET_REDIRECT_URL = "https://www.pianificazionecommesse.it/reset.html";
const GITHUB_PAGES_RESET_REDIRECT_URL = "https://alessenex.github.io/DATABASE_COMMESSE/reset.html";

function getResetRedirectUrl() {
  try {
    const origin = window.location.origin;
    const host = window.location.hostname || "";
    if (!origin || origin === "null") return DEFAULT_RESET_REDIRECT_URL;
    if (host.endsWith("github.io")) return GITHUB_PAGES_RESET_REDIRECT_URL;
    return new URL("/reset.html", origin).toString();
  } catch {
    return DEFAULT_RESET_REDIRECT_URL;
  }
}

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
  },
});

const state = {
  session: null,
  profile: null,
  reparti: [],
  commesse: [],
  filteredCommesse: [],
  commessaRisorseMap: new Map(),
  commessaRisorseToken: 0,
  commesseFiltersActive: false,
  risorse: [],
  utenti: [],
  permessiByRuolo: new Map(),
  permessi: {},
  selected: null,
  canWrite: false,
  realtime: null,
  reportView: "gantt",
  reportActivitiesMap: new Map(),
  reportActivitiesToken: 0,
  reportActivitiesLoading: false,
  todoOverridesMap: new Map(),
  todoOverridesToken: 0,
  todoOverridesLoading: false,
  todoClientFlowMap: new Map(),
  todoClientFlowLoadedCommesse: new Set(),
  todoClientFlowToken: 0,
  todoClientFlowLoading: false,
  reportRangeStart: null,
  reportRangeDays: null,
  reportFilteredCommesse: null,
  reportFilteredToken: 0,
  reportFilteredLoading: false,
  reportFilteredKey: "",
};

const el = (id) => document.getElementById(id);

const authStatus = el("authStatus");
const roleBadge = el("roleBadge");
const setPasswordBtn = el("setPasswordBtn");
const statusMsg = el("statusMsg");
const revisionStamp = el("revisionStamp");

const emailInput = el("emailInput");
const passwordInput = el("passwordInput");
const loginBtn = el("loginBtn");
const logoutBtn = el("logoutBtn");
const authActions = el("authActions");
const authDot = el("authDot");
const resetPasswordBtn = el("resetPasswordBtn");
const resetPasswordFeedback = el("resetPasswordFeedback");
const toastStack = el("toastStack");
const resetWaitOverlay = el("resetWaitOverlay");
const resetWaitCountdown = el("resetWaitCountdown");
const resetWaitCloseBtn = el("resetWaitCloseBtn");
const topbarMenu = el("topbarMenu");
const topbarMenuToggle = el("topbarMenuToggle");
const permissionsLink = el("permissionsLink");

const DEBUG_AUTH = new URLSearchParams(window.location.search).get("debug") === "1";
let debugPanel = null;
const debugLog = (message) => {
  if (!DEBUG_AUTH) return;
  if (!debugPanel) {
    debugPanel = document.createElement("div");
    debugPanel.className = "debug-panel";
    document.body.appendChild(debugPanel);
  }
  const time = new Date().toLocaleTimeString("it-IT", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  debugPanel.textContent = `[${time}] ${message}\n` + debugPanel.textContent;
};

const RESET_COOLDOWN_MS = 300_000;
const RESET_COOLDOWN_KEY = "commesse_reset_cooldown_until";
const THEME_KEY = "commesse_theme";

let resetCooldownTimer = null;
let resetWaitTimer = null;

const setResetFeedback = (message, tone = "") => {
  if (!resetPasswordFeedback) return;
  resetPasswordFeedback.textContent = message;
  resetPasswordFeedback.classList.remove("hidden", "error");
  if (tone === "error") resetPasswordFeedback.classList.add("error");
  resetPasswordFeedback.classList.remove("hidden");
};

function applyTheme(theme) {
  document.body.classList.toggle("theme-dark", theme === "dark");
  if (themeToggleBtn) {
    themeToggleBtn.textContent = theme === "dark" ? "Dark" : "Light";
    themeToggleBtn.title = theme === "dark" ? "Tema notte" : "Tema giorno";
    themeToggleBtn.setAttribute("aria-label", theme === "dark" ? "Tema notte" : "Tema giorno");
  }
}

function initTheme() {
  const stored = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = stored || (prefersDark ? "dark" : "light");
  applyTheme(theme);
}

function toggleTheme() {
  const isDark = document.body.classList.contains("theme-dark");
  const next = isDark ? "light" : "dark";
  localStorage.setItem(THEME_KEY, next);
  applyTheme(next);
}

function getDetailSnapshot() {
  return JSON.stringify({
    anno: d.anno?.value || "",
    numero: d.numero?.value || "",
    titolo: d.titolo?.value || "",
    cliente: d.cliente?.value || "",
    stato: d.stato?.value || "",
    priorita: d.priorita?.value || "",
    data_ingresso: d.data_ingresso?.value || "",
    data_ordine_telaio: d.data_ordine_telaio?.value || "",
    data_arrivo_kit_cavi: d.data_arrivo_kit_cavi?.value || "",
    data_prelievo: d.data_prelievo_materiali?.value || "",
    data_consegna: d.data_consegna?.value || "",
    note: d.note?.value || "",
    telaio_consegnato: Boolean(d.telaio_ordinato?.dataset.ordered === "true"),
    data_consegna_telaio_effettiva: d.data_consegna_telaio_effettiva?.value || "",
  });
}

function setDetailSnapshot() {
  detailSnapshot = getDetailSnapshot();
  detailDirty = false;
}

function updateDetailDirty() {
  detailDirty = getDetailSnapshot() !== detailSnapshot;
}

function showToast(message, tone = "info", timeout = 2600) {
  if (!toastStack || !message) return;
  const toast = document.createElement("div");
  toast.className = `toast ${tone}`;
  toast.textContent = message;
  toastStack.appendChild(toast);
  requestAnimationFrame(() => {
    toast.classList.add("show");
  });
  window.setTimeout(() => {
    toast.classList.remove("show");
    window.setTimeout(() => toast.remove(), 200);
  }, timeout);
}

function setTelaioOrdinatoButton(button, isOrdered, showFeedback = false, plannedInput = null) {
  if (!button) return;
  button.dataset.ordered = isOrdered ? "true" : "false";
  button.classList.toggle("is-ordered", isOrdered);
  button.setAttribute("aria-pressed", isOrdered ? "true" : "false");
  button.textContent = isOrdered ? "Telaio ordinato" : "Telaio da ordinare";
  if (plannedInput) {
    plannedInput.disabled = isOrdered;
    plannedInput.title = isOrdered ? "Unmark frame ordered to edit the planned date." : "";
  }
  if (showFeedback) {
    setStatus(isOrdered ? "Frame marked as ordered." : "Frame marked as to-order.", "ok");
  }
}

const getResetCooldownUntil = () => {
  const raw = localStorage.getItem(RESET_COOLDOWN_KEY);
  const value = raw ? Number(raw) : 0;
  return Number.isFinite(value) ? value : 0;
};

const setResetCooldownUntil = (timestamp) => {
  localStorage.setItem(RESET_COOLDOWN_KEY, String(timestamp));
};

const clearResetCooldown = () => {
  localStorage.removeItem(RESET_COOLDOWN_KEY);
};

const updateResetCooldownUI = (showMessage = false) => {
  if (!resetPasswordBtn) return;
  const now = Date.now();
  const until = getResetCooldownUntil();
  const remaining = Math.max(0, until - now);

  if (remaining <= 0) {
    if (resetCooldownTimer) {
      clearInterval(resetCooldownTimer);
      resetCooldownTimer = null;
    }
    resetPasswordBtn.disabled = false;
    if (resetPasswordFeedback && resetPasswordFeedback.textContent.includes("Try again in")) {
      resetPasswordFeedback.classList.add("hidden");
    }
    return;
  }

  resetPasswordBtn.disabled = true;
  const seconds = Math.ceil(remaining / 1000);
  if (showMessage) {
    setResetFeedback(`You've already requested a reset. Try again in ${seconds}s.`, "error");
  }
  if (!resetCooldownTimer) {
    resetCooldownTimer = setInterval(() => updateResetCooldownUI(false), 1000);
  }
};

const formatCountdown = (ms) => {
  const totalSeconds = Math.max(0, Math.ceil(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, "0")}`;
};

const hideResetWaitOverlay = () => {
  if (resetWaitTimer) {
    clearInterval(resetWaitTimer);
    resetWaitTimer = null;
  }
  if (resetWaitOverlay) resetWaitOverlay.classList.add("hidden");
  document.body.classList.remove("reset-waiting");
};

const showResetWaitOverlay = (durationMs = RESET_COOLDOWN_MS) => {
  if (!resetWaitOverlay) return;
  const endAt = Date.now() + durationMs;
  document.body.classList.add("reset-waiting");
  resetWaitOverlay.classList.remove("hidden");

  const tick = () => {
    const remaining = endAt - Date.now();
    if (resetWaitCountdown) resetWaitCountdown.textContent = formatCountdown(remaining);
    if (remaining <= 0) hideResetWaitOverlay();
  };

  tick();
  if (resetWaitTimer) clearInterval(resetWaitTimer);
  resetWaitTimer = setInterval(tick, 1000);
};

const commesseList = el("commesseList");
const descriptionFilter = el("descriptionFilter");
const numberFilter = el("numberFilter");
const yearFilter = el("yearFilter");
const openImportBtn = el("openImportBtn");
const openImportDatesBtn = el("openImportDatesBtn");
const openImportProductionBtn = el("openImportProductionBtn");
const commessePanel = el("commessePanel");
const commessePanelBody = el("commessePanelBody");
const commesseToggleBtn = el("commesseToggleBtn");
const commesseTitle = el("commesseTitle");
const matrixPanel = el("matrixPanel");
const matrixTitle = el("matrixTitle");
const calendarPanel = el("calendarPanel");
const calendarTitle = el("calendarTitle");
const reportPanel = el("reportPanel");
const reportTitle = el("reportTitle");
const todoTitle = el("todoTitle");
const openNewCommessaBtn = el("openNewCommessaBtn");

const detailForm = el("detailForm");
const selectedCode = el("selectedCode");
const updateBtn = el("updateBtn");
const applyBtn = el("applyBtn");
const commessaDetailModal = el("commessaDetailModal");
const closeDetailBtn = el("closeDetailBtn");
const detailDeleteBtn = el("detailDeleteBtn");

const newForm = el("newForm");
const newRepartiChecks = el("newRepartiChecks");
const commessaCreateModal = el("commessaCreateModal");
const closeCreateBtn = el("closeCreateBtn");

const repartiList = el("repartiList");
const addRepartoForm = el("addRepartoForm");
const addRepartoSelect = el("addRepartoSelect");
const openResourcesBtn = el("openResourcesBtn");
const resourcesModal = el("resourcesModal");
const closeResourcesBtn = el("closeResourcesBtn");
const openProgressBtn = el("openProgressBtn");
const progressModal = el("progressModal");
const closeProgressBtn = el("closeProgressBtn");
const progressYearCurrent = el("progressYearCurrent");
const progressYearPrev = el("progressYearPrev");
const progressSearch = el("progressSearch");
const progressResults = el("progressResults");
const progressMeta = el("progressMeta");
const progressList = el("progressList");
const progressTimelineDays = el("progressTimelineDays");
const workloadModal = el("workloadModal");
const workloadCloseBtn = el("workloadCloseBtn");
const workloadWeekStart = el("workloadWeekStart");
const workloadWeekEnd = el("workloadWeekEnd");
const workloadRunBtn = el("workloadRunBtn");
const workloadSummary = el("workloadSummary");
const workloadList = el("workloadList");
const matrixQuickMenu = el("matrixQuickMenu");
const matrixQuickToggleBtn = el("matrixQuickToggleBtn");
const matrixQuickInfo = el("matrixQuickInfo");
const resourcesList = el("resourcesList");
const usersList = el("usersList");
const resourceForm = el("resourceForm");
const resourceName = el("resourceName");
const resourceReparto = el("resourceReparto");
const resourceActive = el("resourceActive");

const importModal = el("importModal");
const importTextarea = el("importTextarea");
const importPreview = el("importPreview");
const importCommitBtn = el("importCommitBtn");
const closeImportBtn = el("closeImportBtn");
const importDatesModal = el("importDatesModal");
const importDatesTextarea = el("importDatesTextarea");
const importDatesPreview = el("importDatesPreview");
const importDatesCommitBtn = el("importDatesCommitBtn");
const closeImportDatesBtn = el("closeImportDatesBtn");
const importProductionModal = el("importProductionModal");
const importProductionTextarea = el("importProductionTextarea");
const importProductionPreview = el("importProductionPreview");
const importProductionCommitBtn = el("importProductionCommitBtn");
const closeImportProductionBtn = el("closeImportProductionBtn");
const backupModal = el("backupModal");
const openBackupBtn = el("openBackupBtn");
const closeBackupBtn = el("closeBackupBtn");
const backupExportBtn = el("backupExportBtn");
const backupExportStatus = el("backupExportStatus");
const backupImportInput = el("backupImportInput");
const backupImportConfirm = el("backupImportConfirm");
const backupImportSummary = el("backupImportSummary");
const backupImportBtn = el("backupImportBtn");
const themeToggleBtn = el("themeToggleBtn");
const milestonePicker = el("milestonePicker");
const milestonePrevBtn = el("milestonePrevBtn");
const milestoneNextBtn = el("milestoneNextBtn");
const milestoneLabel = el("milestoneLabel");
const milestoneDays = el("milestoneDays");
const milestoneClearBtn = el("milestoneClearBtn");
const milestoneOrderBtn = el("milestoneOrderBtn");
const milestoneTransportVanBtn = el("milestoneTransportVanBtn");
const milestoneTransportShipBtn = el("milestoneTransportShipBtn");
const milestoneTitle = el("milestoneTitle");
const milestoneDragLabel = el("milestoneDragLabel");
const reportDeptMenu = el("reportDeptMenu");
const todoSection = el("todoSection");
const todoGrid = el("todoGrid");
const todoHeaderTableHost = el("todoHeaderTableHost");
const todoStatusMenu = el("todoStatusMenu");
const todoGlobalStatusMenu = el("todoGlobalStatusMenu");
const todoFocusOverlay = el("todoFocusOverlay");
const todoDescFilter = el("todoDescFilter");
const todoNumberFilter = el("todoNumberFilter");
const todoYearFilter = el("todoYearFilter");
const todoPriorityField = el("todoPriorityField");
const todoPriorityBtn = el("todoPriorityBtn");
const todoSortBtn = el("todoSortBtn");
const todoCompleteBtn = el("todoCompleteBtn");
const todoInfoText = el("todoInfoText");
const todoFilterControls = [
  todoDescFilter,
  todoNumberFilter,
  todoYearFilter,
  todoPriorityField,
  todoPriorityBtn,
  todoSortBtn,
  todoCompleteBtn,
].filter(Boolean);
const sectionToolbar = el("sectionToolbar");
const sectionLinks = Array.from(document.querySelectorAll(".section-link"));
const MILESTONE_MIN_YEAR = 2024;
const MILESTONE_MAX_YEAR = 2030;
const calendarDate = el("calendarDate");
const calendarView = el("calendarView");
const calPrevBtn = el("calPrevBtn");
const calNextBtn = el("calNextBtn");
const calendarGrid = el("calendarGrid");
const matrixDate = el("matrixDate");
const matrixPrevBtn = el("matrixPrevBtn");
const matrixNextBtn = el("matrixNextBtn");
const matrixTodayBtn = el("matrixTodayBtn");
const matrixZoomInBtn = el("matrixZoomInBtn");
const matrixZoomOutBtn = el("matrixZoomOutBtn");
const matrixToggleListBtn = el("matrixToggleListBtn");
const matrixShiftToggle = el("matrixShiftToggle");
const matrixColorByActivityBtn = el("matrixColorByActivityBtn");
const matrixColorByCommessaBtn = el("matrixColorByCommessaBtn");
const matrixTrash = el("matrixTrash");
const matrixGrid = el("matrixGrid");
const matrixCommessa = el("matrixCommessa");
const matrixAttivita = el("matrixAttivita");
const matrixCommessaPicker = el("matrixCommessaPicker");
const matrixCommessaPickerToggleBtn = el("matrixCommessaPickerToggleBtn");
const matrixCommessaPickerList = el("matrixCommessaPickerList");
const matrixAttivitaPicker = el("matrixAttivitaPicker");
const matrixAttivitaToggleBtn = el("matrixAttivitaToggleBtn");
const matrixAttivitaList = el("matrixAttivitaList");
const matrixWrap = el("matrixWrap");
const matrixCommessePanel = el("matrixCommessePanel");
const matrixCommesseList = el("matrixCommesseList");
const matrixCommessaSearch = el("matrixCommessaSearch");
const matrixAssenzaItem = el("matrixAssenzaItem");
const matrixAltroItem = el("matrixAltroItem");
const matrixPedItem = el("matrixPedItem");
const matrixSegnaturaItem = el("matrixSegnaturaItem");
const matrixOpOrdiniItem = el("matrixOpOrdiniItem");
const commessaFocusOverlay = el("commessaFocusOverlay");
const commessaQuickMenu = el("commessaQuickMenu");
const commessaQuickSeeMoreBtn = el("commessaQuickSeeMoreBtn");
const matrixReportGrid = el("matrixReportGrid");
const matrixReportRange = el("matrixReportRange");
const matrixReportSortBtn = el("matrixReportSortBtn");
const reportDescFilter = el("reportDescFilter");
const reportNumberFilter = el("reportNumberFilter");
const reportYearFilter = el("reportYearFilter");
const reportHidePast = el("reportHidePast");
const reportOrderDue = el("reportOrderDue");
const reportWorkloadBtn = el("reportWorkloadBtn");
const reportViewToggleBtn = el("reportViewToggleBtn");
const REPORT_MAX_ITEMS = 200;
const REPORT_DEFAULT_DAYS = 84;
const REPORT_MIN_DAYS = 14;
const REPORT_MAX_DAYS = 365;
let reportPanZoom = null;

const assignModal = el("assignModal");
const assignActivities = el("assignActivities");
const assignConfirmBtn = el("assignConfirmBtn");
const closeAssignBtn = el("closeAssignBtn");
const assignAssenzaWrap = el("assignAssenzaWrap");
const assignAssenzaOre = el("assignAssenzaOre");
const assignAssenzaFullBtn = el("assignAssenzaFullBtn");
const assignAssenzaHalfBtn = el("assignAssenzaHalfBtn");
const assignAltroNoteWrap = el("assignAltroNoteWrap");
const assignAltroNote = el("assignAltroNote");
const matrixFilterModal = el("matrixFilterModal");
const matrixFilterTitle = el("matrixFilterTitle");
const matrixFilterCloseBtn = el("matrixFilterCloseBtn");
const matrixFilterYear = el("matrixFilterYear");
const matrixFilterSearch = el("matrixFilterSearch");
const matrixFilterList = el("matrixFilterList");
const matrixFilterClearBtn = el("matrixFilterClearBtn");
const matrixFilterResetBtn = el("matrixFilterResetBtn");
const matrixFilterApplyBtn = el("matrixFilterApplyBtn");
const activityModal = el("activityModal");
const activityCloseBtn = el("activityCloseBtn");
const activityGanttBtn = el("activityGanttBtn");
const activityMeta = el("activityMeta");
const activityTitleList = el("activityTitleList");
const activityAssenzaWrap = el("activityAssenzaWrap");
const activityAssenzaOre = el("activityAssenzaOre");
const activityAssenzaFullBtn = el("activityAssenzaFullBtn");
const activityAssenzaHalfBtn = el("activityAssenzaHalfBtn");
const activityAltroNoteWrap = el("activityAltroNoteWrap");
const activityAltroNote = el("activityAltroNote");
const activityDurationInput = el("activityDurationInput");
const activityDeleteBtn = el("activityDeleteBtn");
const activitySaveBtn = el("activitySaveBtn");
const confirmModal = el("confirmModal");
const confirmMessage = el("confirmMessage");
const confirmCancelBtn = el("confirmCancelBtn");
const confirmOkBtn = el("confirmOkBtn");
const deleteCommessaModal = el("deleteCommessaModal");
const deleteCommessaMessage = el("deleteCommessaMessage");
const deleteCommessaInput = el("deleteCommessaInput");
const deleteCommessaCancelBtn = el("deleteCommessaCancelBtn");
const deleteCommessaConfirmBtn = el("deleteCommessaConfirmBtn");
const confirmDefaults = {
  okLabel: confirmOkBtn ? confirmOkBtn.textContent : "",
  cancelLabel: confirmCancelBtn ? confirmCancelBtn.textContent : "",
};

const d = {
  anno: el("d_anno"),
  numero: el("d_numero"),
  titolo: el("d_titolo"),
  cliente: el("d_cliente"),
  stato: el("d_stato"),
  priorita: el("d_priorita"),
  data_ingresso: el("d_data_ingresso"),
  data_ordine_telaio: el("d_data_ordine_telaio"),
  telaio_ordinato: el("d_telaio_ordinato"),
  data_consegna_telaio_effettiva: el("d_data_consegna_telaio_effettiva"),
  data_arrivo_kit_cavi: el("d_data_arrivo_kit_cavi"),
  data_prelievo_materiali: el("d_data_prelievo_materiali"),
  data_consegna: el("d_data_consegna"),
  note: el("d_note"),
};

const n = {
  anno: el("n_anno"),
  numero: el("n_numero"),
  titolo: el("n_titolo"),
  cliente: el("n_cliente"),
  stato: el("n_stato"),
  priorita: el("n_priorita"),
  data_ingresso: el("n_data_ingresso"),
  data_ordine_telaio: el("n_data_ordine_telaio"),
  telaio_consegnato: el("n_telaio_consegnato"),
  data_consegna_telaio_effettiva: el("n_data_consegna_telaio_effettiva"),
  data_consegna: el("n_data_consegna"),
  note: el("n_note"),
};

const newNumeroWarning = el("n_numeroWarning");
const detailNumeroWarning = el("d_numeroWarning");
const newCodicePreview = el("n_codicePreview");

const calendarState = {
  view: "week",
  date: new Date(),
  attivita: [],
};

const milestonePickerState = {
  commessaId: null,
  field: null,
  month: new Date(),
  selected: null,
  saving: false,
  compareTarget: null,
  tone: "",
};

let milestoneDrag = null;
let milestoneSuppressClickUntil = 0;

const matrixState = {
  date: new Date(),
  attivita: [],
  risorse: [],
  draggingId: null,
  view: "two",
  pendingDrop: null,
  colorMode: "none",
  selectedAttivita: new Set(),
  selectedCommesse: new Set(),
  filterMode: null,
  filterDraft: new Set(),
  editingAttivita: null,
  editingTitles: new Set(),
  resizing: null,
  suppressClickUntil: 0,
  autoShift: false,
  reportSort: "asc",
  collapsedReparti: new Set(),
};

let commessaHighlightId = null;
let commessaHighlightTimer = null;
let commessaHighlightAt = 0;
let commessaLongPressSuppress = false;
let commessaItemLongPressSuppress = false;
let commessaQuickTarget = null;
let commessaQuickTimer = null;
let commessaQuickFocused = null;
let quickMenuAttivita = null;
let confirmResolver = null;
let confirmAction = null;
let todoMenuTarget = null;
let todoLongPressTimer = null;
let todoLongPressStart = null;
let todoLongPressSuppress = false;
let todoFocusedRowKey = null;
let todoMenuAnchorEl = null;
let todoStatusMenuManualPosition = false;
let todoStatusMenuDrag = null;
let todoStatusMenuDragSuppressClickUntil = 0;
let todoGlobalMenuTarget = null;
let todoGlobalMenuAnchorEl = null;
let todoGlobalMenuSuppress = false;
let todoLastRendered = [];
let todoScrollSyncLock = false;
let todoFilterHeightLock = 0;
let todoFilterUnlockTimer = null;
const todoStatusDatePickerState = {
  action: null,
  sentMonth: startOfDay(new Date()),
  sentSelected: null,
  dueMonth: startOfDay(new Date()),
  dueSelected: null,
  confirmMonth: startOfDay(new Date()),
  confirmSelected: null,
};
const TODO_VISUAL_REFRESH_MS = 30_000;
let todoVisualRefreshTimer = null;
let deleteCommessaId = null;
let detailSnapshot = "";
let detailDirty = false;
let backupPayload = null;
let matrixPan = null;

const BACKUP_TABLES = [
  { name: "reparti", conflict: "id" },
  { name: "risorse", conflict: "id" },
  { name: "utenti", conflict: "id" },
  { name: "commesse", conflict: "id" },
  { name: "commessa_schede", conflict: "commessa_id" },
  { name: "commesse_reparti", conflict: "commessa_id,reparto_id" },
  { name: "assegnazioni", conflict: "commessa_id,utente_id" },
  { name: "attivita", conflict: "id" },
  { name: "commessa_attivita_override", conflict: "commessa_id,titolo" },
  { name: "commessa_attivita_cliente", conflict: "commessa_id,titolo,reparto" },
  { name: "whitelist_email", conflict: "email" },
  { name: "log_commessa", conflict: "id" },
];

function updateCommessaInState(commessaId, patch) {
  if (!commessaId || !patch) return null;
  const index = state.commesse.findIndex((c) => c.id === commessaId);
  if (index < 0) return null;
  const target = state.commesse[index];
  Object.assign(target, patch);
  if (state.selected && state.selected.id === commessaId) {
    Object.assign(state.selected, patch);
  }
  if (state.filteredCommesse && state.filteredCommesse.length) {
    const filtered = state.filteredCommesse.find((c) => c.id === commessaId);
    if (filtered) Object.assign(filtered, patch);
  }
  if (state.reportFilteredCommesse && state.reportFilteredCommesse.length) {
    const filtered = state.reportFilteredCommesse.find((c) => c.id === commessaId);
    if (filtered) Object.assign(filtered, patch);
  }
  return target;
}

function isPlannerRole() {
  if (!state.profile) return false;
  return String(state.profile.ruolo || "").trim().toLowerCase() === "planner";
}

async function updateCommessaPlannerDates(commessaId, patch) {
  const commessa = state.commesse.find((c) => c.id === commessaId);
  const ordine = patch.data_ordine_telaio ?? commessa?.data_ordine_telaio ?? null;
  const consegna = patch.data_consegna_macchina ?? commessa?.data_consegna_macchina ?? null;
  const kit = patch.data_arrivo_kit_cavi ?? commessa?.data_arrivo_kit_cavi ?? null;
  const prelievo = patch.data_prelievo ?? commessa?.data_prelievo ?? null;
  const { error } = await supabase.rpc("update_commessa_planner_dates", {
    p_commessa_id: commessaId,
    p_data_ordine_telaio: ordine || null,
    p_data_consegna_macchina: consegna || null,
    p_data_arrivo_kit_cavi: kit || null,
    p_data_prelievo: prelievo || null,
  });
  if (error) return error;
  updateCommessaInState(commessaId, {
    data_ordine_telaio: ordine || null,
    data_consegna_macchina: consegna || null,
    data_arrivo_kit_cavi: kit || null,
    data_prelievo: prelievo || null,
  });
  return null;
}

function applyCommessaHighlight(commessaId) {
  if (!matrixGrid) return;
  if (!commessaId) return;
  commessaHighlightId = commessaId;
  commessaHighlightAt = Date.now();
  commessaLongPressSuppress = true;
  const bars = matrixGrid.querySelectorAll(".matrix-activity-bar");
  bars.forEach((bar) => {
    const same = bar.dataset && bar.dataset.commessaId === String(commessaId);
    bar.classList.toggle("commessa-highlight", same);
    bar.classList.toggle("commessa-dim", !same);
  });
}

function updateMatrixCommessaPickerLabel() {
  if (!matrixCommessaPickerToggleBtn) return;
  const count = matrixState.selectedCommesse.size;
  if (count === 0) {
    matrixCommessaPickerToggleBtn.textContent = "Filtra commesse";
    matrixCommessaPickerToggleBtn.classList.remove("active");
    return;
  }
  if (count === 1) {
    const only = Array.from(matrixState.selectedCommesse)[0];
    const commessa = state.commesse.find((c) => String(c.id) === String(only));
    matrixCommessaPickerToggleBtn.textContent = commessa?.codice || "Commessa selezionata";
  } else {
    matrixCommessaPickerToggleBtn.textContent = `${count} commesse`;
  }
  matrixCommessaPickerToggleBtn.classList.add("active");
}

function setMatrixSelectedCommesse(ids = []) {
  const normalized = ids.map((id) => String(id)).filter(Boolean);
  matrixState.selectedCommesse = new Set(normalized);
  if (matrixCommessa && normalized.length === 1) {
    matrixCommessa.value = normalized[0];
  } else if (matrixCommessa) {
    matrixCommessa.value = "";
  }
  updateMatrixCommessaPickerLabel();
}

function clearCommessaHighlight() {
  if (!matrixGrid) return;
  if (!commessaHighlightId) return;
  const bars = matrixGrid.querySelectorAll(".matrix-activity-bar");
  bars.forEach((bar) => {
    bar.classList.remove("commessa-highlight", "commessa-dim");
  });
  commessaHighlightId = null;
}

function cancelCommessaHighlightTimer() {
  if (commessaHighlightTimer) {
    clearTimeout(commessaHighlightTimer);
    commessaHighlightTimer = null;
  }
}

function openMatrixQuickMenu(attivita, x, y) {
  if (!matrixQuickMenu || !matrixQuickToggleBtn) return;
  if (!state.canWrite) return;
  quickMenuAttivita = attivita;
  const isDone = attivita && attivita.stato === "completata";
  matrixQuickToggleBtn.textContent = isDone ? "Not done" : "Done";
  if (matrixQuickInfo) {
    const commessa = attivita?.commessa_id ? state.commesse.find((c) => c.id === attivita.commessa_id) : null;
    const titolo = attivita?.titolo || "Attivita";
    const descrizione = attivita?.descrizione ? ` \u2014 ${attivita.descrizione}` : "";
    const commessaTitolo = commessa?.titolo ? `${commessa.titolo}` : "";
    const lines = [`${titolo}${descrizione}`, commessaTitolo].filter(Boolean);
    matrixQuickInfo.textContent = lines.join("\n");
  }
  matrixQuickMenu.classList.remove("hidden");
  const padding = 10;
  let left = x + 8;
  let top = y + 8;
  matrixQuickMenu.style.left = `${left}px`;
  matrixQuickMenu.style.top = `${top}px`;
  requestAnimationFrame(() => {
    const rect = matrixQuickMenu.getBoundingClientRect();
    if (rect.right > window.innerWidth - padding) {
      left = Math.max(padding, window.innerWidth - rect.width - padding);
    }
    if (rect.bottom > window.innerHeight - padding) {
      top = Math.max(padding, window.innerHeight - rect.height - padding);
    }
    matrixQuickMenu.style.left = `${left}px`;
    matrixQuickMenu.style.top = `${top}px`;
  });
  commessaLongPressSuppress = true;
}

function closeMatrixQuickMenu() {
  if (!matrixQuickMenu) return;
  matrixQuickMenu.classList.add("hidden");
  quickMenuAttivita = null;
}

function openCommessaQuickMenu(commessa, x, y, itemEl) {
  if (!commessaQuickMenu || !commessaQuickSeeMoreBtn || !commessaFocusOverlay) return;
  commessaQuickTarget = commessa;
  commessaQuickFocused = itemEl;
  commessaQuickSeeMoreBtn.onclick = () => {
    if (!commessaQuickTarget) return;
    window.open(`commessa.html?id=${commessaQuickTarget.id}`, "_blank");
    closeCommessaQuickMenu();
  };
  if (commessaQuickFocused) commessaQuickFocused.classList.add("commessa-item-focused");
  commessaFocusOverlay.classList.remove("hidden");
  commessaQuickMenu.classList.remove("hidden");
  const padding = 10;
  let left = x + 8;
  let top = y + 8;
  commessaQuickMenu.style.left = `${left}px`;
  commessaQuickMenu.style.top = `${top}px`;
  requestAnimationFrame(() => {
    const rect = commessaQuickMenu.getBoundingClientRect();
    if (rect.right > window.innerWidth - padding) {
      left = Math.max(padding, window.innerWidth - rect.width - padding);
    }
    if (rect.bottom > window.innerHeight - padding) {
      top = Math.max(padding, window.innerHeight - rect.height - padding);
    }
    commessaQuickMenu.style.left = `${left}px`;
    commessaQuickMenu.style.top = `${top}px`;
  });
  commessaItemLongPressSuppress = true;
}

function closeCommessaQuickMenu() {
  if (!commessaQuickMenu || !commessaFocusOverlay) return;
  commessaQuickMenu.classList.add("hidden");
  commessaFocusOverlay.classList.add("hidden");
  if (commessaQuickFocused) commessaQuickFocused.classList.remove("commessa-item-focused");
  commessaQuickFocused = null;
  commessaQuickTarget = null;
}

async function toggleQuickMenuDone() {
  if (!quickMenuAttivita) return;
  if (!state.canWrite) return;
  const nextState = quickMenuAttivita.stato === "completata" ? "pianificata" : "completata";
  const { error } = await supabase
    .from("attivita")
    .update({ stato: nextState })
    .eq("id", quickMenuAttivita.id);
  if (error) {
    setStatus(`Update error: ${error.message}`, "error");
    return;
  }
  const isTelaio = (quickMenuAttivita.titolo || "").trim().toLowerCase() === "telaio";
  let commessaUpdated = false;
  if (nextState === "completata" && isTelaio && quickMenuAttivita.commessa_id) {
    const endDateRaw = quickMenuAttivita.data_fine ? new Date(quickMenuAttivita.data_fine) : null;
    const endDate = endDateRaw && !Number.isNaN(endDateRaw.getTime()) ? startOfDay(endDateRaw) : null;
    const existingCommessa = state.commesse.find((c) => c.id === quickMenuAttivita.commessa_id);
    const payload = {
      telaio_consegnato: true,
    };
    if (!existingCommessa?.data_consegna_telaio_effettiva && endDate) {
      payload.data_consegna_telaio_effettiva = formatDateInput(endDate);
    }
    const { error: commessaError } = await supabase
      .from("commesse")
      .update(payload)
      .eq("id", quickMenuAttivita.commessa_id);
    if (commessaError) {
      setStatus(`Activity completed, but failed to update delivery flag: ${commessaError.message}`, "error");
    } else {
      updateCommessaInState(quickMenuAttivita.commessa_id, payload);
      if (state.selected && state.selected.id === quickMenuAttivita.commessa_id) {
        if (d.telaio_ordinato) {
          setTelaioOrdinatoButton(d.telaio_ordinato, true, false, d.data_consegna_telaio_effettiva || null);
        }
        if (d.data_consegna_telaio_effettiva) {
          d.data_consegna_telaio_effettiva.value = payload.data_consegna_telaio_effettiva || "";
        }
      }
      applyFilters();
      renderMatrixReport();
      commessaUpdated = true;
    }
  }
  if (!commessaUpdated) {
    setStatus(nextState === "completata" ? "Activity completed." : "Activity reopened.", "ok");
  }
  closeMatrixQuickMenu();
  await loadMatrixAttivita();
  if (commessaUpdated) {
    setStatus("Activity completed. Delivery flagged.", "ok");
  }
}

function setMatrixViewLabel() {
  return;
}

function hashString(value) {
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash * 31 + value.charCodeAt(i)) | 0;
  }
  return Math.abs(hash);
}

function colorForKey(key) {
  const seed = (hashString(key) * 2654435761) >>> 0;
  let hue = seed % 360;
  if (hue >= 250 && hue <= 290) {
    hue = (hue + 45) % 360;
  }
  const sat = 50 + ((seed >>> 8) % 40);
  const light = 78 + ((seed >>> 16) % 14);
  const borderSat = Math.max(35, sat - 20);
  const borderLight = Math.max(40, light - 25);
  return {
    bg: `hsl(${hue} ${sat}% ${light}%)`,
    border: `hsl(${hue} ${borderSat}% ${borderLight}%)`,
  };
}

function setMatrixColorMode(mode) {
  matrixState.colorMode = mode;
  if (matrixColorByActivityBtn) matrixColorByActivityBtn.classList.toggle("active", mode === "activity");
  if (matrixColorByCommessaBtn) matrixColorByCommessaBtn.classList.toggle("active", mode === "commessa");
  renderMatrix();
}

function setMatrixAutoShift(enabled) {
  matrixState.autoShift = enabled;
  if (matrixShiftToggle) {
    matrixShiftToggle.checked = enabled;
    matrixShiftToggle.setAttribute("aria-checked", enabled ? "true" : "false");
  }
}

function getMatrixAttivitaOptions() {
  if (!matrixAttivita) return [];
  return Array.from(matrixAttivita.options)
    .map((o) => o.value)
    .filter((v) => v);
}

function getSelectedMatrixAttivita() {
  return Array.from(matrixState.selectedAttivita);
}

function getSelectedMatrixCommesse() {
  return Array.from(matrixState.selectedCommesse);
}

function getCommessaYear(commessa) {
  const code = (commessa.codice || "").trim();
  const match = code.match(/^(\d{4})/);
  return match ? match[1] : "";
}

function openMatrixFilter(mode) {
  if (!matrixFilterModal || !matrixFilterSearch || !matrixFilterTitle || !matrixFilterList) return;
  matrixState.filterMode = mode;
  matrixState.filterDraft = new Set(
    mode === "commessa" ? Array.from(matrixState.selectedCommesse) : Array.from(matrixState.selectedAttivita)
  );
  matrixFilterTitle.textContent = mode === "commessa" ? "Filtra commesse" : "Filtra attivita";
  matrixFilterSearch.value = "";
  if (matrixFilterYear) {
    if (mode === "commessa") {
      const years = Array.from(
        new Set(state.commesse.map((c) => getCommessaYear(c)).filter((y) => y))
      ).sort((a, b) => b.localeCompare(a));
      matrixFilterYear.innerHTML = `<option value="">Tutti gli anni</option>`;
      years.forEach((y) => {
        const opt = document.createElement("option");
        opt.value = y;
        opt.textContent = y;
        matrixFilterYear.appendChild(opt);
      });
      matrixFilterYear.classList.remove("hidden");
    } else {
      matrixFilterYear.classList.add("hidden");
      matrixFilterYear.value = "";
    }
  }
  renderMatrixFilterList();
  matrixFilterModal.classList.remove("hidden");
  matrixFilterSearch.focus();
}

function closeMatrixFilter() {
  if (!matrixFilterModal) return;
  matrixFilterModal.classList.add("hidden");
}

function renderMatrixFilterList() {
  if (!matrixFilterList || !matrixFilterSearch) return;
  const q = matrixFilterSearch.value.trim().toLowerCase();
  matrixFilterList.innerHTML = "";
  if (matrixState.filterMode === "commessa") {
    const yearFilter = matrixFilterYear ? matrixFilterYear.value : "";
    if (matrixFilterTitle) {
      const count = matrixState.selectedCommesse.size;
      if (count === 0) {
        matrixFilterTitle.textContent = "Filtra commesse";
      } else if (count === 1) {
        const only = Array.from(matrixState.selectedCommesse)[0];
        const commessa = state.commesse.find((c) => String(c.id) === String(only));
        matrixFilterTitle.textContent = commessa?.codice || "Commessa selezionata";
      } else {
        matrixFilterTitle.textContent = `${count} commesse`;
      }
    }
    const options = state.commesse.filter((c) => {
      if (!q) return true;
      const hay = `${c.codice} ${c.titolo || ""}`.toLowerCase();
      return hay.includes(q);
    });
    const filtered = (yearFilter
      ? options.filter((c) => getCommessaYear(c) === yearFilter)
      : options
    ).slice().sort(compareCommesse);
    if (!filtered.length) {
      const empty = document.createElement("div");
      empty.className = "matrix-empty";
      empty.textContent = "Nessuna commessa.";
      matrixFilterList.appendChild(empty);
      return;
    }
    filtered.forEach((c) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "filter-pill";
      btn.textContent = `${c.codice}${c.titolo ? " - " + c.titolo : ""}`;
      if (matrixState.filterDraft.has(c.id)) btn.classList.add("active");
      btn.addEventListener("click", () => {
        if (matrixState.filterDraft.has(c.id)) {
          matrixState.filterDraft.delete(c.id);
          btn.classList.remove("active");
        } else {
          matrixState.filterDraft.add(c.id);
          btn.classList.add("active");
        }
      });
      matrixFilterList.appendChild(btn);
    });
    return;
  }

  const options = getMatrixAttivitaOptions().filter((name) => name.toLowerCase().includes(q));
  if (!options.length) {
    const empty = document.createElement("div");
    empty.className = "matrix-empty";
    empty.textContent = "Nessuna attivita.";
    matrixFilterList.appendChild(empty);
    return;
  }
  options.forEach((name) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "filter-pill";
    btn.textContent = name;
    if (matrixState.filterDraft.has(name)) btn.classList.add("active");
    btn.addEventListener("click", () => {
      if (matrixState.filterDraft.has(name)) {
        matrixState.filterDraft.delete(name);
        btn.classList.remove("active");
      } else {
        matrixState.filterDraft.add(name);
        btn.classList.add("active");
      }
    });
    matrixFilterList.appendChild(btn);
  });
}

function applyMatrixFilter() {
  if (matrixState.filterMode === "commessa") {
    setMatrixSelectedCommesse(Array.from(matrixState.filterDraft));
  } else if (matrixState.filterMode === "attivita") {
    matrixState.selectedAttivita = new Set(matrixState.filterDraft);
  }
  renderMatrix();
  closeMatrixFilter();
}

function openActivityModal(attivita) {
  if (!activityModal || !activityDurationInput || !activityMeta) return;
  matrixState.editingAttivita = attivita;
  matrixState.editingTitles = new Set([attivita.titolo]);
  if (activityAltroNote && activityAltroNoteWrap) {
    activityAltroNoteWrap.classList.remove("hidden");
    activityAltroNote.value = attivita.descrizione || "";
  }
  if (activityAssenzaWrap && activityAssenzaOre) {
    const isAssente = isAssenteTitle(attivita.titolo);
    activityAssenzaWrap.classList.toggle("hidden", !isAssente);
    activityAssenzaOre.value = isAssente && attivita.ore_assenza != null ? String(attivita.ore_assenza) : "";
  }
  const commessaLabel = getCommessaLabel(attivita.commessa_id);
  const start = new Date(attivita.data_inizio);
  const end = new Date(attivita.data_fine);
  const durationDays = businessDaysBetweenInclusive(start, end);
  activityDurationInput.value = String(durationDays);
  if (activityTitleList) {
    activityTitleList.innerHTML = "";
    const options = getMatrixAttivitaOptions();
    const allowed = options.filter((name) => isActivityAllowedForRisorsa(name, attivita.risorsa_id));
    const list = options.includes(attivita.titolo)
      ? allowed.includes(attivita.titolo)
        ? allowed
        : [...allowed, attivita.titolo]
      : allowed;
    list.forEach((name) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "filter-pill";
      if (isLavenderActivity(name)) btn.classList.add("lavender-pill");
      btn.textContent = name;
      if (matrixState.editingTitles.has(name)) btn.classList.add("active");
      btn.addEventListener("click", () => {
        if (matrixState.editingTitles.has(name)) {
          matrixState.editingTitles.delete(name);
          btn.classList.remove("active");
        } else {
          matrixState.editingTitles.add(name);
          btn.classList.add("active");
        }
        if (activityAssenzaWrap && activityAssenzaOre) {
          const hasAssente = Array.from(matrixState.editingTitles).some((t) => isAssenteTitle(t));
          activityAssenzaWrap.classList.toggle("hidden", !hasAssente);
          if (!hasAssente) activityAssenzaOre.value = "";
        }
      });
      activityTitleList.appendChild(btn);
    });
  }
  activityMeta.textContent = commessaLabel
    ? `${attivita.titolo} \u2022 ${commessaLabel} \u2022 ${formatDateLocal(start)}`
    : `${attivita.titolo} \u2022 ${formatDateLocal(start)}`;
  activityModal.classList.remove("hidden");
  activityDurationInput.focus();
}

function closeActivityModal() {
  if (!activityModal) return;
  activityModal.classList.add("hidden");
  matrixState.editingAttivita = null;
}

function jumpToReportActivity(attivita) {
  if (!attivita) return;
  if (!attivita.commessa_id) {
    setStatus("This activity is not linked to a commessa.", "error");
    return;
  }
  const date = attivita.data_inizio ? new Date(attivita.data_inizio) : null;
  if (!date || Number.isNaN(date.getTime())) {
    setStatus("Activity date not available.", "error");
    return;
  }
  if (reportPanel && reportPanel.classList.contains("collapsed")) {
    reportPanel.classList.remove("collapsed");
  }
  if (state.reportView !== "gantt") {
    state.reportView = "gantt";
    if (reportViewToggleBtn) {
      reportViewToggleBtn.textContent = "Vista: Gantt";
    }
  }
  const focusDay = normalizeBusinessDay(startOfDay(date), 1);
  const rangeDays = state.reportRangeDays || REPORT_DEFAULT_DAYS;
  state.reportRangeDays = rangeDays;
  state.reportRangeStart = addBusinessDays(focusDay, -Math.floor(rangeDays / 2));
  renderMatrixReport();
  requestAnimationFrame(() => {
    reportPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
    const track = matrixReportGrid?.querySelector(
      `.report-gantt-track[data-commessa-id="${String(attivita.commessa_id)}"]`
    );
    if (!track) {
      setStatus("Commessa not visible with current filters.", "error");
      return;
    }
    track.classList.add("report-gantt-jump-glow");
    track.scrollIntoView({ behavior: "smooth", block: "center" });
    window.setTimeout(() => track.classList.remove("report-gantt-jump-glow"), 2600);
  });
}

async function saveActivityDuration() {
  const attivita = matrixState.editingAttivita;
  if (!attivita) return;
  const days = Math.max(1, Number(activityDurationInput.value || 1));
  const titles = Array.from(matrixState.editingTitles);
  if (!titles.length) {
    setStatus("Select at least one activity.", "error");
    return;
  }
  const hasAssente = titles.some((t) => isAssenteTitle(t));
  if (hasAssente && titles.length > 1) {
    setStatus("ASSENTE must be assigned alone.", "error");
    return;
  }
  const notAllowed = titles.filter((t) => !isActivityAllowedForRisorsa(t, attivita.risorsa_id));
  if (notAllowed.length) {
    setStatus("Some activities are not available for this department.", "error");
    return;
  }
  const note = activityAltroNote ? activityAltroNote.value.trim() : "";
  const assenzaOreRaw = activityAssenzaOre ? activityAssenzaOre.value : "";
  const assenzaOre = Number(assenzaOreRaw);
  if (hasAssente && (!assenzaOreRaw || Number.isNaN(assenzaOre) || assenzaOre <= 0)) {
    setStatus("Enter absence hours.", "error");
    return;
  }
  const start = new Date(attivita.data_inizio);
  const end = new Date(attivita.data_fine);
  const newEndDay = addBusinessDays(startOfDay(start), days - 1);
  const newEnd = new Date(newEndDay);
  newEnd.setHours(end.getHours(), end.getMinutes(), end.getSeconds(), end.getMilliseconds());
  const oldEndDay = startOfDay(end);
  const newEndDayValue = startOfDay(newEnd);
  const deltaDays = businessDayDiff(oldEndDay, newEndDayValue);
  const firstTitle = titles[0];
  const keyForChain = isPhaseKey(firstTitle)
    ? {
        id: attivita.id,
        commessa_id: isAssenteTitle(firstTitle) ? null : attivita.commessa_id,
        titolo: firstTitle,
        data_fine: attivita.data_fine,
      }
    : null;
  let dependentActivities = null;
  let dependentToShift = null;
  if (matrixState.autoShift && keyForChain && deltaDays !== 0) {
    dependentActivities = await getDependentActivitiesForKey(keyForChain);
    if (dependentActivities == null) return;
    dependentToShift = getDependentActivitiesToShift(keyForChain, deltaDays, dependentActivities);
  }
  const excludeIds =
    keyForChain && dependentToShift && dependentToShift.length
      ? [attivita.id, ...dependentToShift.map((a) => a.id)]
      : [attivita.id];
  if (matrixState.autoShift && deltaDays > 0) {
    const extStart = addDays(oldEndDay, 1);
    const overlap = await hasOverlapRange(attivita.risorsa_id, extStart, newEndDayValue, attivita.id);
    if (overlap == null) return;
    if (overlap) {
      const ok = await shiftRisorsa(attivita.risorsa_id, extStart, deltaDays, excludeIds);
      if (!ok) return;
    }
  } else if (matrixState.autoShift && deltaDays < 0) {
    const cutDate = addDays(oldEndDay, 1);
    const overlap = await hasOverlapRange(attivita.risorsa_id, cutDate, oldEndDay, attivita.id);
    if (overlap == null) return;
    if (overlap) {
      const ok = await shiftRisorsa(attivita.risorsa_id, cutDate, deltaDays, excludeIds);
      if (!ok) return;
    }
  }
  const { error: updError } = await supabase
    .from("attivita")
    .update({
      data_fine: newEnd.toISOString(),
      titolo: firstTitle,
      descrizione: note || null,
      ore_assenza: isAssenteTitle(firstTitle) ? assenzaOre || null : null,
      commessa_id: isAssenteTitle(firstTitle) ? null : attivita.commessa_id,
    })
    .eq("id", attivita.id);
  if (updError) {
    setStatus(`Update error: ${updError.message}`, "error");
    return;
  }
  if (titles.length > 1) {
    const rows = titles.slice(1).map((title) => ({
      commessa_id: isAssenteTitle(title) ? null : attivita.commessa_id,
      titolo: title,
      descrizione: note || null,
      ore_assenza: isAssenteTitle(title) ? assenzaOre || null : null,
      risorsa_id: attivita.risorsa_id,
      data_inizio: attivita.data_inizio,
      data_fine: newEnd.toISOString(),
      stato: attivita.stato,
      reparto_id: attivita.reparto_id,
    }));
    const { error: insError } = await supabase.from("attivita").insert(rows);
    if (insError) {
      setStatus(`Create error: ${insError.message}`, "error");
      return;
    }
  }
  if (keyForChain && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
    const ok = await shiftDependentPhases(keyForChain, deltaDays, dependentToShift);
    if (!ok) return;
  }
  setStatus(titles.length > 1 ? "Activities updated." : "Activity updated.", "ok");
  closeActivityModal();
  await loadMatrixAttivita();
}

async function deleteActivity() {
  const attivita = matrixState.editingAttivita;
  if (!attivita) return;
  if (!canDeleteMatrixActivity(attivita)) {
    setStatus("Non autorizzato a eliminare attivita.", "error");
    return;
  }
  await openConfirmModal("Eliminare questa attivita?", async () => {
    const { error } = await supabase.from("attivita").delete().eq("id", attivita.id);
    if (error) {
      setStatus(`Delete error: ${error.message}`, "error");
      return;
    }
    setStatus("Activity removed.", "ok");
    closeActivityModal();
    await loadMatrixAttivita();
  });
}

async function deleteActivityById(id) {
  if (!id) return;
  const attivita = matrixState.attivita.find((x) => String(x.id) === String(id));
  if (!canDeleteMatrixActivity(attivita)) {
    setStatus("Non autorizzato a eliminare attivita.", "error");
    return;
  }
  await openConfirmModal("Eliminare questa attivita?", async () => {
    const { error } = await supabase.from("attivita").delete().eq("id", id);
    if (error) {
      setStatus(`Delete error: ${error.message}`, "error");
      return;
    }
    setStatus("Activity removed.", "ok");
    await loadMatrixAttivita();
  });
}

async function commitResize() {
  const ctx = matrixState.resizing;
  if (!ctx) return;
  if (ctx.previewBar) {
    ctx.previewBar.remove();
    ctx.previewBar = null;
  }
  if (ctx.originalBar) {
    ctx.originalBar.classList.remove("resizing");
  }
  matrixState.resizing = null;
  const targetKey = ctx.targetDayKey;
  if (!targetKey) return;
  const targetDate = dateFromKey(targetKey);
  if (targetDate < ctx.startDate) return;
  const newEnd = new Date(targetDate);
  newEnd.setHours(ctx.endTime.h, ctx.endTime.m, ctx.endTime.s, ctx.endTime.ms);
  const oldEndDay = ctx.oldEndDay || startOfDay(new Date(ctx.attivita.data_fine));
  const newEndDay = startOfDay(newEnd);
  const deltaDays = businessDayDiff(oldEndDay, newEndDay);
  let dependentActivities = null;
  let dependentToShift = null;
  if (matrixState.autoShift && isPhaseKey(ctx.attivita.titolo) && deltaDays !== 0) {
    dependentActivities = await getDependentActivitiesForKey(ctx.attivita);
    if (dependentActivities == null) return;
    dependentToShift = getDependentActivitiesToShift(ctx.attivita, deltaDays, dependentActivities);
  }
  const excludeIds =
    dependentToShift && dependentToShift.length
      ? [ctx.attivita.id, ...dependentToShift.map((a) => a.id)]
      : [ctx.attivita.id];
  if (matrixState.autoShift && deltaDays > 0) {
    const extStart = addDays(oldEndDay, 1);
    const overlap = await hasOverlapRange(ctx.attivita.risorsa_id, extStart, newEndDay, ctx.attivita.id);
    if (overlap == null) return;
    if (overlap) {
      const ok = await shiftRisorsa(ctx.attivita.risorsa_id, extStart, deltaDays, excludeIds);
      if (!ok) return;
    }
  } else if (matrixState.autoShift && deltaDays < 0) {
    const cutDate = addDays(oldEndDay, 1);
    const overlap = await hasOverlapRange(ctx.attivita.risorsa_id, cutDate, oldEndDay, ctx.attivita.id);
    if (overlap == null) return;
    if (overlap) {
      const ok = await shiftRisorsa(ctx.attivita.risorsa_id, cutDate, deltaDays, excludeIds);
      if (!ok) return;
    }
  }
  const { error } = await supabase
    .from("attivita")
    .update({ data_fine: newEnd.toISOString() })
    .eq("id", ctx.attivita.id);
  if (error) {
    setStatus(`Update error: ${error.message}`, "error");
    return;
  }
  if (isPhaseKey(ctx.attivita.titolo) && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
    const ok = await shiftDependentPhases(ctx.attivita, deltaDays, dependentToShift);
    if (!ok) return;
  }
  setStatus("Duration updated.", "ok");
  await loadMatrixAttivita();
}

function openAssignModal() {
  if (!assignModal) return;
  assignModal.classList.remove("hidden");
}

function closeAssignModal() {
  if (!assignModal) return;
  assignModal.classList.add("hidden");
  if (assignActivities) assignActivities.innerHTML = "";
  matrixState.pendingDrop = null;
}

function openResourcesModal() {
  if (!resourcesModal) return;
  resourcesModal.classList.remove("hidden");
  loadUsers();
}

function closeResourcesModal() {
  if (!resourcesModal) return;
  resourcesModal.classList.add("hidden");
}

function openProgressModal() {
  if (!progressModal) return;
  progressModal.classList.remove("hidden");
  if (progressMeta) progressMeta.textContent = "Select a commessa.";
  if (progressList) progressList.innerHTML = "";
  if (progressResults) progressResults.innerHTML = "";
  const currentYear = new Date().getFullYear();
  if (progressYearCurrent && progressYearPrev) {
    progressYearCurrent.dataset.year = String(currentYear);
    progressYearPrev.dataset.year = String(currentYear - 1);
    progressYearCurrent.textContent = String(currentYear);
    progressYearPrev.textContent = String(currentYear - 1);
    progressYearCurrent.classList.add("active");
    progressYearPrev.classList.remove("active");
  }
  if (progressSearch) progressSearch.value = "";
  renderProgressResults();
  if (progressSearch) progressSearch.focus();
}

function closeProgressModal() {
  if (!progressModal) return;
  progressModal.classList.add("hidden");
}

function getSelectedProgressYear() {
  if (progressYearCurrent && progressYearCurrent.classList.contains("active")) {
    return Number(progressYearCurrent.dataset.year);
  }
  if (progressYearPrev && progressYearPrev.classList.contains("active")) {
    return Number(progressYearPrev.dataset.year);
  }
  return null;
}

function renderProgressList(commessa, activities) {
  if (!progressList || !progressMeta) return;
  if (!commessa) {
    progressMeta.textContent = "Commessa non trovata.";
    progressList.innerHTML = "";
    if (progressTimelineDays) progressTimelineDays.innerHTML = "";
    return;
  }
  progressMeta.textContent = `${commessa.codice}${commessa.titolo ? " - " + commessa.titolo : ""}`;
  if (!activities.length) {
    progressList.innerHTML = `<div class="matrix-empty">Nessuna attivita trovata.</div>`;
    if (progressTimelineDays) progressTimelineDays.innerHTML = "";
    return;
  }
  const risorseById = new Map((state.risorse || []).map((r) => [r.id, r.nome]));
  progressList.innerHTML = "";
  if (progressTimelineDays) {
    const sorted = activities.slice().sort((a, b) => new Date(a.data_inizio) - new Date(b.data_inizio));
    const start = startOfDay(new Date(sorted[0].data_inizio));
    const end = sorted.reduce((max, a) => {
      const endDay = startOfDay(new Date(a.data_fine));
      return endDay > max ? endDay : max;
    }, start);
    const days = [];
    let cursor = new Date(start);
    while (cursor <= end) {
      if (!isWeekend(cursor)) days.push(new Date(cursor));
      cursor = addDays(cursor, 1);
    }
    if (!days.length) days.push(new Date(start));
    const timelineRoot = progressTimelineDays.closest(".progress-timeline");
    const containerWidth = timelineRoot ? timelineRoot.clientWidth : 900;
    const dayWidth = Math.max(1, Math.floor(containerWidth / days.length));
    const template = `repeat(${days.length}, minmax(${dayWidth}px, 1fr))`;
    progressTimelineDays.style.gridTemplateColumns = template;
    progressList.style.gridTemplateColumns = template;
    progressTimelineDays.innerHTML = "";
    days.forEach((day) => {
      const label = document.createElement("div");
      label.className = "progress-timeline-day";
      label.textContent = formatDateHumanNoYear(day);
      progressTimelineDays.appendChild(label);
    });
    const indexByKey = new Map(days.map((d, i) => [formatDateInput(d), i]));
    const normalizeToBusinessDay = (date, direction) => {
      let d = startOfDay(new Date(date));
      while (isWeekend(d)) d = addDays(d, direction);
      return d;
    };
    const laneEnds = [];
    sorted.forEach((a) => {
      const titleKey = (a.titolo || "").trim().toLowerCase();
      const displayTitle =
        titleKey === "altro" && a.descrizione ? `Altro: ${a.descrizione}` : a.titolo || "Attivita";
      const color = colorForKey(a.titolo || "attivita");
      const item = document.createElement("div");
      item.className = "progress-item";
      item.innerHTML = `
        <div class="progress-color" style="--progress-color: ${color.border};"></div>
        <div class="progress-body">
          <div class="progress-title">${displayTitle}</div>
          <div class="progress-meta">${risorseById.get(a.risorsa_id) || "Risorsa n/d"}</div>
        </div>
      `;
      let startDay = normalizeToBusinessDay(a.data_inizio, 1);
      let endDay = normalizeToBusinessDay(a.data_fine, -1);
      if (endDay < startDay) endDay = startDay;
      let startIndex = indexByKey.get(formatDateInput(startDay));
      let endIndex = indexByKey.get(formatDateInput(endDay));
      if (startIndex == null || endIndex == null) {
        progressList.appendChild(item);
        return;
      }
      if (endIndex < startIndex) {
        const tmp = startIndex;
        startIndex = endIndex;
        endIndex = tmp;
      }
      let lane = 0;
      while (lane < laneEnds.length && startIndex <= laneEnds[lane]) lane += 1;
      if (lane === laneEnds.length) {
        laneEnds.push(endIndex);
      } else {
        laneEnds[lane] = endIndex;
      }
      item.style.gridColumn = `${startIndex + 1} / ${endIndex + 2}`;
      item.style.gridRow = `${lane + 1}`;
      progressList.appendChild(item);
    });
    return;
  }
  activities.forEach((a) => {
    const titleKey = (a.titolo || "").trim().toLowerCase();
    const displayTitle =
      titleKey === "altro" && a.descrizione ? `Altro: ${a.descrizione}` : a.titolo || "Attivita";
    const color = colorForKey(a.titolo || "attivita");
    const item = document.createElement("div");
    item.className = "progress-item";
    item.innerHTML = `
      <div class="progress-color" style="--progress-color: ${color.border};"></div>
      <div class="progress-body">
        <div class="progress-title">${displayTitle}</div>
        <div class="progress-meta">${risorseById.get(a.risorsa_id) || "Risorsa n/d"}</div>
      </div>
    `;
    progressList.appendChild(item);
  });
}

let progressSelectedId = null;

function renderProgressResults() {
  if (!progressResults) return;
  const year = getSelectedProgressYear();
  const q = (progressSearch ? progressSearch.value : "").trim().toLowerCase();
  const tokens = q
    ? q
        .split(/[,\s;]+/)
        .map((t) => t.trim())
        .filter(Boolean)
    : [];
  progressResults.innerHTML = "";
  if (!year) return;
  let list = state.commesse.filter((c) => Number(c.anno) === year);
  const sortDate = (c) => {
    const d = c.data_ingresso || c.data_consegna_prevista || c.data_richiesta;
    return d ? new Date(d).getTime() : Number.POSITIVE_INFINITY;
  };
  if (tokens.length) {
    list = list.filter((c) => {
      const num = String(c.numero || "");
      const code = String(c.codice || "");
      const title = String(c.titolo || "");
      const codeLower = code.toLowerCase();
      const titleLower = title.toLowerCase();
      return tokens.some((token) => {
        const isNumber = /^\d+$/.test(token);
        if (isNumber) {
          return num === token || codeLower.includes(`_${token}`) || codeLower.includes(token);
        }
        return codeLower.includes(token) || titleLower.includes(token);
      });
    });
  }
  list = list.slice().sort((a, b) => {
    const da = sortDate(a);
    const db = sortDate(b);
    if (da !== db) return da - db;
    return compareCommesse(a, b);
  });
  if (!list.length) {
    const empty = document.createElement("div");
    empty.className = "matrix-empty";
    empty.textContent = "Nessuna commessa.";
    progressResults.appendChild(empty);
    return;
  }
  list.forEach((c) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "progress-result";
    btn.innerHTML = `
      <div class="progress-result-title">${c.codice}${c.titolo ? " - " + c.titolo : ""}</div>
    `;
    btn.dataset.id = c.id;
    if (progressSelectedId === c.id) btn.classList.add("active");
    btn.addEventListener("click", async () => {
      progressSelectedId = c.id;
      Array.from(progressResults.querySelectorAll(".progress-result")).forEach((el) =>
        el.classList.toggle("active", el.dataset.id === String(c.id))
      );
      const { data, error } = await supabase
        .from("attivita")
        .select("*")
        .eq("commessa_id", c.id)
        .order("data_inizio", { ascending: true });
      if (error) {
        setStatus(`Progress error: ${error.message}`, "error");
        return;
      }
      renderProgressList(c, data || []);
    });
    progressResults.appendChild(btn);
  });
}

let statusTimer = null;
function setStatus(message, type = "info") {
  statusMsg.textContent = message;
  statusMsg.classList.remove("ok", "error");
  if (type === "ok") statusMsg.classList.add("ok");
  if (type === "error") statusMsg.classList.add("error");
  statusMsg.classList.remove("hidden");
  if (statusTimer) clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    statusMsg.classList.remove("ok", "error");
    statusMsg.classList.add("hidden");
  }, 12000);
}

const withTimeout = (promise, ms) =>
  new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("timeout")), ms);
    promise
      .then((result) => {
        clearTimeout(timer);
        resolve(result);
      })
      .catch((err) => {
        clearTimeout(timer);
        reject(err);
      });
  });

const getStoredAuth = () => {
  const raw = localStorage.getItem(AUTH_STORAGE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
};

async function safeSetSession(accessToken, refreshToken, userFallback = null) {
  let error = null;
  let session = null;
  try {
    const result = await withTimeout(
      supabase.auth.setSession({ access_token: accessToken, refresh_token: refreshToken }),
      8000
    );
    error = result.error;
    session = result.data?.session || null;
  } catch (err) {
    if (String(err?.name || "") === "AbortError") {
      await new Promise((resolve) => setTimeout(resolve, 300));
      try {
        const retry = await withTimeout(
          supabase.auth.setSession({ access_token: accessToken, refresh_token: refreshToken }),
          8000
        );
        error = retry.error;
        session = retry.data?.session || null;
      } catch (retryErr) {
        error = retryErr;
      }
    } else {
      error = err;
    }
  }
  if (error) {
    setStatus("Unable to initialize session. Reload the page.", "error");
    return null;
  }
  if (!session && accessToken && refreshToken) {
    session = { access_token: accessToken, refresh_token: refreshToken, user: userFallback };
  }
  return session;
}

async function ensureSessionFromStorage() {
  const stored = getStoredAuth();
  if (!stored?.access_token || !stored?.refresh_token) return null;
  return await safeSetSession(stored.access_token, stored.refresh_token, stored.user || null);
}

let authSyncInProgress = false;
async function syncSignedIn(session) {
  if (!session || authSyncInProgress) return;
  authSyncInProgress = true;
  state.session = session;
  debugLog(`syncSignedIn start: user=${session.user?.id || "n/a"}`);
  try {
    debugLog("syncSignedIn step: loadProfile");
    const ok = await loadProfile(session.user);
    if (!ok || !state.profile) {
      syncSignedOut();
      return;
    }
    debugLog("syncSignedIn step: loadPermessi");
    await loadPermessi();
    debugLog("syncSignedIn step: loadReparti");
    await loadReparti();
    debugLog("syncSignedIn step: loadRisorse");
    await loadRisorse();
    debugLog("syncSignedIn step: loadUtenti");
    await loadUtenti();
    debugLog("syncSignedIn step: loadCommesse");
    await loadCommesse();
    calendarDate.value = formatDateInput(calendarState.date);
    debugLog("syncSignedIn step: loadAttivita");
    await loadAttivita();
    matrixDate.value = formatDateInput(matrixState.date);
    debugLog("syncSignedIn step: loadMatrixAttivita");
    await loadMatrixAttivita();
    debugLog("syncSignedIn step: setupRealtime");
    await setupRealtime();
    startTodoVisualRefreshLoop();
    debugLog("syncSignedIn complete");
  } catch (err) {
    debugLog(`syncSignedIn error: ${err?.message || err}`);
    setStatus(`Login sync error: ${err?.message || err}`, "error");
  } finally {
    authSyncInProgress = false;
  }
}

function syncSignedOut() {
  debugLog("syncSignedOut");
  stopTodoVisualRefreshLoop();
  state.session = null;
  state.profile = null;
  state.commesse = [];
  state.reportActivitiesMap = new Map();
  state.todoOverridesMap = new Map();
  state.todoClientFlowMap = new Map();
  state.todoClientFlowLoadedCommesse = new Set();
  authStatus.textContent = "Non autenticato";
  logoutBtn.classList.add("hidden");
  if (setPasswordBtn) setPasswordBtn.classList.add("hidden");
  if (openResourcesBtn) openResourcesBtn.classList.add("hidden");
  if (permissionsLink) permissionsLink.classList.add("hidden");
  if (authActions) authActions.classList.remove("is-authenticated");
  if (authDot) authDot.classList.add("hidden");
  setRoleBadge("");
  setWriteAccess(false);
  commesseList.innerHTML = "";
  clearSelection();
  calendarGrid.innerHTML = "";
  matrixGrid.innerHTML = "";
  setStatus("Session ended.");
}

function setAuthLoading(isLoading) {
  loginBtn.disabled = isLoading;
  emailInput.disabled = isLoading;
  if (passwordInput) passwordInput.disabled = isLoading;
  if (resetPasswordBtn) resetPasswordBtn.disabled = isLoading;
  loginBtn.textContent = isLoading ? "Signing in..." : "Log in";
}

function updateRevisionStamp() {
  if (!revisionStamp) return;
  const raw = document.lastModified;
  if (!raw) return;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return;
  const formatted = date.toLocaleString("it-IT", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
  revisionStamp.textContent = `Revisione: ${formatted}`;
}

function setRoleBadge(text) {
  if (!text) {
    roleBadge.classList.add("hidden");
    roleBadge.textContent = "";
    return;
  }
  roleBadge.classList.remove("hidden");
  roleBadge.textContent = text.toUpperCase();
}

function canPermission(key) {
  if (!state.profile) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  if (role === "admin") return true;
  return Boolean(state.permessi?.[key]);
}

function getOwnRisorsaId() {
  if (!state.profile) return null;
  const match = state.risorse.find((r) => String(r.utente_id || "") === String(state.profile.id));
  return match ? match.id : null;
}

function getOwnRepartoId() {
  if (!state.profile) return null;
  if (state.profile.reparto_id) return state.profile.reparto_id;
  const match = state.risorse.find((r) => String(r.utente_id || "") === String(state.profile.id));
  return match ? match.reparto_id || null : null;
}

function canMoveMatrixActivity(activity, targetRisorsaId = null) {
  if (!state.profile) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  if (role === "admin" || role === "responsabile") return Boolean(state.permessi?.can_move_matrix);
  if (role === "operatore") {
    if (!state.permessi?.can_move_matrix) return false;
    const ownId = getOwnRisorsaId();
    if (!ownId) return false;
    const activityOk = activity ? String(activity.risorsa_id) === String(ownId) : true;
    const targetOk = targetRisorsaId ? String(targetRisorsaId) === String(ownId) : true;
    return activityOk && targetOk;
  }
  return false;
}

function canDeleteMatrixActivity(activity) {
  if (!state.profile) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  if (role === "admin" || role === "responsabile") return Boolean(state.permessi?.can_delete_matrix);
  if (role === "operatore") {
    if (!state.permessi?.can_delete_matrix) return false;
    const ownId = getOwnRisorsaId();
    if (!ownId) return false;
    return activity ? String(activity.risorsa_id) === String(ownId) : false;
  }
  return false;
}

function canAssignCommessaToRisorsa(risorsaId) {
  if (!state.profile || !risorsaId) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  if (role === "operatore") {
    const ownId = getOwnRisorsaId();
    return Boolean(ownId) && String(ownId) === String(risorsaId);
  }
  return canMoveMatrixActivity(null, risorsaId);
}

function canEditGanttMilestone(field) {
  if (!state.profile) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  if (!canPermission("can_move_gantt")) return false;
  if (role === "admin") return true;
  if (role === "planner")
    return field === "ordine" || field === "consegna" || field === "kit_cavi" || field === "prelievo";
  if (field === "ordine" || field === "consegna") return false;
  if (field === "kit_cavi") return role === "responsabile";
  if (field === "ordine_pianificato") return role === "responsabile";
  return true;
}

function setWriteAccess(canWrite) {
  state.canWrite = canWrite;
  const disabled = !canWrite;
  const inputs = document.querySelectorAll(
    "#detailForm input, #detailForm select, #detailForm textarea, #newForm input, #newForm select, #newForm textarea, #addRepartoForm select, #resourceForm input, #resourceForm select, #resourceForm button, #resourcesList input, #resourcesList select, #resourcesList button"
  );
  inputs.forEach((input) => (input.disabled = disabled));
  updateBtn.disabled = disabled;
  if (addRepartoForm) addRepartoForm.querySelector("button").disabled = disabled;
  newForm.querySelector("button").disabled = disabled;
  importCommitBtn.disabled = disabled;
  openImportBtn.disabled = disabled;
  if (importDatesCommitBtn) importDatesCommitBtn.disabled = disabled;
  if (openImportDatesBtn) openImportDatesBtn.disabled = disabled;
  if (importProductionCommitBtn) importProductionCommitBtn.disabled = disabled;
  if (openImportProductionBtn) openImportProductionBtn.disabled = disabled;
  if (openNewCommessaBtn) openNewCommessaBtn.disabled = disabled;
  if (matrixCommessa) matrixCommessa.disabled = disabled;
  if (matrixAttivita) matrixAttivita.disabled = disabled;
}

async function signIn() {
  const email = emailInput.value.trim();
  if (!email) {
    setStatus("Enter a valid email.", "error");
    return;
  }
  const password = passwordInput ? passwordInput.value : "";
  if (!password) {
    setStatus("Enter the password.", "error");
    return;
  }
  setAuthLoading(true);
  debugLog("signIn start");
  try {
    let error = null;
    let session = null;
    try {
      const result = await withTimeout(
        supabase.auth.signInWithPassword({
          email,
          password,
        }),
        8000
      );
      error = result.error;
      session = result.data?.session || null;
    } catch (innerErr) {
      if (String(innerErr?.name || "") === "AbortError") {
        await new Promise((resolve) => setTimeout(resolve, 300));
        const retry = await withTimeout(
          supabase.auth.signInWithPassword({
            email,
            password,
          }),
          8000
        );
        error = retry.error;
        session = retry.data?.session || null;
      } else if (String(innerErr?.message || "") === "timeout") {
        debugLog("signIn timeout");
        setStatus(
          "Login pending. Close other open app tabs, then reload the page.",
          "error"
        );
        return;
      } else {
        throw innerErr;
      }
    }
    if (error) {
      debugLog(`signIn error: ${error.message}`);
      setStatus(`Login error: ${error.message}`, "error");
      return;
    }
    setStatus("Signed in.", "ok");
    if (!session) {
      const { data } = await supabase.auth.getSession();
      session = data?.session || null;
    }
    if (!session) {
      session = await ensureSessionFromStorage();
    }
    if (session) {
      debugLog("signIn session ready");
      await syncSignedIn(session);
      return;
    }
    debugLog("signIn session missing");
    setStatus("Signed in but session unavailable. Reload the page.", "error");
  } catch (err) {
    debugLog(`signIn fatal: ${err?.message || err}`);
    if (String(err?.name || "") === "AbortError") {
      setStatus("Sign-in blocked by a lock. Reload the page and try again.", "error");
    } else {
      setStatus("Unexpected login error. Try again.", "error");
    }
  } finally {
    setAuthLoading(false);
  }
}

async function resetPassword() {
  const email = emailInput.value.trim();
  if (!email) {
    setStatus("Enter a valid email.", "error");
    return;
  }
  updateResetCooldownUI(true);
  if (resetPasswordBtn && resetPasswordBtn.disabled) return;
  setAuthLoading(true);
  const { error } = await supabase.auth.resetPasswordForEmail(email, {
    redirectTo: "https://pianificazionecommesse.it/reset.html",
  });
  setAuthLoading(false);
  if (error) {
    const msg = String(error.message || "");
    if (msg.toLowerCase().includes("rate limit")) {
      setResetFeedback("You've hit the request limit. Wait a few minutes and try again.", "error");
    } else {
      setResetFeedback(`Reset error: ${error.message}`, "error");
    }
    return;
  }
  const cooldownUntil = Date.now() + RESET_COOLDOWN_MS;
  setResetCooldownUntil(cooldownUntil);
  setResetFeedback(
    "Reset link sent. Check your email.",
    ""
  );
  updateResetCooldownUI(false);
  showResetWaitOverlay(RESET_COOLDOWN_MS);
}

async function changePassword() {
  if (!state.session) {
    setStatus("Sign in first.", "error");
    return;
  }
  const password = window.prompt("Nuova password (min 8 caratteri):", "");
  if (!password) return;
  if (password.trim().length < 8) {
    setStatus("Password troppo corta (min 8 caratteri).", "error");
    return;
  }
  const confirm = window.prompt("Conferma nuova password:", "");
  if (confirm !== password) {
    setStatus("Le password non coincidono.", "error");
    return;
  }
  if (setPasswordBtn) setPasswordBtn.disabled = true;
  try {
    const { error } = await supabase.auth.updateUser({ password });
    if (error) {
      setStatus(`Errore aggiornamento password: ${error.message}`, "error");
      return;
    }
    setStatus("Password aggiornata.", "ok");
  } finally {
    if (setPasswordBtn) setPasswordBtn.disabled = false;
  }
}

async function signOut() {
  await supabase.auth.signOut();
}

function clearSelection() {
  state.selected = null;
  selectedCode.textContent = "Nessuna selezionata";
  detailForm.reset();
  if (d.telaio_ordinato) {
    setTelaioOrdinatoButton(d.telaio_ordinato, false, false, d.data_consegna_telaio_effettiva || null);
  }
  setDetailSnapshot();
  if (repartiList) repartiList.innerHTML = "";
  if (commessaDetailModal) commessaDetailModal.classList.add("hidden");
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  return dateStr.split("T")[0];
}

function openImportModal() {
  if (!state.canWrite) return;
  importModal.classList.remove("hidden");
  importTextarea.focus();
  updateImportPreview();
}

function openImportDatesModal() {
  if (!state.canWrite) return;
  if (!importDatesModal) return;
  importDatesModal.classList.remove("hidden");
  if (importDatesTextarea) importDatesTextarea.focus();
  updateImportDatesPreview();
}

function openImportProductionModal() {
  if (!state.canWrite) return;
  if (!importProductionModal) return;
  importProductionModal.classList.remove("hidden");
  if (importProductionTextarea) importProductionTextarea.focus();
  updateImportProductionPreview();
}

function openCommessaCreateModal() {
  if (!state.canWrite) return;
  if (!commessaCreateModal) return;
  commessaCreateModal.classList.remove("hidden");
  if (newForm) newForm.reset();
  if (n.anno && !n.anno.value) n.anno.value = String(new Date().getFullYear());
  updateNumeroWarning(n.anno, n.numero, newNumeroWarning);
  updateCodicePreview(n.anno, n.numero, newCodicePreview);
  const firstInput = commessaCreateModal.querySelector("input, textarea, select");
  if (firstInput) firstInput.focus();
}

function closeCommessaCreateModal() {
  if (!commessaCreateModal) return;
  commessaCreateModal.classList.add("hidden");
}

function openCommessaDetailModal() {
  if (!commessaDetailModal) return;
  commessaDetailModal.classList.remove("hidden");
}

function closeCommessaDetailModal() {
  if (!commessaDetailModal) return;
  commessaDetailModal.classList.add("hidden");
}

async function requestCloseCommessaDetailModal() {
  if (!commessaDetailModal || commessaDetailModal.classList.contains("hidden")) return;
  if (!detailDirty) {
    closeCommessaDetailModal();
    return;
  }
  const confirmed = await openConfirmModal("You have unsaved changes. Close without saving?", null, {
    okLabel: "Discard",
    cancelLabel: "Keep editing",
  });
  if (!confirmed) return;
  if (state.selected) {
    selectCommessa(state.selected.id);
  }
  closeCommessaDetailModal();
}

function openConfirmModal(message, onConfirm, options = {}) {
  if (!confirmModal || !confirmMessage) return Promise.resolve(false);
  confirmMessage.textContent = message;
  confirmAction = typeof onConfirm === "function" ? onConfirm : null;
  if (confirmOkBtn) {
    confirmOkBtn.textContent = options.okLabel || confirmDefaults.okLabel;
  }
  if (confirmCancelBtn) {
    confirmCancelBtn.textContent = options.cancelLabel || confirmDefaults.cancelLabel;
  }
  confirmModal.classList.remove("hidden");
  return new Promise((resolve) => {
    confirmResolver = resolve;
  });
}

function closeConfirmModal(result) {
  if (!confirmModal) return;
  confirmModal.classList.add("hidden");
  if (confirmOkBtn) confirmOkBtn.textContent = confirmDefaults.okLabel;
  if (confirmCancelBtn) confirmCancelBtn.textContent = confirmDefaults.cancelLabel;
  if (confirmResolver) {
    confirmResolver(Boolean(result));
    confirmResolver = null;
  }
  if (result && confirmAction) {
    const action = confirmAction;
    confirmAction = null;
    action();
  } else {
    confirmAction = null;
  }
}

function updateDeleteCommessaConfirmState() {
  if (!deleteCommessaConfirmBtn || !deleteCommessaInput) return;
  const ok = deleteCommessaInput.value.trim().toUpperCase() === "DELETE";
  deleteCommessaConfirmBtn.disabled = !ok;
}

function openDeleteCommessaModal() {
  if (!deleteCommessaModal || !state.canWrite || !state.selected) return;
  deleteCommessaId = state.selected.id;
  const code = state.selected.codice || `${state.selected.anno}_${state.selected.numero}`;
  if (deleteCommessaMessage) {
    deleteCommessaMessage.textContent = `Type DELETE to remove commessa ${code}. This cannot be undone.`;
  }
  if (deleteCommessaInput) {
    deleteCommessaInput.value = "";
  }
  updateDeleteCommessaConfirmState();
  deleteCommessaModal.classList.remove("hidden");
  if (deleteCommessaInput) deleteCommessaInput.focus();
}

function closeDeleteCommessaModal() {
  if (!deleteCommessaModal) return;
  deleteCommessaModal.classList.add("hidden");
  deleteCommessaId = null;
  if (deleteCommessaConfirmBtn) {
    deleteCommessaConfirmBtn.disabled = true;
    deleteCommessaConfirmBtn.textContent = "Delete";
  }
}

async function handleDeleteCommessaConfirm() {
  if (!deleteCommessaId || !state.canWrite) return;
  if (deleteCommessaInput && deleteCommessaInput.value.trim().toUpperCase() !== "DELETE") {
    updateDeleteCommessaConfirmState();
    return;
  }
  if (deleteCommessaConfirmBtn) {
    deleteCommessaConfirmBtn.disabled = true;
    deleteCommessaConfirmBtn.textContent = "Deleting...";
  }
  try {
    let { error: commErr } = await supabase.from("commesse").delete().eq("id", deleteCommessaId);
    if (commErr) {
      const msg = String(commErr.message || "").toLowerCase();
      const needsCleanup =
        msg.includes("foreign key") || msg.includes("violates") || msg.includes("referenced");
      if (!needsCleanup) throw commErr;
      const { error: actErr } = await supabase.from("attivita").delete().eq("commessa_id", deleteCommessaId);
      if (actErr) throw actErr;
      const retry = await supabase.from("commesse").delete().eq("id", deleteCommessaId);
      if (retry.error) throw retry.error;
    }
    setStatus("Commessa deleted.", "ok");
    showToast("Commessa deleted.", "ok");
    closeDeleteCommessaModal();
    closeCommessaDetailModal();
    clearSelection();
    state.reportFilteredCommesse = null;
    state.reportFilteredKey = "";
    state.reportActivitiesMap.delete(String(deleteCommessaId));
    await loadCommesse();
    await loadMatrixAttivita();
  } catch (err) {
    const msg = `Delete error: ${err?.message || err}`;
    setStatus(msg, "error");
    showToast(msg, "error");
  } finally {
    if (deleteCommessaConfirmBtn) {
      deleteCommessaConfirmBtn.disabled = false;
      deleteCommessaConfirmBtn.textContent = "Delete";
    }
  }
}

function openWorkloadModal() {
  if (!workloadModal) return;
  if (!state.session) {
    setStatus("Sign in to view workload.", "error");
    return;
  }
  const today = new Date();
  if (workloadWeekStart && !workloadWeekStart.value) {
    workloadWeekStart.value = toWeekInputValue(today);
  }
  if (workloadWeekEnd && !workloadWeekEnd.value) {
    workloadWeekEnd.value = toWeekInputValue(today);
  }
  workloadModal.classList.remove("hidden");
  calculateWorkload();
}

function closeWorkloadModal() {
  if (!workloadModal) return;
  workloadModal.classList.add("hidden");
}

function calculateWorkload() {
  if (!workloadSummary || !workloadList) return;
  if (!state.session) {
    workloadSummary.textContent = "Sign in to calculate workload.";
    workloadList.innerHTML = "";
    return;
  }
  const startWeek = workloadWeekStart ? workloadWeekStart.value : "";
  const endWeek = workloadWeekEnd ? workloadWeekEnd.value : "";
  const startDate = weekInputToDate(startWeek);
  const endDate = weekInputToDate(endWeek || startWeek);
  if (!startDate || !endDate) {
    workloadSummary.textContent = "Select a valid start and end week.";
    workloadList.innerHTML = "";
    return;
  }
  let rangeStart = startOfWeek(startDate);
  let rangeEnd = startOfWeek(endDate);
  if (rangeEnd < rangeStart) {
    const tmp = rangeStart;
    rangeStart = rangeEnd;
    rangeEnd = tmp;
  }
  const rangeFinish = addDays(rangeEnd, 4);
  const isDateOk = (d) =>
    d && !Number.isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
  const entries = (state.commesse || [])
    .map((c) => {
      const raw = c.data_ordine_telaio || "";
      const target = raw ? new Date(raw) : null;
      if (!isDateOk(target)) return null;
      const targetDay = startOfDay(target);
      if (targetDay < rangeStart || targetDay > rangeFinish) return null;
      const plannedRaw = c.data_consegna_telaio_effettiva || "";
      const planned = plannedRaw ? new Date(plannedRaw) : null;
      const plannedDay = isDateOk(planned) ? startOfDay(planned) : null;
      return {
        commessa: c,
        target: targetDay,
        planned: plannedDay,
        ordered: Boolean(c.telaio_consegnato),
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.target.getTime() - b.target.getTime());

  if (!entries.length) {
    workloadSummary.textContent =
      "No frame orders scheduled in the selected weeks.";
    workloadList.innerHTML = "";
    return;
  }

  const withPlanned = entries.filter((e) => e.planned).length;
  const withoutPlanned = entries.length - withPlanned;
  const orderedCount = entries.filter((e) => e.ordered).length;
  workloadSummary.textContent =
    `Range ${formatDateDMY(rangeStart)} \u2192 ${formatDateDMY(rangeFinish)}. ` +
    `Frames expected: ${entries.length}. Planned dates: ${withPlanned}. ` +
    `Missing planned dates: ${withoutPlanned}. Ordered: ${orderedCount}.`;

  const groups = new Map();
  entries.forEach((entry) => {
    const week = getIsoWeekNumber(entry.target);
    const year = getIsoWeekYear(entry.target);
    const key = `${year}-W${String(week).padStart(2, "0")}`;
    const list = groups.get(key) || [];
    list.push(entry);
    groups.set(key, list);
  });

  const groupHtml = Array.from(groups.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, list]) => {
      const weekStart = weekInputToDate(key);
      const weekEnd = weekStart ? addDays(weekStart, 4) : null;
      const headerRange = weekStart && weekEnd
        ? `${formatDateDMY(weekStart)} \u2192 ${formatDateDMY(weekEnd)}`
        : "";
      const itemsHtml = list
        .map((entry) => {
          const code = entry.commessa.codice || `${entry.commessa.anno}_${entry.commessa.numero}`;
          const title = entry.commessa.titolo || "Senza titolo";
          const plannedText = entry.planned ? formatDateDMY(entry.planned) : "-";
          const plannedPill = entry.planned
            ? `<span class="workload-pill is-ok">Planned</span>`
            : `<span class="workload-pill">No planned</span>`;
          const orderedPill = entry.ordered
            ? `<span class="workload-pill is-ok">Ordered</span>`
            : `<span class="workload-pill">Not ordered</span>`;
          return `
            <div class="workload-item">
              <div class="workload-code">${escapeHtml(code)}</div>
              <div>${escapeHtml(title)}</div>
              <div>T: ${escapeHtml(formatDateDMY(entry.target))}</div>
              <div>P: ${escapeHtml(plannedText)}</div>
              <div>${plannedPill} ${orderedPill}</div>
            </div>
          `;
        })
        .join("");
      return `
        <div class="workload-group">
          <div class="workload-group-header">
            <span>Week ${escapeHtml(key.split("-W")[1])}</span>
            ${headerRange ? `<span class="workload-group-range">${escapeHtml(headerRange)}</span>` : ""}
          </div>
          ${itemsHtml}
        </div>
      `;
    })
    .join("");

  workloadList.innerHTML = groupHtml;
}

async function jumpToMatrixActivity(commessaId, startTs, activityId, risorsaId) {
  const ts = startTs ? Number(startTs) : NaN;
  const date = Number.isFinite(ts) ? new Date(ts) : null;
  if (!date || Number.isNaN(date.getTime())) return;
  if (matrixPanel && matrixPanel.classList.contains("collapsed")) {
    matrixPanel.classList.remove("collapsed");
  }
  if (risorsaId) {
    const resource = state.risorse.find((r) => String(r.id) === String(risorsaId));
    if (resource) {
      const reparto = state.reparti.find((rep) => String(rep.id) === String(resource.reparto_id));
      if (reparto && matrixState.collapsedReparti.has(reparto.nome)) {
        matrixState.collapsedReparti.delete(reparto.nome);
      }
    }
  }
  matrixState.date = startOfWeek(date);
  if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
  const activityKey = activityId ? String(activityId) : "";
  let targetBar = null;
  if (activityKey && matrixGrid) {
    targetBar = matrixGrid.querySelector(
      `.matrix-activity-bar[data-activity-id="${activityKey.replace(/"/g, '\\"')}"]`
    );
  }
  if (targetBar) {
    document
      .querySelectorAll(".matrix-activity-bar.matrix-jump-glow")
      .forEach((el) => el.classList.remove("matrix-jump-glow"));
    targetBar.classList.add("matrix-jump-glow");
    targetBar.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
    window.setTimeout(() => targetBar.classList.remove("matrix-jump-glow"), 2600);
  }
  if (commessaId) {
    applyCommessaHighlight(commessaId);
  }
}

function closeReportDeptMenu() {
  if (!reportDeptMenu) return;
  reportDeptMenu.classList.add("hidden");
  reportDeptMenu.innerHTML = "";
}

function showReportDeptMenu(x, y, deptLabel, schedule, deptKey, commessaId) {
  if (!reportDeptMenu) return;
  const padding = 12;
  const byPerson = new Map();
  (schedule || [])
    .filter((entry) => entry.dept === deptKey)
    .forEach((entry) => {
      const name = entry.risorsa || "Risorsa n/d";
      const list = byPerson.get(name) || [];
      list.push({
        id: entry.id,
        risorsaId: entry.risorsaId,
        titolo: entry.titolo || "Attivita",
        start: entry.start ? new Date(entry.start) : null,
      });
      byPerson.set(name, list);
    });

  if (!byPerson.size) {
    reportDeptMenu.innerHTML = `
      <div class="report-dept-menu-title">${escapeHtml(deptLabel)}</div>
      <div class="report-dept-menu-sub">No participants</div>
    `;
  } else {
    const items = Array.from(byPerson.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([name, acts]) => {
        const sorted = acts
          .slice()
          .sort((a, b) => {
            const aTime = a.start ? a.start.getTime() : 0;
            const bTime = b.start ? b.start.getTime() : 0;
            if (aTime !== bTime) return aTime - bTime;
            return String(a.titolo || "").localeCompare(String(b.titolo || ""));
          });
        const actsText = sorted
          .map((a) => {
            const startText = a.start ? formatDateLocal(a.start) : "-";
            const startTs = a.start ? a.start.getTime() : "";
            return `
              <div class="report-dept-activity-row">
                <span class="report-dept-activity-title">${escapeHtml(a.titolo)}</span>
                <span class="report-dept-activity-date">${escapeHtml(startText)}</span>
                <button class="report-dept-activity-arrow" data-start-ts="${escapeHtml(String(startTs))}" data-activity-id="${escapeHtml(String(a.id || ""))}" data-risorsa-id="${escapeHtml(String(a.risorsaId || ""))}" data-commessa="${escapeHtml(String(commessaId || ""))}" title="Go to matrix">&rarr;</button>
              </div>
            `;
          })
          .join("");
        return `
          <div class="report-dept-menu-item">
            <div class="report-dept-menu-name">${escapeHtml(name)}</div>
            <div class="report-dept-menu-acts">${actsText || "<div class=\"report-dept-activity-row\">-</div>"}</div>
          </div>
        `;
      })
      .join("");
    reportDeptMenu.innerHTML = `
      <div class="report-dept-menu-title">${escapeHtml(deptLabel)}</div>
      <div class="report-dept-menu-sub">${byPerson.size} people</div>
      ${items}
    `;
  }

  reportDeptMenu.classList.remove("hidden");
  reportDeptMenu.style.visibility = "hidden";
  requestAnimationFrame(() => {
    const rect = reportDeptMenu.getBoundingClientRect();
    let left = x + 12;
    let top = y + 12;
    if (left + rect.width > window.innerWidth - padding) {
      left = x - rect.width - 12;
    }
    if (top + rect.height > window.innerHeight - padding) {
      top = y - rect.height - 12;
    }
    left = Math.max(padding, Math.min(left, window.innerWidth - rect.width - padding));
    top = Math.max(padding, Math.min(top, window.innerHeight - rect.height - padding));
    reportDeptMenu.style.left = `${left}px`;
    reportDeptMenu.style.top = `${top}px`;
    reportDeptMenu.style.visibility = "visible";
  });

  reportDeptMenu.querySelectorAll(".report-dept-activity-arrow").forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      e.stopPropagation();
      const startTs = btn.dataset.startTs;
      const commessa = btn.dataset.commessa;
      const activityId = btn.dataset.activityId;
      const risorsaId = btn.dataset.risorsaId;
      if (!startTs || !commessa) return;
      await jumpToMatrixActivity(commessa, startTs, activityId, risorsaId);
      closeReportDeptMenu();
    });
  });
}

function openBackupModal() {
  if (!backupModal) return;
  if (!state.session) {
    setStatus("Sign in to access backups.", "error");
    return;
  }
  if (backupExportStatus) backupExportStatus.textContent = "Ready to export.";
  if (backupImportSummary) backupImportSummary.textContent = "No file selected.";
  if (backupImportInput) backupImportInput.value = "";
  if (backupImportConfirm) backupImportConfirm.value = "";
  backupPayload = null;
  updateBackupImportState();
  backupModal.classList.remove("hidden");
}

function closeBackupModal() {
  if (!backupModal) return;
  backupModal.classList.add("hidden");
}

function updateBackupImportState() {
  if (!backupImportBtn) return;
  const confirmed = backupImportConfirm && backupImportConfirm.value.trim().toUpperCase() === "IMPORT";
  backupImportBtn.disabled = !backupPayload || !confirmed || !state.canWrite;
}

function downloadJson(data, filename) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function fetchBackupTable(name) {
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from(name).select("*"),
        FETCH_TIMEOUT_MS * 4
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`backup ${name} timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    data = await fetchTableViaRest(name, "select=*");
  }
  if (!data) {
    return { error: error?.message || "No access", rows: [] };
  }
  return { rows: data, error: null };
}

async function handleBackupExport() {
  if (!state.session) {
    setStatus("Sign in to export backup.", "error");
    return;
  }
  if (!backupExportBtn || !backupExportStatus) return;
  backupExportBtn.disabled = true;
  backupExportBtn.textContent = "Exporting...";
  const tables = {};
  const warnings = [];
  try {
    for (let i = 0; i < BACKUP_TABLES.length; i += 1) {
      const table = BACKUP_TABLES[i].name;
      backupExportStatus.textContent = `Exporting ${table} (${i + 1}/${BACKUP_TABLES.length})...`;
      const result = await fetchBackupTable(table);
      if (result.error) {
        warnings.push(`${table}: ${result.error}`);
      } else {
        tables[table] = result.rows || [];
      }
    }
    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      tables,
      warnings,
    };
    const stamp = new Date().toISOString().slice(0, 10);
    const filename = `commesse-backup-${stamp}.json`;
    downloadJson(payload, filename);
    const exportedCount = Object.keys(tables).length;
    backupExportStatus.textContent = `Backup ready. Tables exported: ${exportedCount}.`;
    if (warnings.length) {
      backupExportStatus.textContent += ` Warnings: ${warnings.length}.`;
    }
    showToast("Backup exported.", "ok");
  } catch (err) {
    backupExportStatus.textContent = `Export failed: ${err?.message || err}`;
    showToast("Backup export failed.", "error");
  } finally {
    backupExportBtn.disabled = false;
    backupExportBtn.textContent = "Download backup";
  }
}

function summarizeBackupPayload(payload) {
  const tables = payload?.tables || {};
  const lines = [];
  let totalRows = 0;
  BACKUP_TABLES.forEach((table) => {
    const rows = Array.isArray(tables[table.name]) ? tables[table.name].length : 0;
    if (rows) {
      lines.push(`${table.name}: ${rows}`);
      totalRows += rows;
    }
  });
  if (!lines.length) {
    return "No usable tables found in the backup.";
  }
  return `Tables: ${lines.join(", ")}. Total rows: ${totalRows}.`;
}

async function handleBackupFileChange() {
  if (!backupImportInput || !backupImportSummary) return;
  const file = backupImportInput.files && backupImportInput.files[0];
  if (!file) {
    backupPayload = null;
    backupImportSummary.textContent = "No file selected.";
    updateBackupImportState();
    return;
  }
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (!parsed || !parsed.tables) {
      backupPayload = null;
      backupImportSummary.textContent = "Invalid backup file.";
      updateBackupImportState();
      return;
    }
    backupPayload = parsed;
    backupImportSummary.textContent = summarizeBackupPayload(parsed);
  } catch (err) {
    backupPayload = null;
    backupImportSummary.textContent = `Invalid JSON: ${err?.message || err}`;
  }
  updateBackupImportState();
}

function chunkArray(list, size) {
  const chunks = [];
  for (let i = 0; i < list.length; i += size) {
    chunks.push(list.slice(i, i + size));
  }
  return chunks;
}

async function handleBackupImport() {
  if (!state.canWrite) {
    setStatus("You don't have permission to import backups.", "error");
    showToast("You don't have permission to import backups.", "error");
    return;
  }
  if (!backupPayload) return;
  if (backupImportConfirm && backupImportConfirm.value.trim().toUpperCase() !== "IMPORT") {
    updateBackupImportState();
    return;
  }
  if (backupImportBtn) {
    backupImportBtn.disabled = true;
    backupImportBtn.textContent = "Importing...";
  }
  if (backupImportSummary) backupImportSummary.textContent = "Import in progress...";
  const failures = [];
  try {
    const tables = backupPayload.tables || {};
    for (let i = 0; i < BACKUP_TABLES.length; i += 1) {
      const table = BACKUP_TABLES[i];
      const rows = Array.isArray(tables[table.name]) ? tables[table.name] : [];
      if (!rows.length) continue;
      const chunks = chunkArray(rows, 400);
      for (let c = 0; c < chunks.length; c += 1) {
        if (backupImportSummary) {
          backupImportSummary.textContent = `Importing ${table.name} (${c + 1}/${chunks.length})...`;
        }
        const { error } = await supabase
          .from(table.name)
          .upsert(chunks[c], { onConflict: table.conflict });
        if (error) {
          failures.push(`${table.name}: ${error.message}`);
          break;
        }
      }
    }
    if (failures.length) {
      backupImportSummary.textContent = `Import completed with ${failures.length} errors.`;
      showToast("Import completed with errors.", "error");
    } else {
      backupImportSummary.textContent = "Import completed.";
      showToast("Backup imported.", "ok");
    }
    await loadReparti();
    await loadRisorse();
    await loadCommesse();
    await loadMatrixAttivita();
    await loadAttivita();
  } catch (err) {
    backupImportSummary.textContent = `Import failed: ${err?.message || err}`;
    showToast("Backup import failed.", "error");
  } finally {
    if (backupImportBtn) {
      backupImportBtn.disabled = false;
      backupImportBtn.textContent = "Import backup";
    }
    updateBackupImportState();
  }
}

function closeImportModal() {
  importModal.classList.add("hidden");
}

function closeImportDatesModal() {
  if (!importDatesModal) return;
  importDatesModal.classList.add("hidden");
}

function closeImportProductionModal() {
  if (!importProductionModal) return;
  importProductionModal.classList.add("hidden");
}

function detectDelimiter(text) {
  if (text.includes("\t")) return "\t";
  if (text.includes(";") && !text.includes(",")) return ";";
  return ",";
}

function parseImportRows(text) {
  const cleaned = text.trim();
  if (!cleaned) return { rows: [], errors: [], total: 0 };
  const delimiter = detectDelimiter(cleaned);
  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim() !== "");
  let startIndex = 0;
  if (lines.length) {
    const first = lines[0].toLowerCase();
    if (first.includes("anno") || first.includes("numero") || first.includes("descr")) {
      startIndex = 1;
    }
  }

  const rows = [];
  const errors = [];
  for (let i = startIndex; i < lines.length; i++) {
    const cols = lines[i]
      .split(delimiter)
      .map((c) => c.trim().replace(/^"(.*)"$/, "$1"));

    if (cols.length < 3) {
      errors.push(`Riga ${i + 1}: servono anno, numero, descrizione`);
      continue;
    }

    const annoRaw = cols[0];
    const numeroRaw = cols[1];
    const descrizione = cols.slice(2).join(" ").trim();

    if (!/^\d{4}$/.test(annoRaw)) {
      errors.push(`Riga ${i + 1}: anno non valido`);
      continue;
    }
    if (!/^\d+$/.test(numeroRaw)) {
      errors.push(`Riga ${i + 1}: numero non valido`);
      continue;
    }

    const anno = Number(annoRaw);
    const numero = Number(numeroRaw);
    const codice = `${anno}_${numero}`;
    const titolo = descrizione || null;

    rows.push({ codice, titolo, anno, numero });
  }

  return { rows, errors, total: lines.length - startIndex };
}

function updateImportPreview() {
  const { rows, errors, total } = parseImportRows(importTextarea.value);
  const sample = rows.slice(0, 3).map((r) => `${r.codice} - ${r.titolo || "Senza titolo"}`);
  const errorText = errors.length ? `Errori: ${errors.slice(0, 3).join(" | ")}` : "Nessun errore rilevato.";
  const summary =
    `Righe riconosciute: ${rows.length}/${total}. ` +
    `Esempi: ${sample.join(" - ") || "-"}. ${errorText}`;
  const table = buildRawImportTable(importTextarea.value);
  importPreview.innerHTML = `${escapeHtml(summary)}${table}`;
}

function formatDateInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatDateLocal(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildRawImportTable(text, maxRows = 8) {
  const cleaned = (text || "").trim();
  if (!cleaned) return "";
  const delimiter = detectDelimiter(cleaned);
  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim() !== "");
  if (!lines.length) return "";
  const rows = lines.slice(0, maxRows).map((line) =>
    line
      .split(delimiter)
      .map((c) => c.trim().replace(/^"(.*)"$/, "$1"))
  );
  const maxCols = Math.max(...rows.map((r) => r.length), 1);
  const headerHtml = Array.from({ length: maxCols })
    .map((_, i) => `<th>${escapeHtml(`Col ${i + 1}`)}</th>`)
    .join("");
  const bodyHtml = rows
    .map((row) => {
      const cells = Array.from({ length: maxCols }).map((_, i) => row[i] || "");
      return `<tr>${cells.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`;
    })
    .join("");
  return `<div class="import-preview-table-wrap"><table class="import-preview-table"><thead><tr>${headerHtml}</tr></thead><tbody>${bodyHtml}</tbody></table></div>`;
}

function updateImportDatesPreview() {
  if (!importDatesPreview) return;
  const { rows, errors, warnings, total } = parseImportDateRows(importDatesTextarea ? importDatesTextarea.value : "");
  if (!rows.length && !errors.length) {
    importDatesPreview.textContent = "Nessun dato incollato.";
    return;
  }
  const commessaMap = new Map(
    (state.commesse || []).map((c) => [`${c.anno}_${c.numero}`, c])
  );
  const toUpdate = [];
  const missing = [];
  const overwrites = [];
  let skipped = 0;
  rows.forEach((row) => {
    const key = row.codeInfo ? row.codeInfo.codice : "";
    const commessa = commessaMap.get(key);
    if (!commessa) {
      missing.push(key);
      return;
    }
    let willUpdate = false;
    if (row.ingresso && commessa.data_ingresso) {
      const existing = formatDate(commessa.data_ingresso);
      const incoming = formatDateInput(row.ingresso);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (ingresso ordine)`);
      }
    }
    if (row.ordine && commessa.data_ordine_telaio) {
      const existing = formatDate(commessa.data_ordine_telaio);
      const incoming = formatDateInput(row.ordine);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (ordine telaio)`);
      }
    }
    if (row.stimato && commessa.data_consegna_telaio_effettiva) {
      const existing = formatDate(commessa.data_consegna_telaio_effettiva);
      const incoming = formatDateInput(row.stimato);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (ordine telaio stimato)`);
      }
    }
    if (row.prelievo && commessa.data_prelievo) {
      const existing = formatDate(commessa.data_prelievo);
      const incoming = formatDateInput(row.prelievo);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (prelievo materiali)`);
      }
    }
    if (row.ordinato != null && commessa.telaio_consegnato != null) {
      const existing = Boolean(commessa.telaio_consegnato);
      const incoming = Boolean(row.ordinato);
      if (existing !== incoming) {
        overwrites.push(`${key} (ordinato)`);
      }
    }
    if (row.consegna && commessa.data_consegna_macchina) {
      const existing = formatDate(commessa.data_consegna_macchina);
      const incoming = formatDateInput(row.consegna);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (consegna macchina)`);
      }
    }
    if (row.ingresso || row.ordine || row.stimato || row.prelievo || row.consegna || row.ordinato != null) {
      willUpdate = true;
      toUpdate.push(key);
    } else {
      skipped += 1;
    }
  });

  const sampleList = (list, max = 20) => {
    if (!list.length) return "";
    if (list.length <= max) return list.join(", ");
    const head = list.slice(0, max).join(", ");
    return `${head} +${list.length - max}`;
  };

  const lines = [];
  lines.push(`Righe riconosciute: ${rows.length}/${total}.`);
  if (errors.length) lines.push(`Errori: ${errors.length} (es: ${errors[0]}).`);
  if (warnings.length) lines.push(`Avvisi: ${warnings.length} (es: ${warnings[0]}).`);
  if (missing.length) lines.push(`Codici non trovati: ${missing.length} (${sampleList(missing)}).`);
  if (toUpdate.length) lines.push(`Commesse da aggiornare: ${toUpdate.length}.`);
  if (skipped) lines.push(`Righe saltate (senza date valide): ${skipped}.`);
  if (overwrites.length) lines.push(`Sovrascritture: ${overwrites.length} (${sampleList(overwrites)}).`);
  const hasIngresso = rows.some((r) => r.ingressoRaw);
  const hasPrelievo = rows.some((r) => r.prelievoRaw);
  const hasOrdinato = rows.some((r) => r.ordinatoRaw || r.ordinatoRaw === "0");
  const headerCols = [
    "Codice",
    ...(hasIngresso ? ["Ingresso ordine"] : []),
    "Ordine telaio target",
    "Ordine telaio stimato",
    ...(hasPrelievo ? ["Prelievo materiali"] : []),
    ...(hasOrdinato ? ["Ordinato"] : []),
    "Consegna macchina",
  ];
  const previewRows = rows.slice(0, 8);
  const tableHtml = previewRows
    .map((row) => {
      const cells = [
        row.codeInfo?.codice || "",
        ...(hasIngresso ? [row.ingressoRaw || ""] : []),
        row.ordineRaw || "",
        row.stimatoRaw || "",
        ...(hasPrelievo ? [row.prelievoRaw || ""] : []),
        ...(hasOrdinato ? [row.ordinatoRaw || ""] : []),
        row.consegnaRaw || "",
      ];
      return `<tr>${cells.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`;
    })
    .join("");
  const headerHtml = headerCols.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  const table =
    previewRows.length > 0
      ? `<div class="import-preview-table-wrap"><table class="import-preview-table"><thead><tr>${headerHtml}</tr></thead><tbody>${tableHtml}</tbody></table></div>`
      : buildRawImportTable(importDatesTextarea ? importDatesTextarea.value : "");
  importDatesPreview.innerHTML = `${escapeHtml(lines.join(" "))}${table}`;
}

function updateImportProductionPreview() {
  if (!importProductionPreview) return;
  const { rows, errors, warnings, total } = parseImportProductionRows(
    importProductionTextarea ? importProductionTextarea.value : ""
  );
  if (!rows.length && !errors.length) {
    importProductionPreview.textContent = "Nessun dato incollato.";
    return;
  }
  const commessaMap = new Map((state.commesse || []).map((c) => [`${c.anno}_${c.numero}`, c]));
  const toUpdate = [];
  const missing = [];
  const overwrites = [];
  let skipped = 0;
  rows.forEach((row) => {
    const key = row.codeInfo ? row.codeInfo.codice : "";
    const commessa = commessaMap.get(key);
    if (!commessa) {
      missing.push(key);
      return;
    }
    if (row.ordine && commessa.data_ordine_telaio) {
      const existing = formatDate(commessa.data_ordine_telaio);
      const incoming = formatDateInput(row.ordine);
      if (existing && incoming && existing !== incoming) overwrites.push(`${key} (ordine telaio)`);
    }
    if (row.prelievo && commessa.data_prelievo) {
      const existing = formatDate(commessa.data_prelievo);
      const incoming = formatDateInput(row.prelievo);
      if (existing && incoming && existing !== incoming) overwrites.push(`${key} (prelievo materiali)`);
    }
    if (row.kit && commessa.data_arrivo_kit_cavi) {
      const existing = formatDate(commessa.data_arrivo_kit_cavi);
      const incoming = formatDateInput(row.kit);
      if (existing && incoming && existing !== incoming) overwrites.push(`${key} (arrivo kit cavi)`);
    }
    if (row.consegna && commessa.data_consegna_macchina) {
      const existing = formatDate(commessa.data_consegna_macchina);
      const incoming = formatDateInput(row.consegna);
      if (existing && incoming && existing !== incoming) overwrites.push(`${key} (consegna macchina)`);
    }
    if (row.ordine || row.prelievo || row.kit || row.consegna) toUpdate.push(key);
    else skipped += 1;
  });

  const sampleList = (list, max = 20) => {
    if (!list.length) return "";
    if (list.length <= max) return list.join(", ");
    const head = list.slice(0, max).join(", ");
    return `${head} +${list.length - max}`;
  };

  const lines = [];
  lines.push(`Righe riconosciute: ${rows.length}/${total}.`);
  if (errors.length) lines.push(`Errori: ${errors.length} (es: ${errors[0]}).`);
  if (warnings.length) lines.push(`Avvisi: ${warnings.length} (es: ${warnings[0]}).`);
  if (missing.length) lines.push(`Codici non trovati: ${missing.length} (${sampleList(missing)}).`);
  if (toUpdate.length) lines.push(`Commesse da aggiornare: ${toUpdate.length}.`);
  if (skipped) lines.push(`Righe saltate (senza date valide): ${skipped}.`);
  if (overwrites.length) lines.push(`Sovrascritture: ${overwrites.length} (${sampleList(overwrites)}).`);
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();

  const headerCols = ["Codice", "Ordine telaio target", "Prelievo materiali", "Arrivo kit cavi", "Consegna macchina"];
  const previewRows = rows.slice(0, 8);
  const tableHtml = previewRows
    .map((row) => {
      const cells = [
        row.codeInfo?.codice || "",
        row.ordineRaw || "",
        row.prelievoRaw || "",
        row.kitRaw || "",
        row.consegnaRaw || "",
      ];
      return `<tr>${cells.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`;
    })
    .join("");
  const headerHtml = headerCols.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  const table =
    previewRows.length > 0
      ? `<div class="import-preview-table-wrap"><table class="import-preview-table"><thead><tr>${headerHtml}</tr></thead><tbody>${tableHtml}</tbody></table></div>`
      : buildRawImportTable(importProductionTextarea ? importProductionTextarea.value : "");
  importProductionPreview.innerHTML = `${escapeHtml(lines.join(" "))}${table}`;
}

async function handleImportDates() {
  if (!state.canWrite) return;
  if (importDatesCommitBtn) {
    importDatesCommitBtn.disabled = true;
    importDatesCommitBtn.textContent = "Importing...";
  }
  if (importDatesPreview) importDatesPreview.textContent = "Import in progress...";
  const { rows, errors, warnings } = parseImportDateRows(importDatesTextarea ? importDatesTextarea.value : "");
  if (!rows.length) {
    setStatus("No valid rows to import.", "error");
    showToast("No valid rows to import.", "error");
    if (importDatesPreview) importDatesPreview.textContent = "No valid rows to import.";
    if (importDatesCommitBtn) {
      importDatesCommitBtn.disabled = false;
      importDatesCommitBtn.textContent = "Import";
    }
    return;
  }
  if (errors.length) {
    setStatus(`Import blocked: fix the errors (e.g. ${errors[0]}).`, "error");
    showToast(`Import blocked: ${errors[0]}`, "error");
    if (importDatesPreview) importDatesPreview.textContent = `Import blocked: ${errors[0]}`;
    if (importDatesCommitBtn) {
      importDatesCommitBtn.disabled = false;
      importDatesCommitBtn.textContent = "Import";
    }
    return;
  }

  const commessaMap = new Map(
    (state.commesse || []).map((c) => [`${c.anno}_${c.numero}`, c])
  );
  const updates = [];
  const skipped = [];
  rows.forEach((row) => {
    const key = row.codeInfo ? row.codeInfo.codice : "";
    const commessa = commessaMap.get(key);
    if (!commessa) {
      skipped.push(`${key} non trovata`);
      return;
    }
    const payload = {};
    if (row.ingresso) payload.data_ingresso = formatDateInput(row.ingresso);
    if (row.ordine) payload.data_ordine_telaio = formatDateInput(row.ordine);
    if (row.stimato) payload.data_consegna_telaio_effettiva = formatDateInput(row.stimato);
    if (row.prelievo) payload.data_prelievo = formatDateInput(row.prelievo);
    if (row.ordinato != null) payload.telaio_consegnato = row.ordinato;
    if (row.consegna) payload.data_consegna_macchina = formatDateInput(row.consegna);
    if (!Object.keys(payload).length) {
      skipped.push(`${key} senza date valide`);
      return;
    }
    updates.push({ id: commessa.id, codice: key, payload });
  });

  if (!updates.length) {
    setStatus("No commesse to update.", "error");
    showToast("No commesse to update.", "error");
    if (importDatesPreview) importDatesPreview.textContent = "No commesse to update.";
    if (importDatesCommitBtn) {
      importDatesCommitBtn.disabled = false;
      importDatesCommitBtn.textContent = "Import";
    }
    return;
  }

  const failures = [];
  for (let i = 0; i < updates.length; i += 1) {
    const upd = updates[i];
    const { error } = await supabase.from("commesse").update(upd.payload).eq("id", upd.id);
    if (error) failures.push(`${upd.codice}: ${error.message}`);
    const pct = Math.round(((i + 1) / updates.length) * 100);
    if (importDatesCommitBtn) importDatesCommitBtn.textContent = `Importing... ${pct}%`;
    if (importDatesPreview) {
      importDatesPreview.textContent = `Importing ${i + 1}/${updates.length} (${pct}%)...`;
    }
  }

  if (failures.length) {
    setStatus(`Partial import: ${failures.length} errors.`, "error");
    showToast(`Partial import: ${failures.length} errors.`, "error");
    console.error("Import dates errors", failures);
  } else {
    const skipInfo = skipped.length ? ` (saltate ${skipped.length})` : "";
    setStatus(`Import completed: ${updates.length} commesse updated${skipInfo}.`, "ok");
    showToast(`Import completed: ${updates.length} commesse updated${skipInfo}.`, "ok");
  }
  if (warnings.length) {
    console.warn("Import dates warnings", warnings);
  }
  if (importDatesPreview) {
    const lines = [];
    lines.push(`Updated: ${updates.length}.`);
    if (skipped.length) lines.push(`Skipped: ${skipped.length}.`);
    if (warnings.length) lines.push(`Warnings: ${warnings.length}.`);
    if (failures.length) lines.push(`Errors: ${failures.length}.`);
    importDatesPreview.textContent = lines.join(" ");
  }
  await loadCommesse();
  if (!failures.length) {
    closeImportDatesModal();
    if (importDatesTextarea) importDatesTextarea.value = "";
  }
  if (importDatesCommitBtn) {
    importDatesCommitBtn.disabled = false;
    importDatesCommitBtn.textContent = "Import";
  }
}

async function handleImportProduction() {
  if (!state.canWrite) return;
  if (importProductionCommitBtn) {
    importProductionCommitBtn.disabled = true;
    importProductionCommitBtn.textContent = "Importing...";
  }
  if (importProductionPreview) importProductionPreview.textContent = "Import in progress...";
  const { rows, errors, warnings } = parseImportProductionRows(
    importProductionTextarea ? importProductionTextarea.value : ""
  );
  if (!rows.length) {
    setStatus("No valid rows to import.", "error");
    showToast("No valid rows to import.", "error");
    if (importProductionPreview) importProductionPreview.textContent = "No valid rows to import.";
    if (importProductionCommitBtn) {
      importProductionCommitBtn.disabled = false;
      importProductionCommitBtn.textContent = "Import";
    }
    return;
  }
  if (errors.length) {
    setStatus(`Import blocked: fix the errors (e.g. ${errors[0]}).`, "error");
    showToast(`Import blocked: ${errors[0]}`, "error");
    if (importProductionPreview) importProductionPreview.textContent = `Import blocked: ${errors[0]}`;
    if (importProductionCommitBtn) {
      importProductionCommitBtn.disabled = false;
      importProductionCommitBtn.textContent = "Import";
    }
    return;
  }

  const commessaMap = new Map((state.commesse || []).map((c) => [`${c.anno}_${c.numero}`, c]));
  const updates = [];
  const skipped = [];
  rows.forEach((row) => {
    const key = row.codeInfo ? row.codeInfo.codice : "";
    const commessa = commessaMap.get(key);
    if (!commessa) {
      skipped.push(`${key} non trovata`);
      return;
    }
    const payload = {};
    if (row.ordine) payload.data_ordine_telaio = formatDateInput(row.ordine);
    if (row.prelievo) payload.data_prelievo = formatDateInput(row.prelievo);
    if (row.kit) payload.data_arrivo_kit_cavi = formatDateInput(row.kit);
    if (row.consegna) payload.data_consegna_macchina = formatDateInput(row.consegna);
    if (!Object.keys(payload).length) {
      skipped.push(`${key} senza date valide`);
      return;
    }
    updates.push({ id: commessa.id, codice: key, payload });
  });

  if (!updates.length) {
    setStatus("No commesse to update.", "error");
    showToast("No commesse to update.", "error");
    if (importProductionPreview) importProductionPreview.textContent = "No commesse to update.";
    if (importProductionCommitBtn) {
      importProductionCommitBtn.disabled = false;
      importProductionCommitBtn.textContent = "Import";
    }
    return;
  }

  const failures = [];
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();
  for (let i = 0; i < updates.length; i += 1) {
    const upd = updates[i];
    let error = null;
    if (role === "planner") {
      const plannerPayload = {
        data_ordine_telaio: upd.payload.data_ordine_telaio || null,
        data_consegna_macchina: upd.payload.data_consegna_macchina || null,
        data_arrivo_kit_cavi: upd.payload.data_arrivo_kit_cavi || null,
        data_prelievo: upd.payload.data_prelievo || null,
      };
      error = await updateCommessaPlannerDates(upd.id, plannerPayload);
    } else {
      const res = await supabase.from("commesse").update(upd.payload).eq("id", upd.id);
      error = res.error || null;
    }
    if (error) failures.push(`${upd.codice}: ${error.message || error}`);
    const pct = Math.round(((i + 1) / updates.length) * 100);
    if (importProductionCommitBtn) importProductionCommitBtn.textContent = `Importing... ${pct}%`;
    if (importProductionPreview) {
      importProductionPreview.textContent = `Importing ${i + 1}/${updates.length} (${pct}%)...`;
    }
  }

  if (failures.length) {
    setStatus(`Partial import: ${failures.length} errors.`, "error");
    showToast(`Partial import: ${failures.length} errors.`, "error");
    console.error("Import production errors", failures);
  } else {
    const skipInfo = skipped.length ? ` (saltate ${skipped.length})` : "";
    setStatus(`Import completed: ${updates.length} commesse updated${skipInfo}.`, "ok");
    showToast(`Import completed: ${updates.length} commesse updated${skipInfo}.`, "ok");
  }
  if (warnings.length) console.warn("Import production warnings", warnings);
  if (importProductionPreview) {
    const lines = [];
    lines.push(`Updated: ${updates.length}.`);
    if (skipped.length) lines.push(`Skipped: ${skipped.length}.`);
    if (warnings.length) lines.push(`Warnings: ${warnings.length}.`);
    if (failures.length) lines.push(`Errors: ${failures.length}.`);
    importProductionPreview.textContent = lines.join(" ");
  }
  await loadCommesse();
  if (!failures.length) {
    closeImportProductionModal();
    if (importProductionTextarea) importProductionTextarea.value = "";
  }
  if (importProductionCommitBtn) {
    importProductionCommitBtn.disabled = false;
    importProductionCommitBtn.textContent = "Import";
  }
}

function normalizeCommessaCode(raw) {
  const text = (raw || "").trim();
  if (!text) return null;
  const match = text.match(/^(\d{4})\s*[_-]?\s*(\d{1,3})$/);
  if (!match) return null;
  const anno = Number(match[1]);
  const numero = Number(match[2]);
  if (!Number.isInteger(anno) || !Number.isInteger(numero) || numero < 1) return null;
  return { anno, numero, codice: `${anno}_${numero}` };
}

function parseDateCell(raw) {
  const text = (raw || "").trim();
  if (!text) return { date: null, reason: "empty" };
  const lowered = text.toLowerCase();
  if (["0", "0.0", "00/00/0000", "00-00-0000", "null", "n/a", "-"].includes(lowered)) {
    return { date: null, reason: "empty" };
  }
  const isInRange = (date) => {
    const y = date.getFullYear();
    return y >= 2000 && y <= 2100;
  };
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
    const [y, m, d] = text.split("-").map((v) => Number(v));
    const date = new Date(y, m - 1, d);
    if (Number.isNaN(date.getTime())) return { date: null, reason: "invalid" };
    if (!isInRange(date)) return { date: null, reason: "range" };
    return { date, reason: "" };
  }
  if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(text)) {
    const parts = text.split(/[\/.\-]/).map((v) => Number(v));
    let [d, m, y] = parts;
    if (y < 100) y += 2000;
    const date = new Date(y, m - 1, d);
    if (Number.isNaN(date.getTime())) return { date: null, reason: "invalid" };
    if (!isInRange(date)) return { date: null, reason: "range" };
    return { date, reason: "" };
  }
  if (/^\d+$/.test(text)) {
    const serial = Number(text);
    if (serial >= 20000 && serial <= 60000) {
      const base = new Date(1899, 11, 30);
      const date = addDays(base, serial);
      if (!isInRange(date)) return { date: null, reason: "range" };
      return { date, reason: "" };
    }
  }
  return { date: null, reason: "invalid" };
}

function parseBooleanCell(raw) {
  const text = (raw || "").trim().toLowerCase();
  if (!text) return { value: null, reason: "empty" };
  if (["vero", "true", "1", "yes", "y", "si"].includes(text)) return { value: true, reason: "" };
  if (["falso", "false", "0", "no", "n"].includes(text)) return { value: false, reason: "" };
  return { value: null, reason: "invalid" };
}

function parseImportDateRows(text) {
  const cleaned = text.trim();
  if (!cleaned) return { rows: [], errors: [], warnings: [], total: 0 };
  const delimiter = detectDelimiter(cleaned);
  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim() !== "");
  let startIndex = 0;
  let headerMap = null;
  if (lines.length) {
    const first = lines[0].toLowerCase();
    if (first.includes("codice") || first.includes("telaio") || first.includes("consegna")) {
      startIndex = 1;
      const headerCols = lines[0]
        .split(delimiter)
        .map((c) => c.trim().replace(/^"(.*)"$/, "$1").toLowerCase());
      const map = {};
      headerCols.forEach((col, index) => {
        if (!col) return;
        if (col.includes("prelievo")) {
          map.prelievo = index;
          return;
        }
        if (col.includes("ingresso")) {
          map.ingresso = index;
          return;
        }
        if (col.includes("stim")) {
          map.stimato = index;
          return;
        }
        if (col.includes("ordinato")) {
          map.ordinato = index;
          return;
        }
        if (col.includes("consegna") || col.includes("delivery")) {
          map.consegna = index;
          return;
        }
        if (col.includes("ordine")) {
          map.ordine = index;
        }
      });
      if (Object.keys(map).length) headerMap = map;
    }
  }

  const rows = [];
  const errors = [];
  const warnings = [];
  for (let i = startIndex; i < lines.length; i++) {
    const cols = lines[i]
      .split(delimiter)
      .map((c) => c.trim().replace(/^"(.*)"$/, "$1"));
    if (cols.length < 1) {
      errors.push(`Riga ${i + 1}: codice commessa mancante`);
      continue;
    }
    const codeInfo = normalizeCommessaCode(cols[0]);
    if (!codeInfo) {
      errors.push(`Riga ${i + 1}: codice commessa non valido`);
      continue;
    }
    let ingressoCell = "";
    let ordineCell = "";
    let stimatoCell = "";
    let prelievoCell = "";
    let ordinatoCell = "";
    let consegnaCell = "";

    if (headerMap) {
      ingressoCell = headerMap.ingresso != null ? cols[headerMap.ingresso] || "" : "";
      ordineCell = headerMap.ordine != null ? cols[headerMap.ordine] || "" : "";
      stimatoCell = headerMap.stimato != null ? cols[headerMap.stimato] || "" : "";
      prelievoCell = headerMap.prelievo != null ? cols[headerMap.prelievo] || "" : "";
      ordinatoCell = headerMap.ordinato != null ? cols[headerMap.ordinato] || "" : "";
      consegnaCell = headerMap.consegna != null ? cols[headerMap.consegna] || "" : "";
    } else {
      const hasIngresso = cols.length >= 6;
      const hasPrelievo = cols.length >= 7;
      const flagIndex = hasIngresso ? (hasPrelievo ? 5 : 4) : 3;
      const flagProbe = parseBooleanCell(cols[flagIndex] || "");
      const hasFlag = flagProbe.reason !== "invalid";

      ingressoCell = hasIngresso ? cols[1] || "" : "";
      ordineCell = hasIngresso ? cols[2] || "" : cols[1] || "";
      stimatoCell = cols.length >= 4 ? (hasIngresso ? cols[3] || "" : cols[2] || "") : "";
      if (hasIngresso && hasPrelievo) {
        prelievoCell = cols[4] || "";
        ordinatoCell = hasFlag ? cols[5] || "" : "";
        consegnaCell = hasFlag ? cols[6] || "" : cols[5] || "";
      } else if (hasIngresso) {
        ordinatoCell = hasFlag ? cols[4] || "" : "";
        consegnaCell = hasFlag ? cols[5] || "" : cols[4] || "";
      } else {
        ordinatoCell = hasFlag ? cols[3] || "" : "";
        consegnaCell = hasFlag
          ? cols[4] || ""
          : cols.length >= 4
          ? cols[3] || ""
          : cols[2] || "";
        if (cols.length >= 6) {
          prelievoCell = cols[5] || "";
        } else if (!hasFlag && cols.length >= 5) {
          prelievoCell = cols[4] || "";
        }
      }
    }
    const ingressoParsed = parseDateCell(ingressoCell);
    const ordineParsed = parseDateCell(ordineCell);
    const stimatoParsed = parseDateCell(stimatoCell);
    const prelievoParsed = parseDateCell(prelievoCell);
    const ordinatoParsed = parseBooleanCell(ordinatoCell);
    const consegnaParsed = parseDateCell(consegnaCell);
    if (ordineCell && ordineParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data ordine telaio non valida`);
    } else if (ordineCell && ordineParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data ordine telaio vuota/zero`);
    } else if (ordineCell && ordineParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data ordine telaio fuori range (2000-2100)`);
    }
    if (ingressoCell && ingressoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data ingresso ordine non valida`);
    } else if (ingressoCell && ingressoParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data ingresso ordine vuota/zero`);
    } else if (ingressoCell && ingressoParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data ingresso ordine fuori range (2000-2100)`);
    }
    if (stimatoCell && stimatoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data ordine telaio stimato non valida`);
    } else if (stimatoCell && stimatoParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data ordine telaio stimato vuota/zero`);
    } else if (stimatoCell && stimatoParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data ordine telaio stimato fuori range (2000-2100)`);
    }
    if (prelievoCell && prelievoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali non valida`);
    } else if (prelievoCell && prelievoParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali vuota/zero`);
    } else if (prelievoCell && prelievoParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali fuori range (2000-2100)`);
    }
    if (ordinatoCell && ordinatoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: valore ordinato non valido (usa VERO/FALSO)`);
    }
    if (consegnaCell && consegnaParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data consegna macchina non valida`);
    } else if (consegnaCell && consegnaParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data consegna macchina vuota/zero`);
    } else if (consegnaCell && consegnaParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data consegna macchina fuori range (2000-2100)`);
    }
    rows.push({
      rawIndex: i + 1,
      codeInfo,
      ingresso: ingressoParsed.date,
      ordine: ordineParsed.date,
      stimato: stimatoParsed.date,
      prelievo: prelievoParsed.date,
      ordinato: ordinatoParsed.value,
      consegna: consegnaParsed.date,
      ingressoReason: ingressoParsed.reason,
      ordineReason: ordineParsed.reason,
      stimatoReason: stimatoParsed.reason,
      prelievoReason: prelievoParsed.reason,
      ordinatoReason: ordinatoParsed.reason,
      consegnaReason: consegnaParsed.reason,
      ingressoRaw: ingressoCell,
      ordineRaw: ordineCell,
      stimatoRaw: stimatoCell,
      prelievoRaw: prelievoCell,
      ordinatoRaw: ordinatoCell,
      consegnaRaw: consegnaCell,
    });
  }
  return { rows, errors, warnings, total: lines.length };
}

function parseImportProductionRows(text) {
  const cleaned = text.trim();
  if (!cleaned) return { rows: [], errors: [], warnings: [], total: 0 };
  const delimiter = detectDelimiter(cleaned);
  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim() !== "");
  let startIndex = 0;
  let headerMap = null;
  if (lines.length) {
    const first = lines[0].toLowerCase();
    if (first.includes("codice") || first.includes("kit") || first.includes("prelievo")) {
      startIndex = 1;
      const headerCols = lines[0]
        .split(delimiter)
        .map((c) => c.trim().replace(/^"(.*)"$/, "$1").toLowerCase());
      const map = {};
      headerCols.forEach((col, index) => {
        if (!col) return;
        if (col.includes("prelievo")) {
          map.prelievo = index;
          return;
        }
        if (col.includes("kit")) {
          map.kit = index;
          return;
        }
        if (col.includes("consegna") || col.includes("delivery")) {
          map.consegna = index;
          return;
        }
        if (col.includes("ordine")) {
          map.ordine = index;
        }
      });
      if (Object.keys(map).length) headerMap = map;
    }
  }

  const rows = [];
  const errors = [];
  const warnings = [];
  for (let i = startIndex; i < lines.length; i++) {
    const cols = lines[i]
      .split(delimiter)
      .map((c) => c.trim().replace(/^"(.*)"$/, "$1"));
    if (cols.length < 1) {
      errors.push(`Riga ${i + 1}: codice commessa mancante`);
      continue;
    }
    const codeInfo = normalizeCommessaCode(cols[0]);
    if (!codeInfo) {
      errors.push(`Riga ${i + 1}: codice commessa non valido`);
      continue;
    }
    let ordineCell = "";
    let prelievoCell = "";
    let kitCell = "";
    let consegnaCell = "";
    if (headerMap) {
      ordineCell = headerMap.ordine != null ? cols[headerMap.ordine] || "" : "";
      prelievoCell = headerMap.prelievo != null ? cols[headerMap.prelievo] || "" : "";
      kitCell = headerMap.kit != null ? cols[headerMap.kit] || "" : "";
      consegnaCell = headerMap.consegna != null ? cols[headerMap.consegna] || "" : "";
    } else {
      ordineCell = cols[1] || "";
      prelievoCell = cols[2] || "";
      kitCell = cols[3] || "";
      consegnaCell = cols[4] || "";
    }
    const ordineParsed = parseDateCell(ordineCell);
    const prelievoParsed = parseDateCell(prelievoCell);
    const kitParsed = parseDateCell(kitCell);
    const consegnaParsed = parseDateCell(consegnaCell);
    if (ordineCell && ordineParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data ordine telaio non valida`);
    } else if (ordineCell && ordineParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data ordine telaio vuota/zero`);
    } else if (ordineCell && ordineParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data ordine telaio fuori range (2000-2100)`);
    }
    if (prelievoCell && prelievoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali non valida`);
    } else if (prelievoCell && prelievoParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali vuota/zero`);
    } else if (prelievoCell && prelievoParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data prelievo materiali fuori range (2000-2100)`);
    }
    if (kitCell && kitParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data arrivo kit cavi non valida`);
    } else if (kitCell && kitParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data arrivo kit cavi vuota/zero`);
    } else if (kitCell && kitParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data arrivo kit cavi fuori range (2000-2100)`);
    }
    if (consegnaCell && consegnaParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data consegna macchina non valida`);
    } else if (consegnaCell && consegnaParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data consegna macchina vuota/zero`);
    } else if (consegnaCell && consegnaParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data consegna macchina fuori range (2000-2100)`);
    }
    rows.push({
      rawIndex: i + 1,
      codeInfo,
      ordine: ordineParsed.date,
      prelievo: prelievoParsed.date,
      kit: kitParsed.date,
      consegna: consegnaParsed.date,
      ordineReason: ordineParsed.reason,
      prelievoReason: prelievoParsed.reason,
      kitReason: kitParsed.reason,
      consegnaReason: consegnaParsed.reason,
      ordineRaw: ordineCell,
      prelievoRaw: prelievoCell,
      kitRaw: kitCell,
      consegnaRaw: consegnaCell,
    });
  }
  return { rows, errors, warnings, total: lines.length };
}

function formatDateDMY(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

function formatDateHumanNoYear(date) {
  return date.toLocaleDateString("it-IT", {
    weekday: "long",
    day: "numeric",
    month: "long",
  });
}

function normalizeBusinessDay(date, direction = 1) {
  let d = new Date(date);
  while (isWeekend(d)) {
    d = addDays(d, direction >= 0 ? 1 : -1);
  }
  return d;
}

function getIsoWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

function getIsoWeekYear(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  return d.getUTCFullYear();
}

function toWeekInputValue(date) {
  const year = getIsoWeekYear(date);
  const week = String(getIsoWeekNumber(date)).padStart(2, "0");
  return `${year}-W${week}`;
}

function weekInputToDate(value) {
  const raw = (value || "").trim();
  const match = raw.match(/^(\d{4})-W(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const week = Number(match[2]);
  if (!year || !week) return null;
  const jan4 = new Date(year, 0, 4);
  const day = jan4.getDay() || 7;
  const monday = new Date(jan4);
  monday.setDate(jan4.getDate() - (day - 1) + (week - 1) * 7);
  monday.setHours(0, 0, 0, 0);
  return monday;
}

function startOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay() || 7;
  if (day !== 1) d.setDate(d.getDate() - (day - 1));
  d.setHours(0, 0, 0, 0);
  return d;
}

function weekdayDates(start) {
  const days = [];
  for (let i = 0; i < 7; i++) {
    const d = addDays(start, i);
    const day = d.getDay();
    if (day !== 0 && day !== 6) days.push(d);
  }
  return days;
}

function weekDaysForView(baseDate, view) {
  const start = startOfWeek(baseDate);
  if (view === "two") {
    return [...weekdayDates(start), ...weekdayDates(addDays(start, 7))];
  }
  if (view === "three") {
    return [
      ...weekdayDates(start),
      ...weekdayDates(addDays(start, 7)),
      ...weekdayDates(addDays(start, 14)),
    ];
  }
  if (view === "six") {
    const days = [];
    for (let i = 0; i < 6; i += 1) {
      days.push(...weekdayDates(addDays(start, i * 7)));
    }
    return days;
  }
  return weekdayDates(start);
}

function endOfDay(date) {
  const d = new Date(date);
  d.setHours(23, 59, 59, 999);
  return d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function dateFromKey(dayKey) {
  return new Date(`${dayKey}T00:00:00`);
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function addBusinessDays(date, days) {
  if (!days) return new Date(date);
  const step = days > 0 ? 1 : -1;
  let remaining = Math.abs(days);
  const d = new Date(date);
  while (remaining > 0) {
    d.setDate(d.getDate() + step);
    if (!isWeekend(d)) remaining -= 1;
  }
  return d;
}

function businessDayDiff(start, end) {
  const sNum = dayNumberUTC(start);
  const eNum = dayNumberUTC(end);
  if (sNum === eNum) return 0;

  const forward = sNum < eNum;
  const fromNum = forward ? sNum : eNum;
  const toNum = forward ? eNum : sNum;
  const totalDays = toNum - fromNum;
  const fullWeeks = Math.floor(totalDays / 7);
  let businessDays = fullWeeks * 5;
  const remaining = totalDays % 7;
  const fromDate = new Date(fromNum * 86400000);
  const startDay = fromDate.getUTCDay();
  for (let i = 1; i <= remaining; i += 1) {
    const day = (startDay + i) % 7;
    if (day !== 0 && day !== 6) businessDays += 1;
  }
  return forward ? businessDays : -businessDays;
}

function daysBetween(start, end) {
  const s = dayNumberUTC(start);
  const e = dayNumberUTC(end);
  return e - s;
}

function getBusinessDiffInfo(today, targetDate) {
  if (!targetDate || Number.isNaN(targetDate.getTime())) {
    return { label: "n/d", cls: "" };
  }
  const y = targetDate.getFullYear();
  if (y < 2000 || y > 2100) {
    return { label: "n/d", cls: "" };
  }
  const diff = businessDayDiff(today, targetDate);
  if (diff > 0) {
    return { label: `mancano ${diff} gg lav.`, cls: "is-ok" };
  }
  if (diff < 0) {
    return { label: `scaduta da ${Math.abs(diff)} gg lav.`, cls: "is-late" };
  }
  return { label: "oggi", cls: "is-today" };
}

function showMilestonePicker({ commessaId, field, date, anchorRect }) {
  if (!milestonePicker || !milestoneDays || !milestoneLabel) return;
  if (!canEditGanttMilestone(field)) {
    if (field === "ordine" || field === "consegna") {
      setStatus("Solo admin e planner possono modificare produzione e consegna macchina.", "error");
      showToast("Solo admin e planner possono modificare produzione e consegna macchina.", "error");
    } else if (field === "kit_cavi") {
      setStatus("Solo admin, planner e responsabili possono modificare arrivo kit cavi.", "error");
      showToast("Solo admin, planner e responsabili possono modificare arrivo kit cavi.", "error");
    } else if (field === "ordine_pianificato") {
      setStatus("Solo admin e responsabili possono modificare i T pianificati.", "error");
      showToast("Solo admin e responsabili possono modificare i T pianificati.", "error");
    } else {
      setStatus("Non hai permessi per modificare questo elemento.", "error");
      showToast("Non hai permessi per modificare questo elemento.", "error");
    }
    return;
  }
  milestonePickerState.commessaId = commessaId;
  milestonePickerState.field = field;
  milestonePickerState.selected = date ? startOfDay(date) : null;
  milestonePickerState.compareTarget = null;
  milestonePickerState.tone = "";
  if (field === "ordine_pianificato" && commessaId) {
    const commessa = state.commesse.find((c) => c.id === commessaId);
    const targetRaw = commessa?.data_ordine_telaio || "";
    const target = targetRaw ? new Date(targetRaw) : null;
    if (target && !Number.isNaN(target.getTime())) {
      milestonePickerState.compareTarget = startOfDay(target);
    }
    if (milestonePickerState.selected && milestonePickerState.compareTarget) {
      milestonePickerState.tone =
        milestonePickerState.selected.getTime() > milestonePickerState.compareTarget.getTime()
          ? "is-telaio-late"
          : "is-telaio-on-time";
    }
  }
  const base = date ? new Date(date.getFullYear(), date.getMonth(), 1) : new Date();
  let clamped = new Date(base);
  if (clamped.getFullYear() < MILESTONE_MIN_YEAR) clamped = new Date(MILESTONE_MIN_YEAR, 0, 1);
  if (clamped.getFullYear() > MILESTONE_MAX_YEAR) clamped = new Date(MILESTONE_MAX_YEAR, 11, 1);
  milestonePickerState.month = clamped;
  renderMilestonePicker();
  const padding = 8;
  let left = anchorRect.left + anchorRect.width / 2;
  let top = anchorRect.bottom + 8;
  milestonePicker.classList.remove("hidden");
  milestonePicker.style.visibility = "hidden";
  requestAnimationFrame(() => {
    const rect = milestonePicker.getBoundingClientRect();
    left = Math.min(Math.max(padding, left - rect.width / 2), window.innerWidth - rect.width - padding);
    if (top + rect.height > window.innerHeight - padding) {
      top = anchorRect.top - rect.height - 8;
    }
    top = Math.max(padding, top);
    milestonePicker.style.left = `${left}px`;
    milestonePicker.style.top = `${top}px`;
    milestonePicker.style.visibility = "visible";
  });
}

function closeMilestonePicker() {
  if (!milestonePicker) return;
  milestonePicker.classList.add("hidden");
  milestonePickerState.commessaId = null;
  milestonePickerState.field = null;
  milestonePickerState.selected = null;
  milestonePickerState.compareTarget = null;
  milestonePickerState.tone = "";
}

function renderMilestonePicker() {
  if (!milestonePicker || !milestoneDays || !milestoneLabel) return;
  let month = milestonePickerState.month;
  if (month.getFullYear() < MILESTONE_MIN_YEAR) month = new Date(MILESTONE_MIN_YEAR, 0, 1);
  if (month.getFullYear() > MILESTONE_MAX_YEAR) month = new Date(MILESTONE_MAX_YEAR, 11, 1);
  milestonePickerState.month = month;
  if (milestonePicker) {
    milestonePicker.classList.remove("is-telaio-on-time", "is-telaio-late", "is-telaio-target");
    if (
      milestonePickerState.field === "ordine" ||
      milestonePickerState.field === "prelievo" ||
      milestonePickerState.field === "kit_cavi"
    ) {
      milestonePicker.classList.add("is-telaio-target");
    } else if (milestonePickerState.tone) {
      milestonePicker.classList.add(milestonePickerState.tone);
    }
  }
  if (milestoneTitle) {
    if (milestonePickerState.field === "ordine") {
      milestoneTitle.textContent = "Production target date";
    } else if (milestonePickerState.field === "kit_cavi") {
      milestoneTitle.textContent = "Kit cables arrival date";
    } else if (milestonePickerState.field === "prelievo") {
      milestoneTitle.textContent = "Materials pickup date";
    } else if (milestonePickerState.field === "ordine_pianificato") {
      milestoneTitle.textContent = "Tech planned date";
    } else if (milestonePickerState.field === "consegna") {
      milestoneTitle.textContent = "Delivery date";
    } else {
      milestoneTitle.textContent = "Date picker";
    }
  }
  const year = month.getFullYear();
  const monthIndex = month.getMonth();
  milestoneLabel.textContent = month.toLocaleDateString("it-IT", { month: "long", year: "numeric" });
  if (milestonePrevBtn) {
    milestonePrevBtn.disabled = year <= MILESTONE_MIN_YEAR && monthIndex === 0;
  }
  if (milestoneNextBtn) {
    milestoneNextBtn.disabled = year >= MILESTONE_MAX_YEAR && monthIndex === 11;
  }
  milestoneDays.innerHTML = "";
  const commessa = milestonePickerState.commessaId
    ? state.commesse.find((c) => c.id === milestonePickerState.commessaId)
    : null;
  const isOrdered = Boolean(commessa?.telaio_consegnato);
  const transport = commessa?.trasporto_consegna || "van";
  const first = new Date(year, monthIndex, 1);
  const last = new Date(year, monthIndex + 1, 0);
  const startOffset = (first.getDay() + 6) % 7; // monday=0
  for (let i = 0; i < startOffset; i += 1) {
    const empty = document.createElement("div");
    empty.className = "milestone-day empty";
    milestoneDays.appendChild(empty);
  }
  const today = startOfDay(new Date());
  for (let d = 1; d <= last.getDate(); d += 1) {
    const date = new Date(year, monthIndex, d);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "milestone-day";
    btn.textContent = String(d);
    if (milestonePickerState.selected && date.getTime() === milestonePickerState.selected.getTime()) {
      btn.classList.add("selected");
    }
    if (date.getTime() === today.getTime()) {
      btn.classList.add("today");
    }
    const day = date.getDay();
    if (day === 0 || day === 6) btn.classList.add("weekend");
    btn.addEventListener("click", async () => {
      if (milestonePickerState.saving) return;
      if (!state.canWrite) {
        setStatus("You don't have permission to edit this commessa.", "error");
        showToast("You don't have permission to edit this commessa.", "error");
        return;
      }
      if (milestonePickerState.field === "ordine_pianificato" && isOrdered) {
        const msg = "Frame marked as ordered. Unmark it to edit the planned date.";
        setStatus(msg, "error");
        showToast(msg, "error");
        return;
      }
      const commessaId = milestonePickerState.commessaId;
      const field = milestonePickerState.field;
      if (!commessaId || !field) return;
      milestonePickerState.saving = true;
      const payload = {};
      if (field === "ordine") payload.data_ordine_telaio = formatDateInput(date);
      if (field === "prelievo") payload.data_prelievo = formatDateInput(date);
      if (field === "kit_cavi") payload.data_arrivo_kit_cavi = formatDateInput(date);
      if (field === "consegna") payload.data_consegna_macchina = formatDateInput(date);
      if (field === "ordine_pianificato") payload.data_consegna_telaio_effettiva = formatDateInput(date);
      const role = String(state.profile?.ruolo || "").trim().toLowerCase();
      let error = null;
      if (
        role === "planner" &&
        (field === "ordine" || field === "consegna" || field === "kit_cavi" || field === "prelievo")
      ) {
        error = await updateCommessaPlannerDates(commessaId, payload);
      } else {
        const res = await supabase.from("commesse").update(payload).eq("id", commessaId);
        error = res.error || null;
        if (!error) updateCommessaInState(commessaId, payload);
      }
      milestonePickerState.saving = false;
      if (error) {
        setStatus(`Update error: ${error.message}`, "error");
        showToast(`Update error: ${error.message}`, "error");
        return;
      }
      if (field === "ordine" && state.selected?.id === commessaId && d.data_ordine_telaio) {
        d.data_ordine_telaio.value = payload.data_ordine_telaio || "";
      }
      if (field === "prelievo" && state.selected?.id === commessaId && d.data_prelievo_materiali) {
        d.data_prelievo_materiali.value = payload.data_prelievo || "";
      }
      if (
        field === "ordine_pianificato" &&
        state.selected?.id === commessaId &&
        d.data_consegna_telaio_effettiva
      ) {
        d.data_consegna_telaio_effettiva.value = payload.data_consegna_telaio_effettiva || "";
      }
      applyFilters();
      renderMatrixReport();
      setStatus("Date updated.", "ok");
      closeMilestonePicker();
      await loadCommesse();
    });
    milestoneDays.appendChild(btn);
  }

  if (milestoneOrderBtn) {
    if (milestonePickerState.field !== "ordine_pianificato") {
      milestoneOrderBtn.classList.add("hidden");
      milestoneOrderBtn.onclick = null;
    } else {
      milestoneOrderBtn.classList.remove("hidden");
      milestoneOrderBtn.textContent = isOrdered ? "Ordered" : "To be ordered";
      milestoneOrderBtn.classList.toggle("is-ordered", isOrdered);
      milestoneOrderBtn.onclick = async () => {
        if (!state.canWrite) {
          setStatus("You don't have permission to edit this commessa.", "error");
          showToast("You don't have permission to edit this commessa.", "error");
          return;
        }
        const commessaId = milestonePickerState.commessaId;
        if (!commessaId) return;
        if (!isOrdered) {
          const hasPlanned =
            Boolean(milestonePickerState.selected) ||
            Boolean(commessa?.data_consegna_telaio_effettiva);
          if (!hasPlanned) {
            const msg = "Add a planned frame order date before marking as ordered.";
            setStatus(msg, "error");
            showToast(msg, "error");
            return;
          }
        }
        const nextState = !isOrdered;
        const { error } = await supabase
          .from("commesse")
          .update({ telaio_consegnato: nextState })
          .eq("id", commessaId);
        if (error) {
          setStatus(`Update error: ${error.message}`, "error");
          return;
        }
        updateCommessaInState(commessaId, { telaio_consegnato: nextState });
        if (state.selected?.id === commessaId && d.telaio_ordinato) {
          setTelaioOrdinatoButton(
            d.telaio_ordinato,
            nextState,
            false,
            d.data_consegna_telaio_effettiva || null
          );
        }
        applyFilters();
        renderMatrixReport();
        setStatus(nextState ? "Frame marked as ordered." : "Frame unmarked as ordered.", "ok");
        renderMilestonePicker();
        await loadCommesse();
      };
    }
  }

  if (milestoneTransportVanBtn && milestoneTransportShipBtn) {
    if (milestonePickerState.field !== "consegna") {
      milestoneTransportVanBtn.classList.add("hidden");
      milestoneTransportShipBtn.classList.add("hidden");
    } else {
      milestoneTransportVanBtn.classList.remove("hidden");
      milestoneTransportShipBtn.classList.remove("hidden");
      milestoneTransportVanBtn.classList.toggle("is-active", transport === "van");
      milestoneTransportShipBtn.classList.toggle("is-active", transport === "ship");
      const setTransport = async (next) => {
        if (!state.canWrite) {
          setStatus("You don't have permission to edit this commessa.", "error");
          showToast("You don't have permission to edit this commessa.", "error");
          return;
        }
        const commessaId = milestonePickerState.commessaId;
        if (!commessaId) return;
        const { error } = await supabase
          .from("commesse")
          .update({ trasporto_consegna: next })
          .eq("id", commessaId);
        if (error) {
          setStatus(`Update error: ${error.message}`, "error");
          showToast(`Update error: ${error.message}`, "error");
          return;
        }
        updateCommessaInState(commessaId, { trasporto_consegna: next });
        renderMatrixReport();
        setStatus(`Transport set to ${next}.`, "ok");
        renderMilestonePicker();
      };
      milestoneTransportVanBtn.onclick = () => setTransport("van");
      milestoneTransportShipBtn.onclick = () => setTransport("ship");
    }
  }
}

function showMilestoneDragLabel(date, anchorX, anchorY, tone = "") {
  if (!milestoneDragLabel || !date) return;
  milestoneDragLabel.textContent = formatDateHumanNoYear(date);
  milestoneDragLabel.classList.remove("is-telaio-on-time", "is-telaio-late");
  if (tone) milestoneDragLabel.classList.add(tone);
  milestoneDragLabel.classList.remove("hidden");
  milestoneDragLabel.style.visibility = "hidden";
  requestAnimationFrame(() => {
    const rect = milestoneDragLabel.getBoundingClientRect();
    const padding = 8;
    let left = anchorX - rect.width / 2;
    let top = anchorY - rect.height - 10;
    left = Math.min(Math.max(padding, left), window.innerWidth - rect.width - padding);
    if (top < padding) top = anchorY + 10;
    milestoneDragLabel.style.left = `${left}px`;
    milestoneDragLabel.style.top = `${top}px`;
    milestoneDragLabel.style.visibility = "visible";
  });
}

function hideMilestoneDragLabel() {
  if (!milestoneDragLabel) return;
  milestoneDragLabel.classList.add("hidden");
}

function beginMilestoneDrag(e, { commessaId, field, date, track, line, marker, compareTarget }) {
  if (!track) return;
  if (!state.canWrite) {
    setStatus("You don't have permission to edit this commessa.", "error");
    showToast("You don't have permission to edit this commessa.", "error");
    return;
  }
  if (!canEditGanttMilestone(field)) {
    if (field === "ordine" || field === "consegna") {
      setStatus("Solo admin e planner possono modificare produzione e consegna macchina.", "error");
      showToast("Solo admin e planner possono modificare produzione e consegna macchina.", "error");
    } else if (field === "kit_cavi") {
      setStatus("Solo admin, planner e responsabili possono modificare arrivo kit cavi.", "error");
      showToast("Solo admin, planner e responsabili possono modificare arrivo kit cavi.", "error");
    } else if (field === "ordine_pianificato") {
      setStatus("Solo admin e responsabili possono modificare i T pianificati.", "error");
      showToast("Solo admin e responsabili possono modificare i T pianificati.", "error");
    } else {
      setStatus("Non hai permessi per modificare questo elemento.", "error");
      showToast("Non hai permessi per modificare questo elemento.", "error");
    }
    return;
  }
  closeMilestonePicker();
  const rect = track.getBoundingClientRect();
  const totalDays = Number(track.dataset.totalDays || "0");
  const rangeStartRaw = track.dataset.rangeStart;
  if (!rect.width || !totalDays || !rangeStartRaw) return;
  milestoneDrag = {
    commessaId,
    field,
    date,
    track,
    line,
    marker,
    compareTarget,
    rect,
    totalDays,
    rangeStart: new Date(`${rangeStartRaw}T00:00:00`),
    startX: e.clientX,
    moved: false,
  };
  marker.classList.add("dragging");
  const markerRect = marker.getBoundingClientRect();
  const tone =
    field === "ordine_pianificato" && compareTarget && date
      ? date.getTime() > compareTarget.getTime()
        ? "is-telaio-late"
        : "is-telaio-on-time"
      : "";
  showMilestoneDragLabel(date, markerRect.left + markerRect.width / 2, markerRect.top, tone);
  document.addEventListener("mousemove", handleMilestoneDragMove);
  document.addEventListener("mouseup", handleMilestoneDragEnd, { once: true });
}

function handleMilestoneDragMove(e) {
  if (!milestoneDrag) return;
  const dx = Math.abs(e.clientX - milestoneDrag.startX);
  if (dx > 3) milestoneDrag.moved = true;
  const info = getMilestoneDragInfo(milestoneDrag, e.clientX);
  if (!info) return;
  milestoneDrag.previewDate = info.date;
  milestoneDrag.marker.style.left = `${info.visualLeftPct}%`;
  if (milestoneDrag.line) milestoneDrag.line.style.left = `${info.visualLeftPct}%`;
  const markerRect = milestoneDrag.marker.getBoundingClientRect();
  const tone =
    milestoneDrag.field === "ordine_pianificato" && milestoneDrag.compareTarget && info.date
      ? info.date.getTime() > milestoneDrag.compareTarget.getTime()
        ? "is-telaio-late"
        : "is-telaio-on-time"
      : "";
  showMilestoneDragLabel(info.date, markerRect.left + markerRect.width / 2, markerRect.top, tone);
}

async function handleMilestoneDragEnd(e) {
  if (!milestoneDrag) return;
  document.removeEventListener("mousemove", handleMilestoneDragMove);
  const drag = milestoneDrag;
  milestoneDrag = null;
  drag.marker.classList.remove("dragging");
  hideMilestoneDragLabel();
  if (!drag.moved) {
    return;
  }
  milestoneSuppressClickUntil = Date.now() + 300;
  const info = getMilestoneDragInfo(drag, e.clientX);
  const newDate = info ? info.date : drag.previewDate;
  if (info) {
    drag.marker.style.left = `${info.leftPct}%`;
    if (drag.line) drag.line.style.left = `${info.leftPct}%`;
  }
  if (!newDate) return;
  const payload = {};
  if (drag.field === "ordine") payload.data_ordine_telaio = formatDateInput(newDate);
  if (drag.field === "kit_cavi") payload.data_arrivo_kit_cavi = formatDateInput(newDate);
  if (drag.field === "prelievo") payload.data_prelievo = formatDateInput(newDate);
  if (drag.field === "consegna") payload.data_consegna_macchina = formatDateInput(newDate);
  if (drag.field === "ordine_pianificato") payload.data_consegna_telaio_effettiva = formatDateInput(newDate);
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();
  if (
    role === "planner" &&
    (drag.field === "ordine" || drag.field === "consegna" || drag.field === "kit_cavi" || drag.field === "prelievo")
  ) {
    const error = await updateCommessaPlannerDates(drag.commessaId, payload);
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      showToast(`Update error: ${error.message}`, "error");
      return;
    }
  } else {
    const { error } = await supabase.from("commesse").update(payload).eq("id", drag.commessaId);
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      showToast(`Update error: ${error.message}`, "error");
      return;
    }
    updateCommessaInState(drag.commessaId, payload);
  }
  if (drag.field === "ordine" && state.selected?.id === drag.commessaId && d.data_ordine_telaio) {
    d.data_ordine_telaio.value = payload.data_ordine_telaio || "";
  }
  if (drag.field === "kit_cavi" && state.selected?.id === drag.commessaId && d.data_arrivo_kit_cavi) {
    d.data_arrivo_kit_cavi.value = payload.data_arrivo_kit_cavi || "";
  }
  if (drag.field === "prelievo" && state.selected?.id === drag.commessaId && d.data_prelievo_materiali) {
    d.data_prelievo_materiali.value = payload.data_prelievo || "";
  }
  if (
    drag.field === "ordine_pianificato" &&
    state.selected?.id === drag.commessaId &&
    d.data_consegna_telaio_effettiva
  ) {
    d.data_consegna_telaio_effettiva.value = payload.data_consegna_telaio_effettiva || "";
  }
  applyFilters();
  renderMatrixReport();
  setStatus("Date updated.", "ok");
  await loadCommesse();
}

function getMilestoneDragInfo(drag, clientX) {
  if (!drag || !drag.rect || !drag.totalDays) return null;
  const x = Math.min(Math.max(clientX, drag.rect.left), drag.rect.right);
  const percent = (x - drag.rect.left) / drag.rect.width;
  const rawIndexFloat = percent * drag.totalDays;
  const maxIndex = drag.totalDays - 1;
  const snapIndex = Math.max(0, Math.min(maxIndex, Math.round(rawIndexFloat - 0.5)));
  const snapCenter = snapIndex + 0.5;
  const distance = Math.abs(rawIndexFloat - snapCenter);
  const snapThreshold = 0.22;
  const rawLeftPct = (rawIndexFloat / drag.totalDays) * 100;
  const snapLeftPct = (snapCenter / drag.totalDays) * 100;
  const visualLeftPct =
    distance <= snapThreshold
      ? snapLeftPct
      : rawLeftPct * 0.7 + snapLeftPct * 0.3;
  const date = addBusinessDays(drag.rangeStart, snapIndex);
  const leftPct = snapLeftPct;
  return { index: snapIndex, date, leftPct, visualLeftPct };
}

function getMatrixRange() {
  const start = startOfWeek(matrixState.date);
  const end = matrixState.view === "six"
    ? endOfDay(addDays(start, 41))
    : matrixState.view === "three"
    ? endOfDay(addDays(start, 20))
    : matrixState.view === "two"
    ? endOfDay(addDays(start, 13))
    : endOfDay(addDays(start, 6));
  return { start, end };
}

function overlapsRange(attivita, start, end) {
  const aStart = new Date(attivita.data_inizio);
  const aEnd = new Date(attivita.data_fine);
  return aStart <= end && aEnd >= start;
}

function getInitials(name) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  const first = parts[0][0] || "";
  const second = parts.length > 1 ? parts[1][0] || "" : "";
  return (first + second).toUpperCase();
}

const PHASE_DEPENDENTS = {
  preliminare: ["3d", "telaio", "carenatura", "costruttivi 1-5", "consegna totale"],
  "3d": ["telaio", "carenatura", "costruttivi 1-5", "consegna totale"],
};

function normalizePhaseKey(title) {
  return (title || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function isPhaseKey(title) {
  return Object.prototype.hasOwnProperty.call(PHASE_DEPENDENTS, normalizePhaseKey(title));
}

function getPhaseDependents(title) {
  return PHASE_DEPENDENTS[normalizePhaseKey(title)] || [];
}

const ACTIVITY_DEPT_LIMITS = {
  "progettazione termodinamica": ["TERMODINAMICI"],
  "progettazione elettrica": ["ELETTRICI"],
  "ordine kit cavi": ["ELETTRICI"],
};

const DEPT_ACTIVITY_ALLOWLISTS = {
  TERMODINAMICI: ["Progettazione termodinamica", "PED E TARGHETTE", "SEGNATURA ELETTRICA", "OP ORDINI"],
  ELETTRICI: [
    "Progettazione elettrica",
    "Ordine KIT cavi",
    "PED E TARGHETTE",
    "SEGNATURA ELETTRICA",
    "OP ORDINI",
  ],
};

const LAVENDER_ACTIVITIES = new Set([
  "altro",
  "ped e targhette",
  "segnatura elettrica",
  "op ordini",
]);

function normalizeDeptKey(name) {
  const base = (name || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (!base) return "";
  const compact = base.replace(/[^A-Z]/g, "");
  if (compact.includes("TERMO")) return "TERMODINAMICI";
  if (compact.includes("ELETT")) return "ELETTRICI";
  if (compact.includes("CAD")) return "CAD";
  return base;
}

function getRisorsaDeptName(risorsaId) {
  if (!risorsaId) return "";
  const risorsa = (state.risorse || []).find((r) => String(r.id) === String(risorsaId));
  const reparto = (state.reparti || []).find((rep) => String(rep.id) === String(risorsa?.reparto_id));
  return reparto ? reparto.nome : "";
}

function isActivityAllowedForRisorsa(title, risorsaId) {
  if (isAssenteTitle(title)) return true;
  const key = normalizePhaseKey(title);
  if (key === "altro") return true;
  const dept = normalizeDeptKey(getRisorsaDeptName(risorsaId));
  const allowlist = DEPT_ACTIVITY_ALLOWLISTS[dept];
  if (allowlist) {
    return allowlist.map(normalizePhaseKey).includes(key);
  }
  const allowedDepts = ACTIVITY_DEPT_LIMITS[key];
  if (!allowedDepts) return true;
  if (!dept) return false;
  return allowedDepts.map(normalizeDeptKey).includes(dept);
}

function isLavenderActivity(title) {
  return LAVENDER_ACTIVITIES.has(normalizePhaseKey(title));
}

function isActivityAllowedForRepartoName(title, repartoName) {
  if (isAssenteTitle(title)) return false;
  const key = normalizePhaseKey(title);
  if (key === "altro") return false;
  const dept = normalizeDeptKey(repartoName || "");
  const allowlist = DEPT_ACTIVITY_ALLOWLISTS[dept];
  if (allowlist) {
    return allowlist.map(normalizePhaseKey).includes(key);
  }
  const allowedDepts = ACTIVITY_DEPT_LIMITS[key];
  if (!allowedDepts) return true;
  if (!dept) return false;
  return allowedDepts.map(normalizeDeptKey).includes(dept);
}

function getTodoActivityCatalog() {
  const options = getMatrixAttivitaOptions();
  const filtered = options.filter(
    (name) => name && !isAssenteTitle(name) && !isLavenderActivity(name) && normalizePhaseKey(name) !== "altro"
  );
  const seen = new Set();
  const unique = [];
  filtered.forEach((name) => {
    const key = normalizePhaseKey(name);
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(name);
  });
  return unique;
}

function getTodoActivityGroups() {
  const catalog = getTodoActivityCatalog();
  const reparti = (state.reparti || [])
    .slice()
    .filter((r) => {
      const key = normalizeDeptKey(r.nome || "");
      return key && !key.includes("TUTTI");
    })
    .sort((a, b) => (a.nome || "").localeCompare(b.nome || ""));
  const groups = [];
  reparti.forEach((r) => {
    const acts = catalog.filter((name) => isActivityAllowedForRepartoName(name, r.nome));
    if (!acts.length) return;
    groups.push({ reparto: r.nome, attivita: acts });
  });
  return groups;
}

function getDeptBucketByRisorsa(risorsaId) {
  const dept = normalizeDeptKey(getRisorsaDeptName(risorsaId));
  if (dept.includes("CAD")) return "CAD";
  if (dept.includes("TERMO")) return "TERMODINAMICI";
  if (dept.includes("ELETTR")) return "ELETTRICI";
  return "ALTRO";
}

function getDeptBucketByRepartoId(repartoId) {
  if (repartoId == null) return "ALTRO";
  const reparto = (state.reparti || []).find((r) => String(r.id) === String(repartoId));
  const dept = normalizeDeptKey(reparto ? reparto.nome : "");
  if (dept.includes("CAD")) return "CAD";
  if (dept.includes("TERMO")) return "TERMODINAMICI";
  if (dept.includes("ELETTR")) return "ELETTRICI";
  return "ALTRO";
}

function compareAttivitaStable(a, b) {
  const aStart = new Date(a.data_inizio).getTime();
  const bStart = new Date(b.data_inizio).getTime();
  if (aStart !== bStart) return aStart - bStart;
  const aEnd = new Date(a.data_fine).getTime();
  const bEnd = new Date(b.data_fine).getTime();
  if (aEnd !== bEnd) return aEnd - bEnd;
  const aCreated = a.created_at ? new Date(a.created_at).getTime() : 0;
  const bCreated = b.created_at ? new Date(b.created_at).getTime() : 0;
  if (aCreated !== bCreated) return aCreated - bCreated;
  return String(a.id || "").localeCompare(String(b.id || ""));
}

function formatTimeRange(startIso, endIso) {
  const start = new Date(startIso);
  const end = new Date(endIso);
  const opts = { hour: "2-digit", minute: "2-digit" };
  return `${start.toLocaleTimeString([], opts)} - ${end.toLocaleTimeString([], opts)}`;
}

function startOfDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function dayNumberUTC(date) {
  return Math.floor(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000);
}

function daysBetweenInclusive(start, end) {
  const s = dayNumberUTC(start);
  const e = dayNumberUTC(end);
  return Math.max(1, e - s + 1);
}

function overlapsDay(attivita, day) {
  const dayStart = startOfDay(day);
  const dayEnd = endOfDay(day);
  const start = new Date(attivita.data_inizio);
  const end = new Date(attivita.data_fine);
  return start <= dayEnd && end >= dayStart;
}

function isAssenteTitle(title) {
  return (title || "").trim().toLowerCase() === "assente";
}

function businessDaysBetweenInclusive(start, end) {
  const diff = businessDayDiff(start, end);
  return Math.max(1, Math.abs(diff) + 1);
}

function applyTime(baseDay, sourceDate) {
  const d = new Date(baseDay);
  d.setHours(
    sourceDate.getHours(),
    sourceDate.getMinutes(),
    sourceDate.getSeconds(),
    sourceDate.getMilliseconds()
  );
  return d;
}

function rangesOverlap(startA, endA, startB, endB) {
  return startA <= endB && endA >= startB;
}

function overlapsAny(start, end, ranges) {
  return ranges.some((r) => rangesOverlap(start, end, r.start, r.end));
}

function buildRangeFromStartDay(startDay, durationDays, startTime, endTime) {
  const start = applyTime(startDay, startTime);
  const endDay = addBusinessDays(startOfDay(start), durationDays - 1);
  const end = applyTime(endDay, endTime);
  return { start, end };
}

function resolveNextAvailableRange(startDay, durationDays, startTime, endTime, blockedRanges) {
  let day = startOfDay(startDay);
  for (let i = 0; i < 2000; i += 1) {
    const candidate = buildRangeFromStartDay(day, durationDays, startTime, endTime);
    if (!overlapsAny(candidate.start, candidate.end, blockedRanges)) {
      return candidate;
    }
    day = addBusinessDays(day, 1);
  }
  return null;
}

async function resolveRangeForActivity(activity, startDay, durationDays, excludeIds = []) {
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .eq("risorsa_id", activity.risorsa_id);
  if (error) {
    setStatus(`Absence check error: ${error.message}`, "error");
    return null;
  }
  const blocked = (data || [])
    .filter((a) => !excludeIds.includes(a.id))
    .map((a) => ({
      start: new Date(a.data_inizio),
      end: new Date(a.data_fine),
    }));
  return resolveNextAvailableRange(
    startDay,
    durationDays,
    new Date(activity.data_inizio),
    new Date(activity.data_fine),
    blocked
  );
}

async function shiftRisorsa(risorsaId, cutDate, deltaDays, excludeIds = []) {
  if (!deltaDays) return true;
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .eq("risorsa_id", risorsaId)
    .order("data_inizio", { ascending: true });
  if (error) {
    setStatus(`Shift error: ${error.message}`, "error");
    return false;
  }
  const cut = startOfDay(cutDate);
  const targets = [];
  const blocked = [];
  (data || []).forEach((a) => {
    const endDay = startOfDay(new Date(a.data_fine));
    if (excludeIds.includes(a.id) || isAssenteTitle(a.titolo) || endDay < cut) {
      blocked.push({ start: new Date(a.data_inizio), end: new Date(a.data_fine) });
      return;
    }
    targets.push(a);
  });
  if (!targets.length) return true;
  targets.sort((a, b) => new Date(a.data_inizio) - new Date(b.data_inizio));
  const placed = [];
  const updates = [];
  for (const a of targets) {
    const start = new Date(a.data_inizio);
    const end = new Date(a.data_fine);
    const durationDays = businessDaysBetweenInclusive(start, end);
    const desiredDay = addBusinessDays(startOfDay(start), deltaDays);
    const candidate = resolveNextAvailableRange(
      desiredDay,
      durationDays,
      start,
      end,
      blocked.concat(placed)
    );
    if (!candidate) {
      setStatus("Unable to find free space for the shift.", "error");
      return false;
    }
    placed.push({ start: candidate.start, end: candidate.end });
    updates.push({ id: a.id, start: candidate.start, end: candidate.end });
  }
  if (!updates.length) return true;
  const results = await Promise.all(
    updates.map((u) =>
      supabase
        .from("attivita")
        .update({ data_inizio: u.start.toISOString(), data_fine: u.end.toISOString() })
        .eq("id", u.id)
    )
  );
  const firstError = results.find((r) => r.error);
  if (firstError && firstError.error) {
    setStatus(`Shift error: ${firstError.error.message}`, "error");
    return false;
  }
  return true;
}

async function shiftRisorsaAll(risorsaId, deltaDays, excludeIds = [], limitDate = null) {
  if (!deltaDays) return true;
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .eq("risorsa_id", risorsaId)
    .order("data_inizio", { ascending: true });
  if (error) {
    setStatus(`Shift error: ${error.message}`, "error");
    return false;
  }
  const targets = (data || [])
    .filter((a) => !excludeIds.includes(a.id))
    .filter((a) => !isAssenteTitle(a.titolo))
    .filter((a) => {
      if (!limitDate) return true;
      return new Date(a.data_inizio) <= limitDate;
    });
  if (!targets.length) return true;
  const updates = targets.map((a) => {
    const newStart = addBusinessDays(new Date(a.data_inizio), deltaDays);
    const newEnd = addBusinessDays(new Date(a.data_fine), deltaDays);
    return supabase
      .from("attivita")
      .update({ data_inizio: newStart.toISOString(), data_fine: newEnd.toISOString() })
      .eq("id", a.id);
  });
  const results = await Promise.all(updates);
  const firstError = results.find((r) => r.error);
  if (firstError && firstError.error) {
    setStatus(`Shift error: ${firstError.error.message}`, "error");
    return false;
  }
  return true;
}

async function hasOverlapRange(risorsaId, start, end, excludeId) {
  const { data, error } = await supabase
    .from("attivita")
    .select("id")
    .eq("risorsa_id", risorsaId)
    .lte("data_inizio", end.toISOString())
    .gte("data_fine", start.toISOString())
    .neq("id", excludeId)
    .limit(1);
  if (error) {
    setStatus(`Overlap check error: ${error.message}`, "error");
    return null;
  }
  return (data || []).length > 0;
}

async function getDependentActivitiesForKey(keyActivity) {
  if (!keyActivity || !keyActivity.commessa_id) return [];
  const dependents = getPhaseDependents(keyActivity.titolo);
  if (!dependents.length) return [];
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .eq("commessa_id", keyActivity.commessa_id);
  if (error) {
    setStatus(`Dependency error: ${error.message}`, "error");
    return null;
  }
  return (data || []).filter(
    (a) => a.id !== keyActivity.id && dependents.includes(normalizePhaseKey(a.titolo))
  );
}

function shouldShiftDependent(keyActivity, deltaDays, dep) {
  if (!deltaDays || !keyActivity || !dep) return false;
  if (!keyActivity.data_fine) return false;
  const oldEndDay = startOfDay(new Date(keyActivity.data_fine));
  const newEndDay = addBusinessDays(oldEndDay, deltaDays);
  const depStartDay = startOfDay(new Date(dep.data_inizio));
  return oldEndDay < depStartDay && newEndDay >= depStartDay;
}

function getDependentActivitiesToShift(keyActivity, deltaDays, dependentActivities = []) {
  return dependentActivities.filter((dep) => shouldShiftDependent(keyActivity, deltaDays, dep));
}

async function shiftDependentPhases(keyActivity, deltaDays, dependentActivities = null) {
  if (!matrixState.autoShift || !deltaDays) return true;
  if (!keyActivity || !keyActivity.commessa_id) return true;
  let activities = dependentActivities;
  if (!activities) {
    activities = await getDependentActivitiesForKey(keyActivity);
    if (activities == null) return false;
  }
  if (!activities.length) return true;
  activities.sort((a, b) => new Date(a.data_inizio) - new Date(b.data_inizio));

  const shiftedIds = new Set();

  for (const dep of activities) {
    if (shiftedIds.has(dep.id)) continue;
    const start = new Date(dep.data_inizio);
    const end = new Date(dep.data_fine);
    const durationDays = businessDaysBetweenInclusive(start, end);
    const targetDay = addBusinessDays(startOfDay(start), deltaDays);
    const targetRange = buildRangeFromStartDay(targetDay, durationDays, start, end);

    if (deltaDays > 0) {
      const overlap = await hasOverlapRange(dep.risorsa_id, targetRange.start, targetRange.end, dep.id);
      if (overlap == null) return false;
      if (overlap) {
        const ok = await shiftRisorsa(dep.risorsa_id, targetDay, deltaDays, [dep.id, keyActivity.id]);
        if (!ok) return false;
        const cut = startOfDay(targetDay);
        activities.forEach((a) => {
          if (a.risorsa_id === dep.risorsa_id && startOfDay(new Date(a.data_fine)) >= cut) {
            shiftedIds.add(a.id);
          }
        });
      }
    } else if (deltaDays < 0) {
      const cutDate = addDays(startOfDay(end), 1);
      const ok = await shiftRisorsa(dep.risorsa_id, cutDate, deltaDays, [dep.id, keyActivity.id]);
      if (!ok) return false;
      const cut = startOfDay(cutDate);
      activities.forEach((a) => {
        if (a.risorsa_id === dep.risorsa_id && startOfDay(new Date(a.data_fine)) >= cut) {
          shiftedIds.add(a.id);
        }
      });
    }

    let newStart = targetRange.start;
    let newEnd = targetRange.end;
    const stillOverlap = await hasOverlapRange(dep.risorsa_id, newStart, newEnd, dep.id);
    if (stillOverlap == null) return false;
    if (stillOverlap) {
      const excludeIds = new Set([keyActivity.id, dep.id]);
      activities.forEach((a) => {
        if (a.id !== dep.id && !shiftedIds.has(a.id)) {
          excludeIds.add(a.id);
        }
      });
      const candidate = await resolveRangeForActivity(
        dep,
        targetDay,
        durationDays,
        Array.from(excludeIds)
      );
      if (!candidate) return false;
      newStart = candidate.start;
      newEnd = candidate.end;
    }
    const { error: updError } = await supabase
      .from("attivita")
      .update({ data_inizio: newStart.toISOString(), data_fine: newEnd.toISOString() })
      .eq("id", dep.id);
    if (updError) {
      setStatus(`Dependency error: ${updError.message}`, "error");
      return false;
    }
    shiftedIds.add(dep.id);
  }

  return true;
}

function hasGapBefore(list, targetStart) {
  if (!list.length) return false;
  const sorted = list
    .slice()
    .sort((a, b) => new Date(a.data_inizio) - new Date(b.data_inizio));
  const tStart = startOfDay(targetStart);
  const prev = sorted.filter((a) => startOfDay(new Date(a.data_inizio)) < tStart);
  if (!prev.length) return false;
  const last = prev[prev.length - 1];
  const lastEnd = startOfDay(new Date(last.data_fine));
  const expectedPrevEnd = addBusinessDays(tStart, -1);
  return lastEnd.getTime() < expectedPrevEnd.getTime();
}

async function handleMatrixDropOnCell(cell, e) {
  if (!cell || !cell.dataset) return;
  const risorsaId = cell.dataset.risorsa;
  const dayKey = cell.dataset.day;
  if (!risorsaId || !dayKey) return;
  const r = matrixState.risorse.find((x) => String(x.id) === String(risorsaId));
  if (!r) return;
  const commessaId = e.dataTransfer.getData("application/x-commessa-id");
  if (commessaId) {
    if (!canAssignCommessaToRisorsa(r.id)) {
      setStatus("Non autorizzato ad assegnare su questa riga.", "error");
      return;
    }
  } else if (!canMoveMatrixActivity(null, r.id)) {
    setStatus("Non autorizzato a modificare questa riga.", "error");
    return;
  }
  const d = dateFromKey(dayKey);
  const createGenericActivity = async (title, options = {}) => {
    const shouldOpen = options.openModal !== false;
    if (!isActivityAllowedForRisorsa(title, r.id)) {
      setStatus("This activity is not available for this department.", "error");
      return true;
    }
    const startAt = new Date(d);
    startAt.setHours(8, 0, 0, 0);
    const endAt = new Date(d);
    endAt.setHours(17, 0, 0, 0);
    if (matrixState.autoShift) {
      const hasOverlap = matrixState.attivita.some((a) => a.risorsa_id === r.id && overlapsDay(a, d));
      if (hasOverlap) {
        const ok = await shiftRisorsa(r.id, d, 1);
        if (!ok) return true;
      }
    }
    const { data, error } = await supabase
      .from("attivita")
      .insert({
        commessa_id: null,
        titolo: title,
        descrizione: null,
        risorsa_id: r.id,
        data_inizio: startAt.toISOString(),
        data_fine: endAt.toISOString(),
        stato: "pianificata",
      })
      .select("*")
      .single();
    if (error) {
      setStatus(`Activity error: ${error.message}`, "error");
      return true;
    }
    setStatus("Activity created.", "ok");
    if (shouldOpen) openActivityModal(data);
    await loadMatrixAttivita();
    return true;
  };
  const assenzaToken = e.dataTransfer.getData("application/x-assenza");
  if (assenzaToken) {
    const startAt = new Date(d);
    startAt.setHours(8, 0, 0, 0);
    const endAt = new Date(d);
    endAt.setHours(17, 0, 0, 0);
    if (matrixState.autoShift) {
      const hasOverlap = matrixState.attivita.some((a) => a.risorsa_id === r.id && overlapsDay(a, d));
      if (hasOverlap) {
        const ok = await shiftRisorsa(r.id, d, 1);
        if (!ok) return;
      }
    }
    const { data, error } = await supabase
      .from("attivita")
      .insert({
        commessa_id: null,
        titolo: "ASSENTE",
        ore_assenza: 8,
        risorsa_id: r.id,
        data_inizio: startAt.toISOString(),
        data_fine: endAt.toISOString(),
        stato: "pianificata",
      })
      .select("*")
      .single();
    if (error) {
      setStatus(`Absence error: ${error.message}`, "error");
      return;
    }
    setStatus("Absence created.", "ok");
    openActivityModal(data);
    await loadMatrixAttivita();
    return;
  }
  const altroToken = e.dataTransfer.getData("application/x-altro");
  if (altroToken) {
    await createGenericActivity("Altro", { openModal: false });
    return;
  }
  const pedToken = e.dataTransfer.getData("application/x-ped");
  if (pedToken) {
    await createGenericActivity("PED E TARGHETTE", { openModal: false });
    return;
  }
  const segnaturaToken = e.dataTransfer.getData("application/x-segnatura");
  if (segnaturaToken) {
    await createGenericActivity("SEGNATURA ELETTRICA", { openModal: false });
    return;
  }
  const opOrdiniToken = e.dataTransfer.getData("application/x-op-ordini");
  if (opOrdiniToken) {
    await createGenericActivity("OP ORDINI", { openModal: false });
    return;
  }
  if (commessaId) {
    matrixState.pendingDrop = { commessaId, risorsaId: r.id, day: d };
    renderAssignActivities();
    openAssignModal();
    return;
  }
  const dragId = matrixState.draggingId || e.dataTransfer.getData("text/plain");
  if (!dragId) return;
  const item = matrixState.attivita.find((x) => x.id === dragId);
  if (!item) return;
  if (!canMoveMatrixActivity(item, r.id)) {
    setStatus("Non autorizzato a spostare questa attivita.", "error");
    return;
  }

  const targetDate = new Date(d);
  const start = new Date(item.data_inizio);
  const end = new Date(item.data_fine);
  const newStart = new Date(targetDate);
  newStart.setHours(start.getHours(), start.getMinutes(), 0, 0);
  const durationDays = businessDaysBetweenInclusive(start, end);
  const newEndDay = addBusinessDays(startOfDay(newStart), durationDays - 1);
  const newEnd = new Date(newEndDay);
  newEnd.setHours(end.getHours(), end.getMinutes(), end.getSeconds(), end.getMilliseconds());
  const deltaDays = businessDayDiff(start, newStart);

        const copy = e.altKey || e.ctrlKey || e.metaKey;
        if (copy) {
          if (!isActivityAllowedForRisorsa(item.titolo, r.id)) {
            setStatus("This activity is not available for this department.", "error");
            return;
          }
          if (matrixState.autoShift) {
            const hasOverlap = matrixState.attivita.some((a) => a.risorsa_id === r.id && overlapsDay(a, targetDate));
            if (hasOverlap) {
              const ok = await shiftRisorsa(r.id, targetDate, durationDays);
              if (!ok) return;
            }
          }
    const { error } = await supabase.from("attivita").insert({
      commessa_id: item.commessa_id,
      titolo: item.titolo,
      descrizione: item.descrizione,
      risorsa_id: r.id,
      reparto_id: item.reparto_id,
      stato: item.stato,
      data_inizio: newStart.toISOString(),
      data_fine: newEnd.toISOString(),
    });
    if (error) {
      setStatus(`Copy error: ${error.message}`, "error");
      return;
    }
    setStatus("Activity copied.", "ok");
        } else {
          if (!isActivityAllowedForRisorsa(item.titolo, r.id)) {
            setStatus("This activity is not available for this department.", "error");
            return;
          }
          let dependentActivities = null;
          let dependentToShift = null;
          if (matrixState.autoShift && isPhaseKey(item.titolo) && deltaDays !== 0) {
            dependentActivities = await getDependentActivitiesForKey(item);
            if (dependentActivities == null) return;
            dependentToShift = getDependentActivitiesToShift(item, deltaDays, dependentActivities);
          }
          const excludeIds =
            dependentToShift && dependentToShift.length
              ? [item.id, ...dependentToShift.map((a) => a.id)]
              : [item.id];
    if (matrixState.autoShift && deltaDays < 0) {
      const cutDate = addDays(startOfDay(end), 1);
      const ok = await shiftRisorsa(r.id, cutDate, deltaDays, excludeIds);
      if (!ok) return;
    }
    if (matrixState.autoShift && deltaDays > 0) {
      const overlap = await hasOverlapRange(r.id, newStart, newEnd, item.id);
      if (overlap == null) return;
      if (overlap) {
        const cutDate = startOfDay(newStart);
        const ok = await shiftRisorsa(r.id, cutDate, deltaDays, excludeIds);
        if (!ok) return;
      }
    }
    if (!canPermission("can_move_matrix")) {
      setStatus("Non autorizzato a spostare attivita.", "error");
      return;
    }
    const { error } = await supabase
      .from("attivita")
      .update({
        risorsa_id: r.id,
        data_inizio: newStart.toISOString(),
        data_fine: newEnd.toISOString(),
      })
      .eq("id", item.id);
    if (error) {
      setStatus(`Move error: ${error.message}`, "error");
      return;
    }
          if (isPhaseKey(item.titolo) && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
            const ok = await shiftDependentPhases(item, deltaDays, dependentToShift);
            if (!ok) return;
          }
    setStatus("Activity moved.", "ok");
  }
  await loadMatrixAttivita();
  await refreshTodoRealtimeSnapshot();
}

function getCommessaLabel(commessaId) {
  if (!commessaId) return "";
  const c = state.commesse.find((x) => x.id === commessaId);
  return c ? `${c.codice}${c.titolo ? " - " + c.titolo : ""}` : commessaId;
}

function parseCommessaCode(code) {
  const text = (code || "").trim();
  const match = text.match(/^(\d{4})[_-]?(\d+)/);
  if (!match) return { year: null, num: null, raw: text };
  return { year: Number(match[1]), num: Number(match[2]), raw: text };
}

function normalizeCommessaParts(annoRaw, numeroRaw) {
  const anno = Number(annoRaw);
  const numero = Number(numeroRaw);
  if (!Number.isInteger(anno) || anno < 2000) return null;
  if (!Number.isInteger(numero) || numero < 1) return null;
  return {
    codice: `${anno}_${numero}`,
    anno,
    numero,
  };
}

function findDuplicateCommessa(anno, numero, excludeId) {
  if (!Number.isInteger(anno) || !Number.isInteger(numero)) return null;
  return state.commesse.find(
    (c) =>
      Number(c.anno) === anno &&
      Number(c.numero) === numero &&
      (!excludeId || String(c.id) !== String(excludeId))
  );
}

function updateNumeroWarning(annoInput, numeroInput, warningEl, excludeId) {
  if (!warningEl || !annoInput || !numeroInput) return;
  const anno = Number(annoInput.value);
  const numero = Number(numeroInput.value);
  const dup = findDuplicateCommessa(anno, numero, excludeId);
  if (dup) {
    warningEl.textContent = `Attenzione: ${anno}_${numero} esiste gia (${dup.titolo || "senza descrizione"}).`;
    warningEl.classList.remove("hidden");
  } else {
    warningEl.textContent = "";
    warningEl.classList.add("hidden");
  }
}

function updateCodicePreview(annoInput, numeroInput, previewEl) {
  if (!previewEl || !annoInput || !numeroInput) return;
  const normalized = normalizeCommessaParts(annoInput.value, numeroInput.value);
  previewEl.textContent = normalized ? normalized.codice : "-";
}

function compareCommesse(a, b) {
  const pa = parseCommessaCode(a.codice);
  const pb = parseCommessaCode(b.codice);
  if (pa.year != null && pb.year != null) {
    if (pa.year !== pb.year) return pa.year - pb.year;
    if (pa.num != null && pb.num != null && pa.num !== pb.num) return pa.num - pb.num;
  }
  return (a.codice || "").localeCompare(b.codice || "");
}

function compareCommesseDesc(a, b) {
  const pa = parseCommessaCode(a.codice);
  const pb = parseCommessaCode(b.codice);
  const ay = pa.year != null ? pa.year : -1;
  const by = pb.year != null ? pb.year : -1;
  if (ay !== by) return by - ay;
  const an = pa.num != null ? pa.num : -1;
  const bn = pb.num != null ? pb.num : -1;
  if (an !== bn) return bn - an;
  return (b.codice || "").localeCompare(a.codice || "");
}

function renderCommesse(list) {
  commesseList.innerHTML = "";
  if (!list.length) {
    const empty = document.createElement("div");
    empty.className = "mono";
    empty.textContent = "Nessuna commessa trovata.";
    commesseList.appendChild(empty);
    return;
  }
  const grouped = !state.commesseFiltersActive;
  const sorted = list.slice().sort(grouped ? compareCommesseDesc : compareCommesse);
  let lastYear = null;
  sorted.forEach((c) => {
    if (grouped) {
      const year = c.anno != null ? String(c.anno) : "Senza anno";
      if (year !== lastYear) {
        const header = document.createElement("div");
        header.className = "commessa-year-header";
        header.textContent = year;
        commesseList.appendChild(header);
        lastYear = year;
      }
    }
    const item = document.createElement("div");
    item.className = `item${state.selected && state.selected.id === c.id ? " selected" : ""}`;
    item.dataset.id = c.id;
    const codeInfo = parseCommessaCode(c.codice || "");
    const annoLabel = c.anno ?? codeInfo.year ?? "";
    const numeroLabel = c.numero ?? codeInfo.num ?? c.codice ?? "-";
    const annoHtml = annoLabel ? `<span class="commessa-year">${annoLabel}</span>` : "";
    const risorse = state.commessaRisorseMap.get(c.id) || [];
    const shownRisorse = risorse.slice(0, 3);
    const extraRisorse = risorse.length > 3 ? ` +${risorse.length - 3}` : "";
    const risorseHtml = shownRisorse
      .map((r) => {
        const acts = r.attivita && r.attivita.length ? r.attivita.join(", ") : "";
        const label = acts ? `${r.nome} \u00B7 ${acts}` : r.nome;
        return `<span class="risorsa-tag ${r.deptClass}">${label}</span>`;
      })
      .join("");
    item.innerHTML = `
      <div class="item-title">
        <span class="commessa-pill">${numeroLabel}</span>
        ${annoHtml}
        <span class="commessa-name">${c.titolo || "Senza titolo"}</span>
      </div>
      <div class="item-meta">
        <span>${c.cliente || "Cliente n/d"}</span>
        <span>Stato: ${c.stato}</span>
        <span>Ingresso: ${formatDate(c.data_ingresso) || "-"}</span>
        ${
          shownRisorse.length
            ? `<div class="commessa-risorse">${risorseHtml}${
                extraRisorse ? `<span class="risorsa-extra">${extraRisorse}</span>` : ""
              }</div>`
            : ""
        }
      </div>
    `;
    item.addEventListener("click", (e) => {
      if (commessaItemLongPressSuppress) {
        commessaItemLongPressSuppress = false;
        e.stopPropagation();
        return;
      }
      closeCommessaQuickMenu();
      selectCommessa(c.id, { openModal: true });
    });
    item.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      closeCommessaQuickMenu();
      const pressX = e.clientX;
      const pressY = e.clientY;
      if (commessaQuickTimer) clearTimeout(commessaQuickTimer);
      commessaQuickTimer = setTimeout(() => {
        commessaQuickTimer = null;
        openCommessaQuickMenu(c, pressX, pressY, item);
      }, 450);
    });
    item.addEventListener("mouseup", () => {
      if (commessaQuickTimer) {
        clearTimeout(commessaQuickTimer);
        commessaQuickTimer = null;
      }
    });
    item.addEventListener("mouseleave", () => {
      if (commessaQuickTimer) {
        clearTimeout(commessaQuickTimer);
        commessaQuickTimer = null;
      }
    });
    commesseList.appendChild(item);
  });
}

function renderReparti(commessa) {
  if (!repartiList) return;
  repartiList.innerHTML = "";
  if (!commessa) return;

  const current = Array.isArray(commessa.reparti) ? commessa.reparti : [];
  if (!current.length) {
    repartiList.innerHTML = `<div class="mono">Nessun reparto associato.</div>`;
    return;
  }

  current.forEach((r) => {
    const row = document.createElement("div");
    row.className = "reparto-row";

    const name = document.createElement("div");
    name.textContent = r.reparto_nome;

    const stato = document.createElement("select");
    ["da_fare", "in_corso", "fatto"].forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s.replace("_", " ");
      if (r.stato === s) opt.selected = true;
      stato.appendChild(opt);
    });

    const note = document.createElement("input");
    note.type = "text";
    note.placeholder = "Note reparto";
    note.value = r.note_reparto || "";

    const save = document.createElement("button");
    save.type = "button";
    save.className = "ghost";
    save.textContent = "Salva";
    save.addEventListener("click", async () => {
      if (!state.canWrite) return;
      const { error } = await supabase
        .from("commesse_reparti")
        .update({ stato: stato.value, note_reparto: note.value })
        .eq("commessa_id", commessa.id)
        .eq("reparto_id", r.reparto_id);
      if (error) {
        setStatus(`Department error: ${error.message}`, "error");
        return;
      }
  setStatus("Department updated.", "ok");
      await loadCommesse();
      selectCommessa(commessa.id);
    });

    row.appendChild(name);
    row.appendChild(stato);
    row.appendChild(note);
    row.appendChild(save);
    repartiList.appendChild(row);
  });
}

function renderRepartiChecks() {
  if (!newRepartiChecks) return;
  newRepartiChecks.innerHTML = "";
  state.reparti.forEach((r) => {
    const wrapper = document.createElement("label");
    wrapper.className = "check-item";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = r.id;
    const span = document.createElement("span");
    span.textContent = r.nome;
    wrapper.appendChild(input);
    wrapper.appendChild(span);
    newRepartiChecks.appendChild(wrapper);
  });
}

function renderAddRepartoSelect() {
  if (!addRepartoSelect) return;
  addRepartoSelect.innerHTML = "";
  state.reparti.forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.id;
    opt.textContent = r.nome;
    addRepartoSelect.appendChild(opt);
  });
}

function renderResourceRepartoSelect(selectEl, selectedId = "") {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Reparto</option>`;
  state.reparti.forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.id;
    opt.textContent = r.nome;
    if (selectedId && String(selectedId) === String(r.id)) opt.selected = true;
    selectEl.appendChild(opt);
  });
}

function renderResourceUserSelect(selectEl, selectedId = "") {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Email utente</option>`;
  state.utenti.forEach((u) => {
    const opt = document.createElement("option");
    opt.value = u.id;
    opt.textContent = u.email || u.id;
    if (selectedId && String(selectedId) === String(u.id)) opt.selected = true;
    selectEl.appendChild(opt);
  });
}

function renderResourcesPanel() {
  if (!resourcesList) return;
  resourcesList.innerHTML = "";
  if (!state.risorse.length) {
    const empty = document.createElement("div");
    empty.className = "matrix-empty";
    empty.textContent = "Nessuna risorsa.";
    resourcesList.appendChild(empty);
    return;
  }
  state.risorse.slice().sort((a, b) => a.nome.localeCompare(b.nome)).forEach((r) => {
    const row = document.createElement("div");
    row.className = "resource-row";
    row.dataset.id = r.id;

    const name = document.createElement("input");
    name.type = "text";
    name.className = "resource-name";
    name.value = r.nome || "";

    const reparto = document.createElement("select");
    reparto.className = "resource-reparto";
    renderResourceRepartoSelect(reparto, r.reparto_id);

    const userSelect = document.createElement("select");
    userSelect.className = "resource-user";
    renderResourceUserSelect(userSelect, r.utente_id);

    const activeLabel = document.createElement("label");
    activeLabel.className = "resource-active";
    const activeInput = document.createElement("input");
    activeInput.type = "checkbox";
    activeInput.checked = r.attiva;
    const activeText = document.createElement("span");
    activeText.textContent = "Attiva";
    activeLabel.appendChild(activeInput);
    activeLabel.appendChild(activeText);

    const save = document.createElement("button");
    save.type = "button";
    save.className = "ghost";
    save.textContent = "Salva";
    save.addEventListener("click", async () => {
      if (!state.canWrite) return;
      const nome = name.value.trim();
      if (!nome) {
        setStatus("Enter a resource name.", "error");
        return;
      }
      const repartoId = reparto.value || null;
      const utenteId = userSelect.value || null;
      if (utenteId) {
        const conflict = state.risorse.find(
          (other) => other.id !== r.id && String(other.utente_id || "") === String(utenteId)
        );
        if (conflict) {
          setStatus("Email gia assegnata a un'altra risorsa.", "error");
          return;
        }
      }
      const { error } = await supabase
        .from("risorse")
        .update({ nome, reparto_id: repartoId, attiva: activeInput.checked, utente_id: utenteId })
        .eq("id", r.id);
      if (error) {
        setStatus(`Resource error: ${error.message}`, "error");
        return;
      }
      setStatus("Resource updated.", "ok");
      await loadRisorse();
    });

    const del = document.createElement("button");
    del.type = "button";
    del.className = "danger";
    del.textContent = "Elimina";
    del.addEventListener("click", async () => {
      if (!state.canWrite) return;
      const ok = confirm(`Eliminare la risorsa ${r.nome}?`);
      if (!ok) return;
      const { error } = await supabase.from("risorse").delete().eq("id", r.id);
      if (error) {
        setStatus(`Delete error: ${error.message}`, "error");
        return;
      }
      setStatus("Resource deleted.", "ok");
      await loadRisorse();
    });

    row.appendChild(name);
    row.appendChild(reparto);
    row.appendChild(userSelect);
    row.appendChild(activeLabel);
    row.appendChild(save);
    row.appendChild(del);
    if (!state.canWrite) {
      row.querySelectorAll("input, select, button").forEach((el) => (el.disabled = true));
    }
    resourcesList.appendChild(row);
  });
}

function applyFilters() {
  const desc = descriptionFilter ? descriptionFilter.value.trim().toLowerCase() : "";
  const numRaw = numberFilter ? numberFilter.value.trim() : "";
  const numVal = numRaw ? Number.parseInt(numRaw, 10) : null;
  const year = yearFilter ? yearFilter.value : "";
  state.commesseFiltersActive = Boolean(desc || numRaw || year);
  const filtered = state.commesse.filter((c) => {
    if (year && String(c.anno || "") !== year) return false;
    if (numVal != null && Number.isFinite(numVal)) {
      if (Number(c.numero ?? -1) !== numVal) return false;
    }
    if (desc) {
      const hay = `${c.titolo || ""}`.toLowerCase();
      if (!hay.includes(desc)) return false;
    }
    return true;
  });
  state.filteredCommesse = filtered;
  renderCommesse(filtered);
  loadCommesseRisorseFor(filtered);
}

function renderYearFilter() {
  if (!yearFilter) return;
  const years = Array.from(new Set(state.commesse.map((c) => c.anno).filter(Boolean))).sort();
  const current = yearFilter.value;
  yearFilter.innerHTML = `<option value="">Tutti gli anni</option>`;
  years.forEach((y) => {
    const opt = document.createElement("option");
    opt.value = String(y);
    opt.textContent = String(y);
    yearFilter.appendChild(opt);
  });
  if (current && years.includes(Number(current))) {
    yearFilter.value = current;
  }
}

function renderReportYearFilter() {
  if (!reportYearFilter) return;
  const years = Array.from(new Set(state.commesse.map((c) => c.anno).filter(Boolean))).sort((a, b) => b - a);
  const current = reportYearFilter.value;
  reportYearFilter.innerHTML = `<option value=\"\">Tutti gli anni</option>`;
  years.forEach((y) => {
    const opt = document.createElement("option");
    opt.value = String(y);
    opt.textContent = String(y);
    reportYearFilter.appendChild(opt);
  });
  if (current && years.includes(Number(current))) {
    reportYearFilter.value = current;
  }
}

function renderTodoYearFilter() {
  if (!todoYearFilter) return;
  const years = Array.from(new Set(state.commesse.map((c) => c.anno).filter(Boolean))).sort((a, b) => b - a);
  const current = todoYearFilter.value;
  todoYearFilter.innerHTML = `<option value="">Tutti gli anni</option>`;
  years.forEach((y) => {
    const opt = document.createElement("option");
    opt.value = String(y);
    opt.textContent = String(y);
    todoYearFilter.appendChild(opt);
  });
  if (current && years.includes(Number(current))) {
    todoYearFilter.value = current;
    return;
  }
  const thisYear = String(new Date().getFullYear());
  if (years.includes(Number(thisYear))) {
    todoYearFilter.value = thisYear;
  }
}

function isTodoFilterControlTarget(target) {
  if (!target || typeof target.closest !== "function") return false;
  if (todoFilterControls.includes(target)) return true;
  return Boolean(target.closest("#todoSection .todo-filters"));
}

function lockTodoGridHeightForFilters() {
  if (!todoGrid) return;
  const active = document.activeElement;
  if (!isTodoFilterControlTarget(active)) return;
  const height = Math.ceil(todoGrid.getBoundingClientRect().height);
  if (!Number.isFinite(height) || height <= 0) return;
  todoFilterHeightLock = Math.max(todoFilterHeightLock, height);
  todoGrid.style.minHeight = `${todoFilterHeightLock}px`;
}

function releaseTodoGridHeightLock() {
  if (!todoGrid) return;
  const active = document.activeElement;
  if (isTodoFilterControlTarget(active)) return;
  todoFilterHeightLock = 0;
  todoGrid.style.minHeight = "";
}

function scheduleTodoGridHeightUnlock() {
  if (todoFilterUnlockTimer) clearTimeout(todoFilterUnlockTimer);
  todoFilterUnlockTimer = setTimeout(() => {
    todoFilterUnlockTimer = null;
    releaseTodoGridHeightLock();
  }, 0);
}

function renderMatrixReportFromTodoFilters() {
  const prevX = window.scrollX;
  const prevY = window.scrollY;
  lockTodoGridHeightForFilters();
  renderMatrixReport();
  requestAnimationFrame(() => {
    window.scrollTo(prevX, prevY);
    scheduleTodoGridHeightUnlock();
  });
}

function getTodoFilters() {
  const desc = todoDescFilter ? todoDescFilter.value.trim().toLowerCase() : "";
  const numRaw = todoNumberFilter ? todoNumberFilter.value.trim() : "";
  const numVal = numRaw ? Number.parseInt(numRaw, 10) : null;
  const year = todoYearFilter ? todoYearFilter.value : "";
  const priorityField = todoPriorityField ? todoPriorityField.value : "";
  const priorityMode = todoPriorityBtn?.dataset.mode || "off";
  const sortMode = todoSortBtn?.dataset.mode || "off";
  const showComplete = todoCompleteBtn ? !todoCompleteBtn.classList.contains("active") : true;
  return { desc, numVal, year, priorityField, priorityMode, sortMode, showComplete };
}

function renderTodoSortButtonLabel() {
  if (!todoSortBtn) return;
  const mode = todoSortBtn.dataset.mode || "off";
  todoSortBtn.classList.toggle("active", mode !== "off");
  if (mode === "asc") {
    todoSortBtn.textContent = "Ordine numero: crescente";
  } else if (mode === "desc") {
    todoSortBtn.textContent = "Ordine numero: decrescente";
  } else {
    todoSortBtn.textContent = "Ordine numero: OFF";
  }
}

function renderTodoCompleteButtonLabel() {
  if (!todoCompleteBtn) return;
  todoCompleteBtn.textContent = todoCompleteBtn.classList.contains("active")
    ? "Complete: nascoste"
    : "Complete: in mostra";
}

function renderTodoPriorityButtonLabel() {
  if (!todoPriorityBtn) return;
  const mode = todoPriorityBtn.dataset.mode || "off";
  todoPriorityBtn.classList.toggle("active", mode !== "off");
  todoPriorityBtn.textContent = mode === "asc" ? "Priorità: crescente" : "Priorità: OFF";
}

function getTodoPriorityDate(commessa, field) {
  const parse = (raw) => {
    if (!raw) return null;
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return null;
    if (d.getFullYear() < 2000 || d.getFullYear() > 2100) return null;
    return d.getTime();
  };
  if (field === "telaio_target") return parse(commessa.data_ordine_telaio);
  if (field === "telaio_stimato") return parse(commessa.data_consegna_telaio_effettiva);
  if (field === "prelievo") return parse(commessa.data_prelievo);
  if (field === "kit_cavi") return parse(commessa.data_arrivo_kit_cavi);
  if (field === "consegna") {
    return parse(commessa.data_consegna_macchina || commessa.data_consegna_prevista || commessa.data_consegna);
  }
  return null;
}

function applyTodoFilters(list, columns) {
  const { desc, numVal, year, priorityField, priorityMode, sortMode, showComplete } = getTodoFilters();
  const isComplete = (commessa) => {
    if (!columns.length) return false;
    return columns.every((col) => getTodoStatusFor(commessa.id, col.titolo, col.reparto) === "fatta");
  };
  const isReleasedOrClosed = (commessa) => {
    const global = getTodoGlobalCommessaStatus(commessa).label;
    return global === "Rilasciata a prod." || global === "Evasa";
  };
  let commesse = list.slice();
  if (year) {
    commesse = commesse.filter((c) => String(c.anno || getCommessaYear(c) || "") === year);
  }
  if (numVal != null && Number.isFinite(numVal)) {
    commesse = commesse.filter((c) => Number(c.numero ?? -1) === numVal);
  }
  if (desc) {
    commesse = commesse.filter((c) => {
      const text = `${c.titolo || ""}`.toLowerCase();
      return text.includes(desc);
    });
  }
  if (!showComplete) {
    commesse = commesse.filter((c) => !isComplete(c) && !isReleasedOrClosed(c));
  }
  if (priorityMode === "asc" && priorityField) {
    commesse.sort((a, b) => {
      const aa = getTodoPriorityDate(a, priorityField);
      const bb = getTodoPriorityDate(b, priorityField);
      if (aa == null && bb == null) return compareCommesseDesc(a, b);
      if (aa == null) return 1;
      if (bb == null) return -1;
      if (aa !== bb) return aa - bb;
      return compareCommesseDesc(a, b);
    });
  }
  if (sortMode === "asc") {
    commesse.sort((a, b) => {
      const aa = Number(a.numero ?? Number.POSITIVE_INFINITY);
      const bb = Number(b.numero ?? Number.POSITIVE_INFINITY);
      if (aa !== bb) return aa - bb;
      return compareCommesse(a, b);
    });
  } else if (sortMode === "desc") {
    commesse.sort((a, b) => {
      const aa = Number(a.numero ?? Number.NEGATIVE_INFINITY);
      const bb = Number(b.numero ?? Number.NEGATIVE_INFINITY);
      if (aa !== bb) return bb - aa;
      return compareCommesseDesc(a, b);
    });
  }
  return commesse;
}

async function loadCommesseRisorseFor(list) {
  if (!list.length) {
    state.commessaRisorseMap = new Map();
    return;
  }
  const token = ++state.commessaRisorseToken;
  const ids = Array.from(new Set(list.map((c) => c.id))).filter(Boolean);
  if (!ids.length) {
    state.commessaRisorseMap = new Map();
    return;
  }

  let rows = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("attivita").select("commessa_id, risorsa_id, titolo").in("commessa_id", ids),
        FETCH_TIMEOUT_MS
      );
      rows = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadCommesseRisorse timeout: ${err?.message || err}`);
    }
  }

  if (!rows || error) {
    debugLog("loadCommesseRisorse fallback REST");
    const inList = ids.join(",");
    rows = await fetchTableViaRest(
      "attivita",
      `select=commessa_id,risorsa_id,titolo&commessa_id=in.(${inList})`
    );
  }

  if (!rows || token !== state.commessaRisorseToken) return;

  const risorseById = new Map((state.risorse || []).map((r) => [String(r.id), r.nome]));
  const map = new Map();
  rows.forEach((row) => {
    if (!row.commessa_id || row.risorsa_id == null) return;
    const key = String(row.commessa_id);
    const byRisorsa = map.get(key) || new Map();
    const risorsaKey = String(row.risorsa_id);
    const entry = byRisorsa.get(risorsaKey) || new Set();
    const titolo = (row.titolo || "").trim();
    if (titolo) entry.add(titolo);
    byRisorsa.set(risorsaKey, entry);
    map.set(key, byRisorsa);
  });

  const resolved = new Map();
  map.forEach((byRisorsa, commessaId) => {
    const items = Array.from(byRisorsa.entries())
      .map(([id, titles]) => {
        const nome = risorseById.get(id);
        if (!nome) return null;
        const dept = normalizeDeptKey(getRisorsaDeptName(id));
        let deptClass = "dept-default";
        if (dept.includes("ELETTR")) deptClass = "dept-elettrici";
        else if (dept.includes("TERMO")) deptClass = "dept-termodinamici";
        else if (dept.includes("CAD")) deptClass = "dept-cad";
        const attivita = Array.from(titles).sort((a, b) => a.localeCompare(b));
        return { nome, deptClass, attivita };
      })
      .filter(Boolean)
      .sort((a, b) => a.nome.localeCompare(b.nome));
    resolved.set(commessaId, items);
  });

  if (token !== state.commessaRisorseToken) return;
  state.commessaRisorseMap = resolved;
  renderCommesse(state.filteredCommesse || list);
}

function selectCommessa(id, options = {}) {
  const commessa = state.commesse.find((c) => c.id === id);
  if (!commessa) return;
  state.selected = commessa;
  selectedCode.textContent = commessa.codice;

  if (d.anno) d.anno.value = commessa.anno != null ? String(commessa.anno) : "";
  if (d.numero) d.numero.value = commessa.numero != null ? String(commessa.numero) : "";
  d.titolo.value = commessa.titolo || "";
  d.cliente.value = commessa.cliente || "";
  d.stato.value = commessa.stato || "nuova";
  d.priorita.value = commessa.priorita || "";
  d.data_ingresso.value = formatDate(commessa.data_ingresso);
  if (d.data_ordine_telaio) d.data_ordine_telaio.value = formatDate(commessa.data_ordine_telaio);
  if (d.telaio_ordinato) {
    setTelaioOrdinatoButton(
      d.telaio_ordinato,
      Boolean(commessa.telaio_consegnato),
      false,
      d.data_consegna_telaio_effettiva || null
    );
  }
  if (d.data_consegna_telaio_effettiva) {
    d.data_consegna_telaio_effettiva.value = formatDate(commessa.data_consegna_telaio_effettiva);
  }
  if (d.data_arrivo_kit_cavi) {
    d.data_arrivo_kit_cavi.value = formatDate(commessa.data_arrivo_kit_cavi);
  }
  if (d.data_prelievo_materiali) {
    d.data_prelievo_materiali.value = formatDate(commessa.data_prelievo);
  }
  d.data_consegna.value = formatDate(commessa.data_consegna_macchina || commessa.data_consegna_prevista);
  d.note.value = commessa.note_generali || "";

  renderReparti(commessa);
  renderCommesse(state.commesse);
  updateNumeroWarning(d.anno, d.numero, detailNumeroWarning, commessa.id);
  setDetailSnapshot();
  if (options.openModal) openCommessaDetailModal();
}

const fetchProfileViaRest = async (userId) => {
  const stored = getStoredAuth();
  const token = stored?.access_token;
  if (!token) return null;
  const res = await fetch(`${SUPABASE_URL}/rest/v1/utenti?id=eq.${userId}&select=*`, {
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${token}`,
    },
  });
  if (!res.ok) {
    debugLog(`loadProfile REST error: ${res.status}`);
    return null;
  }
  const data = await res.json();
  return Array.isArray(data) ? data[0] || null : null;
};

const fetchTableViaRest = async (table, query = "") => {
  const stored = getStoredAuth();
  const token = stored?.access_token;
  if (!token) return null;
  const url = `${SUPABASE_URL}/rest/v1/${table}?${query}`;
  const res = await fetch(url, {
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${token}`,
    },
  });
  if (!res.ok) {
    debugLog(`REST ${table} error: ${res.status}`);
    return null;
  }
  const data = await res.json();
  return Array.isArray(data) ? data : [];
};

async function loadProfile(userFromSession = null) {
  let user = userFromSession;
  if (!user) {
    const {
      data: { user: fetchedUser },
    } = await supabase.auth.getUser();
    user = fetchedUser;
  }
  if (!user) return false;

  debugLog(`loadProfile query: ${user.id}`);
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("utenti").select("*").eq("id", user.id).single(),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadProfile timeout: ${err?.message || err}`);
    }
  }

  if (!data || error) {
    debugLog("loadProfile fallback REST");
    data = await fetchProfileViaRest(user.id);
  }

  if (!data) {
    debugLog(`loadProfile error: ${error?.message || "no data"}`);
    setStatus("Access denied: you're not whitelisted or profile not created.", "error");
    return false;
  }

  state.profile = data;
  debugLog(`loadProfile ok: ${data.email || data.id}`);
  setRoleBadge(data.ruolo);
  if (setPasswordBtn) {
    const role = String(data.ruolo || "").trim().toLowerCase();
    if (role === "admin") setPasswordBtn.classList.remove("hidden");
    else setPasswordBtn.classList.add("hidden");
  }
  if (openResourcesBtn) {
    const role = String(data.ruolo || "").trim().toLowerCase();
    if (role === "admin" || role === "responsabile") openResourcesBtn.classList.remove("hidden");
    else openResourcesBtn.classList.add("hidden");
  }
  if (permissionsLink) {
    const role = String(data.ruolo || "").trim().toLowerCase();
    if (role) permissionsLink.classList.remove("hidden");
    else permissionsLink.classList.add("hidden");
  }
  setWriteAccess(data.ruolo !== "viewer");
  authStatus.textContent = `Connesso come ${data.email}`;
  logoutBtn.classList.remove("hidden");
  if (authActions) authActions.classList.add("is-authenticated");
  if (authDot) authDot.classList.remove("hidden");
  return true;
}

async function loadReparti() {
  debugLog("loadReparti query");
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("reparti").select("*").order("nome"),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadReparti timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadReparti fallback REST");
    data = await fetchTableViaRest("reparti", "select=*&order=nome.asc");
  }
  if (!data) {
    setStatus(`Departments error: ${error?.message || "no data"}`, "error");
    return;
  }
  state.reparti = data || [];
  renderRepartiChecks();
  renderAddRepartoSelect();
  renderResourceRepartoSelect(resourceReparto);
  renderResourcesPanel();
}

async function loadRisorse() {
  debugLog("loadRisorse query");
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("risorse").select("*").order("nome"),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadRisorse timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadRisorse fallback REST");
    data = await fetchTableViaRest("risorse", "select=*&order=nome.asc");
  }
  if (!data) {
    setStatus(`Resources error: ${error?.message || "no data"}`, "error");
    return;
  }
  state.risorse = data || [];
  matrixState.risorse = (data || []).filter((r) => r.attiva);
  renderMatrix();
  renderMatrixReport();
  renderResourcesPanel();
}

async function loadPermessi() {
  debugLog("loadPermessi query");
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("permessi_ruolo").select("*"),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadPermessi timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadPermessi fallback REST");
    data = await fetchTableViaRest("permessi_ruolo", "select=*");
  }
  if (!data) {
    setStatus(`Permissions error: ${error?.message || "no data"}`, "error");
    state.permessiByRuolo = new Map();
    state.permessi = {};
    return;
  }
  state.permessiByRuolo = new Map(
    data.map((row) => [String(row.ruolo || "").toLowerCase(), row])
  );
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();
  state.permessi = state.permessiByRuolo.get(role) || {};
}

async function loadUtenti() {
  debugLog("loadUtenti query");
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("utenti").select("id,email").order("email"),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadUtenti timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadUtenti fallback REST");
    data = await fetchTableViaRest("utenti", "select=id,email&order=email.asc");
  }
  if (!data) {
    setStatus(`Users error: ${error?.message || "no data"}`, "error");
    return;
  }
  state.utenti = data || [];
  renderResourcesPanel();
}

async function loadCommesse() {
  debugLog("loadCommesse query");
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("v_commesse").select("*").order("created_at", { ascending: false }),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadCommesse timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadCommesse fallback REST");
    data = await fetchTableViaRest("v_commesse", "select=*&order=created_at.desc");
  }
  if (!data) {
    setStatus(`Commesse error: ${error?.message || "no data"}`, "error");
    return;
  }
  state.commesse = data || [];
  renderYearFilter();
  renderReportYearFilter();
  renderTodoYearFilter();
  applyFilters();
  renderMatrixCommesse();
  renderMatrixCommesseColumn();
  renderMatrixReport();
  if (state.selected) {
    const stillThere = state.commesse.find((c) => c.id === state.selected.id);
    if (stillThere) selectCommessa(stillThere.id);
    else clearSelection();
  }
}

function renderMatrixHeader(days) {
  matrixGrid.innerHTML = "";
  const header = document.createElement("div");
  header.className = "matrix-row matrix-header-row";
  header.style.gridTemplateColumns = `160px repeat(${days.length}, minmax(0, 1fr))`;

  const empty = document.createElement("div");
  empty.className = "matrix-cell matrix-header matrix-resource";
  empty.textContent = "Risorsa";
  header.appendChild(empty);

  const todayKey = formatDateLocal(new Date());
  days.forEach((d) => {
    const dayKey = formatDateLocal(d);
    const cell = document.createElement("div");
    cell.className = "matrix-cell matrix-header";
    if (dayKey === todayKey) {
      cell.classList.add("today-col");
    }
    if (d.getDay() === 5) {
      cell.classList.add("week-sep");
    }
    cell.textContent = d.toLocaleDateString("it-IT", { weekday: "short", day: "2-digit", month: "short" });
    header.appendChild(cell);
  });

  matrixGrid.appendChild(header);
  setupMatrixPan(header);
}

function setupMatrixPan(headerRow) {
  if (!headerRow || !matrixGrid) return;
  headerRow.onpointerdown = (e) => {
    if (e.button !== 0) return;
    if (e.target.closest("button, input, select, textarea")) return;
    e.preventDefault();
    const cells = headerRow.querySelectorAll(".matrix-cell");
    if (!cells || cells.length < 2) return;
    const dayRect = cells[1].getBoundingClientRect();
    const dayWidth = Math.max(1, dayRect.width);
    const panTargets = Array.from(
      matrixGrid.querySelectorAll(".matrix-cell:not(.matrix-resource), .matrix-activity-bar")
    );
    matrixPan = {
      startX: e.clientX,
      startY: e.clientY,
      baseDate: startOfWeek(matrixState.date),
      dayWidth,
      baseView: matrixState.view,
      moved: false,
      mode: null,
      panTargets,
    };
    headerRow.setPointerCapture?.(e.pointerId);
    matrixGrid.classList.add("is-panning");
  };
  headerRow.onpointermove = (e) => {
    if (!matrixPan || !matrixGrid) return;
    const dx = e.clientX - matrixPan.startX;
    const dy = e.clientY - matrixPan.startY;
    if (Math.abs(dx) > 3) matrixPan.moved = true;
    if (!matrixPan.mode) {
      if (Math.abs(dy) > 3 && Math.abs(dy) >= Math.abs(dx)) matrixPan.mode = "zoom";
      else if (Math.abs(dx) > 3) matrixPan.mode = "pan";
    }
    if (matrixPan.mode === "pan") {
      matrixPan.panTargets.forEach((el) => {
        el.style.transform = `translateX(${dx}px)`;
      });
    }
  };
  headerRow.onpointerup = async (e) => {
    if (!matrixPan) return;
    headerRow.releasePointerCapture?.(e.pointerId);
    const dx = e.clientX - matrixPan.startX;
    const dy = e.clientY - matrixPan.startY;
    const dayWidth = matrixPan.dayWidth;
    const shiftDays = dayWidth ? Math.round(-dx / dayWidth) : 0;
    matrixPan.panTargets.forEach((el) => {
      el.style.transform = "";
    });
    matrixGrid.classList.remove("is-panning");
    const mode = matrixPan.mode;
    const shouldMove = matrixPan.moved && mode === "pan" && shiftDays !== 0;
    matrixPan = null;
    if (mode === "zoom") {
      const zoomSteps = Math.round(dy / 18);
      if (zoomSteps === 0) return;
      const views = ["week", "two", "three", "six"];
      const currentIndex = Math.max(0, views.indexOf(matrixState.view));
      const nextIndex = Math.min(views.length - 1, Math.max(0, currentIndex + zoomSteps));
      matrixState.view = views[nextIndex];
      matrixState.date = startOfWeek(matrixState.date);
      setMatrixViewLabel();
      await loadMatrixAttivita();
      return;
    }
    if (!shouldMove) return;
    matrixState.date = addDays(matrixState.date, shiftDays);
    if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
    await loadMatrixAttivita();
  };
  headerRow.onpointerleave = () => {
    if (!matrixPan) return;
    matrixPan.panTargets.forEach((el) => {
      el.style.transform = "";
    });
    matrixGrid.classList.remove("is-panning");
    matrixPan = null;
  };
}

function renderMatrixCommesse() {
  if (!matrixCommessa) return;
  matrixCommessa.innerHTML = `<option value="">Commessa</option>`;
  state.commesse.slice().sort(compareCommesse).forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = `${c.codice}${c.titolo ? " - " + c.titolo : ""}`;
    matrixCommessa.appendChild(opt);
  });
}

function renderMatrixCommesseColumn() {
  if (!matrixCommesseList) return;
  const q = (matrixCommessaSearch.value || "").trim().toLowerCase();
  const items = state.commesse.filter((c) => {
    if (!q) return true;
    const code = String(c.codice || "").toLowerCase();
    const titolo = String(c.titolo || "").toLowerCase();
    const num = c.numero != null ? String(c.numero) : "";
    const isNumber = /^\d+$/.test(q);
    if (isNumber) {
      return num === q;
    }
    return titolo.includes(q) || code.includes(q);
  }).slice().sort(compareCommesse);
  matrixCommesseList.innerHTML = "";
  items.forEach((c) => {
    const div = document.createElement("div");
    div.className = "matrix-commessa-item";
    div.textContent = `${c.codice}${c.titolo ? " - " + c.titolo : ""}`;
    div.draggable = true;
    div.addEventListener("dragstart", (e) => {
      e.dataTransfer.setData("application/x-commessa-id", c.id);
    });
    matrixCommesseList.appendChild(div);
  });
}

function renderAssignActivities() {
  if (!assignActivities) return;
  assignActivities.innerHTML = "";
  const risorsaId = matrixState.pendingDrop ? matrixState.pendingDrop.risorsaId : null;
  const options = Array.from(matrixAttivita.options)
    .map((o) => o.value)
    .filter((v) => v && !isAssenteTitle(v))
    .filter((v) => !isLavenderActivity(v) || normalizePhaseKey(v) === "altro")
    .filter((v) => !risorsaId || isActivityAllowedForRisorsa(v, risorsaId));
  options.forEach((name) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "filter-pill";
    btn.textContent = name;
    btn.addEventListener("click", () => {
      btn.classList.toggle("active");
      if (assignAltroNoteWrap && assignAltroNote) {
        const hasAltro = Array.from(assignActivities.querySelectorAll(".filter-pill.active")).some(
          (el) => el.textContent.trim().toLowerCase() === "altro"
        );
        assignAltroNoteWrap.classList.toggle("hidden", !hasAltro);
        if (!hasAltro) assignAltroNote.value = "";
      }
      if (assignAssenzaWrap && assignAssenzaOre) {
        const hasAssente = Array.from(assignActivities.querySelectorAll(".filter-pill.active")).some((el) =>
          isAssenteTitle(el.textContent)
        );
        assignAssenzaWrap.classList.toggle("hidden", !hasAssente);
        if (!hasAssente) assignAssenzaOre.value = "";
      }
    });
    assignActivities.appendChild(btn);
  });
  if (assignAltroNoteWrap && assignAltroNote) {
    assignAltroNoteWrap.classList.add("hidden");
    assignAltroNote.value = "";
  }
  if (assignAssenzaWrap && assignAssenzaOre) {
    assignAssenzaWrap.classList.add("hidden");
    assignAssenzaOre.value = "";
  }
}

function renderMatrix() {
  if (!matrixGrid) return;
  if (!state.session) {
    matrixGrid.innerHTML = `<div class="matrix-empty">Accedi per vedere la matrice.</div>`;
    return;
  }

  setMatrixViewLabel();
  const matrixPanel = document.querySelector(".matrix-panel");
  const matrixPanelHeader = matrixPanel ? matrixPanel.querySelector(".panel-header") : null;
  if (matrixPanel && matrixPanelHeader) {
    matrixPanel.style.setProperty("--matrix-panel-header-height", `${matrixPanelHeader.offsetHeight}px`);
  }
  const days = weekDaysForView(matrixState.date, matrixState.view);
  renderMatrixHeader(days);

  if (!matrixState.risorse.length) {
    const empty = document.createElement("div");
    empty.className = "matrix-empty";
    empty.textContent = "Nessuna risorsa.";
    matrixGrid.appendChild(empty);
    return;
  }

  const dayKeys = days.map((d) => formatDateLocal(d));
  const dayIndex = new Map(dayKeys.map((key, idx) => [key, idx]));
  matrixState.dayIndex = dayIndex;

  const repartoById = new Map(state.reparti.map((rep) => [String(rep.id), rep.nome]));
  const labelForRisorsa = (r) => repartoById.get(String(r.reparto_id)) || "Senza reparto";
  const groups = new Map();
  matrixState.risorse.forEach((r) => {
    const label = labelForRisorsa(r);
    if (!groups.has(label)) groups.set(label, []);
    groups.get(label).push(r);
  });
  const desiredOrder = ["CAD", "TERMODINAMICI", "ELETTRICI"];
  const orderedLabels = [];
  desiredOrder.forEach((name) => {
    const match = Array.from(groups.keys()).find((k) => k.toUpperCase() === name);
    if (match) orderedLabels.push(match);
  });
  const remaining = Array.from(groups.keys())
    .filter((k) => !orderedLabels.includes(k))
    .sort((a, b) => a.localeCompare(b));
  orderedLabels.push(...remaining);

  orderedLabels.forEach((label) => {
    const sectionRow = document.createElement("div");
    sectionRow.className = "matrix-row matrix-section-row";
    sectionRow.style.gridTemplateColumns = `160px repeat(${days.length}, minmax(0, 1fr))`;
    const sectionCell = document.createElement("div");
    sectionCell.className = "matrix-section-cell";
    sectionCell.style.gridColumn = `1 / span ${days.length + 1}`;
    const isCollapsed = matrixState.collapsedReparti.has(label);
    sectionCell.innerHTML = `
      <button class="matrix-section-toggle" type="button" aria-expanded="${!isCollapsed}">
        ${isCollapsed ? "\u25B8" : "\u25BE"}
      </button>
      <span class="matrix-section-title">${label}</span>
    `;
    sectionCell.querySelector("button").addEventListener("click", () => {
      if (matrixState.collapsedReparti.has(label)) {
        matrixState.collapsedReparti.delete(label);
      } else {
        matrixState.collapsedReparti.add(label);
      }
      renderMatrix();
    });
    sectionRow.appendChild(sectionCell);
    matrixGrid.appendChild(sectionRow);

    if (isCollapsed) return;

    groups.get(label).forEach((r) => {
      const row = document.createElement("div");
      row.className = "matrix-row";
      row.style.gridTemplateColumns = `160px repeat(${days.length}, minmax(0, 1fr))`;
      row.style.position = "relative";

      const nameCell = document.createElement("div");
      nameCell.className = "matrix-cell matrix-resource";
      const reparto = state.reparti.find((rep) => String(rep.id) === String(r.reparto_id));
      nameCell.innerHTML = `
        <div class="resource-name">${r.nome}</div>
        <div class="resource-dept">${reparto ? reparto.nome : "-"}</div>
      `;
      row.appendChild(nameCell);

    const rowActivities = matrixState.attivita.filter((a) => a.risorsa_id === r.id);
    const filterAttivita = matrixState.selectedAttivita.size > 0;
    const filterCommesse = matrixState.selectedCommesse.size > 0;
    const visibleActivities = rowActivities.filter((a) => {
      if (filterAttivita && !matrixState.selectedAttivita.has(a.titolo)) return false;
      if (filterCommesse && !matrixState.selectedCommesse.has(a.commessa_id)) return false;
      return true;
    });

    const occupiedDays = new Set();
    visibleActivities.forEach((a) => {
      const start = startOfDay(new Date(a.data_inizio));
      const end = startOfDay(new Date(a.data_fine));
      let current = new Date(start);
      while (current <= end) {
        const key = formatDateLocal(current);
        if (dayIndex.has(key)) occupiedDays.add(key);
        current = addDays(current, 1);
      }
    });

    days.forEach((d) => {
      const dayKey = formatDateLocal(d);
      const cell = document.createElement("div");
      cell.className = "matrix-cell";
      cell.dataset.risorsa = r.id;
      cell.dataset.day = dayKey;
      if (dayKey === formatDateLocal(new Date())) {
        cell.classList.add("today-col");
      }
      if (d.getDay() === 5) {
        cell.classList.add("week-sep");
      }

      if (!occupiedDays.has(dayKey)) {
        // leave empty cell without placeholder
      }

      cell.addEventListener("dragover", (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
      });
      cell.addEventListener("dragenter", (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        cell.classList.add("drag-over");
      });
      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over");
      });
      cell.addEventListener("drop", async (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        cell.classList.remove("drag-over");
        await handleMatrixDropOnCell(cell, e);
      });

      cell.addEventListener("click", async () => {
        if (!state.canWrite) return;
        const commesseSel = getSelectedMatrixCommesse();
        const attivitaList = getSelectedMatrixAttivita();
        if (!attivitaList.length) {
          setStatus("Select at least one activity.", "error");
          return;
        }
        const hasAssente = attivitaList.some((t) => isAssenteTitle(t));
        if (hasAssente && attivitaList.length > 1) {
          setStatus("ASSENTE must be assigned alone.", "error");
          return;
        }
        if (!hasAssente && commesseSel.length === 0) {
          setStatus("Select a commessa.", "error");
          return;
        }
        if (!hasAssente && commesseSel.length > 1) {
          setStatus("Select a single commessa to assign.", "error");
          return;
        }
        const notAllowed = attivitaList.filter((t) => !isActivityAllowedForRisorsa(t, r.id));
        if (notAllowed.length) {
          setStatus("Some activities are not available for this department.", "error");
          return;
        }
        const commessaId = hasAssente ? null : commesseSel[0];
        const startAt = new Date(d);
        startAt.setHours(8, 0, 0, 0);
        const endAt = new Date(d);
        endAt.setHours(17, 0, 0, 0);
        if (matrixState.autoShift) {
          const hasOverlap = matrixState.attivita.some((a) => a.risorsa_id === r.id && overlapsDay(a, d));
          if (hasOverlap) {
            const totalDays = hasAssente ? 1 : attivitaList.length;
            const ok = await shiftRisorsa(r.id, d, totalDays);
            if (!ok) return;
          }
        }
        const rows = attivitaList.map((title) => ({
          commessa_id: isAssenteTitle(title) ? null : commessaId,
          titolo: title,
          risorsa_id: r.id,
          data_inizio: startAt.toISOString(),
          data_fine: endAt.toISOString(),
          stato: "pianificata",
        }));
        const { error } = await supabase.from("attivita").insert(rows);
        if (error) {
          setStatus(`Assignment error: ${error.message}`, "error");
          return;
        }
        setStatus(rows.length > 1 ? "Activities assigned." : "Activity assigned.", "ok");
        await loadMatrixAttivita();
        await refreshTodoRealtimeSnapshot();
      });

      row.appendChild(cell);
    });

    const barLayer = document.createElement("div");
    barLayer.className = "matrix-bar-layer";
    barLayer.style.gridTemplateColumns = row.style.gridTemplateColumns;
    barLayer.style.gridTemplateRows = "1fr";
    row.appendChild(barLayer);

    const viewStart = days[0];
    const viewEnd = days[days.length - 1];
    const lanes = [];
    const bars = [];
    const sorted = visibleActivities.slice().sort(compareAttivitaStable);
    sorted.forEach((a) => {
      const aStart = startOfDay(new Date(a.data_inizio));
      const aEnd = startOfDay(new Date(a.data_fine));
      if (aEnd < viewStart || aStart > viewEnd) return;
      let laneIndex = lanes.findIndex((end) => aStart > end);
      if (laneIndex === -1) {
        laneIndex = lanes.length;
        lanes.push(aEnd);
      } else {
        lanes[laneIndex] = aEnd;
      }
      const startKey = formatDateLocal(aStart < viewStart ? viewStart : aStart);
      const endKey = formatDateLocal(aEnd > viewEnd ? viewEnd : aEnd);
      const startIdx = dayIndex.get(startKey);
      const endIdx = dayIndex.get(endKey);
      if (startIdx == null || endIdx == null) return;
      bars.push({ a, laneIndex, startIdx, endIdx, startKey });
    });

    const laneHeight = 54;
    const laneOffset = 6;
    if (lanes.length) {
      row.style.minHeight = `${70 + laneOffset + Math.max(0, lanes.length - 1) * laneHeight}px`;
    }

    bars.forEach((b) => {
      const bar = document.createElement("div");
      bar.className = "matrix-activity-bar";
      bar.style.gridColumn = `${2 + b.startIdx} / ${2 + b.endIdx + 1}`;
      bar.style.marginTop = `${laneOffset + b.laneIndex * laneHeight}px`;
      const commessa = state.commesse.find((c) => c.id === b.a.commessa_id);
      const titleKey = (b.a.titolo || "").trim().toLowerCase();
      const displayTitle =
        titleKey === "altro" && b.a.descrizione
          ? `Altro: ${b.a.descrizione}`
          : isAssenteTitle(b.a.titolo)
          ? `Assente${b.a.ore_assenza ? ` (${b.a.ore_assenza}h)` : ""}`
          : b.a.titolo;
      if (titleKey === "preliminare") bar.classList.add("activity-preliminare");
      if (titleKey === "3d") bar.classList.add("activity-3d");
      if (isAssenteTitle(b.a.titolo)) bar.classList.add("activity-assente");
      const isLavender = !b.a.commessa_id && isLavenderActivity(b.a.titolo);
      if (isLavender) bar.classList.add("activity-lavender");
      const code = commessa ? commessa.codice || "" : "";
      const parsed = parseCommessaCode(code);
      const commessaHtml =
        parsed.year != null && parsed.num != null
          ? `<strong class="commessa-num">${parsed.num}</strong><span class="commessa-year">${parsed.year}</span>`
          : code
          ? `<strong class="commessa-num">${code}</strong>`
          : "";
      bar.innerHTML = commessaHtml
        ? `
        <div class="commessa-line">${commessaHtml}</div>
        <span class="activity-title">${displayTitle || ""}</span>
      `
        : `<strong class="commessa-num">${displayTitle || ""}</strong>`;
      bar.dataset.commessaId = b.a.commessa_id ? String(b.a.commessa_id) : "";
      bar.dataset.activityId = b.a.id ? String(b.a.id) : "";
      bar.dataset.risorsaId = b.a.risorsa_id ? String(b.a.risorsa_id) : "";
      bar.dataset.attivita = b.a.titolo || "";
      bar.dataset.reparto = getRisorsaDeptName(b.a.risorsa_id) || "";
      bar.classList.toggle("is-done", b.a.stato === "completata");
      applyMatrixBarClientFlowClasses(bar, b.a, bar.dataset.reparto || "");
      if (matrixState.colorMode !== "none" && !isLavender) {
        const key =
          matrixState.colorMode === "activity"
            ? b.a.titolo || "attivita"
            : commessa
            ? commessa.codice || commessa.id
            : b.a.commessa_id || "commessa";
        const colors = colorForKey(key);
        bar.classList.add("colorized");
        bar.style.setProperty("--matrix-activity-bg", colors.bg);
        bar.style.setProperty("--matrix-activity-border", colors.border);
      } else {
        bar.classList.remove("colorized");
        bar.style.removeProperty("--matrix-activity-bg");
        bar.style.removeProperty("--matrix-activity-border");
      }
      bar.draggable = true;
      bar.addEventListener("dragstart", (e) => {
        if (!state.canWrite) return;
        if (!canMoveMatrixActivity(b.a, b.a.risorsa_id)) {
          setStatus("Non autorizzato a spostare attivita.", "error");
          e.preventDefault();
          return;
        }
        if (matrixState.resizing) {
          e.preventDefault();
          return;
        }
        cancelCommessaHighlightTimer();
        closeMatrixQuickMenu();
        matrixState.suppressClickUntil = Date.now() + 300;
        matrixState.draggingId = b.a.id;
        if (matrixGrid) matrixGrid.classList.add("matrix-dragging");
        bar.classList.add("dragging");
        e.dataTransfer.setData("text/plain", b.a.id);
      });
      bar.addEventListener("dragend", () => {
        matrixState.draggingId = null;
        if (matrixGrid) matrixGrid.classList.remove("matrix-dragging");
        bar.classList.remove("dragging");
      });
      bar.addEventListener("mousedown", (e) => {
        if (!state.canWrite) return;
        if (matrixState.resizing) return;
        if (e.button !== 0) return;
        cancelCommessaHighlightTimer();
        closeMatrixQuickMenu();
        commessaHighlightTimer = setTimeout(() => {
          commessaHighlightTimer = null;
          if (matrixState.draggingId) return;
          closeTodoStatusMenu();
          closeTodoGlobalStatusMenu();
          clearCommessaHighlight();
          openActivityModal(b.a);
          matrixState.suppressClickUntil = Date.now() + 600;
        }, 450);
      });
      bar.addEventListener("mouseup", () => {
        cancelCommessaHighlightTimer();
      });
      bar.addEventListener("mouseleave", () => {
        cancelCommessaHighlightTimer();
      });
      bar.addEventListener("dragover", (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
      });
      bar.addEventListener("drop", async (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        e.stopPropagation();
        const elements = document.elementsFromPoint(e.clientX, e.clientY);
        const cell = elements.find((el) => el.classList && el.classList.contains("matrix-cell"));
        if (cell) {
          await handleMatrixDropOnCell(cell, e);
        }
      });
      bar.addEventListener("click", async (e) => {
        e.stopPropagation();
        if (!state.canWrite) return;
        if (Date.now() < matrixState.suppressClickUntil) return;
        if (commessaLongPressSuppress) {
          commessaLongPressSuppress = false;
          return;
        }
        closeMatrixQuickMenu();
        clearCommessaHighlight();
        if (!b.a.commessa_id || !b.a.titolo) {
          openActivityModal(b.a);
          return;
        }
        mergeReportActivitiesFromMatrix(b.a.commessa_id);
        openTodoStatusMenu(bar, e.clientX, e.clientY);
      });
      const handle = document.createElement("span");
      handle.className = "matrix-resize-handle";
      handle.title = "Trascina per estendere";
      handle.addEventListener("mousedown", (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        e.stopPropagation();
        matrixState.suppressClickUntil = Date.now() + 400;
        const rowEl = bar.closest(".matrix-row");
        const layerEl = bar.closest(".matrix-bar-layer");
        const previewBar = bar.cloneNode(true);
        previewBar.classList.add("resizing-preview");
        previewBar.classList.remove("dragging");
        previewBar.style.pointerEvents = "none";
        if (layerEl) {
          const layerRect = layerEl.getBoundingClientRect();
          const barRect = bar.getBoundingClientRect();
          previewBar.style.position = "absolute";
          previewBar.style.top = `${barRect.top - layerRect.top}px`;
          previewBar.style.left = `${barRect.left - layerRect.left}px`;
          previewBar.style.width = `${barRect.width}px`;
          previewBar.style.height = `${barRect.height}px`;
          previewBar.style.marginTop = "0";
          previewBar.style.gridColumn = "auto";
        }
        if (layerEl) layerEl.appendChild(previewBar);
        const startDate = startOfDay(new Date(b.a.data_inizio));
        const endDate = new Date(b.a.data_fine);
        const startCell = rowEl ? rowEl.querySelector(`.matrix-cell[data-day="${b.startKey}"]`) : null;
        const startCellRect = startCell ? startCell.getBoundingClientRect() : null;
        const layerRect = layerEl ? layerEl.getBoundingClientRect() : null;
        matrixState.resizing = {
          attivita: b.a,
          startDayKey: b.startKey,
          startDate,
          oldEndDay: startOfDay(new Date(b.a.data_fine)),
          endTime: {
            h: endDate.getHours(),
            m: endDate.getMinutes(),
            s: endDate.getSeconds(),
            ms: endDate.getMilliseconds(),
          },
          targetDayKey: b.startKey,
          startIdx: b.startIdx,
          rowEl,
          layerEl,
          layerRect,
          startCellRect,
          previewBar,
          originalBar: bar,
        };
        bar.classList.add("resizing");
      });
      bar.appendChild(handle);
      barLayer.appendChild(bar);
    });

      matrixGrid.appendChild(row);
    });
  });
}

function renderMatrixReport() {
  if (!matrixReportGrid || !matrixReportRange) return;
  if (!state.session) {
    matrixReportGrid.innerHTML = `<div class="matrix-empty">Accedi per vedere la reportistica.</div>`;
    matrixReportRange.textContent = "";
    return;
  }
  if (reportPanel && reportPanel.classList.contains("collapsed")) {
    return;
  }
  const today = startOfDay(new Date());
  matrixReportRange.textContent = `Oggi: ${formatDateLocal(today)}`;

  const filters = getReportFilters();
  const todoCommesse = state.commesse.slice().sort(compareCommesseDesc);
  const useServerFilter = Boolean(filters.desc || filters.year || (filters.numVal != null && Number.isFinite(filters.numVal)));
  let commesse = [];
  if (useServerFilter) {
    const key = JSON.stringify({ desc: filters.desc, year: filters.year, numVal: filters.numVal });
    if (state.reportFilteredKey !== key) {
      state.reportFilteredKey = key;
      state.reportFilteredCommesse = null;
      loadReportFilteredCommesse(filters);
    }
    if (state.reportFilteredLoading || !state.reportFilteredCommesse) {
      matrixReportGrid.innerHTML = `<div class="matrix-empty">Caricamento reportistica...</div>`;
      return;
    }
    commesse = applyReportFilters(state.reportFilteredCommesse, filters);
  } else {
    commesse = applyReportFilters(state.commesse, filters);
  }
  const showYearHeaders = !filters.orderDue;

  if (!commesse.length) {
    matrixReportGrid.innerHTML = `<div class="matrix-empty">Nessuna commessa trovata.</div>`;
    renderTodoSection(todoCommesse);
    return;
  }

  if (commesse.length > REPORT_MAX_ITEMS) {
    matrixReportGrid.innerHTML =
      `<div class="matrix-empty">Troppe commesse (${commesse.length}). ` +
      `Applica un filtro per visualizzare la reportistica.</div>`;
    renderTodoSection(todoCommesse);
    return;
  }

  commesse = sortReportCommesse(commesse, filters.orderDue);
  const missingSchedules = commesse.some((c) => !state.reportActivitiesMap.has(String(c.id)));
  if (missingSchedules && !state.reportActivitiesLoading) {
    loadReportActivitiesFor(commesse);
  }
  const missingOverrides = commesse.some((c) => !state.todoOverridesMap.has(String(c.id)));
  if (missingOverrides && !state.todoOverridesLoading) {
    loadTodoOverridesFor(commesse);
  }
  const missingClientFlows = commesse.some((c) => !state.todoClientFlowLoadedCommesse.has(String(c.id)));
  if (missingClientFlows && !state.todoClientFlowLoading) {
    loadTodoClientFlowsFor(commesse);
  }
  if (todoCommesse.length && todoCommesse.length <= REPORT_MAX_ITEMS) {
    const todoMissingSchedules = todoCommesse.some((c) => !state.reportActivitiesMap.has(String(c.id)));
    if (todoMissingSchedules && !state.reportActivitiesLoading) {
      loadReportActivitiesFor(todoCommesse);
    }
    const todoMissingOverrides = todoCommesse.some((c) => !state.todoOverridesMap.has(String(c.id)));
    if (todoMissingOverrides && !state.todoOverridesLoading) {
      loadTodoOverridesFor(todoCommesse);
    }
    const todoMissingClientFlows = todoCommesse.some((c) => !state.todoClientFlowLoadedCommesse.has(String(c.id)));
    if (todoMissingClientFlows && !state.todoClientFlowLoading) {
      loadTodoClientFlowsFor(todoCommesse);
    }
  }
  renderTodoSection(todoCommesse);
  if (state.reportView === "gantt") {
    renderReportGantt(commesse, today);
    return;
  }

  matrixReportGrid.classList.remove("report-gantt");
  matrixReportGrid.classList.add("report-list");
  matrixReportGrid.innerHTML = "";

  let lastYear = null;
  commesse.forEach((c) => {
    const yearKey = String(c.anno || getCommessaYear(c) || "Senza anno");
    if (showYearHeaders && yearKey !== lastYear) {
      const header = document.createElement("div");
      header.className = "report-year-header";
      header.textContent = yearKey;
      matrixReportGrid.appendChild(header);
      lastYear = yearKey;
    }

    const targetRaw = c.data_consegna_macchina || c.data_consegna_prevista || c.data_consegna || "";
    const targetDate = targetRaw ? new Date(targetRaw) : null;
    const ordineRaw = c.data_ordine_telaio || "";
    const kitRaw = c.data_arrivo_kit_cavi || "";
    const ordineDate = ordineRaw ? new Date(ordineRaw) : null;
    const kitDate = kitRaw ? new Date(kitRaw) : null;
    const prelievoRaw = c.data_prelievo || "";
    const prelievoDate = prelievoRaw ? new Date(prelievoRaw) : null;
    const isDateOk = (d) => d && !Number.isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
    const safeTarget = isDateOk(targetDate) ? targetDate : null;
    const safeOrdine = isDateOk(ordineDate) ? ordineDate : null;
    const safeKit = isDateOk(kitDate) ? kitDate : null;
    const safePrelievo = isDateOk(prelievoDate) ? prelievoDate : null;
    const targetInfo = getBusinessDiffInfo(today, safeTarget);
    const ordineInfo = getBusinessDiffInfo(today, safeOrdine);
    const missingPrelievo = !safePrelievo;
    const missingKit = !safeKit;
    const missingOrdine = !safeOrdine || missingPrelievo;
    const missingTarget = !safeTarget;

    const item = document.createElement("div");
    item.className = "report-item";
    const scheduleHtml = buildReportScheduleHtml(c.id);
    item.innerHTML = `
      <div class="report-main">
        <button class="report-title report-commessa-link" type="button" data-commessa-id="${c.id}">
          ${c.codice} - ${c.titolo || "Senza titolo"}
        </button>
        <div class="report-meta">Stato: ${c.stato || "-"} - Cliente: ${c.cliente || "-"}</div>
        ${scheduleHtml}
        ${
          missingOrdine || missingTarget
            ? `<div class="report-missing">
                ${!safeOrdine ? `<span class="missing-pill">Manca data ordine telaio</span>` : ""}
                ${missingKit ? `<span class="missing-pill">Manca data arrivo kit cavi</span>` : ""}
                ${missingPrelievo ? `<span class="missing-pill">Manca data prelievo materiali</span>` : ""}
                ${missingTarget ? `<span class="missing-pill">Manca data consegna macchina</span>` : ""}
              </div>`
            : ""
        }
      </div>
      <div class="report-milestone ${ordineInfo.cls}">
        <div class="report-label">Ordine telaio</div>
        <div class="report-value">${safeOrdine ? formatDateDMY(safeOrdine) : "-"}</div>
        <div class="report-sub">${ordineInfo.label}</div>
      </div>
      <div class="report-milestone ${targetInfo.cls}">
        <div class="report-label">Target consegna</div>
        <div class="report-value">${safeTarget ? formatDateDMY(safeTarget) : "-"}</div>
        <div class="report-sub">${targetInfo.label}</div>
      </div>
    `;
    const commessaBtn = item.querySelector(".report-commessa-link");
    if (commessaBtn) {
      commessaBtn.addEventListener("click", () => {
        selectCommessa(c.id, { openModal: true });
      });
    }
    matrixReportGrid.appendChild(item);
  });
}

function getReportFilteredCommesse() {
  const filters = getReportFilters();
  return applyReportFilters(state.commesse.slice(), filters);
}

function getReportFilters() {
  const desc = reportDescFilter ? reportDescFilter.value.trim().toLowerCase() : "";
  const numRaw = reportNumberFilter ? reportNumberFilter.value.trim() : "";
  const numVal = numRaw ? Number.parseInt(numRaw, 10) : null;
  const year = reportYearFilter ? reportYearFilter.value : "";
  const hidePast = reportHidePast ? reportHidePast.classList.contains("active") : false;
  const orderDue = reportOrderDue ? reportOrderDue.classList.contains("active") : false;
  return { desc, numVal, year, hidePast, orderDue };
}

function applyReportFilters(list, filters) {
  const { desc, numVal, year, hidePast } = filters;
  const today = startOfDay(new Date());
  const getTargetDate = (c) => {
    const raw = c.data_consegna_macchina || c.data_consegna_prevista || c.data_consegna || "";
    const d = raw ? new Date(raw) : null;
    if (!d || Number.isNaN(d.getTime())) return null;
    if (d.getFullYear() < 2000 || d.getFullYear() > 2100) return null;
    return startOfDay(d);
  };

  let commesse = list.slice();
  if (hidePast) {
    commesse = commesse.filter((c) => {
      if (c.telaio_consegnato) return false;
      const target = getTargetDate(c);
      if (!target) return true;
      return target >= today;
    });
  }
  if (year) {
    commesse = commesse.filter((c) => String(c.anno || getCommessaYear(c) || "") === year);
  }
  if (numVal != null && Number.isFinite(numVal)) {
    commesse = commesse.filter((c) => Number(c.numero ?? -1) === numVal);
  }
  if (desc) {
    commesse = commesse.filter((c) => `${c.titolo || ""}`.toLowerCase().includes(desc));
  }
  return commesse;
}

function getReportSortData(c) {
  const rawOrdine = c.data_ordine_telaio || "";
  const ordineDate = rawOrdine ? new Date(rawOrdine) : null;
  const rawConsegna = c.data_consegna_macchina || c.data_consegna_prevista || c.data_consegna || "";
  const consegnaDate = rawConsegna ? new Date(rawConsegna) : null;
  const isDateOk = (d) => d && !Number.isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
  const ordine = isDateOk(ordineDate) ? ordineDate : null;
  const consegna = isDateOk(consegnaDate) ? consegnaDate : null;
  if (ordine) {
    return { group: 0, date: ordine, orderKey: ordine.getTime() };
  }
  if (consegna) {
    return { group: 1, date: consegna, orderKey: consegna.getTime() };
  }
  return { group: 2, date: null, orderKey: Number.POSITIVE_INFINITY };
}

function sortReportCommesse(list, orderDueActive) {
  if (!orderDueActive) return list.slice().sort(compareCommesseDesc);
  return list.slice().sort((a, b) => {
    const aa = getReportSortData(a);
    const bb = getReportSortData(b);
    if (aa.group !== bb.group) return aa.group - bb.group;
    if (aa.orderKey !== bb.orderKey) return aa.orderKey - bb.orderKey;
    return compareCommesseDesc(a, b);
  });
}

async function loadReportFilteredCommesse(filters) {
  const { desc, numVal, year } = filters;
  const token = ++state.reportFilteredToken;
  state.reportFilteredLoading = true;
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      let query = supabase.from("v_commesse").select("*");
      if (year) query = query.eq("anno", Number(year));
      if (numVal != null && Number.isFinite(numVal)) query = query.eq("numero", numVal);
      if (desc) query = query.ilike("titolo", `%${desc}%`);
      const result = await withTimeout(query.order("created_at", { ascending: false }), FETCH_TIMEOUT_MS);
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadReportFilteredCommesse timeout: ${err?.message || err}`);
    }
  }
  if (!data || error) {
    debugLog("loadReportFilteredCommesse fallback REST");
    const parts = ["select=*"];
    if (year) parts.push(`anno=eq.${encodeURIComponent(year)}`);
    if (numVal != null && Number.isFinite(numVal)) parts.push(`numero=eq.${encodeURIComponent(numVal)}`);
    if (desc) parts.push(`titolo=ilike.*${encodeURIComponent(desc)}*`);
    parts.push("order=created_at.desc");
    data = await fetchTableViaRest("v_commesse", parts.join("&"));
  }
  if (token !== state.reportFilteredToken) return;
  state.reportFilteredLoading = false;
  if (!data || error) {
    setStatus(`Report error: ${error?.message || "no data"}`, "error");
    state.reportFilteredCommesse = [];
    return;
  }
  state.reportFilteredCommesse = data || [];
  renderMatrixReport();
}

function buildReportScheduleHtml(commessaId) {
  const key = String(commessaId || "");
  if (!key) return "";
  if (!state.reportActivitiesMap.has(key)) {
    return `<div class="report-schedule report-schedule-loading">Schedulazioni in caricamento...</div>`;
  }
  const entries = state.reportActivitiesMap.get(key) || [];
  if (!entries.length) {
    return `<div class="report-schedule report-schedule-empty report-pill report-pill-orange">Nessuna attivita schedulata.</div>`;
  }
  const lines = entries
    .map((e) => {
      const start = e.start ? formatDateLocal(e.start) : "-";
      const end = e.end ? formatDateLocal(e.end) : "-";
      return `<div class="report-schedule-item">${e.titolo} \u2014 ${e.risorsa} \u00B7 ${start} \u2192 ${end}</div>`;
    })
    .join("");
  return `<div class="report-schedule">${lines}</div>`;
}

async function loadReportActivitiesFor(commesse) {
  if (!commesse || !commesse.length) {
    state.reportActivitiesMap = new Map();
    return;
  }
  if (commesse.length > REPORT_MAX_ITEMS) {
    return;
  }
  const token = ++state.reportActivitiesToken;
  state.reportActivitiesLoading = true;
  const ids = Array.from(new Set(commesse.map((c) => c.id))).filter(Boolean);
  if (!ids.length) {
    state.reportActivitiesLoading = false;
    return;
  }

  let rows = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
          supabase
            .from("attivita")
            .select("id, commessa_id, titolo, risorsa_id, reparto_id, assegnato_a, data_inizio, data_fine, stato")
            .in("commessa_id", ids),
        FETCH_TIMEOUT_MS
      );
      rows = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadReportActivities timeout: ${err?.message || err}`);
    }
  }

  if (!rows || error) {
    debugLog("loadReportActivities fallback REST");
    const inList = ids.join(",");
      rows = await fetchTableViaRest(
        "attivita",
        `select=id,commessa_id,titolo,risorsa_id,reparto_id,assegnato_a,data_inizio,data_fine,stato&commessa_id=in.(${inList})`
      );
  }

  if (!rows || token !== state.reportActivitiesToken) return;

  const risorseById = new Map((state.risorse || []).map((r) => [String(r.id), r.nome]));
  const map = new Map();
  ids.forEach((id) => {
    map.set(String(id), []);
  });
  rows.forEach((row) => {
    if (!row.commessa_id) return;
    const key = String(row.commessa_id);
    const list = map.get(key) || [];
    const dept =
      row.risorsa_id != null
        ? getDeptBucketByRisorsa(row.risorsa_id)
        : row.reparto_id != null
        ? getDeptBucketByRepartoId(row.reparto_id)
        : "ALTRO";
      list.push({
        id: row.id,
        risorsaId: row.risorsa_id,
        assegnatoA: row.assegnato_a || null,
        titolo: row.titolo || "Attivita",
        risorsa:
          row.risorsa_id != null
            ? risorseById.get(String(row.risorsa_id)) || `Risorsa ${row.risorsa_id}`
            : "Risorsa n/d",
        dept,
        start: row.data_inizio ? new Date(row.data_inizio) : null,
        end: row.data_fine ? new Date(row.data_fine) : null,
        stato: row.stato || "pianificata",
      });
    map.set(key, list);
  });
  map.forEach((list, key) => {
    list.sort((a, b) => {
      const aTime = a.start ? a.start.getTime() : 0;
      const bTime = b.start ? b.start.getTime() : 0;
      if (aTime !== bTime) return aTime - bTime;
      return String(a.titolo || "").localeCompare(String(b.titolo || ""));
    });
    map.set(key, list);
  });

  if (token !== state.reportActivitiesToken) return;
  const merged = new Map(state.reportActivitiesMap);
  map.forEach((value, key) => {
    merged.set(String(key), value);
  });
  state.reportActivitiesMap = merged;
  state.reportActivitiesLoading = false;
  renderMatrixReport();
}

async function loadTodoOverridesFor(commesse) {
  if (!commesse || !commesse.length) {
    state.todoOverridesMap = new Map();
    return;
  }
  if (commesse.length > REPORT_MAX_ITEMS) {
    return;
  }
  const token = ++state.todoOverridesToken;
  state.todoOverridesLoading = true;
  const ids = Array.from(new Set(commesse.map((c) => c.id))).filter(Boolean);
  if (!ids.length) {
    state.todoOverridesLoading = false;
    return;
  }

  let rows = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase
          .from("commessa_attivita_override")
          .select("commessa_id, titolo, stato")
          .in("commessa_id", ids),
        FETCH_TIMEOUT_MS
      );
      rows = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadTodoOverrides timeout: ${err?.message || err}`);
    }
  }

  if (!rows || error) {
    debugLog("loadTodoOverrides fallback REST");
    const inList = ids.join(",");
    rows = await fetchTableViaRest(
      "commessa_attivita_override",
      `select=commessa_id,titolo,stato&commessa_id=in.(${inList})`
    );
  }

  if (!rows || token !== state.todoOverridesToken) return;

  const map = new Map();
  ids.forEach((id) => {
    map.set(String(id), new Map());
  });
  rows.forEach((row) => {
    if (!row.commessa_id || !row.titolo) return;
    const key = String(row.commessa_id);
    const list = map.get(key) || new Map();
    list.set(normalizePhaseKey(row.titolo), row.stato || "non_necessaria");
    map.set(key, list);
  });
  const merged = new Map(state.todoOverridesMap);
  map.forEach((value, key) => {
    merged.set(String(key), value);
  });
  state.todoOverridesMap = merged;
  state.todoOverridesLoading = false;
  renderTodoSection(todoLastRendered);
}

async function loadTodoClientFlowsFor(commesse) {
  if (!commesse || !commesse.length) {
    state.todoClientFlowMap = new Map();
    state.todoClientFlowLoadedCommesse = new Set();
    return;
  }
  if (commesse.length > REPORT_MAX_ITEMS) {
    return;
  }
  const token = ++state.todoClientFlowToken;
  state.todoClientFlowLoading = true;
  const ids = Array.from(new Set(commesse.map((c) => c.id))).filter(Boolean);
  if (!ids.length) {
    state.todoClientFlowLoading = false;
    return;
  }

  let rows = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase
          .from("commessa_attivita_cliente")
          .select("commessa_id,titolo,reparto,inviato_il,scadenza_il,esito,confermato_il,updated_at")
          .in("commessa_id", ids),
        FETCH_TIMEOUT_MS
      );
      rows = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadTodoClientFlows timeout: ${err?.message || err}`);
    }
  }

  if (!rows || error) {
    debugLog("loadTodoClientFlows fallback REST");
    const inList = ids.join(",");
    rows = await fetchTableViaRest(
      "commessa_attivita_cliente",
      `select=commessa_id,titolo,reparto,inviato_il,scadenza_il,esito,confermato_il,updated_at&commessa_id=in.(${inList})`
    );
  }

  if (!rows || token !== state.todoClientFlowToken) {
    if (token === state.todoClientFlowToken) {
      const loaded = new Set(state.todoClientFlowLoadedCommesse || []);
      ids.forEach((id) => loaded.add(String(id)));
      state.todoClientFlowLoadedCommesse = loaded;
      state.todoClientFlowLoading = false;
    }
    return;
  }

  const idSet = new Set(ids.map((id) => String(id)));
  const merged = new Map(state.todoClientFlowMap);
  Array.from(merged.keys()).forEach((key) => {
    const commessaKey = String(key).split("|")[0];
    if (idSet.has(commessaKey)) merged.delete(key);
  });
  rows.forEach((row) => {
    const normalized = normalizeTodoClientFlowRow(row);
    if (!normalized) return;
    const key = buildTodoClientFlowKey(normalized.commessaId, normalized.titolo, normalized.reparto);
    merged.set(key, normalized);
  });
  state.todoClientFlowMap = merged;
  const loaded = new Set(state.todoClientFlowLoadedCommesse || []);
  ids.forEach((id) => loaded.add(String(id)));
  state.todoClientFlowLoadedCommesse = loaded;
  state.todoClientFlowLoading = false;
  renderTodoSection(todoLastRendered);
}

function getTodoOverrideStatus(commessaId, title) {
  const key = String(commessaId);
  const list = state.todoOverridesMap.get(key);
  if (!list) return null;
  return list.get(normalizePhaseKey(title)) || null;
}

const TODO_CLIENT_FLOW_DEFAULT_DEADLINE_BUSINESS_DAYS = 5;
const TODO_CLIENT_FLOW_OUTCOME = {
  IN_ATTESA: "in_attesa",
  CONFERMATO: "confermato",
  SILENZIO_ASSENSO: "silenzio_assenso",
  RESPINTO: "respinto",
};

function buildTodoClientFlowKey(commessaId, titolo, repartoName = "") {
  return `${String(commessaId || "")}|${normalizePhaseKey(titolo || "")}|${normalizeDeptKey(repartoName || "")}`;
}

const TODO_CLIENT_FLOW_TRACKED_TITLES = new Set(["preliminare", "3d"]);

function isTodoClientFlowTracked(titolo, repartoName = "") {
  const titleKey = normalizePhaseKey(titolo || "");
  const deptKey = normalizeDeptKey(repartoName || "");
  return TODO_CLIENT_FLOW_TRACKED_TITLES.has(titleKey) && deptKey.includes("CAD");
}

function getTodoClientFlowSectionTitle(titolo = "") {
  const cleanTitle = String(titolo || "").trim();
  return cleanTitle ? `Cliente ${cleanTitle}` : "Cliente";
}

function parseIsoDateOnly(value) {
  if (!value) return null;
  if (value instanceof Date) {
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return null;
    return startOfDay(d);
  }
  const text = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  const parsed = new Date(`${text}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return null;
  return startOfDay(parsed);
}

function formatIsoDateOnly(value) {
  const d = parseIsoDateOnly(value);
  if (!d) return null;
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getTodoClientFlow(commessaId, titolo, repartoName = "") {
  const key = buildTodoClientFlowKey(commessaId, titolo, repartoName);
  return state.todoClientFlowMap.get(key) || null;
}

function normalizeTodoClientFlowRow(row) {
  if (!row?.commessa_id || !row?.titolo) return null;
  const reparto = row.reparto || "";
  const sentAt = parseIsoDateOnly(row.inviato_il);
  const dueAt = parseIsoDateOnly(row.scadenza_il);
  const confirmedAt = parseIsoDateOnly(row.confermato_il);
  const outcome = normalizePhaseKey(row.esito || "") || TODO_CLIENT_FLOW_OUTCOME.IN_ATTESA;
  return {
    commessaId: row.commessa_id,
    titolo: row.titolo,
    reparto,
    sentAt,
    dueAt,
    confirmedAt,
    outcome,
    updatedAt: row.updated_at || null,
  };
}

function getTodoClientFlowResolvedStatus(flow, referenceDate = startOfDay(new Date())) {
  if (!flow || !flow.sentAt) return "non_inviato";
  if (flow.outcome === TODO_CLIENT_FLOW_OUTCOME.CONFERMATO) return "confermato";
  if (flow.outcome === TODO_CLIENT_FLOW_OUTCOME.SILENZIO_ASSENSO) return "silenzio_assenso";
  if (flow.outcome === TODO_CLIENT_FLOW_OUTCOME.RESPINTO) return "respinto";
  if (flow.dueAt && businessDayDiff(startOfDay(referenceDate), startOfDay(flow.dueAt)) < 0) return "scaduto";
  return "in_attesa";
}

function getTodoClientFlowAgingDays(flow, referenceDate = startOfDay(new Date())) {
  if (!flow?.sentAt) return null;
  return Math.max(0, dayNumberUTC(referenceDate) - dayNumberUTC(flow.sentAt));
}

function getTodoClientFlowRemainingDays(flow, referenceDate = startOfDay(new Date())) {
  if (!flow?.dueAt) return null;
  return businessDayDiff(startOfDay(referenceDate), startOfDay(flow.dueAt));
}

function getTodoClientFlowMeta(flow, referenceDate = startOfDay(new Date())) {
  if (!flow || !flow.sentAt) return null;
  const status = getTodoClientFlowResolvedStatus(flow, referenceDate);
  const waitDays = getTodoClientFlowAgingDays(flow, referenceDate);
  const remainingDays = getTodoClientFlowRemainingDays(flow, referenceDate);
  const sentLabel = formatDateDMY(flow.sentAt);
  if (status === "confermato") {
    const confLabel = flow.confirmedAt ? formatDateDMY(flow.confirmedAt) : "data n/d";
    return { text: `Confermato ${confLabel} · Inviato ${sentLabel}`, tone: "ok" };
  }
  if (status === "silenzio_assenso") {
    const decisionLabel = flow.confirmedAt ? formatDateDMY(flow.confirmedAt) : flow.dueAt ? formatDateDMY(flow.dueAt) : "data n/d";
    return { text: `Silenzio assenso ${decisionLabel} · Inviato ${sentLabel}`, tone: "warn" };
  }
  if (status === "respinto") {
    return { text: `Inviato ${sentLabel} · Respinto`, tone: "danger" };
  }
  if (status === "scaduto") {
    const overdueDays = remainingDays != null ? Math.max(0, -remainingDays) : null;
    return {
      text: `Inviato ${sentLabel} · Scaduto${overdueDays != null ? ` da ${overdueDays}g` : ""}`,
      tone: "danger",
    };
  }
  const dueLabel = flow.dueAt ? formatDateDMY(flow.dueAt) : "scadenza n/d";
  if (remainingDays === 0) {
    return { text: `Inviato ${sentLabel} · Restano 0g · Scad. ${dueLabel}`, tone: "warn" };
  }
  if (remainingDays != null && remainingDays > 0) {
    return { text: `Inviato ${sentLabel} · Restano ${remainingDays}g · Scad. ${dueLabel}`, tone: "" };
  }
  return { text: `Inviato ${sentLabel} · In attesa ${waitDays}g · Scad. ${dueLabel}`, tone: "" };
}

function applyMatrixBarClientFlowClasses(bar, activity, repartoName = "") {
  if (!bar) return;
  bar.classList.remove("is-client-sent", "is-client-final");
  if (!activity) return;
  const commessaId = activity.commessa_id ?? activity.commessaId ?? "";
  const titolo = activity.titolo || "";
  const statoRaw = activity.stato || (bar.classList.contains("is-done") ? "completata" : "");
  if (!commessaId || normalizePhaseKey(statoRaw) !== "completata") return;
  const reparto =
    repartoName ||
    activity.reparto ||
    (activity.risorsa_id != null ? getRisorsaDeptName(activity.risorsa_id) : "") ||
    "";
  if (!isTodoClientFlowTracked(titolo, reparto)) return;
  const flow = getTodoClientFlow(commessaId, titolo, reparto);
  if (!flow?.sentAt) return;
  const resolved = getTodoClientFlowResolvedStatus(flow, startOfDay(new Date()));
  if (resolved === "confermato" || resolved === "silenzio_assenso") {
    bar.classList.add("is-client-final");
    return;
  }
  bar.classList.add("is-client-sent");
}

function refreshMatrixClientFlowIndicatorsForTarget(target) {
  if (!matrixGrid || !target?.commessaId || !target?.titolo) return;
  const commessaKey = String(target.commessaId);
  const titleKey = normalizePhaseKey(target.titolo || "");
  const deptKey = normalizeDeptKey(target.reparto || "");
  const byId = new Map((matrixState.attivita || []).map((entry) => [String(entry.id), entry]));
  matrixGrid.querySelectorAll(".matrix-activity-bar[data-activity-id]").forEach((bar) => {
    if (String(bar.dataset.commessaId || "") !== commessaKey) return;
    if (normalizePhaseKey(bar.dataset.attivita || "") !== titleKey) return;
    if (deptKey && normalizeDeptKey(bar.dataset.reparto || "") !== deptKey) return;
    const entry = byId.get(String(bar.dataset.activityId || ""));
    const activity = entry || {
      commessa_id: commessaKey,
      titolo: target.titolo,
      stato: bar.classList.contains("is-done") ? "completata" : "pianificata",
      reparto: target.reparto || "",
    };
    applyMatrixBarClientFlowClasses(bar, activity, target.reparto || "");
  });
}

function getTodoClientFlowEditBlockReason(titolo, repartoName = "") {
  if (!isTodoClientFlowTracked(titolo, repartoName)) return "Tracking cliente non disponibile per questa attivita.";
  if (!state.canWrite) return "Permessi insufficienti.";
  const role = getTodoRole();
  if (role === "admin" || role === "planner") return "";
  if (role !== "responsabile") return "Solo admin, responsabile e planner possono modificare il tracking cliente.";
  const userDept = getProfileDeptKey();
  const targetDept = normalizeDeptKey(repartoName || "");
  if (!userDept || !targetDept || userDept !== targetDept) {
    return "Responsabile: puoi modificare solo il tracking del tuo reparto.";
  }
  return "";
}

function getDefaultClientFlowDueDate(sentAt) {
  if (!sentAt) return null;
  return addBusinessDays(sentAt, TODO_CLIENT_FLOW_DEFAULT_DEADLINE_BUSINESS_DAYS);
}

function getTodoRole() {
  return String(state.profile?.ruolo || "").trim().toLowerCase();
}

function getProfileDeptKey() {
  const repartoId = getOwnRepartoId();
  if (!repartoId) return "";
  const reparto = (state.reparti || []).find((r) => String(r.id) === String(repartoId));
  return normalizeDeptKey(reparto?.nome || "");
}

function getTodoEntriesForCell(commessaId, title, repartoName = "") {
  const entries = state.reportActivitiesMap.get(String(commessaId)) || [];
  const titleKey = normalizePhaseKey(title);
  const deptKey = normalizeDeptKey(repartoName || "");
  return entries.filter((entry) => {
    if (normalizePhaseKey(entry.titolo) !== titleKey) return false;
    if (!deptKey) return true;
    return normalizeDeptKey(entry.dept || "") === deptKey;
  });
}

function mergeReportActivitiesFromMatrix(commessaId) {
  if (!commessaId) return;
  const key = String(commessaId);
  const existing = state.reportActivitiesMap.get(key) || [];
  const byId = new Map();
  existing.forEach((entry) => {
    const idKey = entry?.id != null ? String(entry.id) : `${entry?.titolo || ""}|${entry?.risorsaId || ""}|${entry?.start || ""}`;
    byId.set(idKey, entry);
  });
  const risorseById = new Map((state.risorse || []).map((r) => [String(r.id), r.nome]));
  (matrixState.attivita || [])
    .filter((row) => String(row.commessa_id || "") === key)
    .forEach((row) => {
      const idKey = row?.id != null ? String(row.id) : `${row?.titolo || ""}|${row?.risorsa_id || ""}|${row?.data_inizio || ""}`;
      const dept =
        row.risorsa_id != null
          ? getDeptBucketByRisorsa(row.risorsa_id)
          : row.reparto_id != null
          ? getDeptBucketByRepartoId(row.reparto_id)
          : "ALTRO";
      byId.set(idKey, {
        id: row.id,
        risorsaId: row.risorsa_id,
        assegnatoA: row.assegnato_a || null,
        titolo: row.titolo || "Attivita",
        risorsa:
          row.risorsa_id != null
            ? risorseById.get(String(row.risorsa_id)) || `Risorsa ${row.risorsa_id}`
            : "Risorsa n/d",
        dept,
        start: row.data_inizio ? new Date(row.data_inizio) : null,
        end: row.data_fine ? new Date(row.data_fine) : null,
        stato: row.stato || "pianificata",
      });
    });
  const mergedList = Array.from(byId.values()).sort((a, b) => {
    const aTime = a.start ? a.start.getTime() : 0;
    const bTime = b.start ? b.start.getTime() : 0;
    if (aTime !== bTime) return aTime - bTime;
    return String(a.titolo || "").localeCompare(String(b.titolo || ""));
  });
  const mergedMap = new Map(state.reportActivitiesMap);
  mergedMap.set(key, mergedList);
  state.reportActivitiesMap = mergedMap;
}

function getTodoEditableEntries(entries, repartoName = "") {
  const role = getTodoRole();
  if (!Array.isArray(entries) || !entries.length) return [];
  if (role === "admin") return entries.slice();
  if (role === "responsabile") {
    const userDept = getProfileDeptKey();
    const targetDept = normalizeDeptKey(repartoName || "");
    if (!userDept || !targetDept || userDept !== targetDept) return [];
    return entries.slice();
  }
  if (role === "operatore") {
    const ownRisorsaId = getOwnRisorsaId();
    const profileId = String(state.profile?.id || "");
    return entries.filter((entry) => {
      const byRisorsa = ownRisorsaId && String(entry.risorsaId || "") === String(ownRisorsaId);
      const byAssignee = entry.assegnatoA && String(entry.assegnatoA) === profileId;
      return Boolean(byRisorsa || byAssignee);
    });
  }
  return [];
}

function getTodoStatusMenuBlockReason(commessaId, title, repartoName = "") {
  if (!state.canWrite) return "Permessi insufficienti su questa attivita.";
  const role = getTodoRole();
  if (!role || role === "viewer") return "Permessi insufficienti su questa attivita.";
  if (role === "planner") return "Planner: permessi in sola lettura sui To Dos.";
  if (role === "admin") return "";
  const entries = getTodoEntriesForCell(commessaId, title, repartoName);
  if (role === "responsabile") {
    const userDept = getProfileDeptKey();
    const targetDept = normalizeDeptKey(repartoName || "");
    if (!userDept || !targetDept || userDept !== targetDept) {
      return "Responsabile: puoi aggiornare solo attivita del tuo reparto.";
    }
    return "";
  }
  if (role === "operatore") {
    return getTodoEditableEntries(entries, repartoName).length > 0
      ? ""
      : "Operatore: puoi aggiornare solo attivita assegnate a te.";
  }
  return "Permessi insufficienti su questa attivita.";
}

function canOpenTodoStatusMenuForCell(commessaId, title, repartoName = "") {
  return !getTodoStatusMenuBlockReason(commessaId, title, repartoName);
}

function getTodoStatusFor(commessaId, title, repartoName = "") {
  const override = getTodoOverrideStatus(commessaId, title);
  if (override === "non_necessaria") return "non_necessaria";
  const matching = getTodoEntriesForCell(commessaId, title, repartoName).filter(
    (entry) => normalizePhaseKey(entry.stato) !== "annullata"
  );
  if (!matching.length) return "da_schedulare";
  if (matching.some((e) => normalizePhaseKey(e.stato) === "completata")) return "fatta";
  return "schedulata";
}

function getTodoStatusVisual(status) {
  if (status === "fatta") return { label: "DONE", cls: "todo-fatta" };
  if (status === "schedulata") return { label: "PLANNED", cls: "todo-schedulata" };
  if (status === "non_necessaria") return { label: "NOT NEED", cls: "todo-non-necessaria" };
  return { label: "TO PLAN", cls: "todo-da-schedulare" };
}

function findTodoCellByTarget(commessaId, titolo, repartoName = "") {
  if (!todoGrid || !commessaId || !titolo) return null;
  const titleKey = normalizePhaseKey(titolo);
  const deptKey = normalizeDeptKey(repartoName || "");
  const candidates = Array.from(todoGrid.querySelectorAll(`.todo-cell[data-commessa-id="${String(commessaId)}"]`));
  return (
    candidates.find((cell) => {
      const cellTitleKey = normalizePhaseKey(cell.dataset.attivita || "");
      const cellDeptKey = normalizeDeptKey(cell.dataset.reparto || "");
      return cellTitleKey === titleKey && cellDeptKey === deptKey;
    }) || null
  );
}

function refreshTodoCellVisualFromState(target) {
  if (!target) return;
  const { commessaId, titolo, reparto } = target;
  const cell = findTodoCellByTarget(commessaId, titolo, reparto || "");
  if (!cell) return;

  const status = getTodoStatusFor(commessaId, titolo, reparto || "");
  const statusVisual = getTodoStatusVisual(status);
  const statusPill = cell.querySelector(".todo-status");
  if (statusPill) {
    statusPill.textContent = statusVisual.label;
    statusPill.classList.remove("todo-fatta", "todo-schedulata", "todo-non-necessaria", "todo-da-schedulare");
    statusPill.classList.add(statusVisual.cls);
  }

  if (isTodoClientFlowTracked(titolo, reparto || "")) {
    const clientFlow = getTodoClientFlow(commessaId, titolo, reparto || "");
    const clientMeta = getTodoClientFlowMeta(clientFlow, startOfDay(new Date()));
    let metaEl = cell.querySelector(".todo-client-meta");
    if (clientMeta) {
      if (!metaEl) {
        const wrap = cell.querySelector(".todo-status-wrap");
        if (wrap) {
          metaEl = document.createElement("span");
          metaEl.className = "todo-client-meta";
          wrap.appendChild(metaEl);
        }
      }
      if (metaEl) {
        metaEl.textContent = clientMeta.text;
        metaEl.title = clientMeta.text;
        metaEl.classList.remove("todo-client-meta-ok", "todo-client-meta-warn", "todo-client-meta-danger");
        if (clientMeta.tone === "ok") metaEl.classList.add("todo-client-meta-ok");
        if (clientMeta.tone === "warn") metaEl.classList.add("todo-client-meta-warn");
        if (clientMeta.tone === "danger") metaEl.classList.add("todo-client-meta-danger");
      }
    } else if (metaEl) {
      metaEl.remove();
    }
  }
}

const TODO_GLOBAL_STATUS_ORDER = ["In corso", "Rilasciata a prod.", "Evasa"];
const TODO_GLOBAL_STATUS_TO_DB = {
  "In corso": "in_corso",
  "Rilasciata a prod.": "sospesa",
  Evasa: "chiusa",
};

function canEditTodoGlobalStatus() {
  const role = getTodoRole();
  return role === "admin" || role === "responsabile" || role === "planner";
}

function getTodoGlobalCommessaStatus(commessa) {
  const raw = normalizePhaseKey(commessa?.stato || "");
  if (raw === "in_corso") return { label: "In corso", cls: "todo-global-in-corso" };
  if (raw === "sospesa") return { label: "Rilasciata a prod.", cls: "todo-global-rilasciata" };
  if (raw === "chiusa") return { label: "Evasa", cls: "todo-global-evasa" };
  if (raw.includes("evas")) return { label: "Evasa", cls: "todo-global-evasa" };
  if (raw.includes("rilasci")) return { label: "Rilasciata a prod.", cls: "todo-global-rilasciata" };
  if (!raw || raw.includes("corso")) return { label: "In corso", cls: "todo-global-in-corso" };
  return { label: commessa?.stato || "In corso", cls: "todo-global-in-corso" };
}

function updateTodoPanelHeaderOffset() {
  if (!todoSection) return;
  const todoPanelHeader = todoSection.querySelector(".panel-header");
  const todoToolbar = todoSection.querySelector(".todo-header");
  if (todoPanelHeader) {
    todoSection.style.setProperty("--todo-panel-header-height", `${todoPanelHeader.offsetHeight}px`);
  }
  if (todoToolbar) {
    todoSection.style.setProperty("--todo-toolbar-height", `${todoToolbar.offsetHeight}px`);
  }
  if (todoToolbar) {
    const toolbarRect = todoToolbar.getBoundingClientRect();
    const row1Raw = getComputedStyle(todoSection).getPropertyValue("--todo-head-row1-height").trim();
    const row1 = Number.parseFloat(row1Raw);
    const row1Height = Number.isFinite(row1) ? row1 : 36;
    const row1Top = Math.max(0, toolbarRect.bottom);
    const row2Top = Math.max(0, row1Top + row1Height);
    todoSection.style.setProperty("--todo-head-row1-top", `${row1Top}px`);
    todoSection.style.setProperty("--todo-head-row2-top", `${row2Top}px`);
  }
}

function updateTodoInfoBar(commesse, columns) {
  if (!todoInfoText) return;
  const base = "Click o long press su una pill per cambiare stato.";
  if (!Array.isArray(commesse) || !commesse.length || !Array.isArray(columns) || !columns.length) {
    todoInfoText.textContent = base;
    return;
  }
  const trackedColumns = columns.filter((col) => isTodoClientFlowTracked(col.titolo, col.reparto));
  if (!trackedColumns.length) {
    todoInfoText.textContent = base;
    return;
  }
  const seen = new Set();
  let waitingCount = 0;
  let expiredCount = 0;
  let confirmedCount = 0;
  const today = startOfDay(new Date());
  commesse.forEach((commessa) => {
    trackedColumns.forEach((col) => {
      const key = `${commessa.id}|${normalizePhaseKey(col.titolo)}|${normalizeDeptKey(col.reparto)}`;
      if (seen.has(key)) return;
      seen.add(key);
      const activityStatus = getTodoStatusFor(commessa.id, col.titolo, col.reparto);
      if (activityStatus !== "fatta") return;
      const flow = getTodoClientFlow(commessa.id, col.titolo, col.reparto);
      const status = getTodoClientFlowResolvedStatus(flow, today);
      if (status === "in_attesa") waitingCount += 1;
      else if (status === "scaduto") expiredCount += 1;
      else if (status === "confermato") confirmedCount += 1;
    });
  });
  todoInfoText.textContent =
    `${base} · Preliminare/3D CAD in attesa: ${waitingCount}` +
    ` · Scaduti: ${expiredCount}` +
    ` · Confermati: ${confirmedCount}`;
}

async function updateTodoGlobalStatus(commessa, statusLabel) {
  if (!commessa?.id) return false;
  if (!canEditTodoGlobalStatus()) {
    setStatus("Permessi insufficienti per aggiornare lo stato commessa.", "error");
    return false;
  }
  const dbValue = TODO_GLOBAL_STATUS_TO_DB[statusLabel] || statusLabel;
  const { error } = await supabase.from("commesse").update({ stato: dbValue }).eq("id", commessa.id);
  if (error) {
    setStatus(`Update error: ${error.message}`, "error");
    return false;
  }
  updateCommessaInState(commessa.id, { stato: dbValue });
  setStatus("Stato commessa aggiornato.", "ok");
  renderMatrixReport();
  return true;
}

function isTodoMenuOpen(menuEl) {
  return Boolean(menuEl) && !menuEl.classList.contains("hidden");
}

function renderTodoSection(commesse) {
  if (!todoSection || !todoGrid) return;
  todoLastRendered = commesse || [];
  updateTodoPanelHeaderOffset();
  if (!state.session) {
    updateTodoInfoBar([], []);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Accedi per vedere i To Dos.</div>`;
    return;
  }
  if (!commesse || !commesse.length) {
    updateTodoInfoBar([], []);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Nessuna commessa.</div>`;
    return;
  }

  const groups = getTodoActivityGroups();
  if (!groups.length) {
    updateTodoInfoBar([], []);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Nessuna attivita disponibile.</div>`;
    return;
  }
  const targetPhaseKey = "progettazione termodinamica";
  const cadOrderKeys = [
    "preliminare",
    "3d",
    "carenatura",
    "telaio",
    "costruttivi 1-5",
    "consegna totale",
  ];
  const orderActivitiesForGroup = (group) => {
    const deptKey = normalizeDeptKey(group?.reparto || "");
    const list = (group?.attivita || []).slice();
    if (deptKey.includes("CAD")) {
      return list.sort((a, b) => {
        const aKey = normalizePhaseKey(a);
        const bKey = normalizePhaseKey(b);
        const aIdx = cadOrderKeys.indexOf(aKey);
        const bIdx = cadOrderKeys.indexOf(bKey);
        const aRank = aIdx >= 0 ? aIdx : Number.MAX_SAFE_INTEGER;
        const bRank = bIdx >= 0 ? bIdx : Number.MAX_SAFE_INTEGER;
        if (aRank !== bRank) return aRank - bRank;
        return String(a || "").localeCompare(String(b || ""));
      });
    }
    return list.sort((a, b) => {
      const aTarget = normalizePhaseKey(a) === targetPhaseKey ? 0 : 1;
      const bTarget = normalizePhaseKey(b) === targetPhaseKey ? 0 : 1;
      if (aTarget !== bTarget) return aTarget - bTarget;
      return String(a || "").localeCompare(String(b || ""));
    });
  };
  const orderedGroups = groups
    .map((group) => ({
      ...group,
      attivita: orderActivitiesForGroup(group),
    }))
    .sort((a, b) => {
      const aHas = (a.attivita || []).some((name) => normalizePhaseKey(name) === targetPhaseKey);
      const bHas = (b.attivita || []).some((name) => normalizePhaseKey(name) === targetPhaseKey);
      if (aHas !== bHas) return aHas ? -1 : 1;
      return String(a.reparto || "").localeCompare(String(b.reparto || ""));
    });

  const columns = [];
  orderedGroups.forEach((g) => {
    g.attivita.forEach((name) => {
      columns.push({ reparto: g.reparto, titolo: name });
    });
  });
  const filteredCommesse = applyTodoFilters(commesse, columns);
  if (!filteredCommesse.length) {
    updateTodoInfoBar([], columns);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Nessuna commessa trovata con i filtri To Dos.</div>`;
    return;
  }
  if (filteredCommesse.length > REPORT_MAX_ITEMS) {
    updateTodoInfoBar([], columns);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Troppe commesse (${filteredCommesse.length}). Affina i filtri To Dos.</div>`;
    return;
  }
  const missingData = filteredCommesse.some(
    (c) => !state.reportActivitiesMap.has(String(c.id)) || !state.todoOverridesMap.has(String(c.id))
  );
  if (missingData) {
    updateTodoInfoBar([], columns);
    if (!state.reportActivitiesLoading) loadReportActivitiesFor(filteredCommesse);
    if (!state.todoOverridesLoading) loadTodoOverridesFor(filteredCommesse);
    if (todoHeaderTableHost) todoHeaderTableHost.innerHTML = "";
    todoGrid.innerHTML = `<div class="report-empty">Caricamento To Dos...</div>`;
    return;
  }
  const missingClientFlows = filteredCommesse.some((c) => !state.todoClientFlowLoadedCommesse.has(String(c.id)));
  if (missingClientFlows && !state.todoClientFlowLoading) {
    loadTodoClientFlowsFor(filteredCommesse);
  }

  const gridCols = `240px 160px repeat(${columns.length}, minmax(120px, 1fr))`;
  const headerTable = document.createElement("div");
  headerTable.className = "todo-table todo-table-head";
  headerTable.style.gridTemplateColumns = gridCols;
  const bodyTable = document.createElement("div");
  bodyTable.className = "todo-table todo-table-body";
  bodyTable.style.gridTemplateColumns = gridCols;
  const today = startOfDay(new Date());

  const headerBlank = document.createElement("div");
  headerBlank.className = "todo-cell todo-head";
  headerBlank.style.gridRow = "1";
  headerBlank.style.gridColumn = "1";
  headerBlank.textContent = "";
  headerTable.appendChild(headerBlank);

  const statusBlank = document.createElement("div");
  statusBlank.className = "todo-cell todo-head";
  statusBlank.style.gridRow = "1";
  statusBlank.style.gridColumn = "2";
  statusBlank.textContent = "";
  headerTable.appendChild(statusBlank);

  let colStart = 3;
  orderedGroups.forEach((group) => {
    const cell = document.createElement("div");
    cell.className = "todo-cell todo-head todo-group";
    const groupKey = normalizeDeptKey(group.reparto || "");
    if (groupKey.includes("CAD")) cell.classList.add("todo-group-cad");
    if (groupKey.includes("ELETTR")) cell.classList.add("todo-group-elettrico");
    if (groupKey.includes("TERMO")) cell.classList.add("todo-group-termodinamico");
    cell.style.gridRow = "1";
    cell.style.gridColumn = `${colStart} / span ${group.attivita.length}`;
    cell.textContent = group.reparto;
    headerTable.appendChild(cell);
    colStart += group.attivita.length;
  });

  const commessaHeader = document.createElement("div");
  commessaHeader.className = "todo-cell todo-head";
  commessaHeader.style.gridRow = "2";
  commessaHeader.style.gridColumn = "1";
  commessaHeader.textContent = "Commessa";
  headerTable.appendChild(commessaHeader);

  const statusHeader = document.createElement("div");
  statusHeader.className = "todo-cell todo-head";
  statusHeader.style.gridRow = "2";
  statusHeader.style.gridColumn = "2";
  statusHeader.textContent = "Stato";
  headerTable.appendChild(statusHeader);

  colStart = 3;
  columns.forEach((col) => {
    const cell = document.createElement("div");
    cell.className = "todo-cell todo-head todo-activity-head";
    cell.style.gridRow = "2";
    cell.style.gridColumn = String(colStart);
    cell.textContent = col.titolo;
    headerTable.appendChild(cell);
    colStart += 1;
  });

  filteredCommesse.forEach((commessa, idx) => {
    const rowIndex = idx + 1;
    const rowKey = String(commessa.id || "");
    const info = document.createElement("div");
    info.className = "todo-cell todo-commessa todo-commessa-clickable";
    info.style.gridRow = String(rowIndex);
    info.style.gridColumn = "1";
    info.dataset.todoRow = rowKey;
    const numero = commessa.numero != null ? String(commessa.numero) : "-";
    const anno = commessa.anno != null ? String(commessa.anno) : "-";
    const titoloFull = String(commessa.titolo || "");
    const titoloShort = titoloFull.length > 25 ? `${titoloFull.slice(0, 25)}...` : titoloFull;
    info.innerHTML = `<div>${numero} · ${anno}</div><div class="todo-meta" title="${escapeHtml(titoloFull)}">${escapeHtml(titoloShort)}</div>`;
    info.addEventListener("click", () => {
      selectCommessa(commessa.id, { openModal: true });
    });
    bodyTable.appendChild(info);

    const globalStatus = getTodoGlobalCommessaStatus(commessa);
    const statusCell = document.createElement("div");
    statusCell.className = "todo-cell todo-commessa-status";
    statusCell.style.gridRow = String(rowIndex);
    statusCell.style.gridColumn = "2";
    statusCell.dataset.todoRow = rowKey;
    if (canEditTodoGlobalStatus()) statusCell.classList.add("todo-commessa-status-clickable");
    statusCell.innerHTML = `<span class="todo-global-pill ${globalStatus.cls}">${globalStatus.label}</span>`;
    statusCell.addEventListener("click", async () => {
      const pill = statusCell.querySelector(".todo-global-pill");
      const rect = (pill || statusCell).getBoundingClientRect();
      openTodoGlobalStatusMenu(statusCell, commessa, rect.left + rect.width / 2, rect.bottom);
    });
    bodyTable.appendChild(statusCell);

    colStart = 3;
    columns.forEach((col) => {
      const status = getTodoStatusFor(commessa.id, col.titolo, col.reparto);
      const isTrackedFlow = isTodoClientFlowTracked(col.titolo, col.reparto);
      const clientFlow = isTrackedFlow ? getTodoClientFlow(commessa.id, col.titolo, col.reparto) : null;
      const clientMeta = isTrackedFlow && status === "fatta" ? getTodoClientFlowMeta(clientFlow, today) : null;
      const cell = document.createElement("div");
      cell.className = "todo-cell";
      cell.style.gridRow = String(rowIndex);
      cell.style.gridColumn = String(colStart);
      const statusLabel =
        status === "fatta"
          ? "DONE"
          : status === "schedulata"
          ? "PLANNED"
          : status === "non_necessaria"
          ? "NOT NEED"
          : "TO PLAN";
      const statusClass =
        status === "fatta"
          ? "todo-fatta"
          : status === "schedulata"
          ? "todo-schedulata"
          : status === "non_necessaria"
          ? "todo-non-necessaria"
          : "todo-da-schedulare";
      const metaToneClass =
        clientMeta?.tone === "ok"
          ? "todo-client-meta-ok"
          : clientMeta?.tone === "danger"
          ? "todo-client-meta-danger"
          : clientMeta?.tone === "warn"
          ? "todo-client-meta-warn"
          : "";
      cell.innerHTML = `
        <div class="todo-status-wrap">
          <span class="todo-status ${statusClass}">${statusLabel}</span>
          ${clientMeta ? `<span class="todo-client-meta ${metaToneClass}" title="${escapeHtml(clientMeta.text)}">${escapeHtml(clientMeta.text)}</span>` : ""}
        </div>
      `;
      cell.dataset.commessaId = commessa.id;
      cell.dataset.attivita = col.titolo;
      cell.dataset.reparto = col.reparto || "";
      cell.dataset.todoRow = rowKey;
      const statusPill = cell.querySelector(".todo-status");
      const openMenuOnLongPress = (event) => {
        todoLongPressStart = { x: event.clientX, y: event.clientY };
        clearTimeout(todoLongPressTimer);
        todoLongPressTimer = window.setTimeout(() => {
          todoLongPressSuppress = true;
          openTodoStatusMenu(cell, event.clientX, event.clientY);
        }, 550);
      };
      const clearLongPress = () => {
        clearTimeout(todoLongPressTimer);
        todoLongPressStart = null;
      };
      if (statusPill) {
        statusPill.addEventListener("click", (event) => {
          if (todoLongPressSuppress) {
            todoLongPressSuppress = false;
            event.preventDefault();
            event.stopPropagation();
            return;
          }
          event.preventDefault();
          event.stopPropagation();
          const rect = statusPill.getBoundingClientRect();
          const x = Number.isFinite(event.clientX) ? event.clientX : rect.left + rect.width / 2;
          const y = Number.isFinite(event.clientY) ? event.clientY : rect.top + rect.height / 2;
          openTodoStatusMenu(cell, x, y);
        });
        statusPill.addEventListener("pointerdown", openMenuOnLongPress);
        statusPill.addEventListener("pointerup", clearLongPress);
        statusPill.addEventListener("pointerleave", clearLongPress);
        statusPill.addEventListener("pointercancel", clearLongPress);
      }
      cell.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        openTodoStatusMenu(cell, event.clientX, event.clientY);
      });
      bodyTable.appendChild(cell);
      colStart += 1;
    });
  });

  todoGrid.innerHTML = "";
  if (todoHeaderTableHost) {
    todoHeaderTableHost.innerHTML = "";
    todoHeaderTableHost.appendChild(headerTable);
    todoHeaderTableHost.scrollLeft = todoGrid.scrollLeft;
  } else {
    todoGrid.appendChild(headerTable);
  }
  todoGrid.appendChild(bodyTable);
  updateTodoInfoBar(filteredCommesse, columns);
  updateTodoPanelHeaderOffset();
}

function refreshTodoVisualizationNow() {
  if (!state.session) return;
  if (!todoSection || todoSection.classList.contains("collapsed")) return;
  const base =
    todoLastRendered && todoLastRendered.length ? todoLastRendered : (state.commesse || []).slice().sort(compareCommesseDesc);
  renderTodoSection(base);
  if (todoStatusDatePickerState.action === "client_sent" && isTodoMenuOpen(todoStatusMenu)) {
    renderTodoStatusDatePicker();
    repositionTodoStatusMenuToAnchor();
  }
}

function stopTodoVisualRefreshLoop() {
  if (!todoVisualRefreshTimer) return;
  clearInterval(todoVisualRefreshTimer);
  todoVisualRefreshTimer = null;
}

function startTodoVisualRefreshLoop() {
  stopTodoVisualRefreshLoop();
  todoVisualRefreshTimer = setInterval(() => {
    if (document.hidden) return;
    refreshTodoVisualizationNow();
  }, TODO_VISUAL_REFRESH_MS);
}

function resetTodoStatusDatePickerState() {
  todoStatusDatePickerState.action = null;
  todoStatusDatePickerState.sentSelected = null;
  todoStatusDatePickerState.dueSelected = null;
  todoStatusDatePickerState.confirmSelected = null;
  const now = startOfDay(new Date());
  todoStatusDatePickerState.sentMonth = now;
  todoStatusDatePickerState.dueMonth = now;
  todoStatusDatePickerState.confirmMonth = now;
}

function closeTodoStatusDatePicker() {
  if (!todoStatusMenu) return;
  const panel = todoStatusMenu.querySelector(".todo-status-date-panel");
  if (panel) panel.classList.add("hidden");
  resetTodoStatusDatePickerState();
}

function renderTodoStatusDatePicker() {
  if (!todoStatusMenu) return;
  const panel = todoStatusMenu.querySelector(".todo-status-date-panel");
  if (!panel) return;
  const title = panel.querySelector(".todo-status-date-title");
  const applyBtn = panel.querySelector('button[data-action="todo_date_apply"]');
  const dualWrap = panel.querySelector(".todo-status-date-dual");
  const confirmWrap = panel.querySelector(".todo-status-date-confirm");
  if (!title || !dualWrap || !confirmWrap) return;

  const renderCalendar = (field, monthInput, selectedInput) => {
    const root = panel.querySelector(`.todo-status-date-calendar[data-field="${field}"]`);
    if (!root) return;
    const label = root.querySelector(".todo-status-date-label");
    const daysHost = root.querySelector(".todo-status-date-days");
    if (!label || !daysHost) return;
    const selected = selectedInput ? startOfDay(selectedInput) : startOfDay(new Date());
    const month = monthInput ? new Date(monthInput.getFullYear(), monthInput.getMonth(), 1) : new Date(selected.getFullYear(), selected.getMonth(), 1);
    label.textContent = month.toLocaleDateString("it-IT", { month: "long", year: "numeric" });
    daysHost.innerHTML = "";
    const firstWeekday = (month.getDay() + 6) % 7;
    for (let i = 0; i < firstWeekday; i += 1) {
      const empty = document.createElement("button");
      empty.type = "button";
      empty.className = "milestone-day empty";
      empty.disabled = true;
      empty.setAttribute("aria-hidden", "true");
      daysHost.appendChild(empty);
    }
    const today = startOfDay(new Date());
    const cursor = new Date(month);
    while (cursor.getMonth() === month.getMonth()) {
      const day = document.createElement("button");
      day.type = "button";
      day.className = "milestone-day";
      if (cursor.getDay() === 0 || cursor.getDay() === 6) day.classList.add("weekend");
      if (cursor.getTime() === today.getTime()) day.classList.add("today");
      if (cursor.getTime() === selected.getTime()) day.classList.add("selected");
      day.dataset.action = "todo_date_pick";
      day.dataset.field = field;
      day.dataset.date = formatIsoDateOnly(cursor);
      day.textContent = String(cursor.getDate());
      daysHost.appendChild(day);
      cursor.setDate(cursor.getDate() + 1);
    }
  };

  if (todoStatusDatePickerState.action === "client_sent") {
    if (applyBtn) applyBtn.textContent = "Salva invio e scadenza";
    title.textContent = "Invio cliente e scadenza";
    dualWrap.classList.remove("hidden");
    confirmWrap.classList.add("hidden");
    renderCalendar("sent", todoStatusDatePickerState.sentMonth, todoStatusDatePickerState.sentSelected);
    renderCalendar("due", todoStatusDatePickerState.dueMonth, todoStatusDatePickerState.dueSelected);
    const counterValue = panel.querySelector(".todo-status-counter-value");
    if (counterValue) {
      counterValue.classList.remove("is-warn", "is-danger", "is-soon", "is-future");
      const dueDate = parseIsoDateOnly(todoStatusDatePickerState.dueSelected);
      const remaining = dueDate ? businessDayDiff(startOfDay(new Date()), startOfDay(dueDate)) : null;
      if (remaining == null) {
        counterValue.textContent = "n/d";
      } else if (remaining < 0) {
        counterValue.textContent = `Scaduta da ${Math.abs(remaining)} giorni lavorativi`;
        counterValue.classList.add("is-danger");
      } else if (remaining <= 3) {
        counterValue.textContent = "Scade oggi (0 giorni lavorativi)";
        if (remaining > 0) {
          counterValue.textContent = `In scadenza: ${remaining} giorni lavorativi`;
        }
        counterValue.classList.add("is-soon");
      } else {
        counterValue.textContent = `${remaining} giorni lavorativi rimanenti`;
        counterValue.classList.add("is-future");
      }
    }
    return;
  }

  const isSilenceDecision = todoStatusDatePickerState.action === "client_silence";
  title.textContent = isSilenceDecision ? "Data decisione silenzio assenso" : "Data conferma cliente";
  if (applyBtn) applyBtn.textContent = isSilenceDecision ? "Conferma silenzio assenso" : "Conferma data";
  dualWrap.classList.add("hidden");
  confirmWrap.classList.remove("hidden");
  renderCalendar("confirm", todoStatusDatePickerState.confirmMonth, todoStatusDatePickerState.confirmSelected);
}

function openTodoStatusDatePicker(action) {
  if (!todoStatusMenu || !todoMenuTarget) return;
  const panel = todoStatusMenu.querySelector(".todo-status-date-panel");
  if (!panel) return;
  const existing = getTodoClientFlow(todoMenuTarget.commessaId, todoMenuTarget.titolo, todoMenuTarget.reparto);
  const now = startOfDay(new Date());
  todoStatusDatePickerState.action = action;
  if (action === "client_sent") {
    const sent = parseIsoDateOnly(existing?.sentAt) || now;
    const due = parseIsoDateOnly(existing?.dueAt) || getDefaultClientFlowDueDate(sent) || now;
    todoStatusDatePickerState.sentSelected = sent;
    todoStatusDatePickerState.dueSelected = due;
    todoStatusDatePickerState.sentMonth = new Date(sent.getFullYear(), sent.getMonth(), 1);
    todoStatusDatePickerState.dueMonth = new Date(due.getFullYear(), due.getMonth(), 1);
  } else {
    const confirmed = parseIsoDateOnly(existing?.confirmedAt) || now;
    todoStatusDatePickerState.confirmSelected = confirmed;
    todoStatusDatePickerState.confirmMonth = new Date(confirmed.getFullYear(), confirmed.getMonth(), 1);
  }
  panel.classList.remove("hidden");
  renderTodoStatusDatePicker();
  requestAnimationFrame(() => repositionTodoStatusMenuToAnchor());
}

async function handleTodoStatusDatePickerAction(action, button) {
  if (!todoStatusMenu) return false;
  if (action === "todo_date_prev") {
    const field = button?.dataset.field || "";
    if (field === "sent") {
      const month = todoStatusDatePickerState.sentMonth || startOfDay(new Date());
      todoStatusDatePickerState.sentMonth = new Date(month.getFullYear(), month.getMonth() - 1, 1);
    } else if (field === "due") {
      const month = todoStatusDatePickerState.dueMonth || startOfDay(new Date());
      todoStatusDatePickerState.dueMonth = new Date(month.getFullYear(), month.getMonth() - 1, 1);
    } else {
      const month = todoStatusDatePickerState.confirmMonth || startOfDay(new Date());
      todoStatusDatePickerState.confirmMonth = new Date(month.getFullYear(), month.getMonth() - 1, 1);
    }
    renderTodoStatusDatePicker();
    return false;
  }
  if (action === "todo_date_next") {
    const field = button?.dataset.field || "";
    if (field === "sent") {
      const month = todoStatusDatePickerState.sentMonth || startOfDay(new Date());
      todoStatusDatePickerState.sentMonth = new Date(month.getFullYear(), month.getMonth() + 1, 1);
    } else if (field === "due") {
      const month = todoStatusDatePickerState.dueMonth || startOfDay(new Date());
      todoStatusDatePickerState.dueMonth = new Date(month.getFullYear(), month.getMonth() + 1, 1);
    } else {
      const month = todoStatusDatePickerState.confirmMonth || startOfDay(new Date());
      todoStatusDatePickerState.confirmMonth = new Date(month.getFullYear(), month.getMonth() + 1, 1);
    }
    renderTodoStatusDatePicker();
    return false;
  }
  if (action === "todo_date_pick") {
    const iso = button?.dataset.date || "";
    const field = button?.dataset.field || "";
    const picked = parseIsoDateOnly(iso);
    if (!picked) return false;
    if (field === "sent") {
      todoStatusDatePickerState.sentSelected = picked;
      if (!todoStatusDatePickerState.dueSelected) {
        todoStatusDatePickerState.dueSelected = getDefaultClientFlowDueDate(picked) || picked;
      }
    } else if (field === "due") {
      todoStatusDatePickerState.dueSelected = picked;
    } else {
      todoStatusDatePickerState.confirmSelected = picked;
    }
    renderTodoStatusDatePicker();
    return false;
  }
  if (action === "todo_date_cancel") {
    closeTodoStatusDatePicker();
    requestAnimationFrame(() => repositionTodoStatusMenuToAnchor());
    return false;
  }
  if (action === "todo_date_apply") {
    if (!todoMenuTarget || !todoStatusDatePickerState.action) return false;
    if (todoStatusDatePickerState.action === "client_sent") {
      const sent = parseIsoDateOnly(todoStatusDatePickerState.sentSelected) || startOfDay(new Date());
      const due = parseIsoDateOnly(todoStatusDatePickerState.dueSelected) || getDefaultClientFlowDueDate(sent) || sent;
      if (dayNumberUTC(due) < dayNumberUTC(sent)) {
        setStatus("La scadenza non puo essere prima dell'invio.", "error");
        return false;
      }
      const updated = await applyTodoStatusAction(todoStatusDatePickerState.action, todoMenuTarget, {
        selectedDate: formatIsoDateOnly(sent),
        dueDate: formatIsoDateOnly(due),
      });
      return updated;
    }
    const selectedDate = formatIsoDateOnly(todoStatusDatePickerState.confirmSelected || startOfDay(new Date()));
    const updated = await applyTodoStatusAction(todoStatusDatePickerState.action, todoMenuTarget, { selectedDate });
    return updated;
  }
  return false;
}

function positionTodoStatusMenuCentered() {
  if (!todoStatusMenu || todoStatusMenu.classList.contains("hidden")) return;
  const padding = 10;
  const rect = todoStatusMenu.getBoundingClientRect();
  let left = Math.round((window.innerWidth - rect.width) / 2);
  let top = Math.round((window.innerHeight - rect.height) / 2);
  left = Math.max(padding, left);
  top = Math.max(padding, top);
  if (left + rect.width > window.innerWidth - padding) {
    left = Math.max(padding, window.innerWidth - rect.width - padding);
  }
  if (top + rect.height > window.innerHeight - padding) {
    top = Math.max(padding, window.innerHeight - rect.height - padding);
  }
  todoStatusMenu.style.left = `${left}px`;
  todoStatusMenu.style.top = `${top}px`;
}

function positionTodoStatusMenuWithinViewport() {
  if (!todoStatusMenu || todoStatusMenu.classList.contains("hidden")) return;
  const padding = 10;
  const rect = todoStatusMenu.getBoundingClientRect();
  const leftRaw = Number.parseFloat(todoStatusMenu.style.left);
  const topRaw = Number.parseFloat(todoStatusMenu.style.top);
  let left = Number.isFinite(leftRaw) ? leftRaw : rect.left;
  let top = Number.isFinite(topRaw) ? topRaw : rect.top;
  if (left + rect.width > window.innerWidth - padding) {
    left = Math.max(padding, window.innerWidth - rect.width - padding);
  }
  if (top + rect.height > window.innerHeight - padding) {
    top = Math.max(padding, window.innerHeight - rect.height - padding);
  }
  if (left < padding) left = padding;
  if (top < padding) top = padding;
  todoStatusMenu.style.left = `${left}px`;
  todoStatusMenu.style.top = `${top}px`;
}

function isTodoStatusMenuBorderHit(event) {
  if (!todoStatusMenu) return false;
  const rect = todoStatusMenu.getBoundingClientRect();
  const grip = 12;
  const x = event.clientX;
  const y = event.clientY;
  return x - rect.left <= grip || rect.right - x <= grip || y - rect.top <= grip || rect.bottom - y <= grip;
}

function openTodoStatusMenu(cell, x, y) {
  if (!todoStatusMenu || !cell) return;
  closeTodoGlobalStatusMenu();
  const commessaId = cell.dataset.commessaId;
  const titolo = cell.dataset.attivita;
  const reparto = cell.dataset.reparto || "";
  if (!commessaId || !titolo) return;
  const denyReason = getTodoStatusMenuBlockReason(commessaId, titolo, reparto);
  if (denyReason) {
    setStatus(denyReason, "error");
    return;
  }
  const rowKey = cell.dataset.todoRow || "";
  const commessa =
    (state.commesse || []).find((c) => String(c.id) === String(commessaId)) ||
    (todoLastRendered || []).find((c) => String(c.id) === String(commessaId)) ||
    null;
  const commessaNumero = commessa?.numero != null ? String(commessa.numero) : "-";
  const commessaAnno = commessa?.anno != null ? String(commessa.anno) : "-";
  const commessaTitolo = String(commessa?.titolo || "Senza descrizione");

  const role = getTodoRole();
  const entries = getTodoEntriesForCell(commessaId, titolo, reparto);
  const editableEntries = getTodoEditableEntries(entries, reparto);
  const currentStatus = getTodoStatusFor(commessaId, titolo, reparto);
  const isClientTracked = isTodoClientFlowTracked(titolo, reparto);
  const canAccessClientFlow = isClientTracked && currentStatus === "fatta";
  const clientFlow = isClientTracked ? getTodoClientFlow(commessaId, titolo, reparto) : null;
  const clientFlowResolvedStatus = isClientTracked ? getTodoClientFlowResolvedStatus(clientFlow, startOfDay(new Date())) : "";
  const clientFlowBlockReason = isClientTracked ? getTodoClientFlowEditBlockReason(titolo, reparto) : "";
  const statusActions =
    role === "operatore"
      ? [
          { action: "planned", label: "PLANNED" },
          { action: "done", label: "DONE" },
        ]
      : [
          { action: "to_plan", label: "TO PLAN" },
          { action: "planned", label: "PLANNED" },
          { action: "done", label: "DONE" },
          { action: "non_necessaria", label: "NOT NEED" },
        ];
  const clientActions = [];
  if (canAccessClientFlow) {
    clientActions.push(
      { action: "client_sent", label: "CLIENTE: INVIATO..." },
      { action: "client_confirmed", label: "CLIENTE: CONFERMATO" },
      { action: "client_silence", label: "CLIENTE: SILENZIO ASSENSO..." },
      { action: "client_reset", label: "CLIENTE: RESET" }
    );
  }
  const canShowGoMatrix = currentStatus === "schedulata";
  const getBlockReason = (def) => {
    let blockReason = "";
    const requiresEditableRows = def.action === "planned" || def.action === "done";
    if (requiresEditableRows && !editableEntries.length) {
      if (role === "operatore") {
        blockReason = "Operatore: puoi impostare stato solo su task assegnate a te.";
      } else if (role === "responsabile") {
        blockReason = "Responsabile: puoi aggiornare solo task del tuo reparto.";
      } else {
        blockReason = "Attivita non modificabile in questo contesto.";
      }
    }
    if (!blockReason && def.action === "go_matrix" && !getTodoMatrixJumpEntry({ commessaId, titolo, reparto })) {
      blockReason = "Nessuna attivita pianificata trovata in matrice.";
    }
    if (!blockReason && def.action === "to_plan" && role === "operatore") {
      blockReason = "Operatore: TO PLAN non disponibile da questo menu.";
    }
    if (!blockReason && def.action.startsWith("client_") && clientFlowBlockReason) {
      blockReason = clientFlowBlockReason;
    }
    if (!blockReason && def.action === "client_confirmed" && !clientFlow?.sentAt) {
      blockReason = "Prima registra la data di invio al cliente.";
    }
    if (!blockReason && def.action === "client_silence" && !clientFlow?.sentAt) {
      blockReason = "Prima registra la data di invio al cliente.";
    }
    if (!blockReason && def.action === "client_reset" && !clientFlow) {
      blockReason = "Nessun tracking cliente da resettare.";
    }
    return blockReason;
  };
  const goMatrixDef = { action: "go_matrix", label: "\u2197 Apri in matrice", title: "Apri in matrice" };
  const goMatrixBlockReason = canShowGoMatrix ? getBlockReason(goMatrixDef) : "";
  const appendSection = (title, defs, extraClass = "") => {
    if (!defs.length) return;
    const section = document.createElement("div");
    section.className = `todo-status-section ${extraClass}`.trim();
    if (title) {
      const titleEl = document.createElement("div");
      titleEl.className = "todo-status-section-title";
      titleEl.textContent = title;
      section.appendChild(titleEl);
    }
    const row = document.createElement("div");
    row.className = "todo-status-row";
    defs.forEach((def) => {
      const button = document.createElement("button");
      button.type = "button";
      button.dataset.action = def.action;
      button.textContent = def.label;
      if (def.action === "planned" && currentStatus === "schedulata") {
        button.classList.add("is-planned");
      }
      if (def.action === "done" && currentStatus === "fatta") {
        button.classList.add("is-success");
      }
      if (def.action === "non_necessaria" && currentStatus === "non_necessaria") {
        button.classList.add("is-not-needed");
      }
      if (
        def.action === "client_sent" &&
        clientFlow?.sentAt &&
        (clientFlowResolvedStatus === "in_attesa" || clientFlowResolvedStatus === "scaduto")
      ) {
        button.classList.add("is-client-sent");
      }
      if (def.action === "client_confirmed" && clientFlowResolvedStatus === "confermato") {
        button.classList.add("is-success");
      }
      if (def.action === "client_silence" && clientFlowResolvedStatus === "silenzio_assenso") {
        button.classList.add("is-success");
      }
      if (def.iconOnly) {
        button.classList.add("todo-status-btn-icon");
        button.setAttribute("aria-label", def.title || def.label);
        if (def.title) button.title = def.title;
      }
      const blockReason = getBlockReason(def);
      if (blockReason) {
        button.dataset.blocked = "1";
        button.dataset.blockReason = blockReason;
        button.classList.add("is-blocked");
      }
      row.appendChild(button);
    });
    section.appendChild(row);
    todoStatusMenu.appendChild(section);
  };
  const appendInfoSection = (title, text) => {
    const section = document.createElement("div");
    section.className = "todo-status-section";
    if (title) {
      const titleEl = document.createElement("div");
      titleEl.className = "todo-status-section-title";
      titleEl.textContent = title;
      section.appendChild(titleEl);
    }
    const note = document.createElement("div");
    note.className = "todo-status-section-note";
    note.textContent = text;
    section.appendChild(note);
    todoStatusMenu.appendChild(section);
  };
  todoStatusMenu.innerHTML = "";
  const contextBox = document.createElement("div");
  contextBox.className = "todo-status-context";
  const goMatrixButtonHtml = canShowGoMatrix
    ? `<button type="button" data-action="go_matrix" class="todo-status-context-matrix${
        goMatrixBlockReason ? " is-blocked" : ""
      }"${goMatrixBlockReason ? ` data-blocked="1" data-block-reason="${escapeHtml(goMatrixBlockReason)}"` : ""}>${goMatrixDef.label}</button>`
    : "";
  contextBox.innerHTML = `
    <div class="todo-status-context-meta">Commessa ${escapeHtml(commessaNumero)} · ${escapeHtml(commessaAnno)}</div>
    <div class="todo-status-context-main">
      <a href="#" data-action="open_commessa_detail" class="todo-status-context-link" title="${escapeHtml(commessaTitolo)}">${escapeHtml(commessaTitolo)}</a>
      ${goMatrixButtonHtml}
    </div>
  `;
  todoStatusMenu.appendChild(contextBox);
  appendSection("Stato attivita", statusActions);
  const clientSectionTitle = getTodoClientFlowSectionTitle(titolo);
  if (canAccessClientFlow) {
    appendSection(clientSectionTitle, clientActions);
  } else if (isClientTracked) {
    appendInfoSection(clientSectionTitle, "Livello 2 disponibile solo con stato DONE.");
  }
  if (clientActions.length) {
    const datePanel = document.createElement("div");
    datePanel.className = "todo-status-date-panel hidden";
    datePanel.innerHTML = `
      <div class="todo-status-date-title">Seleziona data</div>
      <div class="todo-status-date-dual">
        <div class="todo-status-date-calendar" data-field="sent">
          <div class="todo-status-date-subtitle">Invio</div>
          <div class="todo-status-date-head">
            <button type="button" data-action="todo_date_prev" data-field="sent" aria-label="Mese precedente invio">◀</button>
            <div class="todo-status-date-label"></div>
            <button type="button" data-action="todo_date_next" data-field="sent" aria-label="Mese successivo invio">▶</button>
          </div>
          <div class="milestone-weekdays">
            <span>L</span><span>M</span><span>M</span><span>G</span><span>V</span><span>S</span><span>D</span>
          </div>
          <div class="todo-status-date-days milestone-days"></div>
        </div>
        <div class="todo-status-date-calendar" data-field="due">
          <div class="todo-status-date-subtitle">Scadenza</div>
          <div class="todo-status-date-head">
            <button type="button" data-action="todo_date_prev" data-field="due" aria-label="Mese precedente scadenza">◀</button>
            <div class="todo-status-date-label"></div>
            <button type="button" data-action="todo_date_next" data-field="due" aria-label="Mese successivo scadenza">▶</button>
          </div>
          <div class="milestone-weekdays">
            <span>L</span><span>M</span><span>M</span><span>G</span><span>V</span><span>S</span><span>D</span>
          </div>
          <div class="todo-status-date-days milestone-days"></div>
        </div>
        <div class="todo-status-date-counter">
          <div class="todo-status-counter-label">Tempo rimanente</div>
          <div class="todo-status-counter-value">n/d</div>
        </div>
      </div>
      <div class="todo-status-date-confirm hidden">
        <div class="todo-status-date-calendar" data-field="confirm">
          <div class="todo-status-date-subtitle">Conferma</div>
          <div class="todo-status-date-head">
            <button type="button" data-action="todo_date_prev" data-field="confirm" aria-label="Mese precedente conferma">◀</button>
            <div class="todo-status-date-label"></div>
            <button type="button" data-action="todo_date_next" data-field="confirm" aria-label="Mese successivo conferma">▶</button>
          </div>
          <div class="milestone-weekdays">
            <span>L</span><span>M</span><span>M</span><span>G</span><span>V</span><span>S</span><span>D</span>
          </div>
          <div class="todo-status-date-days milestone-days"></div>
        </div>
      </div>
      <div class="todo-status-date-actions">
        <button type="button" data-action="todo_date_cancel">Annulla</button>
        <button type="button" data-action="todo_date_apply">Conferma data</button>
      </div>
    `;
    todoStatusMenu.appendChild(datePanel);
  }
  closeTodoStatusDatePicker();
  todoMenuTarget = { commessaId, titolo, reparto };
  todoMenuAnchorEl = null;
  todoStatusMenuManualPosition = false;
  todoStatusMenu.classList.remove("hidden");
  setTodoMenuFocus(rowKey);
  positionTodoStatusMenuCentered();
  requestAnimationFrame(() => positionTodoStatusMenuCentered());
}

function repositionTodoStatusMenuToAnchor() {
  if (!todoStatusMenu || todoStatusMenu.classList.contains("hidden")) return;
  if (todoStatusMenuManualPosition) {
    positionTodoStatusMenuWithinViewport();
    return;
  }
  positionTodoStatusMenuCentered();
}

function setTodoMenuFocus(rowKey = "") {
  todoFocusedRowKey = String(rowKey || "");
  document.querySelectorAll(".todo-cell.todo-cell-focus").forEach((cell) => {
    cell.classList.remove("todo-cell-focus");
  });
  if (todoFocusOverlay) todoFocusOverlay.classList.remove("hidden");
  document.body.classList.add("todo-menu-open");
  if (!todoFocusedRowKey) return;
  document.querySelectorAll(".todo-cell[data-todo-row]").forEach((cell) => {
    if (String(cell.dataset.todoRow || "") === todoFocusedRowKey) {
      cell.classList.add("todo-cell-focus");
    }
  });
}

function closeTodoStatusMenu() {
  if (!todoStatusMenu) return;
  todoStatusMenuDrag = null;
  todoStatusMenu.style.cursor = "";
  todoStatusMenuManualPosition = false;
  closeTodoStatusDatePicker();
  todoStatusMenu.classList.add("hidden");
  todoMenuTarget = null;
  todoMenuAnchorEl = null;
  releaseTodoMenuFocusIfNoneOpen();
}

function releaseTodoMenuFocusIfNoneOpen() {
  if (isTodoMenuOpen(todoStatusMenu) || isTodoMenuOpen(todoGlobalStatusMenu)) return;
  todoFocusedRowKey = null;
  document.querySelectorAll(".todo-cell.todo-cell-focus").forEach((cell) => {
    cell.classList.remove("todo-cell-focus");
  });
  document.body.classList.remove("todo-menu-open");
  if (todoFocusOverlay) todoFocusOverlay.classList.add("hidden");
}

function openTodoGlobalStatusMenu(cell, commessa, x, y) {
  if (!todoGlobalStatusMenu || !cell || !commessa?.id) return;
  closeTodoStatusMenu();
  if (!canEditTodoGlobalStatus()) {
    setStatus("Solo admin, responsabile e planner possono aggiornare questo stato.", "error");
    return;
  }
  const current = getTodoGlobalCommessaStatus(commessa).label;
  todoGlobalStatusMenu.innerHTML = "";
  TODO_GLOBAL_STATUS_ORDER.forEach((label) => {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.status = label;
    button.textContent = label.toUpperCase();
    if (label === current) {
      button.classList.add("active");
      button.disabled = true;
    }
    todoGlobalStatusMenu.appendChild(button);
  });
  todoGlobalMenuTarget = { commessaId: commessa.id };
  todoGlobalMenuAnchorEl = cell.querySelector(".todo-global-pill") || cell;
  todoGlobalMenuSuppress = true;
  setTodoMenuFocus(cell.dataset.todoRow || "");
  todoGlobalStatusMenu.classList.remove("hidden");
  const padding = 10;
  let left = x + 8;
  let top = y + 8;
  todoGlobalStatusMenu.style.left = `${left}px`;
  todoGlobalStatusMenu.style.top = `${top}px`;
  requestAnimationFrame(() => {
    const rect = todoGlobalStatusMenu.getBoundingClientRect();
    if (rect.right > window.innerWidth - padding) {
      left = Math.max(padding, window.innerWidth - rect.width - padding);
    }
    if (rect.bottom > window.innerHeight - padding) {
      top = Math.max(padding, window.innerHeight - rect.height - padding);
    }
    todoGlobalStatusMenu.style.left = `${left}px`;
    todoGlobalStatusMenu.style.top = `${top}px`;
  });
}

function closeTodoGlobalStatusMenu() {
  if (!todoGlobalStatusMenu) return;
  todoGlobalStatusMenu.classList.add("hidden");
  todoGlobalMenuTarget = null;
  todoGlobalMenuAnchorEl = null;
  releaseTodoMenuFocusIfNoneOpen();
}

function repositionTodoGlobalStatusMenuToAnchor() {
  if (!todoGlobalStatusMenu || todoGlobalStatusMenu.classList.contains("hidden")) return;
  if (!todoGlobalMenuAnchorEl || !document.body.contains(todoGlobalMenuAnchorEl)) return;
  const anchorRect = todoGlobalMenuAnchorEl.getBoundingClientRect();
  const padding = 10;
  let left = anchorRect.left + anchorRect.width / 2 + 8;
  let top = anchorRect.bottom + 8;
  todoGlobalStatusMenu.style.left = `${left}px`;
  todoGlobalStatusMenu.style.top = `${top}px`;
  const rect = todoGlobalStatusMenu.getBoundingClientRect();
  if (rect.right > window.innerWidth - padding) {
    left = Math.max(padding, window.innerWidth - rect.width - padding);
  }
  if (rect.bottom > window.innerHeight - padding) {
    top = Math.max(padding, anchorRect.top - rect.height - 8);
  }
  if (left < padding) left = padding;
  if (top < padding) top = padding;
  todoGlobalStatusMenu.style.left = `${left}px`;
  todoGlobalStatusMenu.style.top = `${top}px`;
}

function scrollToPanelTop(panelEl) {
  if (!panelEl) return;
  const rect = panelEl.getBoundingClientRect();
  const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
  const gap = typeof getSectionToolbarGap === "function" ? getSectionToolbarGap() : 0;
  const top = rect.top + window.scrollY - toolbarOffset - gap;
  window.scrollTo({ top: Math.max(0, top), left: 0, behavior: "smooth" });
}

function getTodoMatrixJumpEntry(target) {
  if (!target) return null;
  const entries = getTodoEntriesForCell(target.commessaId, target.titolo, target.reparto).filter(
    (entry) => normalizePhaseKey(entry.stato) !== "annullata"
  );
  if (!entries.length) return null;
  const plannedEntries = entries.filter((entry) => normalizePhaseKey(entry.stato) === "pianificata");
  const source = plannedEntries.length ? plannedEntries : entries;
  source.sort((a, b) => {
    const at = a.start instanceof Date ? a.start.getTime() : Number.MAX_SAFE_INTEGER;
    const bt = b.start instanceof Date ? b.start.getTime() : Number.MAX_SAFE_INTEGER;
    return at - bt;
  });
  return source[0] || null;
}

async function openTodoActivityInMatrix(target) {
  const entry = getTodoMatrixJumpEntry(target);
  if (!entry) {
    setStatus("Nessuna attivita pianificata trovata in matrice.", "error");
    return false;
  }
  if (matrixPanel) {
    matrixPanel.classList.remove("collapsed");
    scrollToPanelTop(matrixPanel);
  }
  const startTs = entry.start instanceof Date ? entry.start.getTime() : null;
  await jumpToMatrixActivity(target.commessaId, startTs, entry.id, entry.risorsaId);
  setStatus("Aperta attivita in matrice.", "ok");
  return true;
}

function setTodoClientFlowInState(row) {
  const normalized = normalizeTodoClientFlowRow(row);
  if (!normalized) return;
  const key = buildTodoClientFlowKey(normalized.commessaId, normalized.titolo, normalized.reparto);
  state.todoClientFlowMap.set(key, normalized);
  state.todoClientFlowLoadedCommesse.add(String(normalized.commessaId));
}

function removeTodoClientFlowFromState(target) {
  if (!target?.commessaId || !target?.titolo) return;
  const key = buildTodoClientFlowKey(target.commessaId, target.titolo, target.reparto || "");
  state.todoClientFlowMap.delete(key);
}

async function upsertTodoClientFlow(target, patch) {
  if (!target?.commessaId || !target?.titolo) return { ok: false, error: "Target non valido." };
  const payload = {
    commessa_id: target.commessaId,
    titolo: target.titolo,
    reparto: target.reparto || "",
    updated_by: state.profile?.id || null,
    ...patch,
  };
  const { data, error } = await supabase
    .from("commessa_attivita_cliente")
    .upsert(payload, { onConflict: "commessa_id,titolo,reparto" })
    .select("commessa_id,titolo,reparto,inviato_il,scadenza_il,esito,confermato_il,updated_at")
    .single();
  if (error) {
    return { ok: false, error: error.message };
  }
  if (data) setTodoClientFlowInState(data);
  return { ok: true };
}

async function deleteTodoClientFlow(target) {
  if (!target?.commessaId || !target?.titolo) return { ok: false, error: "Target non valido." };
  const { error } = await supabase
    .from("commessa_attivita_cliente")
    .delete()
    .eq("commessa_id", target.commessaId)
    .eq("titolo", target.titolo)
    .eq("reparto", target.reparto || "");
  if (error) {
    return { ok: false, error: error.message };
  }
  removeTodoClientFlowFromState(target);
  return { ok: true };
}

async function applyTodoClientFlowAction(action, target, options = {}) {
  if (!target) return false;
  const activityStatus = getTodoStatusFor(target.commessaId, target.titolo, target.reparto || "");
  if (activityStatus !== "fatta") {
    setStatus("Livello cliente disponibile solo quando l'attivita e DONE.", "error");
    return false;
  }
  const blockReason = getTodoClientFlowEditBlockReason(target.titolo, target.reparto || "");
  if (blockReason) {
    setStatus(blockReason, "error");
    return false;
  }
  const existing = getTodoClientFlow(target.commessaId, target.titolo, target.reparto);
  if (action === "client_sent") {
    const sentAt = parseIsoDateOnly(options.selectedDate) || startOfDay(new Date());
    const dueAt = parseIsoDateOnly(options.dueDate) || existing?.dueAt || getDefaultClientFlowDueDate(sentAt);
    if (dueAt && dayNumberUTC(dueAt) < dayNumberUTC(sentAt)) {
      setStatus("La scadenza non puo essere prima dell'invio.", "error");
      return false;
    }
    const result = await upsertTodoClientFlow(target, {
      inviato_il: formatIsoDateOnly(sentAt),
      scadenza_il: formatIsoDateOnly(dueAt),
      esito: TODO_CLIENT_FLOW_OUTCOME.IN_ATTESA,
      confermato_il: null,
    });
    if (!result.ok) {
      setStatus(`Update error: ${result.error}`, "error");
      return false;
    }
    refreshTodoVisualizationNow();
    refreshMatrixClientFlowIndicatorsForTarget(target);
    void refreshTodoData();
    setStatus("Tracking cliente aggiornato: inviato.", "ok");
    return true;
  }
  if (action === "client_reset") {
    const result = await deleteTodoClientFlow(target);
    if (!result.ok) {
      setStatus(`Update error: ${result.error}`, "error");
      return false;
    }
    refreshMatrixClientFlowIndicatorsForTarget(target);
    refreshTodoVisualizationNow();
    void refreshTodoData();
    setStatus("Tracking cliente resettato.", "ok");
    return true;
  }
  if (!existing?.sentAt) {
    setStatus("Prima registra la data di invio al cliente.", "error");
    return false;
  }
  if (action === "client_confirmed") {
    const confirmAt = parseIsoDateOnly(options.selectedDate) || startOfDay(new Date());
    const result = await upsertTodoClientFlow(target, {
      inviato_il: formatIsoDateOnly(existing.sentAt),
      scadenza_il: formatIsoDateOnly(existing.dueAt || getDefaultClientFlowDueDate(existing.sentAt)),
      esito: TODO_CLIENT_FLOW_OUTCOME.CONFERMATO,
      confermato_il: formatIsoDateOnly(confirmAt),
    });
    if (!result.ok) {
      setStatus(`Update error: ${result.error}`, "error");
      return false;
    }
    refreshTodoVisualizationNow();
    refreshMatrixClientFlowIndicatorsForTarget(target);
    void refreshTodoData();
    setStatus("Tracking cliente aggiornato: confermato.", "ok");
    return true;
  }
  if (action === "client_silence") {
    const decisionAt = parseIsoDateOnly(options.selectedDate) || startOfDay(new Date());
    const result = await upsertTodoClientFlow(target, {
      inviato_il: formatIsoDateOnly(existing.sentAt),
      scadenza_il: formatIsoDateOnly(existing.dueAt || getDefaultClientFlowDueDate(existing.sentAt)),
      esito: TODO_CLIENT_FLOW_OUTCOME.SILENZIO_ASSENSO,
      confermato_il: formatIsoDateOnly(decisionAt),
    });
    if (!result.ok) {
      setStatus(`Update error: ${result.error}`, "error");
      return false;
    }
    refreshTodoVisualizationNow();
    refreshMatrixClientFlowIndicatorsForTarget(target);
    void refreshTodoData();
    setStatus("Tracking cliente aggiornato: silenzio assenso.", "ok");
    return true;
  }
  return false;
}

async function setTodoOverride(commessaId, titolo, enabled) {
  if (!commessaId || !titolo) return false;
  if (!state.canWrite) return false;
  if (enabled) {
    const { error } = await supabase
      .from("commessa_attivita_override")
      .upsert({ commessa_id: commessaId, titolo, stato: "non_necessaria" }, { onConflict: "commessa_id,titolo" });
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      return false;
    }
  } else {
    const { error } = await supabase
      .from("commessa_attivita_override")
      .delete()
      .eq("commessa_id", commessaId)
      .eq("titolo", titolo);
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      return false;
    }
  }
  const key = String(commessaId);
  const list = state.todoOverridesMap.get(key) || new Map();
  if (enabled) {
    list.set(normalizePhaseKey(titolo), "non_necessaria");
  } else {
    list.delete(normalizePhaseKey(titolo));
  }
  state.todoOverridesMap.set(key, list);
  return true;
}

async function updateTodoActivitiesStatus(ids, stato) {
  if (!ids.length) return null;
  const { error } = await supabase.from("attivita").update({ stato }).in("id", ids);
  return error || null;
}

function applyMatrixActivityStatusPatch(ids, stato) {
  if (!Array.isArray(ids) || !ids.length) return;
  const idSet = new Set(ids.map((id) => String(id)).filter(Boolean));
  if (!idSet.size) return;
  const byId = new Map((matrixState.attivita || []).map((entry) => [String(entry.id), entry]));
  (matrixState.attivita || []).forEach((entry) => {
    const key = entry?.id != null ? String(entry.id) : "";
    if (idSet.has(key)) {
      entry.stato = stato;
    }
  });
  if (!matrixGrid) return;
  const isDone = stato === "completata";
  matrixGrid.querySelectorAll(".matrix-activity-bar[data-activity-id]").forEach((bar) => {
    const key = String(bar.dataset.activityId || "");
    if (!idSet.has(key)) return;
    bar.classList.toggle("is-done", isDone);
    const entry = byId.get(key);
    const fallback = entry || {
      commessa_id: bar.dataset.commessaId || "",
      titolo: bar.dataset.attivita || "",
      stato,
      reparto: bar.dataset.reparto || "",
    };
    applyMatrixBarClientFlowClasses(bar, fallback, bar.dataset.reparto || "");
  });
}

async function refreshTodoData() {
  if (!todoLastRendered || !todoLastRendered.length) {
    renderTodoSection(todoLastRendered);
    return;
  }
  await loadReportActivitiesFor(todoLastRendered);
  await loadTodoOverridesFor(todoLastRendered);
  await loadTodoClientFlowsFor(todoLastRendered);
}

async function applyTodoStatusAction(action, target, options = {}) {
  if (!target) return false;
  if (action === "go_matrix") {
    return openTodoActivityInMatrix(target);
  }
  if (action.startsWith("client_")) {
    return applyTodoClientFlowAction(action, target, options);
  }
  const { commessaId, titolo, reparto } = target;
  const entries = getTodoEntriesForCell(commessaId, titolo, reparto);
  const editableEntries = getTodoEditableEntries(entries, reparto);
  const editableIds = editableEntries.map((entry) => entry.id).filter(Boolean);

  if (action === "planned" || action === "done") {
    if (!editableIds.length) {
      setStatus("Permessi insufficienti su questa attivita.", "error");
      return false;
    }
    const overrideCleared = await setTodoOverride(commessaId, titolo, false);
    if (!overrideCleared) return false;
    const nextState = action === "done" ? "completata" : "pianificata";
    const error = await updateTodoActivitiesStatus(editableIds, nextState);
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      return false;
    }
    editableEntries.forEach((entry) => {
      entry.stato = nextState;
    });
    applyMatrixActivityStatusPatch(editableIds, nextState);
  } else if (action === "to_plan") {
    const overrideCleared = await setTodoOverride(commessaId, titolo, false);
    if (!overrideCleared) return false;
    if (editableIds.length) {
      const error = await updateTodoActivitiesStatus(editableIds, "annullata");
      if (error) {
        setStatus(`Update error: ${error.message}`, "error");
        return false;
      }
      editableEntries.forEach((entry) => {
        entry.stato = "annullata";
      });
      applyMatrixActivityStatusPatch(editableIds, "annullata");
    }
  } else if (action === "non_necessaria") {
    const overrideSet = await setTodoOverride(commessaId, titolo, true);
    if (!overrideSet) return false;
  } else {
    return false;
  }

  refreshTodoCellVisualFromState(target);
  void refreshTodoData();
  setStatus("Status aggiornato.", "ok");
  return true;
}

function renderReportGantt(commesse, today) {
  if (!matrixReportGrid) return;
  matrixReportGrid.classList.remove("report-list");
  matrixReportGrid.classList.add("report-gantt");
  matrixReportGrid.innerHTML = "";

  const isDateOk = (d) => d && !Number.isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
  const getStartDate = (c) => {
    const raw = c.data_ingresso ? new Date(c.data_ingresso) : null;
    return isDateOk(raw) ? raw : today;
  };
  const getTargetDate = (c) => {
    const raw = c.data_consegna_macchina || c.data_consegna_prevista || c.data_consegna || "";
    const d = raw ? new Date(raw) : null;
    return isDateOk(d) ? d : null;
  };
  const getOrdineDate = (c) => {
    const raw = c.data_ordine_telaio || "";
    const d = raw ? new Date(raw) : null;
    return isDateOk(d) ? d : null;
  };
  const getPrelievoDate = (c) => {
    const raw = c.data_prelievo || "";
    const d = raw ? new Date(raw) : null;
    return isDateOk(d) ? d : null;
  };
  const getKitCaviDate = (c) => {
    const raw = c.data_arrivo_kit_cavi || "";
    const d = raw ? new Date(raw) : null;
    return isDateOk(d) ? d : null;
  };
  const getTelaioEffDate = (c) => {
    const raw = c.data_consegna_telaio_effettiva || "";
    const d = raw ? new Date(raw) : null;
    return isDateOk(d) ? d : null;
  };

  let minDate = today;
  let maxDate = today;
  commesse.forEach((c) => {
    const start = getStartDate(c);
    const target = getTargetDate(c);
    const ordine = getOrdineDate(c);
    const prelievo = getPrelievoDate(c);
    const kitCavi = getKitCaviDate(c);
    const ordineDay = ordine ? startOfDay(ordine) : null;
    if (start && start < minDate) minDate = start;
    if (target && target > maxDate) maxDate = target;
    if (ordine && ordine < minDate) minDate = ordine;
    if (ordine && ordine > maxDate) maxDate = ordine;
    if (prelievo && prelievo < minDate) minDate = prelievo;
    if (prelievo && prelievo > maxDate) maxDate = prelievo;
    if (kitCavi && kitCavi < minDate) minDate = kitCavi;
    if (kitCavi && kitCavi > maxDate) maxDate = kitCavi;
  });
  if (!maxDate || maxDate < minDate) {
    maxDate = addDays(minDate, 30);
  }
  const boundsStart = normalizeBusinessDay(startOfDay(minDate), 1);
  const boundsEnd = normalizeBusinessDay(startOfDay(maxDate), 1);
  if (!state.reportRangeStart) {
    state.reportRangeStart = startOfWeek(addDays(today, -7));
  }
  if (!state.reportRangeDays) {
    state.reportRangeDays = REPORT_DEFAULT_DAYS;
  }
  let rangeDays = Math.min(REPORT_MAX_DAYS, Math.max(REPORT_MIN_DAYS, state.reportRangeDays));
  let rangeStart = normalizeBusinessDay(startOfDay(state.reportRangeStart), 1);
  if (rangeStart > boundsEnd) {
    rangeStart = startOfWeek(addDays(boundsEnd, -7));
  }
  let rangeEnd = addBusinessDays(rangeStart, rangeDays - 1);
  if (rangeEnd < boundsStart) {
    rangeStart = startOfWeek(addDays(boundsStart, -7));
    rangeEnd = addBusinessDays(rangeStart, rangeDays - 1);
  }
  state.reportRangeStart = rangeStart;
  state.reportRangeDays = rangeDays;
  const totalDays = rangeDays;

  const header = document.createElement("div");
  header.className = "report-gantt-header";
  header.innerHTML = `
    <div class="report-gantt-label">Commessa</div>
    <div class="report-gantt-track"></div>
  `;
  matrixReportGrid.appendChild(header);

  const trackHeader = header.querySelector(".report-gantt-track");
  let headerContent = null;
  let dayPx = null;
  if (trackHeader) {
    trackHeader.classList.add("report-gantt-track-header");
    trackHeader.style.setProperty("--day-step", `${100 / totalDays}%`);
    headerContent = document.createElement("div");
    headerContent.className = "report-gantt-header-content report-gantt-track-content";
    trackHeader.appendChild(headerContent);
    appendWeekBands(headerContent, rangeStart, rangeEnd, totalDays, true);
    appendMonthLabels(headerContent, rangeStart, rangeEnd, totalDays);
    appendDayNumbers(headerContent, rangeStart, rangeEnd, totalDays);
    const todayKey = normalizeBusinessDay(startOfDay(today), 1);
    if (todayKey >= rangeStart && todayKey <= rangeEnd) {
      const todayLeft = (businessDayDiff(rangeStart, todayKey) / totalDays) * 100;
      const todayWidth = (1 / totalDays) * 100;
      const todayBand = document.createElement("div");
      todayBand.className = "report-gantt-today-band";
      todayBand.style.left = `${todayLeft}%`;
      todayBand.style.width = `${todayWidth}%`;
      headerContent.appendChild(todayBand);
    }
    setupReportPanZoom(trackHeader, headerContent, { rangeStart, rangeDays, boundsStart, boundsEnd });
    const headerWidth = headerContent.getBoundingClientRect().width;
    if (headerWidth > 0) {
      dayPx = headerWidth / totalDays;
    }
  }

  const showYearHeaders = !getReportFilters().orderDue;
  let lastYear = null;
  commesse.forEach((c) => {
    const milestoneStacks = new Map();
    const stackMilestone = (marker, dayKey, base = 2) => {
      if (!marker || !dayKey) return;
      const count = milestoneStacks.get(dayKey) || 0;
      marker.style.bottom = `${base + count * 12}px`;
      milestoneStacks.set(dayKey, count + 1);
    };
    const yearKey = String(c.anno || getCommessaYear(c) || "Senza anno");
    if (showYearHeaders && yearKey !== lastYear) {
      const yearHeader = document.createElement("div");
      yearHeader.className = "report-year-header";
      yearHeader.textContent = yearKey;
      matrixReportGrid.appendChild(yearHeader);
      lastYear = yearKey;
    }
    const start = getStartDate(c);
    const target = getTargetDate(c);
    const ordine = getOrdineDate(c);
    const prelievo = getPrelievoDate(c);
    const kitCavi = getKitCaviDate(c);
    const ordinePianificato = getTelaioEffDate(c);
    const missingPrelievo = !prelievo;
    const missingKitCavi = !kitCavi;
    const missingOrdineTarget = !ordine;
    const missingPlanning = missingOrdineTarget || missingPrelievo;
    const missingTarget = !target;
    const ordineDay = ordine ? normalizeBusinessDay(startOfDay(ordine), 1) : null;
    const prelievoDay = prelievo ? normalizeBusinessDay(startOfDay(prelievo), 1) : null;
    const kitCaviDay = kitCavi ? normalizeBusinessDay(startOfDay(kitCavi), 1) : null;
    const ordinePianificatoDay = ordinePianificato ? normalizeBusinessDay(startOfDay(ordinePianificato), 1) : null;
    const telaioConsegnato = Boolean(c.telaio_consegnato);
    const plannedMatchesTarget = Boolean(
      ordineDay &&
        ordinePianificatoDay &&
        formatDateLocal(ordinePianificatoDay) === formatDateLocal(ordineDay)
    );
    const highlightPlannedOnTarget = plannedMatchesTarget;
    const showPlannedOrdine = Boolean(ordinePianificatoDay && (!ordineDay || !plannedMatchesTarget));
    const plannedIsLate = Boolean(ordineDay && ordinePianificatoDay && ordinePianificatoDay.getTime() > ordineDay.getTime());
    const markTargetDelivered = Boolean(telaioConsegnato && ordineDay && !showPlannedOrdine);
    const startDay = start ? normalizeBusinessDay(startOfDay(start), 1) : rangeStart;
    const targetDay = target ? normalizeBusinessDay(startOfDay(target), 1) : null;

    const row = document.createElement("div");
    row.className = "report-gantt-row";
    const label = document.createElement("div");
    label.className = "report-gantt-label";
    const codeInfo = parseCommessaCode(c.codice || "");
    const annoLabel = c.anno ?? codeInfo.year ?? "";
    const numeroLabel = c.numero ?? codeInfo.num ?? c.codice ?? "-";
    const titoloLabel = c.titolo || "Senza titolo";
    label.innerHTML = `
      <button class="report-title report-commessa-link report-commessa-block" type="button" data-commessa-id="${c.id}">
        <div class="report-commessa-head">
          <span class="report-commessa-num">${escapeHtml(String(numeroLabel))}</span>
          ${annoLabel ? `<span class="report-commessa-year">${escapeHtml(String(annoLabel))}</span>` : ""}
        </div>
        <div class="report-commessa-desc" title="${escapeHtml(titoloLabel)}">${escapeHtml(titoloLabel)}</div>
      </button>
    `;
    if (missingPlanning && missingTarget) {
      label.classList.add("missing-both");
    } else if (missingPlanning && !missingTarget) {
      label.classList.add("missing-ordine");
    }
    const ganttBtn = label.querySelector(".report-commessa-link");
    if (ganttBtn) {
      ganttBtn.addEventListener("click", () => {
        selectCommessa(c.id, { openModal: true });
      });
    }
    const track = document.createElement("div");
    track.className = "report-gantt-track report-gantt-track-pan";
    track.dataset.rangeStart = formatDateLocal(rangeStart);
    track.dataset.totalDays = String(totalDays);
    track.dataset.commessaId = String(c.id);
    track.style.setProperty("--day-step", `${100 / totalDays}%`);
    const trackContent = document.createElement("div");
    trackContent.className = "report-gantt-track-content";
    if (dayPx) {
      trackContent.style.setProperty("--day-px", `${dayPx}px`);
    }
    track.appendChild(trackContent);
    const trackWrap = document.createElement("div");
    trackWrap.className = "report-gantt-track-wrap";
    trackWrap.appendChild(track);
    appendWeekBands(trackContent, rangeStart, rangeEnd, totalDays, false);
    if (headerContent) {
      setupReportPanZoom(track, headerContent, { rangeStart, rangeDays, boundsStart, boundsEnd });
    }
    if (today >= rangeStart && today <= rangeEnd) {
      const todayKey = normalizeBusinessDay(startOfDay(today), 1);
      const todayLeft = (businessDayDiff(rangeStart, todayKey) / totalDays) * 100;
      const todayWidth = (1 / totalDays) * 100;
      const todayBand = document.createElement("div");
      todayBand.className = "report-gantt-today-band";
      todayBand.style.left = `${todayLeft}%`;
      todayBand.style.width = `${todayWidth}%`;
      trackContent.appendChild(todayBand);
    }

    const leftPct = (businessDayDiff(rangeStart, startDay) / totalDays) * 100;
    const offsetDays = businessDayDiff(rangeStart, startDay);
    const barOffsetPx = dayPx ? offsetDays * dayPx : 0;
    if (targetDay) {
      const widthDays = Math.max(0, businessDayDiff(startDay, targetDay) + 0.5);
      const widthPct = Math.max(1, (widthDays / totalDays) * 100);
      const bar = document.createElement("div");
      bar.className = "report-gantt-bar";
      if (missingPlanning) {
        bar.classList.add("is-missing-ordine");
      }
      if (dayPx) {
        bar.style.setProperty("--bar-offset", `${barOffsetPx}px`);
      }
      const diff = businessDayDiff(today, targetDay);
      if (diff < 0) bar.classList.add("is-late");
      bar.style.left = `${leftPct}%`;
      bar.style.width = `${widthPct}%`;
      bar.title = `Target: ${formatDateDMY(targetDay)}`;
      trackContent.appendChild(bar);
    } else if (ordineDay && missingTarget) {
      const widthDays = Math.max(0, businessDayDiff(startDay, ordineDay) + 0.5);
      const widthPct = Math.max(1, (widthDays / totalDays) * 100);
      const bar = document.createElement("div");
      bar.className = "report-gantt-bar is-ordine-only";
      if (missingPlanning) {
        bar.classList.add("is-missing-ordine");
      }
      if (dayPx) {
        bar.style.setProperty("--bar-offset", `${barOffsetPx}px`);
      }
      bar.style.left = `${leftPct}%`;
      bar.style.width = `${widthPct}%`;
      bar.title = `Ordine telaio: ${formatDateDMY(ordineDay)}`;
      trackContent.appendChild(bar);
    } else {
      const dot = document.createElement("div");
      dot.className = "report-gantt-dot";
      dot.style.left = `${leftPct}%`;
      trackContent.appendChild(dot);
    }

    if (ordineDay) {
      if (ordineDay >= rangeStart && ordineDay <= rangeEnd) {
        const ordineIndex = businessDayDiff(rangeStart, ordineDay);
        const ordineLeft = ((ordineIndex + 0.5) / totalDays) * 100;
        const milestone = document.createElement("button");
        milestone.type = "button";
        milestone.className = "report-gantt-milestone";
        if (markTargetDelivered) milestone.classList.add("is-telaio-delivered");
        if (highlightPlannedOnTarget) milestone.classList.add("is-telaio-on-time");
        milestone.style.left = `${ordineLeft}%`;
        milestone.textContent = "T";
        milestone.title = highlightPlannedOnTarget
          ? `Ordine telaio (stimato): ${formatDateDMY(ordineDay)}`
          : `Ordine telaio (target): ${formatDateDMY(ordineDay)}${markTargetDelivered ? " \u2022 consegnato" : ""}`;
        milestone.dataset.commessaId = String(c.id);
        milestone.dataset.field = "ordine";
        if (canEditGanttMilestone("ordine")) {
          milestone.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            beginMilestoneDrag(e, {
              commessaId: c.id,
              field: "ordine",
              date: ordineDay,
              track,
              line: null,
              marker: milestone,
            });
          });
          milestone.addEventListener("click", (e) => {
            e.stopPropagation();
            if (Date.now() < milestoneSuppressClickUntil) return;
            showMilestonePicker({
              commessaId: c.id,
              field: "ordine",
              date: ordineDay,
              anchorRect: milestone.getBoundingClientRect(),
            });
          });
        } else {
          milestone.classList.add("is-locked");
          milestone.setAttribute("aria-disabled", "true");
          milestone.title += " \u2022 bloccato";
        }
        trackContent.appendChild(milestone);
        stackMilestone(milestone, formatDateLocal(ordineDay), 2);
      }
    }
    if (showPlannedOrdine) {
      if (ordinePianificatoDay && ordinePianificatoDay >= rangeStart && ordinePianificatoDay <= rangeEnd) {
        const effIndex = businessDayDiff(rangeStart, ordinePianificatoDay);
        const effLeft = ((effIndex + 0.5) / totalDays) * 100;
        const effMarker = document.createElement("button");
        effMarker.type = "button";
        effMarker.className = "report-gantt-milestone is-telaio-actual";
        effMarker.style.left = `${effLeft}%`;
        effMarker.textContent = "T";
        if (plannedIsLate) {
          effMarker.classList.add("is-telaio-late");
        } else {
          effMarker.classList.add("is-telaio-on-time");
        }
        if (telaioConsegnato) effMarker.classList.add("is-telaio-delivered");
        effMarker.title = `Ordine telaio (stimato): ${formatDateDMY(ordinePianificatoDay)}${
          telaioConsegnato ? " \u2022 ordinato" : ""
        }`;
        effMarker.dataset.commessaId = String(c.id);
        effMarker.dataset.field = "ordine_pianificato";
        if (telaioConsegnato) {
          effMarker.classList.add("is-locked");
          effMarker.setAttribute("aria-disabled", "true");
          effMarker.addEventListener("click", (e) => {
            e.stopPropagation();
            if (Date.now() < milestoneSuppressClickUntil) return;
            showMilestonePicker({
              commessaId: c.id,
              field: "ordine_pianificato",
              date: ordinePianificatoDay,
              anchorRect: effMarker.getBoundingClientRect(),
            });
          });
        } else if (canEditGanttMilestone("ordine_pianificato")) {
          effMarker.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            beginMilestoneDrag(e, {
              commessaId: c.id,
              field: "ordine_pianificato",
              date: ordinePianificatoDay,
              track,
              line: null,
              marker: effMarker,
              compareTarget: ordineDay,
            });
          });
          effMarker.addEventListener("click", (e) => {
            e.stopPropagation();
            if (Date.now() < milestoneSuppressClickUntil) return;
            showMilestonePicker({
              commessaId: c.id,
              field: "ordine_pianificato",
              date: ordinePianificatoDay,
              anchorRect: effMarker.getBoundingClientRect(),
            });
          });
        } else {
          effMarker.classList.add("is-locked");
          effMarker.setAttribute("aria-disabled", "true");
        }
        trackContent.appendChild(effMarker);
        stackMilestone(effMarker, formatDateLocal(ordinePianificatoDay), 2);
      }
    }
    if (kitCaviDay) {
      if (kitCaviDay >= rangeStart && kitCaviDay <= rangeEnd) {
        const kitIndex = businessDayDiff(rangeStart, kitCaviDay);
        const kitLeft = ((kitIndex + 0.5) / totalDays) * 100;
        const kitMarker = document.createElement("button");
        kitMarker.type = "button";
        kitMarker.className = "report-gantt-milestone is-kit-cavi";
        kitMarker.style.left = `${kitLeft}%`;
        kitMarker.textContent = "\uD83D\uDD0C";
        kitMarker.title = `Arrivo kit cavi: ${formatDateDMY(kitCaviDay)}`;
        kitMarker.dataset.commessaId = String(c.id);
        kitMarker.dataset.field = "kit_cavi";
        if (canEditGanttMilestone("kit_cavi")) {
          kitMarker.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            beginMilestoneDrag(e, {
              commessaId: c.id,
              field: "kit_cavi",
              date: kitCaviDay,
              track,
              line: null,
              marker: kitMarker,
            });
          });
          kitMarker.addEventListener("click", (e) => {
            e.stopPropagation();
            if (Date.now() < milestoneSuppressClickUntil) return;
            showMilestonePicker({
              commessaId: c.id,
              field: "kit_cavi",
              date: kitCaviDay,
              anchorRect: kitMarker.getBoundingClientRect(),
            });
          });
        } else {
          kitMarker.classList.add("is-locked");
          kitMarker.setAttribute("aria-disabled", "true");
          kitMarker.title += " \u2022 bloccato";
        }
        trackContent.appendChild(kitMarker);
        stackMilestone(kitMarker, formatDateLocal(kitCaviDay), 2);
      }
    }
    if (prelievoDay) {
      if (prelievoDay >= rangeStart && prelievoDay <= rangeEnd) {
        const prelievoIndex = businessDayDiff(rangeStart, prelievoDay);
        const prelievoLeft = ((prelievoIndex + 0.5) / totalDays) * 100;
        const prelievoMarker = document.createElement("button");
        prelievoMarker.type = "button";
        prelievoMarker.className = "report-gantt-milestone is-prelievo";
        prelievoMarker.style.left = `${prelievoLeft}%`;
        prelievoMarker.textContent = "\uD83D\uDCE6";
        prelievoMarker.title = `Prelievo materiali: ${formatDateDMY(prelievoDay)}`;
        prelievoMarker.dataset.commessaId = String(c.id);
        prelievoMarker.dataset.field = "prelievo";
        if (canEditGanttMilestone("prelievo")) {
          prelievoMarker.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            beginMilestoneDrag(e, {
              commessaId: c.id,
              field: "prelievo",
              date: prelievoDay,
              track,
              line: null,
              marker: prelievoMarker,
            });
          });
          prelievoMarker.addEventListener("click", (e) => {
            e.stopPropagation();
            if (Date.now() < milestoneSuppressClickUntil) return;
            showMilestonePicker({
              commessaId: c.id,
              field: "prelievo",
              date: prelievoDay,
              anchorRect: prelievoMarker.getBoundingClientRect(),
            });
          });
        } else {
          prelievoMarker.classList.add("is-locked");
          prelievoMarker.setAttribute("aria-disabled", "true");
          prelievoMarker.title += " \u2022 bloccato";
        }
        trackContent.appendChild(prelievoMarker);
        stackMilestone(prelievoMarker, formatDateLocal(prelievoDay), 2);
      }
    }
    if (targetDay) {
      const consegnaIndex = businessDayDiff(rangeStart, targetDay);
      const consegnaLeft = ((consegnaIndex + 0.5) / totalDays) * 100;
      const consegna = document.createElement("button");
      consegna.type = "button";
      consegna.className = "report-gantt-milestone is-consegna";
      consegna.style.left = `${consegnaLeft}%`;
      const transport = c.trasporto_consegna || "van";
      consegna.textContent = transport === "ship" ? "\uD83D\uDEA2" : "\uD83D\uDE9A";
      consegna.title =
        (transport === "ship" ? "Delivery by ship" : "Delivery by van") +
        ` \u2022 ${formatDateDMY(targetDay)}`;
      consegna.dataset.commessaId = String(c.id);
      consegna.dataset.field = "consegna";
      consegna.dataset.transport = transport;
      if (canEditGanttMilestone("consegna")) {
        consegna.addEventListener("mousedown", (e) => {
          e.stopPropagation();
          beginMilestoneDrag(e, {
            commessaId: c.id,
            field: "consegna",
            date: targetDay,
            track,
            line: null,
            marker: consegna,
          });
        });
        consegna.addEventListener("click", (e) => {
          e.stopPropagation();
          if (Date.now() < milestoneSuppressClickUntil) return;
          showMilestonePicker({
            commessaId: c.id,
            field: "consegna",
            date: targetDay,
            anchorRect: consegna.getBoundingClientRect(),
          });
        });
      } else {
        consegna.classList.add("is-locked");
        consegna.setAttribute("aria-disabled", "true");
        consegna.title += " \u2022 bloccato";
      }
      trackContent.appendChild(consegna);
      stackMilestone(consegna, formatDateLocal(targetDay), 2);
    }

    const schedule = state.reportActivitiesMap.get(String(c.id));
    if (!schedule) {
      const loading = document.createElement("div");
      loading.className = "report-gantt-empty report-pill report-pill-orange";
      loading.textContent = "Schedulazioni in caricamento";
      label.appendChild(loading);
    } else if (!schedule.length) {
      const empty = document.createElement("div");
      empty.className = "report-gantt-empty report-pill report-pill-orange";
      empty.textContent = "Nessuna attivita schedulata";
      label.appendChild(empty);
    } else {
      const deptRanges = new Map();
      schedule.forEach((a) => {
        if (!a.start) return;
        const rawStart = normalizeBusinessDay(startOfDay(a.start), 1);
        const rawEnd = normalizeBusinessDay(startOfDay(a.end || a.start), -1);
        const safeStart = rawStart > rawEnd ? rawStart : rawStart;
        const safeEnd = rawStart > rawEnd ? rawStart : rawEnd;
        if (safeEnd < rangeStart || safeStart > rangeEnd) return;
        const startClamp = safeStart < rangeStart ? rangeStart : safeStart;
        const endClamp = safeEnd > rangeEnd ? rangeEnd : safeEnd;
        const bucket = a.dept || "ALTRO";
        const current = deptRanges.get(bucket);
        if (!current) {
          deptRanges.set(bucket, { start: startClamp, end: endClamp });
        } else {
          if (startClamp < current.start) current.start = startClamp;
          if (endClamp > current.end) current.end = endClamp;
        }
      });

      if (!deptRanges.size) {
        const empty = document.createElement("div");
        empty.className = "report-gantt-empty report-pill report-pill-orange";
        empty.textContent = "Nessuna attivita schedulata";
        label.appendChild(empty);
      } else {
        const deptOrder = [
          { key: "CAD", cls: "dept-cad", label: "CAD" },
          { key: "TERMODINAMICI", cls: "dept-termo", label: "Termodinamici" },
          { key: "ELETTRICI", cls: "dept-elett", label: "Elettrici" },
          { key: "ALTRO", cls: "dept-other", label: "Altro" },
        ];
        const laneHeight = 12;
        const barHeight = 8;
        const topPadding = 4;
        let laneIndex = 0;
        deptOrder.forEach((dept) => {
          const range = deptRanges.get(dept.key);
          if (!range) return;
          const left = (businessDayDiff(rangeStart, range.start) / totalDays) * 100;
          const widthDays = Math.max(1, businessDayDiff(range.start, range.end) + 1);
          const width = (widthDays / totalDays) * 100;
          const bar = document.createElement("div");
          bar.className = `report-gantt-activity ${dept.cls}`;
          bar.style.left = `${left}%`;
          bar.style.width = `${width}%`;
          bar.style.top = `${topPadding + laneIndex * laneHeight}px`;
          bar.style.height = `${barHeight}px`;
          bar.title = `${dept.label}: ${formatDateDMY(range.start)} \u2192 ${formatDateDMY(range.end)}`;
          bar.addEventListener("click", (e) => {
            e.stopPropagation();
            showReportDeptMenu(e.clientX, e.clientY, dept.label, schedule, dept.key, c.id);
          });
          trackContent.appendChild(bar);
          laneIndex += 1;
        });
        const neededHeight = Math.max(28, topPadding + laneIndex * laneHeight + 10);
        track.style.height = `${neededHeight}px`;
      }
    }

    row.appendChild(label);
    row.appendChild(trackWrap);
    matrixReportGrid.appendChild(row);
  });
}

function appendWeekBands(container, rangeStart, rangeEnd, totalDays, withLabels) {
  const bands = document.createElement("div");
  bands.className = "report-gantt-week-bands";
  container.appendChild(bands);
  let current = startOfWeek(rangeStart);
  let index = 0;
  while (current <= rangeEnd) {
    const weekStart = current;
    const weekEnd = addDays(weekStart, 4);
    const startClamp = weekStart < rangeStart ? rangeStart : weekStart;
    const endClamp = weekEnd > rangeEnd ? rangeEnd : weekEnd;
    const left = (businessDayDiff(rangeStart, startClamp) / totalDays) * 100;
    const widthDays = Math.max(1, businessDayDiff(startClamp, endClamp) + 1);
    const width = (widthDays / totalDays) * 100;
    const band = document.createElement("div");
    band.className = `report-gantt-week-band ${index % 2 ? "is-alt" : "is-base"}`;
    band.style.left = `${left}%`;
    band.style.width = `${width}%`;
    bands.appendChild(band);
    const boundary = document.createElement("div");
    boundary.className = "report-gantt-week-boundary";
    boundary.style.left = `${left}%`;
    bands.appendChild(boundary);
    if (withLabels) {
      const label = document.createElement("div");
      label.className = "report-gantt-week-label";
      const dayStep = 100 / totalDays;
      const labelLeft = Math.min(100, left + dayStep / 2);
      const labelWidth = Math.max(0, width - dayStep / 2);
      label.style.left = `${labelLeft}%`;
      label.style.width = `${labelWidth}%`;
      label.textContent = `week ${getIsoWeekNumber(weekStart)}`;
      bands.appendChild(label);
    }
    current = addDays(weekStart, 7);
    index += 1;
  }
}

function appendMonthLabels(container, rangeStart, rangeEnd, totalDays) {
  const row = document.createElement("div");
  row.className = "report-gantt-months";
  container.appendChild(row);
  let current = new Date(rangeStart);
  let index = 0;
  while (current <= rangeEnd) {
    const monthStartIndex = index;
    const month = current.getMonth();
    const year = current.getFullYear();
    while (current <= rangeEnd && current.getMonth() === month && current.getFullYear() === year) {
      current = addBusinessDays(current, 1);
      index += 1;
    }
    const spanDays = Math.max(1, index - monthStartIndex);
    const left = (monthStartIndex / totalDays) * 100;
    const width = (spanDays / totalDays) * 100;
    const label = document.createElement("div");
    label.className = `report-gantt-month-label ${month % 2 ? "is-alt" : "is-base"}`;
    label.style.left = `${left}%`;
    label.style.width = `${width}%`;
    label.textContent = new Date(year, month, 1).toLocaleDateString("it-IT", { month: "long" });
    row.appendChild(label);
  }
}

function appendDayNumbers(container, rangeStart, rangeEnd, totalDays) {
  const row = document.createElement("div");
  row.className = "report-gantt-day-numbers";
  if (totalDays > 90) {
    row.classList.add("is-sparse");
  }
  container.appendChild(row);
  let current = new Date(rangeStart);
  let index = 0;
  while (current <= rangeEnd) {
    const left = (index / totalDays) * 100;
    const width = (1 / totalDays) * 100;
    const label = document.createElement("div");
    label.className = "report-gantt-day-number";
    label.style.left = `${left}%`;
    label.style.width = `${width}%`;
    label.textContent = String(current.getDate()).padStart(2, "0");
    if (index === 0) {
      label.classList.add("is-range-start");
    }
    if (current.getDay() === 1) {
      label.classList.add("is-monday");
    }
    row.appendChild(label);
    current = addBusinessDays(current, 1);
    index += 1;
  }
}

function setupReportPanZoom(track, headerContent, { rangeStart, rangeDays, boundsStart, boundsEnd }) {
  if (!track || !headerContent) return;
  track.onpointerdown = (e) => {
    if (e.button !== 0) return;
    if (e.target && e.target.closest(".report-gantt-milestone, .report-gantt-activity")) return;
    e.preventDefault();
    const rect = track.getBoundingClientRect();
    const previewTargets = Array.from(
      (matrixReportGrid || document).querySelectorAll(".report-gantt-track-content")
    );
    reportPanZoom = {
      startX: e.clientX,
      startY: e.clientY,
      baseStart: rangeStart,
      baseDays: rangeDays,
      previewTargets,
      rect,
      boundsStart,
      boundsEnd,
      moved: false,
      mode: null,
    };
    track.classList.add("is-dragging");
    track.setPointerCapture?.(e.pointerId);
  };
  track.onpointermove = (e) => {
    if (!reportPanZoom) return;
    const dx = e.clientX - reportPanZoom.startX;
    const dy = e.clientY - reportPanZoom.startY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) reportPanZoom.moved = true;
    if (!reportPanZoom.mode) {
      if (Math.abs(dy) > 3 && Math.abs(dy) >= Math.abs(dx)) reportPanZoom.mode = "zoom";
      else if (Math.abs(dx) > 3) reportPanZoom.mode = "pan";
    }
    const daysPerPx = reportPanZoom.baseDays / Math.max(1, reportPanZoom.rect.width);
    const shiftDays = reportPanZoom.mode === "pan" ? Math.round(-dx * daysPerPx) : 0;
    const zoomSteps = reportPanZoom.mode === "zoom" ? Math.round(dy / 18) : 0;
    let newDays = reportPanZoom.baseDays + zoomSteps * 8;
    newDays = Math.min(REPORT_MAX_DAYS, Math.max(REPORT_MIN_DAYS, newDays));
    const center = addBusinessDays(reportPanZoom.baseStart, Math.round(reportPanZoom.baseDays / 2));
    let newStart = addBusinessDays(center, -Math.round(newDays / 2));
    newStart = addBusinessDays(newStart, shiftDays);
    reportPanZoom.previewStart = normalizeBusinessDay(startOfDay(newStart), 1);
    reportPanZoom.previewDays = newDays;
    applyReportHeaderPreview(reportPanZoom, reportPanZoom.previewStart, reportPanZoom.previewDays);
  };
  track.onpointerup = (e) => {
    if (!reportPanZoom) return;
    track.classList.remove("is-dragging");
    track.releasePointerCapture?.(e.pointerId);
    const drag = reportPanZoom;
    clearReportPreview(drag);
    reportPanZoom = null;
    if (!drag.moved) return;
    if (drag.previewStart) {
      let nextStart = drag.previewStart;
      let nextDays = drag.previewDays || drag.baseDays;
      if (nextStart > drag.boundsEnd) {
        nextStart = startOfWeek(addDays(drag.boundsEnd, -7));
      }
      let nextEnd = addBusinessDays(nextStart, nextDays - 1);
      if (nextEnd < drag.boundsStart) {
        nextStart = startOfWeek(addDays(drag.boundsStart, -7));
        nextEnd = addBusinessDays(nextStart, nextDays - 1);
      }
      state.reportRangeStart = nextStart;
      state.reportRangeDays = nextDays;
      renderMatrixReport();
    }
  };
  track.onpointerleave = () => {
    if (!reportPanZoom) return;
    track.classList.remove("is-dragging");
    clearReportPreview(reportPanZoom);
  };
}

function applyReportHeaderPreview(drag, previewStart, previewDays) {
  if (!drag || !drag.previewTargets || !previewStart || !previewDays) return;
  const baseDays = drag.baseDays;
  const width = Math.max(1, drag.rect.width);
  const shiftDays = businessDayDiff(drag.baseStart, previewStart);
  const scale = baseDays / previewDays;
  const shiftPx = -scale * (shiftDays / baseDays) * width;
  drag.previewTargets.forEach((target) => {
    target.style.transformOrigin = "left center";
    target.style.transform = `translateX(${shiftPx}px) scaleX(${scale})`;
    target.style.setProperty("--preview-scale", scale.toFixed(4));
    target.style.setProperty("--preview-inv-scale", (1 / scale).toFixed(4));
    target.classList.add("is-previewing");
  });
}

function clearReportPreview(drag) {
  if (!drag || !drag.previewTargets) return;
  drag.previewTargets.forEach((target) => {
    target.style.transform = "";
    target.style.transformOrigin = "";
    target.style.removeProperty("--preview-scale");
    target.style.removeProperty("--preview-inv-scale");
    target.classList.remove("is-previewing");
  });
}

async function loadMatrixAttivita() {
  if (!state.session) return;
  debugLog("loadMatrixAttivita query");
  const start = startOfWeek(matrixState.date);
  const end = matrixState.view === "six"
    ? endOfDay(addDays(start, 41))
    : matrixState.view === "three"
    ? endOfDay(addDays(start, 20))
    : matrixState.view === "two"
    ? endOfDay(addDays(start, 13))
    : endOfDay(addDays(start, 6));
  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase
          .from("attivita")
          .select("*")
          .lte("data_inizio", end.toISOString())
          .gte("data_fine", start.toISOString())
          .order("data_inizio", { ascending: true }),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadMatrixAttivita timeout: ${err?.message || err}`);
    }
  }
  if (error || !data) {
    debugLog("loadMatrixAttivita fallback REST");
    const query = [
      "select=*",
      `data_inizio=lte.${encodeURIComponent(end.toISOString())}`,
      `data_fine=gte.${encodeURIComponent(start.toISOString())}`,
      "order=data_inizio.asc",
    ].join("&");
    const rest = await fetchTableViaRest("attivita", query);
    if (!rest) {
      setStatus(`Matrix error: ${error?.message || "no data"}`, "error");
      return;
    }
    matrixState.attivita = rest || [];
  } else {
    matrixState.attivita = data || [];
  }
  renderMatrix();
  renderMatrixReport();
}

async function loadAttivita() {
  if (!state.session) return;
  debugLog("loadAttivita query");
  const view = calendarState.view;
  const baseDate = calendarState.date;
  let start;
  let end;
  if (view === "day") {
    start = new Date(baseDate);
    start.setHours(0, 0, 0, 0);
    end = endOfDay(baseDate);
  } else {
    start = startOfWeek(baseDate);
    end = endOfDay(addDays(start, 6));
  }

  let data = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase
          .from("attivita")
          .select("*")
          .lte("data_inizio", end.toISOString())
          .gte("data_fine", start.toISOString())
          .order("data_inizio", { ascending: true }),
        FETCH_TIMEOUT_MS
      );
      data = result.data;
      error = result.error;
    } catch (err) {
      debugLog(`loadAttivita timeout: ${err?.message || err}`);
    }
  }

  if (error || !data) {
    debugLog("loadAttivita fallback REST");
    const query = [
      "select=*",
      `data_inizio=lte.${encodeURIComponent(end.toISOString())}`,
      `data_fine=gte.${encodeURIComponent(start.toISOString())}`,
      "order=data_inizio.asc",
    ].join("&");
    const rest = await fetchTableViaRest("attivita", query);
    if (!rest) {
      setStatus(`Activity error: ${error?.message || "no data"}`, "error");
      return;
    }
    calendarState.attivita = rest || [];
  } else {
    calendarState.attivita = data || [];
  }
  renderCalendar(start, end);
}

function renderCalendar(start, end) {
  calendarGrid.innerHTML = "";
  if (!state.session) {
    calendarGrid.innerHTML = `<div class="calendar-empty">Accedi per vedere le attivita.</div>`;
    return;
  }

  const days = [];
  let current = new Date(start);
  while (current <= end) {
    days.push(new Date(current));
    current = addDays(current, 1);
  }

  calendarGrid.style.gridTemplateColumns =
    calendarState.view === "day" ? "1fr" : "repeat(7, minmax(0, 1fr))";

  days.forEach((day) => {
    const col = document.createElement("div");
    col.className = "calendar-day";
    const today = startOfDay(new Date());
    if (startOfDay(day).getTime() === today.getTime()) {
      col.classList.add("calendar-today");
    }
    const label = day.toLocaleDateString("it-IT", {
      weekday: "short",
      day: "2-digit",
      month: "short",
    });
    col.innerHTML = `<h4>${label}</h4>`;
    calendarGrid.appendChild(col);
  });

  if (!calendarState.attivita.length) {
    calendarGrid.firstChild.insertAdjacentHTML(
      "beforeend",
      `<div class="calendar-empty">Nessuna attivita in questo periodo.</div>`
    );
    return;
  }

  calendarState.attivita.forEach((a) => {
    const startDate = new Date(a.data_inizio);
    const dayIndex =
      calendarState.view === "day"
        ? 0
        : Math.max(0, Math.min(6, Math.floor((startDate - start) / 86400000)));
    const col = calendarGrid.children[dayIndex];
    if (!col) return;

    const card = document.createElement("div");
    card.className = "calendar-card";
    const deptKey = normalizeDeptKey(getRisorsaDeptName(a.risorsa_id));
    if (deptKey === "CAD") card.classList.add("calendar-cad");
    if (deptKey === "TERMODINAMICI") card.classList.add("calendar-termo");
    if (deptKey === "ELETTRICI") card.classList.add("calendar-elett");
    card.innerHTML = `
      <strong>${a.titolo}</strong>
      <div class="calendar-meta">${formatTimeRange(a.data_inizio, a.data_fine)}</div>
      <div class="calendar-meta">${getCommessaLabel(a.commessa_id)}</div>
    `;
    card.addEventListener("click", () => {
      const commessa = state.commesse.find((c) => c.id === a.commessa_id);
      if (commessa) selectCommessa(commessa.id);
    });
    col.appendChild(card);
  });
}

async function refreshTodoRealtimeSnapshot() {
  if (!state.session) return;
  const base =
    todoLastRendered && todoLastRendered.length ? todoLastRendered : (state.commesse || []).slice(0, REPORT_MAX_ITEMS);
  if (!base.length || base.length > REPORT_MAX_ITEMS) return;
  await loadReportActivitiesFor(base);
  await loadTodoOverridesFor(base);
  await loadTodoClientFlowsFor(base);
}

async function setupRealtime() {
  if (state.realtime) return;
  state.realtime = supabase
    .channel("realtime-commesse")
    .on("postgres_changes", { event: "*", schema: "public", table: "commesse" }, loadCommesse)
    .on("postgres_changes", { event: "*", schema: "public", table: "commesse_reparti" }, loadCommesse)
    .on("postgres_changes", { event: "*", schema: "public", table: "attivita" }, async () => {
      await loadAttivita();
      await loadMatrixAttivita();
      await refreshTodoRealtimeSnapshot();
    })
    .on("postgres_changes", { event: "*", schema: "public", table: "commessa_attivita_override" }, async () => {
      await refreshTodoRealtimeSnapshot();
    })
    .on("postgres_changes", { event: "*", schema: "public", table: "commessa_attivita_cliente" }, async () => {
      await refreshTodoRealtimeSnapshot();
    })
    .subscribe();
}

async function saveCommessa(closeOnSuccess = true) {
  if (!state.canWrite) {
    setStatus("You don't have permission to edit this commessa.", "error");
    showToast("You don't have permission to edit this commessa.", "error");
    return;
  }
  if (updateBtn) {
    updateBtn.disabled = true;
    updateBtn.textContent = "Saving...";
  }
  if (applyBtn) {
    applyBtn.disabled = true;
  }
  try {
    if (!state.selected) {
      setStatus("Select a commessa.", "error");
      return;
    }
    const normalized = normalizeCommessaParts(d.anno.value, d.numero.value);
    if (!normalized) {
      setStatus("Invalid year or number.", "error");
      return;
    }
    const dup = findDuplicateCommessa(normalized.anno, normalized.numero, state.selected?.id);
    if (dup) {
      setStatus("Number already exists for this year.", "error");
      return;
    }
    const telaioEffettivo =
      d.data_consegna_telaio_effettiva ? d.data_consegna_telaio_effettiva.value || null : null;
    const telaioConsegnato = Boolean(d.telaio_ordinato?.dataset.ordered === "true");
    if (telaioConsegnato && !telaioEffettivo) {
      if (d.data_consegna_telaio_effettiva) {
        d.data_consegna_telaio_effettiva.classList.add("is-missing-required");
        setTimeout(() => d.data_consegna_telaio_effettiva.classList.remove("is-missing-required"), 1600);
        d.data_consegna_telaio_effettiva.focus();
      }
      const msg = "Add a planned frame order date before marking as ordered.";
      setStatus(msg, "error");
      showToast(msg, "error");
      return;
    }
    const consegnaValue = d.data_consegna.value || null;
    const payload = {
      codice: normalized.codice,
      anno: normalized.anno,
      numero: normalized.numero,
      titolo: d.titolo.value.trim() || null,
      cliente: d.cliente.value.trim() || null,
      stato: d.stato.value,
      priorita: d.priorita.value || null,
      data_ingresso: d.data_ingresso.value || null,
      data_ordine_telaio: d.data_ordine_telaio ? d.data_ordine_telaio.value || null : null,
      data_arrivo_kit_cavi: d.data_arrivo_kit_cavi ? d.data_arrivo_kit_cavi.value || null : null,
      data_prelievo: d.data_prelievo_materiali ? d.data_prelievo_materiali.value || null : null,
      telaio_consegnato: telaioConsegnato,
      data_consegna_telaio_effettiva: telaioEffettivo,
      data_consegna_prevista: consegnaValue,
      data_consegna_macchina: consegnaValue,
      note_generali: d.note.value.trim() || null,
    };
    const role = String(state.profile?.ruolo || "").trim().toLowerCase();
    if (role === "planner") {
      const plannerPayload = {
        data_ordine_telaio: payload.data_ordine_telaio || null,
        data_consegna_macchina: payload.data_consegna_macchina || null,
        data_arrivo_kit_cavi: payload.data_arrivo_kit_cavi || null,
        data_prelievo: payload.data_prelievo || null,
      };
      const error = await updateCommessaPlannerDates(state.selected.id, plannerPayload);
      if (error) {
        setStatus(`Update error: ${error.message}`, "error");
        showToast(`Update error: ${error.message}`, "error");
        return;
      }
      updateCommessaInState(state.selected.id, plannerPayload);
    } else {
      const { error } = await supabase.from("commesse").update(payload).eq("id", state.selected.id);
      if (error) {
        setStatus(`Update error: ${error.message}`, "error");
        return;
      }
      updateCommessaInState(state.selected.id, payload);
    }
    applyFilters();
    renderMatrixReport();
    setDetailSnapshot();
    setStatus("Commessa updated.", "ok");
    await loadCommesse();
    if (closeOnSuccess) {
      closeCommessaDetailModal();
    }
  } catch (err) {
    setStatus(`Unexpected error: ${err?.message || err}`, "error");
    console.error("handleUpdate error", err);
  } finally {
    if (updateBtn) {
      updateBtn.disabled = false;
      updateBtn.textContent = "Save and close";
    }
    if (applyBtn) {
      applyBtn.disabled = false;
    }
  }
}

async function handleUpdate(e) {
  e.preventDefault();
  await saveCommessa(true);
}

async function handleCreate(e) {
  e.preventDefault();
  if (!n.titolo.value.trim()) {
    setStatus("Description is required.", "error");
    return;
  }
  const normalized = normalizeCommessaParts(n.anno.value, n.numero.value);
  if (!normalized) {
    setStatus("Invalid year or number.", "error");
    return;
  }
  const dup = findDuplicateCommessa(normalized.anno, normalized.numero);
  if (dup) {
    setStatus("Number already exists for this year.", "error");
    return;
  }
  const telaioEffettivo =
    n.data_consegna_telaio_effettiva ? n.data_consegna_telaio_effettiva.value || null : null;
  const telaioConsegnato = Boolean(n.telaio_consegnato?.checked);
  if (telaioConsegnato && !telaioEffettivo) {
    setStatus("Set a planned frame order date before marking as ordered.", "error");
    return;
  }
  const consegnaValue = n.data_consegna.value || null;
  const payload = {
    codice: normalized.codice,
    anno: normalized.anno,
    numero: normalized.numero,
    titolo: n.titolo.value.trim(),
    cliente: n.cliente.value.trim() || null,
    stato: n.stato.value,
    priorita: n.priorita.value || null,
    data_ingresso: n.data_ingresso.value || null,
    data_ordine_telaio: n.data_ordine_telaio ? n.data_ordine_telaio.value || null : null,
    telaio_consegnato: telaioConsegnato,
    data_consegna_telaio_effettiva: telaioEffettivo,
    data_consegna_prevista: consegnaValue,
    data_consegna_macchina: consegnaValue,
    note_generali: n.note.value.trim() || null,
  };
  const { data, error } = await supabase.from("commesse").insert(payload).select().single();
  if (error) {
    setStatus(`Create error: ${error.message}`, "error");
    return;
  }
  newForm.reset();
  setStatus("Commessa created.", "ok");
  closeCommessaCreateModal();
  await loadCommesse();
  selectCommessa(data.id, { openModal: true });
}

async function handleAddReparto(e) {
  e.preventDefault();
  if (!addRepartoSelect) return;
  if (!state.selected) {
    setStatus("Select a commessa.", "error");
    return;
  }
  const repartoId = Number(addRepartoSelect.value);
  const { error } = await supabase.from("commesse_reparti").insert({
    commessa_id: state.selected.id,
    reparto_id: repartoId,
    stato: "da_fare",
  });
  if (error) {
    setStatus(`Department add error: ${error.message}`, "error");
    return;
  }
  setStatus("Department added.", "ok");
  await loadCommesse();
  selectCommessa(state.selected.id);
}

async function handleImport() {
  if (!state.canWrite) return;
  const { rows, errors } = parseImportRows(importTextarea.value);
  if (!rows.length) {
    setStatus("No valid rows to import.", "error");
    return;
  }
  if (errors.length) {
    setStatus(`Import blocked: fix the errors (e.g. ${errors[0]}).`, "error");
    return;
  }

  const { error } = await supabase
    .from("commesse")
    .upsert(rows, { onConflict: "codice" })
    .select("codice");

  if (error) {
    setStatus(`Import error: ${error.message}`, "error");
    importPreview.textContent = `Import error: ${error.message}`;
    console.error("Import error", error);
    return;
  }
  setStatus(`Import completed: ${rows.length} rows.`, "ok");
  closeImportModal();
  importTextarea.value = "";
  await loadCommesse();
}

async function handleAssignConfirm() {
  if (!matrixState.pendingDrop) {
    closeAssignModal();
    return;
  }
  try {
    if (!assignActivities) {
      setStatus("Error: activity list not available.", "error");
      return;
    }
    const selected = Array.from(assignActivities.querySelectorAll(".filter-pill.active")).map((i) => i.textContent.trim());
    const altroNote = assignAltroNote ? assignAltroNote.value.trim() : "";
    const assenzaOreRaw = assignAssenzaOre ? assignAssenzaOre.value : "";
    const assenzaOre = Number(assenzaOreRaw);
    if (selected.some((t) => isAssenteTitle(t)) && (!assenzaOreRaw || Number.isNaN(assenzaOre) || assenzaOre <= 0)) {
      setStatus("Enter absence hours.", "error");
      return;
    }
    if (!selected.length) {
      setStatus("Select at least one activity.", "error");
      return;
    }
    if (assignConfirmBtn) {
      assignConfirmBtn.disabled = true;
      assignConfirmBtn.textContent = "Assegno...";
    }
  const { commessaId, risorsaId, day } = matrixState.pendingDrop;
  if (!canAssignCommessaToRisorsa(risorsaId)) {
    setStatus("Non autorizzato a modificare questa riga.", "error");
    return;
  }
  const startAt = new Date(day);
  startAt.setHours(8, 0, 0, 0);
  const endAt = new Date(day);
  endAt.setHours(17, 0, 0, 0);

  if (matrixState.autoShift) {
    const hasOverlap = matrixState.attivita.some((a) => a.risorsa_id === risorsaId && overlapsDay(a, day));
    if (hasOverlap) {
      const ok = await shiftRisorsa(risorsaId, day, selected.length);
      if (!ok) return;
    }
  }

  const rows = selected.map((title) => ({
    commessa_id: commessaId,
    titolo: title,
    descrizione: title.toLowerCase() === "altro" ? altroNote || null : null,
    ore_assenza: isAssenteTitle(title) ? assenzaOre || null : null,
    risorsa_id: risorsaId,
    data_inizio: startAt.toISOString(),
    data_fine: endAt.toISOString(),
    stato: "pianificata",
  }));

  const { error } = await supabase.from("attivita").insert(rows);
  if (error) {
    setStatus(`Assignment error: ${error.message}`, "error");
    return;
  }
  setStatus("Activities assigned.", "ok");
  closeAssignModal();
  await loadMatrixAttivita();
  await refreshTodoRealtimeSnapshot();
  } catch (err) {
    setStatus(`Assignment error: ${err.message || err}`, "error");
  } finally {
    if (assignConfirmBtn) {
      assignConfirmBtn.disabled = false;
      assignConfirmBtn.textContent = "Assegna";
    }
  }
}

async function init() {
  if (SUPABASE_URL.startsWith("INSERISCI")) {
    setStatus("");
  }

  closeImportModal();
  closeImportDatesModal();

  const { data } = await supabase.auth.getSession();
  state.session = data.session;

  if (state.session) {
    debugLog("init: session found");
    await syncSignedIn(state.session);
    return;
  }
  const storedSession = await ensureSessionFromStorage();
  if (storedSession) {
    debugLog("init: session from storage");
    await syncSignedIn(storedSession);
    return;
  }
  debugLog("init: no session");
  setStatus("");
}

loginBtn.addEventListener("click", signIn);
logoutBtn.addEventListener("click", signOut);
if (resetPasswordBtn) {
  resetPasswordBtn.addEventListener("click", resetPassword);
}
if (setPasswordBtn) {
  setPasswordBtn.addEventListener("click", changePassword);
}
if (resetWaitCloseBtn) {
  resetWaitCloseBtn.addEventListener("click", hideResetWaitOverlay);
}
if (topbarMenu && topbarMenuToggle) {
  topbarMenuToggle.addEventListener("click", () => {
    const isOpen = topbarMenu.classList.toggle("is-open");
    topbarMenuToggle.setAttribute("aria-expanded", String(isOpen));
  });
}

function isAdmin() {
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();
  return role === "admin";
}

async function loadUsers() {
  if (!usersList) return;
  let data = null;
  let error = null;
  try {
    const result = await supabase.from("utenti").select("id,email,nome,ruolo").order("email");
    data = result.data;
    error = result.error;
  } catch (err) {
    error = err;
  }
  if (!data || error) {
    data = await fetchTableViaRest("utenti", "select=id,email,nome,ruolo&order=email.asc");
  }
  if (!data) {
    usersList.innerHTML = `<div class="matrix-empty">Errore caricamento utenti.</div>`;
    return;
  }
  state.utenti = data || [];
  renderResourcesPanel();
  renderUsersPanel(data);
}

function renderUsersPanel(users) {
  if (!usersList) return;
  usersList.innerHTML = "";
  if (!users.length) {
    usersList.innerHTML = `<div class="matrix-empty">Nessun utente.</div>`;
    return;
  }
  const canEdit = isAdmin();
  const currentUserId = state.profile?.id;

  users.forEach((u) => {
    const row = document.createElement("div");
    row.className = "user-row";

    const meta = document.createElement("div");
    meta.className = "user-meta";
    const email = document.createElement("div");
    email.className = "user-email";
    email.textContent = u.email || u.id;
    const name = document.createElement("div");
    name.className = "user-name";
    name.textContent = u.nome || "";
    meta.appendChild(email);
    meta.appendChild(name);

    const role = document.createElement("select");
    ["admin", "responsabile", "planner", "operatore", "viewer"].forEach((r) => {
      const opt = document.createElement("option");
      opt.value = r;
      opt.textContent = r;
      if (String(u.ruolo || "").toLowerCase() === r) opt.selected = true;
      role.appendChild(opt);
    });

    const save = document.createElement("button");
    save.type = "button";
    save.className = "ghost";
    save.textContent = "Salva";
    save.addEventListener("click", async () => {
      if (!isAdmin()) return;
      const { error } = await supabase.from("utenti").update({ ruolo: role.value }).eq("id", u.id);
      if (error) {
        setStatus(`Ruolo non aggiornato: ${error.message}`, "error");
        return;
      }
      setStatus("Ruolo aggiornato.", "ok");
      await loadUsers();
    });

    if (!canEdit || (currentUserId && String(currentUserId) === String(u.id))) {
      role.disabled = true;
      save.disabled = true;
    }

    row.appendChild(meta);
    row.appendChild(role);
    row.appendChild(save);
    usersList.appendChild(row);
  });
}

updateResetCooldownUI();
initTheme();
if (descriptionFilter) descriptionFilter.addEventListener("input", applyFilters);
if (numberFilter) numberFilter.addEventListener("input", applyFilters);
if (yearFilter) yearFilter.addEventListener("change", applyFilters);
if (reportDescFilter) reportDescFilter.addEventListener("input", renderMatrixReport);
if (reportNumberFilter) reportNumberFilter.addEventListener("input", renderMatrixReport);
if (reportYearFilter) reportYearFilter.addEventListener("change", renderMatrixReport);
if (todoDescFilter) todoDescFilter.addEventListener("input", renderMatrixReportFromTodoFilters);
if (todoNumberFilter) todoNumberFilter.addEventListener("input", renderMatrixReportFromTodoFilters);
if (todoYearFilter) todoYearFilter.addEventListener("change", renderMatrixReportFromTodoFilters);
if (todoPriorityField) todoPriorityField.addEventListener("change", renderMatrixReportFromTodoFilters);
if (todoPriorityBtn) {
  renderTodoPriorityButtonLabel();
  todoPriorityBtn.addEventListener("click", () => {
    const mode = todoPriorityBtn.dataset.mode || "off";
    const next = mode === "off" ? "asc" : "off";
    todoPriorityBtn.dataset.mode = next;
    renderTodoPriorityButtonLabel();
    renderMatrixReportFromTodoFilters();
  });
}
if (todoSortBtn) {
  renderTodoSortButtonLabel();
  todoSortBtn.addEventListener("click", () => {
    const mode = todoSortBtn.dataset.mode || "off";
    const next = mode === "off" ? "asc" : mode === "asc" ? "desc" : "off";
    todoSortBtn.dataset.mode = next;
    renderTodoSortButtonLabel();
    renderMatrixReportFromTodoFilters();
  });
}
if (todoCompleteBtn) {
  renderTodoCompleteButtonLabel();
  todoCompleteBtn.addEventListener("click", () => {
    todoCompleteBtn.classList.toggle("active");
    renderTodoCompleteButtonLabel();
    renderMatrixReportFromTodoFilters();
  });
}
todoFilterControls.forEach((control) => {
  control.addEventListener("focus", lockTodoGridHeightForFilters);
  control.addEventListener("blur", scheduleTodoGridHeightUnlock);
});
if (todoGrid && todoHeaderTableHost) {
  todoGrid.addEventListener("scroll", () => {
    if (todoScrollSyncLock) return;
    todoScrollSyncLock = true;
    todoHeaderTableHost.scrollLeft = todoGrid.scrollLeft;
    todoScrollSyncLock = false;
  });
}
if (reportHidePast) {
  reportHidePast.textContent = reportHidePast.classList.contains("active")
    ? "Passato: nascosto"
    : "Passato: visibile";
  reportHidePast.addEventListener("click", () => {
    reportHidePast.classList.toggle("active");
    reportHidePast.textContent = reportHidePast.classList.contains("active")
      ? "Passato: nascosto"
      : "Passato: visibile";
    renderMatrixReport();
  });
}
if (reportOrderDue) {
  reportOrderDue.textContent = reportOrderDue.classList.contains("active")
    ? "Ordina in scadenza"
    : "Ordine standard";
  reportOrderDue.addEventListener("click", () => {
    reportOrderDue.classList.toggle("active");
    reportOrderDue.textContent = reportOrderDue.classList.contains("active")
      ? "Ordina in scadenza"
      : "Ordine standard";
    renderMatrixReport();
  });
}
if (reportViewToggleBtn) {
  reportViewToggleBtn.addEventListener("click", () => {
    state.reportView = state.reportView === "gantt" ? "list" : "gantt";
    reportViewToggleBtn.textContent = state.reportView === "gantt" ? "Vista: Gantt" : "Vista: Elenco";
    renderMatrixReport();
  });
}
if (reportViewToggleBtn) {
  reportViewToggleBtn.textContent = state.reportView === "gantt" ? "Vista: Gantt" : "Vista: Elenco";
}
detailForm.addEventListener("submit", handleUpdate);
detailForm.addEventListener("input", updateDetailDirty);
detailForm.addEventListener("change", updateDetailDirty);
if (d.telaio_ordinato) {
  d.telaio_ordinato.addEventListener("click", () => {
    const isOrdered = d.telaio_ordinato.dataset.ordered === "true";
    const plannedDate = d.data_consegna_telaio_effettiva ? d.data_consegna_telaio_effettiva.value || "" : "";
    if (!isOrdered && !plannedDate) {
      if (d.data_consegna_telaio_effettiva) {
        d.data_consegna_telaio_effettiva.classList.add("is-missing-required");
        setTimeout(() => d.data_consegna_telaio_effettiva.classList.remove("is-missing-required"), 1600);
        d.data_consegna_telaio_effettiva.focus();
      }
      const msg = "Add a planned frame order date before marking as ordered.";
      setStatus(msg, "error");
      showToast(msg, "error");
      return;
    }
    setTelaioOrdinatoButton(
      d.telaio_ordinato,
      !isOrdered,
      true,
      d.data_consegna_telaio_effettiva || null
    );
    updateDetailDirty();
  });
}
newForm.addEventListener("submit", handleCreate);
if (n.anno && n.numero) {
  const handler = () => {
    updateNumeroWarning(n.anno, n.numero, newNumeroWarning);
    updateCodicePreview(n.anno, n.numero, newCodicePreview);
  };
  n.anno.addEventListener("input", handler);
  n.numero.addEventListener("input", handler);
}
if (d.anno && d.numero) {
  const handler = () => updateNumeroWarning(d.anno, d.numero, detailNumeroWarning, state.selected?.id);
  d.anno.addEventListener("input", handler);
  d.numero.addEventListener("input", handler);
}
if (openNewCommessaBtn) {
  openNewCommessaBtn.addEventListener("click", openCommessaCreateModal);
}
if (closeCreateBtn) {
  closeCreateBtn.addEventListener("click", closeCommessaCreateModal);
}
if (commessaCreateModal) {
  commessaCreateModal.addEventListener("click", (e) => {
    if (e.target === commessaCreateModal) closeCommessaCreateModal();
  });
}
if (closeDetailBtn) {
  closeDetailBtn.addEventListener("click", requestCloseCommessaDetailModal);
}
if (detailDeleteBtn) {
  detailDeleteBtn.addEventListener("click", openDeleteCommessaModal);
}
if (commessaDetailModal) {
  commessaDetailModal.addEventListener("click", (e) => {
    if (e.target === commessaDetailModal) requestCloseCommessaDetailModal();
  });
}
if (commesseToggleBtn && commessePanel) {
  const toggle = () => {
    commessePanel.classList.toggle("collapsed");
    const isCollapsed = commessePanel.classList.contains("collapsed");
    commesseToggleBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  };
  commesseToggleBtn.addEventListener("click", toggle);
  if (commesseTitle) commesseTitle.addEventListener("click", toggle);
}

if (matrixPanel && matrixTitle) {
  matrixTitle.addEventListener("click", () => {
    matrixPanel.classList.toggle("collapsed");
  });
}

if (calendarPanel && calendarTitle) {
  calendarTitle.addEventListener("click", () => {
    calendarPanel.classList.toggle("collapsed");
  });
}

if (reportPanel && reportTitle) {
  reportTitle.addEventListener("click", () => {
    reportPanel.classList.toggle("collapsed");
    if (!reportPanel.classList.contains("collapsed")) {
      renderMatrixReport();
    }
  });
}
if (todoSection && todoTitle) {
  todoTitle.addEventListener("click", () => {
    todoSection.classList.toggle("collapsed");
    if (!todoSection.classList.contains("collapsed")) {
      renderMatrixReport();
    }
  });
}
if (addRepartoForm) {
  addRepartoForm.addEventListener("submit", handleAddReparto);
}
if (openResourcesBtn) {
  openResourcesBtn.addEventListener("click", openResourcesModal);
}
if (closeResourcesBtn) {
  closeResourcesBtn.addEventListener("click", closeResourcesModal);
}
if (resourcesModal) {
  resourcesModal.addEventListener("click", (e) => {
    if (e.target === resourcesModal) closeResourcesModal();
  });
}
if (openProgressBtn) {
  openProgressBtn.addEventListener("click", openProgressModal);
}
if (openBackupBtn) {
  openBackupBtn.addEventListener("click", openBackupModal);
}
if (themeToggleBtn) {
  themeToggleBtn.addEventListener("click", toggleTheme);
}
if (reportWorkloadBtn) {
  reportWorkloadBtn.addEventListener("click", openWorkloadModal);
}
if (closeProgressBtn) {
  closeProgressBtn.addEventListener("click", closeProgressModal);
}
if (closeBackupBtn) {
  closeBackupBtn.addEventListener("click", closeBackupModal);
}
if (backupExportBtn) {
  backupExportBtn.addEventListener("click", handleBackupExport);
}
if (backupImportInput) {
  backupImportInput.addEventListener("change", handleBackupFileChange);
}
if (backupImportConfirm) {
  backupImportConfirm.addEventListener("input", updateBackupImportState);
}
if (backupImportBtn) {
  backupImportBtn.addEventListener("click", handleBackupImport);
}
if (progressModal) {
  progressModal.addEventListener("click", (e) => {
    if (e.target === progressModal) closeProgressModal();
  });
}
if (workloadCloseBtn) {
  workloadCloseBtn.addEventListener("click", closeWorkloadModal);
}
if (workloadRunBtn) {
  workloadRunBtn.addEventListener("click", calculateWorkload);
}
if (workloadWeekStart) {
  workloadWeekStart.addEventListener("change", calculateWorkload);
}
if (workloadWeekEnd) {
  workloadWeekEnd.addEventListener("change", calculateWorkload);
}
if (workloadModal) {
  workloadModal.addEventListener("click", (e) => {
    if (e.target === workloadModal) closeWorkloadModal();
  });
}
if (backupModal) {
  backupModal.addEventListener("click", (e) => {
    if (e.target === backupModal) closeBackupModal();
  });
}
if (progressYearCurrent) {
  progressYearCurrent.addEventListener("click", () => {
    progressYearCurrent.classList.add("active");
    if (progressYearPrev) progressYearPrev.classList.remove("active");
    progressSelectedId = null;
    if (progressMeta) progressMeta.textContent = "Select a commessa.";
    if (progressList) progressList.innerHTML = "";
    renderProgressResults();
  });
}
if (progressYearPrev) {
  progressYearPrev.addEventListener("click", () => {
    progressYearPrev.classList.add("active");
    if (progressYearCurrent) progressYearCurrent.classList.remove("active");
    progressSelectedId = null;
    if (progressMeta) progressMeta.textContent = "Select a commessa.";
    if (progressList) progressList.innerHTML = "";
    renderProgressResults();
  });
}
  if (progressSearch) {
    progressSearch.addEventListener("input", () => {
      progressSelectedId = null;
      if (progressMeta) progressMeta.textContent = "Select a commessa.";
      if (progressList) progressList.innerHTML = "";
      if (progressTimelineDays) progressTimelineDays.innerHTML = "";
      renderProgressResults();
    });
  }
if (resourceForm) {
  resourceForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!state.canWrite) return;
    const nome = resourceName.value.trim();
    if (!nome) {
      setStatus("Enter a resource name.", "error");
      return;
    }
    const repartoId = resourceReparto.value;
    if (!repartoId) {
      setStatus("Select a department.", "error");
      return;
    }
    const payload = {
      nome,
      reparto_id: repartoId,
      attiva: resourceActive.checked,
    };
    const { error } = await supabase.from("risorse").insert(payload);
    if (error) {
      setStatus(`Resource creation error: ${error.message}`, "error");
      return;
    }
    resourceName.value = "";
    resourceReparto.value = "";
    resourceActive.checked = true;
    setStatus("Resource added.", "ok");
    await loadRisorse();
  });
}
if (applyBtn) {
  applyBtn.addEventListener("click", () => saveCommessa(false));
}
openImportBtn.addEventListener("click", openImportModal);
if (openImportDatesBtn) {
  openImportDatesBtn.addEventListener("click", openImportDatesModal);
}
if (openImportProductionBtn) {
  openImportProductionBtn.addEventListener("click", openImportProductionModal);
}
closeImportBtn.addEventListener("click", closeImportModal);
if (closeImportDatesBtn) {
  closeImportDatesBtn.addEventListener("click", closeImportDatesModal);
}
if (closeImportProductionBtn) {
  closeImportProductionBtn.addEventListener("click", closeImportProductionModal);
}
importTextarea.addEventListener("input", updateImportPreview);
if (importDatesTextarea) importDatesTextarea.addEventListener("input", updateImportDatesPreview);
if (importProductionTextarea) importProductionTextarea.addEventListener("input", updateImportProductionPreview);
importCommitBtn.addEventListener("click", handleImport);
if (importDatesCommitBtn) importDatesCommitBtn.addEventListener("click", handleImportDates);
if (importProductionCommitBtn) importProductionCommitBtn.addEventListener("click", handleImportProduction);
importModal.addEventListener("click", (e) => {
  if (e.target === importModal) closeImportModal();
});
if (importDatesModal) {
  importDatesModal.addEventListener("click", (e) => {
    if (e.target === importDatesModal) closeImportDatesModal();
  });
}
if (importProductionModal) {
  importProductionModal.addEventListener("click", (e) => {
    if (e.target === importProductionModal) closeImportProductionModal();
  });
}
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (commessaDetailModal && !commessaDetailModal.classList.contains("hidden")) {
      requestCloseCommessaDetailModal();
      return;
    }
    closeImportModal();
    closeImportDatesModal();
    closeAssignModal();
    closeMatrixQuickMenu();
    closeCommessaCreateModal();
    closeConfirmModal(false);
    closeDeleteCommessaModal();
    closeWorkloadModal();
    closeBackupModal();
    closeReportDeptMenu();
    closeMilestonePicker();
  }
});
calendarView.addEventListener("change", async (e) => {
  calendarState.view = e.target.value;
  await loadAttivita();
});
calendarDate.addEventListener("change", async (e) => {
  calendarState.date = new Date(e.target.value);
  await loadAttivita();
});
calPrevBtn.addEventListener("click", async () => {
  calendarState.date =
    calendarState.view === "day" ? addDays(calendarState.date, -1) : addDays(calendarState.date, -7);
  calendarDate.value = formatDateInput(calendarState.date);
  await loadAttivita();
});
calNextBtn.addEventListener("click", async () => {
  calendarState.date =
    calendarState.view === "day" ? addDays(calendarState.date, 1) : addDays(calendarState.date, 7);
  calendarDate.value = formatDateInput(calendarState.date);
  await loadAttivita();
});
matrixDate.addEventListener("change", async (e) => {
  matrixState.date = new Date(e.target.value);
  await loadMatrixAttivita();
});
matrixPrevBtn.addEventListener("click", async () => {
  matrixState.date = addDays(matrixState.date, -7);
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixNextBtn.addEventListener("click", async () => {
  matrixState.date = addDays(matrixState.date, 7);
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixTodayBtn.addEventListener("click", async () => {
  matrixState.date = new Date();
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixZoomInBtn.addEventListener("click", async () => {
  if (matrixState.view === "six") {
    matrixState.view = "three";
  } else if (matrixState.view === "three") {
    matrixState.view = "two";
  } else if (matrixState.view === "two") {
    matrixState.view = "week";
  } else {
    matrixState.view = "week";
  }
  matrixState.date = startOfWeek(matrixState.date);
  setMatrixViewLabel();
  await loadMatrixAttivita();
});
matrixZoomOutBtn.addEventListener("click", async () => {
  if (matrixState.view === "week") {
    matrixState.view = "two";
  } else if (matrixState.view === "two") {
    matrixState.view = "three";
  } else if (matrixState.view === "three") {
    matrixState.view = "six";
  }
  matrixState.date = startOfWeek(matrixState.date);
  setMatrixViewLabel();
  await loadMatrixAttivita();
});
if (matrixColorByActivityBtn) {
  matrixColorByActivityBtn.addEventListener("click", () => {
    setMatrixColorMode(matrixState.colorMode === "activity" ? "none" : "activity");
  });
}
if (matrixColorByCommessaBtn) {
  matrixColorByCommessaBtn.addEventListener("click", () => {
    setMatrixColorMode(matrixState.colorMode === "commessa" ? "none" : "commessa");
  });
}
if (matrixAssenzaItem) {
  matrixAssenzaItem.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("application/x-assenza", "1");
  });
}
if (matrixAltroItem) {
  matrixAltroItem.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("application/x-altro", "1");
  });
}
if (matrixPedItem) {
  matrixPedItem.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("application/x-ped", "1");
  });
}
if (matrixSegnaturaItem) {
  matrixSegnaturaItem.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("application/x-segnatura", "1");
  });
}
if (matrixOpOrdiniItem) {
  matrixOpOrdiniItem.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("application/x-op-ordini", "1");
  });
}
if (commessaFocusOverlay) {
  commessaFocusOverlay.addEventListener("click", () => closeCommessaQuickMenu());
}
if (matrixShiftToggle) {
  matrixShiftToggle.addEventListener("change", (e) => {
    setMatrixAutoShift(e.target.checked);
  });
}
if (matrixReportSortBtn) {
  matrixReportSortBtn.addEventListener("click", () => {
    matrixState.reportSort = matrixState.reportSort === "asc" ? "default" : "asc";
    matrixReportSortBtn.textContent = matrixState.reportSort === "asc" ? "Ordine: Crescente" : "Ordine: Originale";
    renderMatrixReport();
  });
}
if (matrixTrash) {
  matrixTrash.addEventListener("dragover", (e) => {
    if (!state.canWrite) return;
    const dragId = matrixState.draggingId || e.dataTransfer.getData("text/plain");
    const attivita = matrixState.attivita.find((x) => String(x.id) === String(dragId));
    if (!canDeleteMatrixActivity(attivita)) return;
    e.preventDefault();
    matrixTrash.classList.add("is-dragover");
  });
  matrixTrash.addEventListener("dragleave", () => {
    matrixTrash.classList.remove("is-dragover");
  });
  matrixTrash.addEventListener("drop", async (e) => {
    if (!state.canWrite) {
      setStatus("Non autorizzato a eliminare attivita.", "error");
      return;
    }
    const dragId = matrixState.draggingId || e.dataTransfer.getData("text/plain");
    const attivita = matrixState.attivita.find((x) => String(x.id) === String(dragId));
    if (!canDeleteMatrixActivity(attivita)) {
      setStatus("Non autorizzato a eliminare attivita.", "error");
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    matrixTrash.classList.remove("is-dragover");
    if (!dragId) return;
    await deleteActivityById(dragId);
  });
}
if (matrixCommessaPickerToggleBtn) {
  matrixCommessaPickerToggleBtn.addEventListener("click", () => openMatrixFilter("commessa"));
}
if (matrixAttivitaToggleBtn) {
  matrixAttivitaToggleBtn.addEventListener("click", () => openMatrixFilter("attivita"));
}
matrixToggleListBtn.addEventListener("click", () => {
  if (!matrixCommessePanel) return;
  matrixCommessePanel.classList.toggle("collapsed");
  matrixToggleListBtn.textContent = matrixCommessePanel.classList.contains("collapsed")
    ? "Espandi lista"
    : "Comprimi lista";
});
matrixCommessaSearch.addEventListener("input", renderMatrixCommesseColumn);
assignConfirmBtn.addEventListener("click", handleAssignConfirm);
closeAssignBtn.addEventListener("click", closeAssignModal);
assignModal.addEventListener("click", (e) => {
  if (e.target === assignModal) closeAssignModal();
});
if (matrixFilterSearch) {
  matrixFilterSearch.addEventListener("input", renderMatrixFilterList);
}
if (matrixFilterYear) {
  matrixFilterYear.addEventListener("change", renderMatrixFilterList);
}
if (matrixFilterClearBtn) {
  matrixFilterClearBtn.addEventListener("click", () => {
    matrixState.filterDraft.clear();
    renderMatrixFilterList();
  });
}
if (matrixFilterResetBtn) {
  matrixFilterResetBtn.addEventListener("click", () => {
    matrixState.filterDraft.clear();
    if (matrixFilterSearch) matrixFilterSearch.value = "";
    if (matrixFilterYear) matrixFilterYear.value = "";
    renderMatrixFilterList();
  });
}
if (matrixFilterApplyBtn) {
  matrixFilterApplyBtn.addEventListener("click", applyMatrixFilter);
}
if (matrixFilterCloseBtn) {
  matrixFilterCloseBtn.addEventListener("click", closeMatrixFilter);
}
if (matrixFilterModal) {
  matrixFilterModal.addEventListener("click", (e) => {
    if (e.target === matrixFilterModal) closeMatrixFilter();
  });
}
if (matrixQuickToggleBtn) {
  matrixQuickToggleBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleQuickMenuDone();
  });
}
document.addEventListener(
  "scroll",
  () => {
    closeMatrixQuickMenu();
    closeReportDeptMenu();
  },
  true
);
if (activityCloseBtn) {
  activityCloseBtn.addEventListener("click", closeActivityModal);
}
if (activityGanttBtn) {
  activityGanttBtn.addEventListener("click", () => {
    const attivita = matrixState.editingAttivita;
    if (!attivita) return;
    closeActivityModal();
    jumpToReportActivity(attivita);
  });
}
if (activitySaveBtn) {
  activitySaveBtn.addEventListener("click", saveActivityDuration);
}
if (activityDeleteBtn) {
  activityDeleteBtn.addEventListener("click", deleteActivity);
}
if (activityModal) {
  activityModal.addEventListener("click", (e) => {
    if (e.target === activityModal) closeActivityModal();
  });
}
if (confirmCancelBtn) {
  confirmCancelBtn.addEventListener("click", () => closeConfirmModal(false));
}
if (confirmOkBtn) {
  confirmOkBtn.addEventListener("click", () => closeConfirmModal(true));
}
if (confirmModal) {
  confirmModal.addEventListener("click", (e) => {
    if (e.target === confirmModal) closeConfirmModal(false);
  });
}
if (deleteCommessaInput) {
  deleteCommessaInput.addEventListener("input", updateDeleteCommessaConfirmState);
}
if (deleteCommessaCancelBtn) {
  deleteCommessaCancelBtn.addEventListener("click", closeDeleteCommessaModal);
}
if (deleteCommessaConfirmBtn) {
  deleteCommessaConfirmBtn.addEventListener("click", handleDeleteCommessaConfirm);
}
if (deleteCommessaModal) {
  deleteCommessaModal.addEventListener("click", (e) => {
    if (e.target === deleteCommessaModal) closeDeleteCommessaModal();
  });
}
if (assignAssenzaFullBtn && assignAssenzaOre) {
  assignAssenzaFullBtn.addEventListener("click", () => {
    assignAssenzaOre.value = "8";
  });
}
if (assignAssenzaHalfBtn && assignAssenzaOre) {
  assignAssenzaHalfBtn.addEventListener("click", () => {
    assignAssenzaOre.value = "4";
  });
}
if (activityAssenzaFullBtn && activityAssenzaOre) {
  activityAssenzaFullBtn.addEventListener("click", () => {
    activityAssenzaOre.value = "8";
  });
}
if (activityAssenzaHalfBtn && activityAssenzaOre) {
  activityAssenzaHalfBtn.addEventListener("click", () => {
    activityAssenzaOre.value = "4";
  });
}
if (milestonePrevBtn) {
  milestonePrevBtn.addEventListener("click", () => {
    const m = milestonePickerState.month;
    const next = new Date(m.getFullYear(), m.getMonth() - 1, 1);
    milestonePickerState.month = next;
    renderMilestonePicker();
  });
}
if (milestoneNextBtn) {
  milestoneNextBtn.addEventListener("click", () => {
    const m = milestonePickerState.month;
    const next = new Date(m.getFullYear(), m.getMonth() + 1, 1);
    milestonePickerState.month = next;
    renderMilestonePicker();
  });
}
if (milestoneClearBtn) {
  milestoneClearBtn.addEventListener("click", async () => {
    if (milestonePickerState.saving) return;
    if (!state.canWrite) {
      setStatus("You don't have permission to edit this commessa.", "error");
      showToast("You don't have permission to edit this commessa.", "error");
      return;
    }
    const commessaId = milestonePickerState.commessaId;
    const field = milestonePickerState.field;
    if (!commessaId || !field) return;
    milestonePickerState.saving = true;
    const payload = {};
    if (field === "ordine") payload.data_ordine_telaio = null;
    if (field === "kit_cavi") payload.data_arrivo_kit_cavi = null;
    if (field === "prelievo") payload.data_prelievo = null;
    if (field === "consegna") payload.data_consegna_macchina = null;
    if (field === "ordine_pianificato") payload.data_consegna_telaio_effettiva = null;
    const role = String(state.profile?.ruolo || "").trim().toLowerCase();
    let error = null;
      if (role === "planner" && (field === "ordine" || field === "consegna" || field === "kit_cavi")) {
        error = await updateCommessaPlannerDates(commessaId, payload);
      } else {
      const res = await supabase.from("commesse").update(payload).eq("id", commessaId);
      error = res.error || null;
      if (!error) updateCommessaInState(commessaId, payload);
    }
    milestonePickerState.saving = false;
    if (error) {
      setStatus(`Update error: ${error.message}`, "error");
      showToast(`Update error: ${error.message}`, "error");
      return;
    }
    setStatus("Date removed.", "ok");
    closeMilestonePicker();
    await loadCommesse();
  });
}
document.addEventListener("mousemove", (e) => {
  if (!matrixState.resizing) return;
  const elements = document.elementsFromPoint(e.clientX, e.clientY);
  const cell = elements.find((el) => el.classList && el.classList.contains("matrix-cell"));
  if (!cell || !cell.dataset.day) return;
  const dayKey = cell.dataset.day;
  if (dayKey < matrixState.resizing.startDayKey) return;
  matrixState.resizing.targetDayKey = dayKey;
  if (matrixState.resizing.previewBar && matrixState.resizing.layerRect && matrixState.resizing.startCellRect) {
    const endRect = cell.getBoundingClientRect();
    const left = matrixState.resizing.startCellRect.left - matrixState.resizing.layerRect.left;
    const maxRight = matrixState.resizing.layerRect.right - matrixState.resizing.layerRect.left;
    const mouseX = e.clientX - matrixState.resizing.layerRect.left;
    const clampedX = Math.max(left, Math.min(mouseX, maxRight));
    const width = Math.max(0, clampedX - left);
    matrixState.resizing.previewBar.style.left = `${left}px`;
    matrixState.resizing.previewBar.style.width = `${width}px`;
  }
});
document.addEventListener("mouseup", () => {
  if (matrixState.resizing) {
    matrixState.suppressClickUntil = Date.now() + 500;
    commitResize();
  }
  cancelCommessaHighlightTimer();
});
document.addEventListener("dragend", () => {
  if (matrixGrid) matrixGrid.classList.remove("matrix-dragging");
  matrixState.draggingId = null;
  document.querySelectorAll(".matrix-activity-bar.dragging").forEach((el) => el.classList.remove("dragging"));
  closeMatrixQuickMenu();
  if (matrixTrash) matrixTrash.classList.remove("is-dragover");
});
document.addEventListener("drop", () => {
  if (matrixGrid) matrixGrid.classList.remove("matrix-dragging");
  matrixState.draggingId = null;
  document.querySelectorAll(".matrix-activity-bar.dragging").forEach((el) => el.classList.remove("dragging"));
  closeMatrixQuickMenu();
  if (matrixTrash) matrixTrash.classList.remove("is-dragover");
});
sectionLinks.forEach((btn) => {
  btn.addEventListener("click", () => {
    const targetId = btn.dataset.target;
    if (!targetId) return;
    const target = document.getElementById(targetId);
    if (!target) return;
    const rect = target.getBoundingClientRect();
    const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
    const top = rect.top + window.scrollY - toolbarOffset - getSectionToolbarGap();
    window.scrollTo({ top: Math.max(0, top), left: 0, behavior: "smooth" });
  });
});
const updateSectionToolbarOffset = () => {
  const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
  document.documentElement.style.setProperty("--section-toolbar-h", `${toolbarOffset}px`);
  updateTodoPanelHeaderOffset();
};
updateSectionToolbarOffset();
window.addEventListener("resize", updateSectionToolbarOffset);
if (sectionToolbar && typeof ResizeObserver !== "undefined") {
  const toolbarObserver = new ResizeObserver(() => updateSectionToolbarOffset());
  toolbarObserver.observe(sectionToolbar);
}
if (todoStatusMenu) {
  todoStatusMenu.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    if (todoStatusMenu.classList.contains("hidden")) return;
    if (!isTodoStatusMenuBorderHit(event)) return;
    const rect = todoStatusMenu.getBoundingClientRect();
    todoStatusMenuDrag = {
      pointerId: event.pointerId,
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top,
      moved: false,
    };
    todoStatusMenuManualPosition = true;
    todoStatusMenu.style.cursor = "move";
    if (typeof todoStatusMenu.setPointerCapture === "function") {
      todoStatusMenu.setPointerCapture(event.pointerId);
    }
    event.preventDefault();
    event.stopPropagation();
  });
  todoStatusMenu.addEventListener("pointermove", (event) => {
    if (!todoStatusMenuDrag) {
      todoStatusMenu.style.cursor = isTodoStatusMenuBorderHit(event) ? "move" : "";
      return;
    }
    if (event.pointerId !== todoStatusMenuDrag.pointerId) return;
    const padding = 10;
    const rect = todoStatusMenu.getBoundingClientRect();
    let left = event.clientX - todoStatusMenuDrag.offsetX;
    let top = event.clientY - todoStatusMenuDrag.offsetY;
    left = Math.max(padding, Math.min(left, window.innerWidth - rect.width - padding));
    top = Math.max(padding, Math.min(top, window.innerHeight - rect.height - padding));
    todoStatusMenu.style.left = `${left}px`;
    todoStatusMenu.style.top = `${top}px`;
    todoStatusMenuDrag.moved = true;
    event.preventDefault();
    event.stopPropagation();
  });
  const stopTodoStatusMenuDrag = (event) => {
    if (!todoStatusMenuDrag) return;
    if (event.pointerId !== todoStatusMenuDrag.pointerId) return;
    const wasMoved = todoStatusMenuDrag.moved;
    todoStatusMenuDrag = null;
    if (typeof todoStatusMenu.releasePointerCapture === "function") {
      try {
        todoStatusMenu.releasePointerCapture(event.pointerId);
      } catch (_err) {}
    }
    todoStatusMenu.style.cursor = "";
    if (wasMoved) {
      todoStatusMenuDragSuppressClickUntil = Date.now() + 200;
    }
    event.stopPropagation();
  };
  todoStatusMenu.addEventListener("pointerup", stopTodoStatusMenuDrag);
  todoStatusMenu.addEventListener("pointercancel", stopTodoStatusMenuDrag);
  todoStatusMenu.addEventListener("pointerleave", () => {
    if (!todoStatusMenuDrag) todoStatusMenu.style.cursor = "";
  });
  todoStatusMenu.addEventListener("click", async (event) => {
    event.stopPropagation();
    if (Date.now() < todoStatusMenuDragSuppressClickUntil) {
      event.preventDefault();
      return;
    }
    const target = event.target.closest("[data-action]");
    if (!target || !todoMenuTarget) return;
    if (target.tagName === "A") event.preventDefault();
    const action = target.dataset.action;
    if (!action) return;
    if (action === "open_commessa_detail") {
      const commessaId = todoMenuTarget.commessaId;
      if (!commessaId) return;
      selectCommessa(commessaId, { openModal: true });
      closeTodoStatusMenu();
      return;
    }
    if (action.startsWith("todo_date_")) {
      const shouldClose = await handleTodoStatusDatePickerAction(action, target);
      if (shouldClose) closeTodoStatusMenu();
      return;
    }
    if (target.dataset.blocked === "1") {
      setStatus(target.dataset.blockReason || "Azione non disponibile.", "error");
      return;
    }
    if (action === "client_sent" || action === "client_confirmed" || action === "client_silence") {
      openTodoStatusDatePicker(action);
      return;
    }
    const updated = await applyTodoStatusAction(action, todoMenuTarget);
    if (updated) {
      if (action === "done") {
        const prevManual = todoStatusMenuManualPosition;
        const prevLeft = todoStatusMenu?.style.left || "";
        const prevTop = todoStatusMenu?.style.top || "";
        const targetCell = findTodoCellByTarget(
          todoMenuTarget.commessaId,
          todoMenuTarget.titolo,
          todoMenuTarget.reparto || ""
        );
        if (targetCell) {
          openTodoStatusMenu(targetCell, 0, 0);
          if (prevManual && todoStatusMenu) {
            todoStatusMenuManualPosition = true;
            todoStatusMenu.style.left = prevLeft;
            todoStatusMenu.style.top = prevTop;
            positionTodoStatusMenuWithinViewport();
          }
        } else {
          closeTodoStatusMenu();
        }
        return;
      }
      closeTodoStatusMenu();
    }
  });
}
if (todoGlobalStatusMenu) {
  todoGlobalStatusMenu.addEventListener("click", async (event) => {
    event.stopPropagation();
    const target = event.target.closest("button");
    if (!target || !todoGlobalMenuTarget) return;
    const statusLabel = target.dataset.status;
    if (!statusLabel) return;
    const commessa =
      (state.commesse || []).find((c) => String(c.id) === String(todoGlobalMenuTarget.commessaId)) ||
      (todoLastRendered || []).find((c) => String(c.id) === String(todoGlobalMenuTarget.commessaId));
    if (!commessa) return;
    const updated = await updateTodoGlobalStatus(commessa, statusLabel);
    if (updated) closeTodoGlobalStatusMenu();
  });
}
if (todoFocusOverlay) {
  todoFocusOverlay.addEventListener("click", () => {
    closeTodoStatusMenu();
    closeTodoGlobalStatusMenu();
  });
}
window.addEventListener(
  "scroll",
  () => {
    repositionTodoStatusMenuToAnchor();
    repositionTodoGlobalStatusMenuToAnchor();
  },
  { passive: true, capture: true }
);
window.addEventListener("resize", () => {
  updateTodoPanelHeaderOffset();
  repositionTodoStatusMenuToAnchor();
  repositionTodoGlobalStatusMenuToAnchor();
});
window.addEventListener("focus", refreshTodoVisualizationNow);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshTodoVisualizationNow();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeTodoStatusMenu();
    closeTodoGlobalStatusMenu();
  }
});
document.addEventListener("click", (e) => {
  const skipTodoDismiss = todoLongPressSuppress;
  if (todoLongPressSuppress) todoLongPressSuppress = false;
  const skipTodoGlobalDismiss = todoGlobalMenuSuppress;
  if (todoGlobalMenuSuppress) todoGlobalMenuSuppress = false;
  if (matrixQuickMenu && !matrixQuickMenu.classList.contains("hidden")) {
    if (!e.target.closest("#matrixQuickMenu")) {
      closeMatrixQuickMenu();
    }
  }
  if (reportDeptMenu && !reportDeptMenu.classList.contains("hidden")) {
    if (!e.target.closest("#reportDeptMenu")) {
      closeReportDeptMenu();
    }
  }
  if (commessaQuickMenu && !commessaQuickMenu.classList.contains("hidden")) {
    if (!e.target.closest("#commessaQuickMenu")) {
      closeCommessaQuickMenu();
    }
  }
  if (milestonePicker && !milestonePicker.classList.contains("hidden")) {
    if (!e.target.closest("#milestonePicker")) {
      closeMilestonePicker();
    }
  }
  if (todoStatusMenu && !todoStatusMenu.classList.contains("hidden") && !skipTodoDismiss) {
    if (!e.target.closest("#todoStatusMenu")) {
      closeTodoStatusMenu();
    }
  }
  if (todoGlobalStatusMenu && !todoGlobalStatusMenu.classList.contains("hidden") && !skipTodoGlobalDismiss) {
    if (!e.target.closest("#todoGlobalStatusMenu")) {
      closeTodoGlobalStatusMenu();
    }
  }
  if (commessaLongPressSuppress) {
    commessaLongPressSuppress = false;
    return;
  }
  if (!matrixGrid || !commessaHighlightId) return;
  const bar = e.target.closest(".matrix-activity-bar");
  if (!bar) {
    clearCommessaHighlight();
  }
});

setMatrixColorMode("commessa");
setMatrixAutoShift(false);

supabase.auth.onAuthStateChange(async (event, session) => {
  debugLog(`auth event: ${event} session=${session ? "yes" : "no"}`);
  if (session) {
    await syncSignedIn(session);
  } else {
    syncSignedOut();
  }
});

init();
updateRevisionStamp();

})();

const getSectionToolbarGap = () => {
  const raw = getComputedStyle(document.documentElement).getPropertyValue("--section-toolbar-gap").trim();
  const value = Number.parseFloat(raw);
  return Number.isFinite(value) ? value : 0;
};


