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
  reportGanttRowHeight: null,
  machineTypes: [],
};

const el = (id) => document.getElementById(id);

const authStatus = el("authStatus");
const roleBadge = el("roleBadge");
const setPasswordBtn = el("setPasswordBtn");
const statusMsg = el("statusMsg");
const alarmsSection = el("alarmsSection");
const alarmsTitle = el("alarmsTitle");
const alarmsList = el("alarmsList");
const openNotificationsBtn = el("openNotificationsBtn");
const notificationsModal = el("notificationsModal");
const closeNotificationsBtn = el("closeNotificationsBtn");
const refreshNotificationsBtn = el("refreshNotificationsBtn");
const notificationsOnlyMineBtn = el("notificationsOnlyMineBtn");
const notificationsSummary = el("notificationsSummary");
const notificationsClassFilters = el("notificationsClassFilters");
const notificationsList = el("notificationsList");
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
const reportsLink = el("reportsLink");

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
    tipo_macchina: d.tipo_macchina?.value || "",
    variante_macchina: d.variante_macchina?.value || "",
    stato: d.stato?.value || "",
    priorita: d.priorita?.value || "",
    data_ingresso: d.data_ingresso?.value || "",
    data_ordine_telaio: d.data_ordine_telaio?.value || "",
    data_conferma_consegna_telaio: d.data_conferma_consegna_telaio?.value || "",
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

function isDetailFieldFilled(field) {
  if (!field) return false;
  if (field.tagName === "SELECT") {
    return String(field.value || "").trim() !== "";
  }
  const type = String(field.type || "").toLowerCase();
  if (type === "checkbox" || type === "radio") {
    return Boolean(field.checked);
  }
  return String(field.value || "").trim() !== "";
}

function updateDetailFieldCompletionUI() {
  if (!detailForm) return;
  const fields = detailForm.querySelectorAll("input, select, textarea");
  fields.forEach((field) => {
    const type = String(field.type || "").toLowerCase();
    if (type === "hidden") return;
    if (field.id === "d_note") {
      field.classList.remove("detail-field-empty", "detail-field-filled");
      return;
    }
    const filled = isDetailFieldFilled(field);
    field.classList.toggle("detail-field-empty", !filled);
    field.classList.remove("detail-field-filled");
  });
}

function updateDetailDirty() {
  detailDirty = getDetailSnapshot() !== detailSnapshot;
  updateDetailFieldCompletionUI();
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

function formatDbErrorMessage(errorLike) {
  const message = String(errorLike?.message || errorLike || "").trim();
  if (!message) return "Unknown error";
  const lower = message.toLowerCase();
  const missingMachineCols =
    (lower.includes("tipo_macchina") || lower.includes("variante_macchina")) &&
    (lower.includes("schema cache") || lower.includes("could not find"));
  if (missingMachineCols) {
    return "DB non aggiornato: esegui la migration 20260314153000_add_machine_type_to_commesse.sql e poi `NOTIFY pgrst, 'reload schema';` in Supabase SQL Editor.";
  }
  return message;
}

function setTelaioOrdinatoButton(button, isOrdered, showFeedback = false, plannedInput = null) {
  if (!button) return;
  button.dataset.ordered = isOrdered ? "true" : "false";
  button.classList.toggle("is-ordered", isOrdered);
  button.setAttribute("aria-pressed", isOrdered ? "true" : "false");
  let label = "Telaio da ordinare";
  if (isOrdered) {
    const dateRaw = plannedInput ? plannedInput.value || "" : "";
    const dateObj = parseIsoDateOnly(dateRaw);
    const dateLabel = dateObj ? formatDateDMY(dateObj) : "";
    label = dateLabel ? `Telaio ordinato (${dateLabel})` : "Telaio ordinato";
  }
  button.textContent = label;
  if (plannedInput) {
    plannedInput.disabled = false;
    plannedInput.title = "";
  }
  if (showFeedback) {
    setStatus(isOrdered ? "Frame marked as ordered." : "Frame marked as to-order.", "ok");
  }
}

function findCommessaInState(commessaId) {
  if (!commessaId) return null;
  const key = String(commessaId);
  return (
    (state.commesse || []).find((c) => String(c.id) === key) ||
    (state.filteredCommesse || []).find((c) => String(c.id) === key) ||
    (state.reportFilteredCommesse || []).find((c) => String(c.id) === key) ||
    (state.selected && String(state.selected.id) === key ? state.selected : null) ||
    null
  );
}

function ensureTelaioOrderPayloadConsistency(commessaId, payload, options = {}) {
  if (!payload || typeof payload !== "object") return { ok: true, payload };
  const hasOrderedPatch = Object.prototype.hasOwnProperty.call(payload, "telaio_consegnato");
  const hasDatePatch = Object.prototype.hasOwnProperty.call(payload, "data_consegna_telaio_effettiva");
  if (!hasOrderedPatch && !hasDatePatch) return { ok: true, payload };

  const nextPayload = { ...payload };
  if (hasDatePatch) {
    nextPayload.data_consegna_telaio_effettiva = formatIsoDateOnly(nextPayload.data_consegna_telaio_effettiva);
  }

  const current = findCommessaInState(commessaId);
  const nextOrdered = hasOrderedPatch ? Boolean(nextPayload.telaio_consegnato) : Boolean(current?.telaio_consegnato);
  const nextDate = hasDatePatch
    ? nextPayload.data_consegna_telaio_effettiva
    : formatIsoDateOnly(current?.data_consegna_telaio_effettiva);
  if (!(nextOrdered && !nextDate)) return { ok: true, payload: nextPayload };

  const message = options.message || "Set a frame order date before marking as ordered.";
  if (options.focusDateField && state.selected && String(state.selected.id) === String(commessaId)) {
    if (d.data_consegna_telaio_effettiva) {
      d.data_consegna_telaio_effettiva.classList.add("is-missing-required");
      setTimeout(() => d.data_consegna_telaio_effettiva.classList.remove("is-missing-required"), 1600);
      d.data_consegna_telaio_effettiva.focus();
    }
  }
  if (options.notify !== false) setStatus(message, "error");
  if (options.toast) showToast(message, "error");
  return { ok: false, payload: nextPayload, message };
}

function getTelaioOrderAlarms() {
  return [];
}

function renderTelaioOrderAlarms() {
  if (!alarmsSection || !alarmsList) return;
  const alarms = getTelaioOrderAlarms();
  alarmsList.innerHTML = "";
  alarmsSection.classList.toggle("hidden", alarms.length === 0);
  if (alarmsTitle) {
    alarmsTitle.textContent = alarms.length ? `Allarmi (${alarms.length})` : "Allarmi";
  }
  if (!alarms.length) {
    if (!notificationsItemsCache.length) {
      updateNotificationsButtonLabel(0);
    }
    return;
  }

  const maxItems = 8;
  alarms.slice(0, maxItems).forEach((commessa) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "status-alarm-item";
    item.textContent = `${getCommessaLabel(commessa.id)}: ordinato senza data`;
    item.addEventListener("click", () => selectCommessa(commessa.id, { openModal: true }));
    alarmsList.appendChild(item);
  });
  if (alarms.length > maxItems) {
    const more = document.createElement("div");
    more.className = "status-alarm-more";
    more.textContent = `+${alarms.length - maxItems} altri`;
    alarmsList.appendChild(more);
  }
  if (!notificationsItemsCache.length) {
    updateNotificationsButtonLabel(alarms.length);
  }
}

function isCommessaClosedForNotifications(commessa) {
  const raw = normalizePhaseKey(commessa?.stato || "");
  return raw === "chiusa" || raw.includes("evasa");
}

function getCommessaNotificationDueDate(commessa, field) {
  if (!commessa || !field) return null;
  if (field === "data_consegna_macchina") {
    return parseIsoDateOnly(commessa.data_consegna_macchina || commessa.data_consegna_prevista || commessa.data_consegna || null);
  }
  return parseIsoDateOnly(commessa[field] || null);
}

function isNotificationEntryOwnedByCurrentUser(entry, ownRisorsaId, ownProfileId) {
  if (!entry) return false;
  const byRisorsa = ownRisorsaId && String(entry.risorsaId || "") === String(ownRisorsaId);
  const byAssignee = ownProfileId && String(entry.assegnatoA || "") === String(ownProfileId);
  return Boolean(byRisorsa || byAssignee);
}

function getNotificationUrgencyMeta(item, today = startOfDay(new Date())) {
  if (item.type === "allarme") {
    return {
      group: 0,
      sortValue: Number.NEGATIVE_INFINITY,
      text: "Dato incoerente",
      toneClass: "is-danger",
      cardToneClass: "is-danger",
    };
  }
  const dueAt = parseIsoDateOnly(item.dueAt);
  if (!dueAt) {
    return {
      group: 4,
      sortValue: Number.POSITIVE_INFINITY,
      text: "Senza data",
      toneClass: "",
      cardToneClass: "",
    };
  }
  const remaining = businessDayDiff(today, dueAt);
  if (remaining < 0) {
    return {
      group: 0,
      sortValue: remaining,
      text: `Scaduta da ${Math.abs(remaining)}g lav.`,
      toneClass: "is-danger",
      cardToneClass: "is-danger",
    };
  }
  if (remaining === 0) {
    return {
      group: 1,
      sortValue: 0,
      text: "Scade oggi",
      toneClass: "is-warn",
      cardToneClass: "is-warn",
    };
  }
  if (remaining <= 2) {
    return {
      group: 2,
      sortValue: remaining,
      text: `Urgente: ${remaining}g lav.`,
      toneClass: "is-soon",
      cardToneClass: "is-soon",
    };
  }
  return {
    group: 3,
    sortValue: remaining,
    text: `Tra ${remaining}g lav.`,
    toneClass: "",
    cardToneClass: "",
  };
}

function sortNotificationItems(items) {
  return items.slice().sort((a, b) => {
    if (a.urgency.group !== b.urgency.group) return a.urgency.group - b.urgency.group;
    if (a.urgency.sortValue !== b.urgency.sortValue) return a.urgency.sortValue - b.urgency.sortValue;
    const aDue = parseIsoDateOnly(a.dueAt);
    const bDue = parseIsoDateOnly(b.dueAt);
    const aTs = aDue ? aDue.getTime() : Number.POSITIVE_INFINITY;
    const bTs = bDue ? bDue.getTime() : Number.POSITIVE_INFINITY;
    if (aTs !== bTs) return aTs - bTs;
    if (String(a.commessaCodice || "") !== String(b.commessaCodice || "")) {
      return String(a.commessaCodice || "").localeCompare(String(b.commessaCodice || ""));
    }
    return String(a.title || "").localeCompare(String(b.title || ""));
  });
}

async function ensureNotificationsDataLoaded(commesse) {
  if (!Array.isArray(commesse) || !commesse.length) return;
  const missingActivities = commesse.filter((c) => !state.reportActivitiesMap.has(String(c.id)));
  const missingFlows = commesse.filter((c) => !state.todoClientFlowLoadedCommesse.has(String(c.id)));
  if (missingActivities.length) {
    for (const chunk of chunkArray(missingActivities, REPORT_MAX_ITEMS)) {
      await loadReportActivitiesFor(chunk);
    }
  }
  if (missingFlows.length) {
    for (const chunk of chunkArray(missingFlows, REPORT_MAX_ITEMS)) {
      await loadTodoClientFlowsFor(chunk);
    }
  }
}

function buildNotificationItems(commesse) {
  const today = startOfDay(new Date());
  const ownRisorsaId = getOwnRisorsaId();
  const ownProfileId = String(state.profile?.id || "");
  const commessaById = new Map((commesse || []).map((commessa) => [String(commessa.id), commessa]));
  const commessaMineMap = new Map();
  const commessaPeopleMap = new Map();
  const items = [];

  (commesse || []).forEach((commessa) => {
    const commessaId = String(commessa.id || "");
    if (!commessaId) return;
    const entries = state.reportActivitiesMap.get(commessaId) || [];
    const activeEntries = entries.filter((entry) => normalizePhaseKey(entry.stato || "") !== "annullata");
    const people = Array.from(
      new Set(activeEntries.map((entry) => String(entry.risorsa || "").trim()).filter(Boolean))
    );
    const mine = activeEntries.some((entry) => isNotificationEntryOwnedByCurrentUser(entry, ownRisorsaId, ownProfileId));
    commessaPeopleMap.set(commessaId, people);
    commessaMineMap.set(commessaId, mine);

    NOTIFICATION_COMMESSA_DEADLINES.forEach((def) => {
      const dueAt = getCommessaNotificationDueDate(commessa, def.field);
      if (!dueAt) return;
      items.push({
        id: `commessa|${commessaId}|${def.field}`,
        type: "commessa",
        typeLabel: "Commessa",
        phaseKey: "",
        commessaId: commessa.id,
        commessaCodice: commessa.codice || "",
        commessaLabel: getCommessaLabel(commessa.id) || `${commessa.codice || commessa.id}`,
        title: def.label,
        detail: `Stato: ${commessa.stato || "-"}`,
        dueAt,
        people,
        mine,
      });
    });

    activeEntries.forEach((entry) => {
      const status = normalizePhaseKey(entry.stato || "");
      if (status === "completata" || status === "annullata") return;
      if (isAssenteTitle(entry.titolo || "")) return;
      const dueAt = parseIsoDateOnly(entry.end || entry.start || null);
      if (!dueAt) return;
      const people = Array.from(new Set([String(entry.risorsa || "").trim()].filter(Boolean)));
      const mine = isNotificationEntryOwnedByCurrentUser(entry, ownRisorsaId, ownProfileId);
      const dept = normalizeDeptKey(entry.dept || "");
      const deptLabel = dept || "Reparto n/d";
      items.push({
        id: `attivita|${commessaId}|${entry.id || `${entry.titolo}|${dueAt.getTime()}`}`,
        type: "attivita",
        typeLabel: "Attivita",
        phaseKey: normalizePhaseKey(entry.titolo || ""),
        commessaId: commessa.id,
        commessaCodice: commessa.codice || "",
        commessaLabel: getCommessaLabel(commessa.id) || `${commessa.codice || commessa.id}`,
        title: entry.titolo || "Attivita",
        detail: `${deptLabel} · ${status || "pianificata"}`,
        dueAt,
        people,
        mine,
      });
    });
  });

  state.todoClientFlowMap.forEach((flow) => {
    const commessaId = String(flow?.commessaId || "");
    if (!commessaById.has(commessaId)) return;
    const resolved = getTodoClientFlowResolvedStatus(flow, today);
    if (resolved !== "in_attesa" && resolved !== "scaduto") return;
    const dueAt = parseIsoDateOnly(flow.dueAt || null);
    if (!dueAt) return;
    const commessa = commessaById.get(commessaId);
    const relatedEntries = getTodoEntriesForCell(commessaId, flow.titolo, flow.reparto || "");
    const people = Array.from(
      new Set(relatedEntries.map((entry) => String(entry.risorsa || "").trim()).filter(Boolean))
    );
    const mine =
      relatedEntries.some((entry) => isNotificationEntryOwnedByCurrentUser(entry, ownRisorsaId, ownProfileId)) ||
      Boolean(commessaMineMap.get(commessaId));
    const sentLabel = flow.sentAt ? formatDateDMY(flow.sentAt) : "n/d";
    items.push({
      id: `cliente|${commessaId}|${normalizePhaseKey(flow.titolo)}|${normalizeDeptKey(flow.reparto || "")}`,
      type: "cliente",
      typeLabel: "Cliente",
      phaseKey: normalizePhaseKey(flow.titolo || ""),
      commessaId: commessa?.id,
      commessaCodice: commessa?.codice || "",
      commessaLabel: getCommessaLabel(commessa?.id) || `${commessa?.codice || commessa?.id || commessaId}`,
      title: `Conferma cliente ${flow.titolo || ""}`.trim(),
      detail: `Inviato ${sentLabel}`,
      dueAt,
      people,
      mine,
    });
  });

  getTelaioOrderAlarms().forEach((commessa) => {
    const commessaId = String(commessa.id || "");
    items.push({
      id: `allarme|${commessaId}|telaio_ordered_missing_date`,
      type: "allarme",
      typeLabel: "Allarme",
      phaseKey: "",
      commessaId: commessa.id,
      commessaCodice: commessa.codice || "",
      commessaLabel: getCommessaLabel(commessa.id) || `${commessa.codice || commessa.id}`,
      title: "Telaio ordinato senza data",
      detail: "Correggere data ordine telaio effettiva.",
      dueAt: null,
      people: commessaPeopleMap.get(commessaId) || [],
      mine: Boolean(commessaMineMap.get(commessaId)),
    });
  });

  const enriched = items.map((item) => ({
    ...item,
    urgency: getNotificationUrgencyMeta(item, today),
  }));
  return sortNotificationItems(enriched);
}

function isPreliminare3DNotification(item) {
  const phaseKey = normalizePhaseKey(item?.phaseKey || "");
  return NOTIFICATION_PHASE_CLASS_KEYS.has(phaseKey);
}

function getNotificationClassOptions(items = []) {
  const hasAllarme = (items || []).some((item) => item?.type === "allarme");
  const base = [
    { key: "all", label: "Tutte" },
    { key: "preliminare_3d", label: "Preliminare/3D" },
    { key: "commessa", label: "Scadenze commessa" },
    { key: "attivita", label: "Attivita" },
    { key: "cliente", label: "Cliente" },
  ];
  if (hasAllarme) base.push({ key: "allarme", label: "Allarmi" });
  return base;
}

function matchesNotificationClassFilter(item, classKey = "all") {
  if (!item) return false;
  if (!classKey || classKey === "all") return true;
  if (classKey === "preliminare_3d") return isPreliminare3DNotification(item);
  if (classKey === "commessa") return item.type === "commessa";
  if (classKey === "attivita") return item.type === "attivita";
  if (classKey === "cliente") return item.type === "cliente";
  if (classKey === "allarme") return item.type === "allarme";
  return true;
}

function renderNotificationsClassFilters(items = []) {
  if (!notificationsClassFilters) return;
  const options = getNotificationClassOptions(items);
  const validKeys = new Set(options.map((option) => option.key));
  if (!validKeys.has(notificationsClassFilter)) {
    notificationsClassFilter = "all";
  }
  notificationsClassFilters.innerHTML = "";
  options.forEach((option) => {
    const count =
      option.key === "all"
        ? items.length
        : items.filter((item) => matchesNotificationClassFilter(item, option.key)).length;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `notifications-class-btn${notificationsClassFilter === option.key ? " active" : ""}`;
    btn.dataset.classKey = option.key;
    btn.textContent = `${option.label} (${count})`;
    btn.addEventListener("click", () => {
      notificationsClassFilter = option.key;
      renderNotificationsFromCache();
    });
    notificationsClassFilters.appendChild(btn);
  });
}

function updateNotificationsOnlyMineButton() {
  if (!notificationsOnlyMineBtn) return;
  notificationsOnlyMineBtn.dataset.onlyMine = notificationsOnlyMine ? "true" : "false";
  notificationsOnlyMineBtn.textContent = notificationsOnlyMine ? "Solo le mie: ON" : "Solo le mie: OFF";
  notificationsOnlyMineBtn.classList.toggle("active", notificationsOnlyMine);
}

function updateNotificationsButtonLabel(count = 0) {
  if (!openNotificationsBtn) return;
  const safeCount = Number(count) || 0;
  openNotificationsBtn.textContent =
    safeCount > 0 ? `Controlla notifiche (${safeCount})` : "Controlla notifiche";
}

function renderNotificationsFromCache() {
  if (!notificationsList || !notificationsSummary) return;
  const full = Array.isArray(notificationsItemsCache) ? notificationsItemsCache : [];
  const scoped = notificationsOnlyMine ? full.filter((item) => item.mine) : full;
  renderNotificationsClassFilters(scoped);
  const visible = scoped.filter((item) => matchesNotificationClassFilter(item, notificationsClassFilter));
  updateNotificationsButtonLabel(full.length);
  const criticalCount = full.filter((item) => item.urgency.group === 0).length;
  const modeLabel = notificationsOnlyMine ? " (solo mie)" : "";
  const classLabel =
    getNotificationClassOptions(scoped).find((option) => option.key === notificationsClassFilter)?.label || "Tutte";
  notificationsSummary.textContent =
    `${visible.length} notifiche visualizzate${modeLabel} · Classe: ${classLabel} · Critiche: ${criticalCount}.`;
  notificationsList.innerHTML = "";

  if (!visible.length) {
    notificationsList.innerHTML = `<div class="matrix-empty">Nessuna scadenza da mostrare.</div>`;
    return;
  }

  visible.forEach((item) => {
    const wrap = document.createElement("article");
    const cardTone = item.urgency.cardToneClass ? ` ${item.urgency.cardToneClass}` : "";
    wrap.className = `notification-item${cardTone}`;
    const dueAt = parseIsoDateOnly(item.dueAt);
    const dueLabel = dueAt ? formatDateDMY(dueAt) : "n/d";
    const peopleText = item.people && item.people.length ? item.people.slice(0, 4).join(", ") : "n/d";
    const urgencyClass = item.urgency.toneClass ? ` ${item.urgency.toneClass}` : "";
    wrap.innerHTML = `
      <div class="notification-main">
        <div class="notification-top">
          <span class="notification-kind">${escapeHtml(item.typeLabel || "Notifica")}</span>
          <span class="notification-urgency${urgencyClass}">${escapeHtml(item.urgency.text || "")}</span>
        </div>
        <div class="notification-title">${escapeHtml(item.commessaLabel || "-")} · ${escapeHtml(item.title || "-")}</div>
        <div class="notification-meta">Scadenza: ${escapeHtml(dueLabel)} · ${escapeHtml(item.detail || "")}</div>
        <div class="notification-meta">Coinvolti: ${escapeHtml(peopleText)}</div>
      </div>
      <button class="ghost notification-open-btn" type="button">Apri commessa</button>
    `;
    const openBtn = wrap.querySelector(".notification-open-btn");
    if (openBtn) {
      openBtn.addEventListener("click", () => {
        if (!item.commessaId) return;
        selectCommessa(item.commessaId, { openModal: true });
        closeNotificationsModal();
      });
    }
    notificationsList.appendChild(wrap);
  });
}

async function refreshNotificationsPanel() {
  if (!notificationsSummary || !notificationsList) return;
  if (!state.session) {
    notificationsItemsCache = [];
    renderNotificationsClassFilters([]);
    notificationsSummary.textContent = "Accedi per vedere le notifiche.";
    notificationsList.innerHTML = `<div class="matrix-empty">Accedi per vedere le scadenze.</div>`;
    updateNotificationsButtonLabel(0);
    return;
  }
  if (notificationsRefreshInFlight) return;
  notificationsRefreshInFlight = true;
  if (refreshNotificationsBtn) {
    refreshNotificationsBtn.disabled = true;
    refreshNotificationsBtn.textContent = "Aggiorno...";
  }
  try {
    notificationsSummary.textContent = "Caricamento scadenze in corso...";
    notificationsList.innerHTML = `<div class="matrix-empty">Caricamento notifiche...</div>`;
    const commesseScope = (state.commesse || []).filter((commessa) => !isCommessaClosedForNotifications(commessa));
    if (!commesseScope.length) {
      notificationsItemsCache = [];
      renderNotificationsClassFilters([]);
      notificationsSummary.textContent = "Nessuna commessa attiva con scadenze.";
      notificationsList.innerHTML = `<div class="matrix-empty">Nessuna commessa attiva.</div>`;
      updateNotificationsButtonLabel(0);
      return;
    }
    await ensureNotificationsDataLoaded(commesseScope);
    notificationsItemsCache = buildNotificationItems(commesseScope);
    renderNotificationsFromCache();
  } catch (err) {
    const msg = `Errore notifiche: ${err?.message || err}`;
    renderNotificationsClassFilters([]);
    notificationsSummary.textContent = msg;
    notificationsList.innerHTML = `<div class="matrix-empty">${escapeHtml(msg)}</div>`;
    setStatus(msg, "error");
  } finally {
    notificationsRefreshInFlight = false;
    if (refreshNotificationsBtn) {
      refreshNotificationsBtn.disabled = false;
      refreshNotificationsBtn.textContent = "Aggiorna";
    }
  }
}

function openNotificationsModal() {
  if (!notificationsModal) return;
  if (!state.session) {
    setStatus("Accedi per vedere le notifiche.", "error");
    return;
  }
  notificationsModal.classList.remove("hidden");
  updateNotificationsOnlyMineButton();
  void refreshNotificationsPanel();
}

function closeNotificationsModal() {
  if (!notificationsModal) return;
  notificationsModal.classList.add("hidden");
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
const commesseSeqAlert = el("commesseSeqAlert");
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
const importProgressWrap = el("importProgressWrap");
const importProgressBar = el("importProgressBar");
const importProgressText = el("importProgressText");
const importResult = el("importResult");
const importCommitBtn = el("importCommitBtn");
const copyImportHeaderBtn = el("copyImportHeaderBtn");
const closeImportBtn = el("closeImportBtn");
const importDatesModal = el("importDatesModal");
const importDatesTextarea = el("importDatesTextarea");
const importDatesPreview = el("importDatesPreview");
const importDatesCommitBtn = el("importDatesCommitBtn");
const copyImportDatesHeaderBtn = el("copyImportDatesHeaderBtn");
const closeImportDatesBtn = el("closeImportDatesBtn");
const importProductionModal = el("importProductionModal");
const importProductionTextarea = el("importProductionTextarea");
const importProductionPreview = el("importProductionPreview");
const importProductionCommitBtn = el("importProductionCommitBtn");
const copyImportProductionHeaderBtn = el("copyImportProductionHeaderBtn");
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
const todoSyncMatrixFiltersBtn = el("todoSyncMatrixFiltersBtn");
const todoYearFilter = el("todoYearFilter");
const todoPriorityField = el("todoPriorityField");
const todoPriorityBtn = el("todoPriorityBtn");
const todoSortBtn = el("todoSortBtn");
const todoCompleteBtn = el("todoCompleteBtn");
const todoInfoText = el("todoInfoText");
const todoFilterControls = [
  todoDescFilter,
  todoNumberFilter,
  todoSyncMatrixFiltersBtn,
  todoYearFilter,
  todoPriorityField,
  todoPriorityBtn,
  todoSortBtn,
  todoCompleteBtn,
].filter(Boolean);
const sectionToolbar = el("sectionToolbar");
const sectionToolbarDate = el("sectionToolbarDate");
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
const matrixCommessaClearBtn = el("matrixCommessaClearBtn");
const matrixCommessaPickerList = el("matrixCommessaPickerList");
const matrixQuickCommessaWrap = el("matrixQuickCommessaWrap");
const matrixQuickCommessaFilter = el("matrixQuickCommessaFilter");
const matrixAttivitaPicker = el("matrixAttivitaPicker");
const matrixAttivitaToggleBtn = el("matrixAttivitaToggleBtn");
const matrixAttivitaClearBtn = el("matrixAttivitaClearBtn");
const matrixAttivitaList = el("matrixAttivitaList");
const matrixWrap = el("matrixWrap");
const matrixCommessePanel = el("matrixCommessePanel");
const matrixCommesseList = el("matrixCommesseList");
const matrixCommessaSearch = el("matrixCommessaSearch");
const matrixCommessaYear = el("matrixCommessaYear");
const matrixAssenzaItem = el("matrixAssenzaItem");
const matrixAltroItem = el("matrixAltroItem");
const matrixPedItem = el("matrixPedItem");
const matrixSegnaturaItem = el("matrixSegnaturaItem");
const matrixOpOrdiniItem = el("matrixOpOrdiniItem");
const commessaFocusOverlay = el("commessaFocusOverlay");
const commessaQuickMenu = el("commessaQuickMenu");
const commessaQuickSeeMoreBtn = el("commessaQuickSeeMoreBtn");
const matrixDualStatusHost = el("matrixDualStatusHost");
const matrixDualStatusPlaceholder = el("matrixDualStatusPlaceholder");
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
const REPORT_GANTT_ROW_HEIGHT_STORAGE_KEY = "commesse_report_gantt_row_height";
const REPORT_GANTT_ROW_HEIGHT_DEFAULT = 58;
const REPORT_GANTT_ROW_HEIGHT_MIN = 36;
const REPORT_GANTT_ROW_HEIGHT_MAX = 170;
let reportPanZoom = null;
let reportGanttRowResize = null;
const MATRIX_MIN_WEEKS = 1;
const MATRIX_MAX_WEEKS = 12;
const MATRIX_DEFAULT_WEEKS = 3;
const MATRIX_MOVEMENT_ENABLED = true;
const MATRIX_BUTTON_NAV_ENABLED = true;
const MATRIX_ZOOM_BUTTONS_ENABLED = true;
const MATRIX_AUTO_FIT_ON_ACTIVITY_FILTER = false;
const MATRIX_PAN_DEADZONE_PX = 5;
const MATRIX_TRANSPARENT_DRAG_IMAGE_SRC = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
const matrixTransparentDragImage = new Image();
matrixTransparentDragImage.src = MATRIX_TRANSPARENT_DRAG_IMAGE_SRC;

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
const matrixFilterNumero = el("matrixFilterNumero");
const matrixFilterSearch = el("matrixFilterSearch");
const matrixFilterSelectedPills = el("matrixFilterSelectedPills");
const matrixFilterList = el("matrixFilterList");
const matrixFilterClearBtn = el("matrixFilterClearBtn");
const matrixFilterResetBtn = el("matrixFilterResetBtn");
const matrixFilterApplyBtn = el("matrixFilterApplyBtn");
const activityModal = el("activityModal");
const activityModalCard = activityModal ? activityModal.querySelector(".modal-card") : null;
const activityModalDragHandle = activityModal ? activityModal.querySelector(".activity-dual-head") : null;
const activityCloseBtn = el("activityCloseBtn");
const activityGanttBtn = el("activityGanttBtn");
const activityModalCommessaInfo = el("activityModalCommessaInfo");
const activityMeta = el("activityMeta");
const activityTitleList = el("activityTitleList");
const activityAssenzaWrap = el("activityAssenzaWrap");
const activityAssenzaOre = el("activityAssenzaOre");
const activityAssenzaFullBtn = el("activityAssenzaFullBtn");
const activityAssenzaHalfBtn = el("activityAssenzaHalfBtn");
const activityAltroNoteWrap = el("activityAltroNoteWrap");
const activityAltroNote = el("activityAltroNote");
const activityDurationInput = el("activityDurationInput");
const activityHours1Btn = el("activityHours1Btn");
const activityHours4Btn = el("activityHours4Btn");
const activityHours8Btn = el("activityHours8Btn");
const activityHoursQuickButtons = activityModal
  ? activityModal.querySelector(".activity-hours-row .quick-buttons")
  : null;
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
  tipo_macchina: el("d_tipo_macchina"),
  variante_macchina: el("d_variante_macchina"),
  stato: el("d_stato"),
  priorita: el("d_priorita"),
  data_ingresso: el("d_data_ingresso"),
  data_ordine_telaio: el("d_data_ordine_telaio"),
  data_conferma_consegna_telaio: el("d_data_conferma_consegna_telaio"),
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
  tipo_macchina: el("n_tipo_macchina"),
  variante_macchina: el("n_variante_macchina"),
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
const detailTelaioPlannerWarning = el("d_telaioPlannerWarning");
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
  dragPointerOffsetX: null,
  dragPointerOffsetY: null,
  dragFollowerEl: null,
  dragFollowerOffsetX: null,
  dragFollowerOffsetY: null,
  dragSourceEl: null,
  dragDropHandled: false,
  dragSourceRestoreTimer: null,
  view: "three",
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
  customWeeks: null,
  quickCommessaFilterRaw: "",
  quickCommessaFilterIds: new Set(),
  quickCommessaFilterYear: null,
  panVisibleBusinessDays: 0,
  panBufferBusinessDays: 0,
  panDayWidthPx: 0,
  panBaseOffsetPx: 0,
  panResidualDays: 0,
  initialRenderStabilized: false,
  initialRenderStabilizeAttempts: 0,
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
let matrixDualPanelsActive = false;
let activityModalDrag = null;
const todoStatusMenuHomeParent = todoStatusMenu ? todoStatusMenu.parentElement : null;
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
let todoSyncedCommessaIds = new Set();
let todoScrollSyncLock = false;
let todoFilterHeightLock = 0;
let todoFilterUnlockTimer = null;
let sectionToolbarDateKey = "";
let sectionToolbarDateTimer = null;
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
let notificationsOnlyMine = false;
let notificationsClassFilter = "all";
let notificationsItemsCache = [];
let notificationsRefreshInFlight = false;
let deleteCommessaId = null;
let detailSnapshot = "";
let detailDirty = false;
let backupPayload = null;
let matrixPan = null;
let matrixInitialRenderRafId = 0;
let matrixBootstrapSyncInProgress = false;

const IMPORT_HEADER_BASE = "anno\tnumero\tdescrizione\ttipo_macchina";
const IMPORT_CHUNK_SIZE = 200;
const IMPORT_HEADER_DATES =
  "codice_commessa\tingresso_ordine\tordine_telaio_target\tordine_telaio_stimato\tprelievo_materiali\tordinato\tordine_telaio_effettivo\tconsegna_macchina";
const IMPORT_HEADER_PRODUCTION =
  "codice_commessa\tordine_telaio_target\tprelievo_materiali\tdata_arrivo_kit_cavi\tconsegna_macchina";

const DEFAULT_MACHINE_TYPES = [
  "TAGO",
  "MINIBOOSTER SWISS",
  "DRAVA",
  "SENNA",
  "SENNA-XS",
  "NEVA",
  "ELBA",
  "LT-UNIT",
  "AH",
  "GH",
  "YUKON",
];

const CUSTOM_VARIANT_MACHINE_TYPES = ["NEVA", "ELBA", "LT-UNIT", "YUKON"];

const BACKUP_TABLES = [
  { name: "reparti", conflict: "id" },
  { name: "risorse", conflict: "id" },
  { name: "utenti", conflict: "id" },
  { name: "permessi_ruolo", conflict: "ruolo" },
  { name: "commesse", conflict: "id" },
  { name: "commessa_schede", conflict: "commessa_id" },
  { name: "commesse_reparti", conflict: "commessa_id,reparto_id" },
  { name: "assegnazioni", conflict: "commessa_id,utente_id" },
  { name: "attivita", conflict: "id" },
  { name: "commessa_attivita_override", conflict: "commessa_id,titolo" },
  { name: "commessa_attivita_cliente", conflict: "commessa_id,titolo,reparto" },
  { name: "commessa_imponibili", conflict: "commessa_id" },
  { name: "motore_tipologie", conflict: "id" },
  { name: "motore_varianti", conflict: "id" },
  { name: "motore_regole", conflict: "variante_id,attivita_titolo" },
  { name: "motore_canvas_config", conflict: "variante_id" },
  { name: "motore_canvas_flow", conflict: "variante_id,attivita_titolo" },
  { name: "motore_canvas_skill", conflict: "variante_id,risorsa_id" },
  { name: "whitelist_email", conflict: "email" },
  { name: "log_commessa", conflict: "id" },
];
const BACKUP_LOCAL_STORAGE_KEYS = [
  "motore_canvas_structure_v1",
  "motore_flow_layout_v1",
  "motore_sandbox_mode_v1",
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
  if (
    Object.prototype.hasOwnProperty.call(patch, "telaio_consegnato") ||
    Object.prototype.hasOwnProperty.call(patch, "data_consegna_telaio_effettiva")
  ) {
    refreshMatrixTelaioOrderIndicators(commessaId);
    renderTelaioOrderAlarms();
  }
  return target;
}

function resolveMatrixViewportAnchor(anchorEl = null) {
  if (anchorEl && anchorEl.isConnected) return anchorEl;
  if (!matrixGrid) return null;
  return (
    matrixGrid.querySelector(".matrix-activity-bar.commessa-highlight") ||
    matrixGrid.querySelector(".matrix-row[data-risorsa-id]:not(.matrix-row-uninvolved)")
  );
}

function preserveMatrixViewport(anchorEl, applyChange) {
  if (typeof applyChange !== "function") return;
  const prevX = window.scrollX;
  const prevY = window.scrollY;
  const anchor = resolveMatrixViewportAnchor(anchorEl);
  const anchorTop = anchor ? anchor.getBoundingClientRect().top : null;
  applyChange();
  requestAnimationFrame(() => {
    if (anchor && anchor.isConnected && Number.isFinite(anchorTop)) {
      const nextTop = anchor.getBoundingClientRect().top;
      const delta = nextTop - anchorTop;
      if (Math.abs(delta) > 0.5) {
        window.scrollBy(0, delta);
        return;
      }
    }
    window.scrollTo(prevX, prevY);
  });
}

function isPlannerRole() {
  if (!state.profile) return false;
  return String(state.profile.ruolo || "").trim().toLowerCase() === "planner";
}

function canEditTelaioOrderByRole() {
  if (!state.profile) return false;
  const role = String(state.profile.ruolo || "").trim().toLowerCase();
  return role === "admin" || role === "responsabile";
}

function refreshDetailTelaioOrderAccess() {
  const canEdit = state.canWrite && canEditTelaioOrderByRole();
  const isPlanner = isPlannerRole();
  if (d.telaio_ordinato) d.telaio_ordinato.disabled = !canEdit;
  if (d.data_consegna_telaio_effettiva) {
    d.data_consegna_telaio_effettiva.disabled = !canEdit;
    d.data_consegna_telaio_effettiva.title = canEdit
      ? ""
      : "Solo admin e responsabile possono modificare lo stato ordine telaio.";
  }
  if (d.data_conferma_consegna_telaio) {
    d.data_conferma_consegna_telaio.disabled = !canEdit;
    d.data_conferma_consegna_telaio.title = canEdit
      ? ""
      : "Solo admin e responsabile possono modificare la data pianificata telaio.";
  }
  if (d.telaio_ordinato) {
    d.telaio_ordinato.title = canEdit ? "" : "Solo admin e responsabile possono modificare lo stato ordine telaio.";
  }
  if (detailTelaioPlannerWarning) {
    detailTelaioPlannerWarning.classList.toggle("hidden", !isPlanner);
  }
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

function isTelaioActivityTitle(title) {
  return normalizePhaseKey(title) === "telaio";
}

function getLastTelaioEndDayFromRows(rows) {
  if (!Array.isArray(rows) || !rows.length) return null;
  const telaioRows = rows.filter((row) => isTelaioActivityTitle(row?.titolo || ""));
  if (!telaioRows.length) return null;
  let last = null;
  telaioRows.forEach((row) => {
    const end = row?.data_fine ? startOfDay(new Date(row.data_fine)) : null;
    if (!end || Number.isNaN(end.getTime())) return;
    if (!last || end.getTime() > last.getTime()) last = end;
  });
  return last;
}

async function fetchLastTelaioEndDayForCommessa(commessaId) {
  if (!commessaId) return null;
  const { data, error } = await supabase
    .from("attivita")
    .select("titolo,data_fine")
    .eq("commessa_id", commessaId);
  if (error) {
    console.warn("telaio consistency check error", error.message);
    return null;
  }
  return getLastTelaioEndDayFromRows(data || []);
}

function showTelaioPlannedMismatchWarning(commessaId, plannedDay, telaioEndDay) {
  const label = getCommessaLabel(commessaId) || "Commessa";
  if (!telaioEndDay) {
    const msg = `${label}: data telaio pianificata ${formatDateDMY(plannedDay)} ma nessuna attivita TELAIO in Matrice.`;
    showToast(msg, "info");
    return;
  }
  if (plannedDay.getTime() !== telaioEndDay.getTime()) {
    const msg = `${label}: TELAIO in Matrice chiude il ${formatDateDMY(
      telaioEndDay
    )}, ma data telaio pianificata e ${formatDateDMY(plannedDay)}.`;
    showToast(msg, "info");
  }
}

async function warnTelaioPlannedMismatch(commessaId, plannedDateRaw, options = {}) {
  const plannedDay = parseIsoDateOnly(plannedDateRaw || null);
  if (!commessaId || !plannedDay) return;
  const hintedEndDay = options.telaioEndDay || null;
  const telaioEndDay = hintedEndDay || (await fetchLastTelaioEndDayForCommessa(commessaId));
  showTelaioPlannedMismatchWarning(commessaId, plannedDay, telaioEndDay);
}

function applyCommessaHighlight(commessaId, options = {}) {
  if (!matrixGrid) return;
  if (!commessaId) return;
  const { preserveViewport: keepViewport = false, anchorEl = null } = options;
  commessaHighlightId = commessaId;
  commessaHighlightAt = Date.now();
  commessaLongPressSuppress = true;
  const bars = matrixGrid.querySelectorAll(".matrix-activity-bar");
  bars.forEach((bar) => {
    const same = bar.dataset && bar.dataset.commessaId === String(commessaId);
    bar.classList.toggle("commessa-highlight", same);
    bar.classList.toggle("commessa-dim", !same);
  });
  const focusedRisorse = new Set();
  matrixState.attivita.forEach((a) => {
    if (String(a.commessa_id || "") !== String(commessaId)) return;
    if (!a.risorsa_id) return;
    focusedRisorse.add(String(a.risorsa_id));
  });
  applyMatrixResourceCompaction(focusedRisorse, {
    preserveViewport: keepViewport,
    anchorEl,
  });
}

function isMatrixQuickCommessaFilterActive() {
  return Boolean(String(matrixState.quickCommessaFilterRaw || "").trim());
}

function getCommessaYearAndNumber(commessa) {
  const parsed = parseCommessaCode(commessa?.codice || "");
  const yearRaw = Number(commessa?.anno);
  const numberRaw = Number(commessa?.numero);
  const year = Number.isInteger(yearRaw) && yearRaw > 0 ? yearRaw : parsed.year;
  const num = Number.isInteger(numberRaw) && numberRaw > 0 ? numberRaw : parsed.num;
  return { year, num };
}

function resolveMatrixQuickCommessaFilter(numero) {
  if (!Number.isFinite(numero) || numero < 1) {
    return { ids: new Set(), year: null };
  }
  const currentYear = new Date().getFullYear();
  const previousYear = currentYear - 1;
  const byYear = (targetYear) =>
    (state.commesse || [])
      .filter((c) => {
        const parts = getCommessaYearAndNumber(c);
        return parts.year === targetYear && parts.num === numero;
      })
      .map((c) => String(c.id));
  const currentMatches = byYear(currentYear);
  if (currentMatches.length) {
    return { ids: new Set(currentMatches), year: currentYear };
  }
  const previousMatches = byYear(previousYear);
  if (previousMatches.length) {
    return { ids: new Set(previousMatches), year: previousYear };
  }
  return { ids: new Set(), year: null };
}

function getActiveMatrixCommessaFilter() {
  if (isMatrixQuickCommessaFilterActive()) {
    return {
      active: true,
      ids: new Set(matrixState.quickCommessaFilterIds || []),
      source: "quick",
    };
  }
  return {
    active: matrixState.selectedCommesse.size > 0,
    ids: new Set(matrixState.selectedCommesse || []),
    source: "advanced",
  };
}

function updateMatrixQuickCommessaUi() {
  const quickActive = isMatrixQuickCommessaFilterActive();
  const hasMatch = (matrixState.quickCommessaFilterIds || new Set()).size > 0;
  if (matrixQuickCommessaWrap) {
    matrixQuickCommessaWrap.classList.toggle("is-active", quickActive);
    matrixQuickCommessaWrap.classList.toggle("is-no-match", quickActive && !hasMatch);
  }
  if (matrixCommessaPickerToggleBtn) {
    matrixCommessaPickerToggleBtn.classList.toggle("is-disabled", quickActive);
    matrixCommessaPickerToggleBtn.setAttribute("aria-disabled", quickActive ? "true" : "false");
    if (quickActive) {
      matrixCommessaPickerToggleBtn.title = "Disattiva il filtro rapido per usare \"Filtra commesse\"";
    } else {
      matrixCommessaPickerToggleBtn.removeAttribute("title");
    }
  }
  if (matrixQuickCommessaFilter) {
    const expected = String(matrixState.quickCommessaFilterRaw || "");
    if (matrixQuickCommessaFilter.value !== expected) matrixQuickCommessaFilter.value = expected;
    if (!quickActive) {
      matrixQuickCommessaFilter.title = "Filtro rapido per numero commessa";
      return;
    }
    if (hasMatch) {
      const annoLabel = matrixState.quickCommessaFilterYear ? `anno ${matrixState.quickCommessaFilterYear}` : "anno attivo";
      matrixQuickCommessaFilter.title = `Filtro rapido attivo (${annoLabel})`;
      return;
    }
    matrixQuickCommessaFilter.title = "Nessuna commessa trovata su anno corrente o precedente";
  }
}

function setMatrixQuickCommessaFilterRaw(rawValue, options = {}) {
  const shouldRender = options.render !== false;
  const normalized = String(rawValue || "")
    .replace(/[^\d]/g, "")
    .slice(0, 6);
  matrixState.quickCommessaFilterRaw = normalized;
  if (!normalized) {
    matrixState.quickCommessaFilterIds = new Set();
    matrixState.quickCommessaFilterYear = null;
    updateMatrixCommessaPickerLabel();
    updateMatrixQuickCommessaUi();
    if (shouldRender) renderMatrix();
    return;
  }
  const numero = Number.parseInt(normalized, 10);
  const resolved = resolveMatrixQuickCommessaFilter(numero);
  matrixState.quickCommessaFilterIds = resolved.ids;
  matrixState.quickCommessaFilterYear = resolved.year;
  updateMatrixCommessaPickerLabel();
  updateMatrixQuickCommessaUi();
  if (shouldRender) renderMatrix();
}

function updateMatrixCommessaPickerLabel() {
  if (!matrixCommessaPickerToggleBtn) return;
  if (isMatrixQuickCommessaFilterActive()) {
    matrixCommessaPickerToggleBtn.textContent = "Filtra commesse";
    matrixCommessaPickerToggleBtn.classList.remove("active");
    if (matrixCommessaClearBtn) matrixCommessaClearBtn.classList.add("hidden");
    return;
  }
  const count = matrixState.selectedCommesse.size;
  if (count === 0) {
    matrixCommessaPickerToggleBtn.textContent = "Filtra commesse";
    matrixCommessaPickerToggleBtn.classList.remove("active");
    if (matrixCommessaClearBtn) matrixCommessaClearBtn.classList.add("hidden");
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
  if (matrixCommessaClearBtn) matrixCommessaClearBtn.classList.remove("hidden");
}

function updateMatrixAttivitaPickerLabel() {
  if (!matrixAttivitaToggleBtn) return;
  const count = matrixState.selectedAttivita.size;
  if (count === 0) {
    matrixAttivitaToggleBtn.textContent = "Filtra attivita";
    matrixAttivitaToggleBtn.classList.remove("active");
    if (matrixAttivitaClearBtn) matrixAttivitaClearBtn.classList.add("hidden");
    return;
  }
  if (count === 1) {
    const only = Array.from(matrixState.selectedAttivita)[0];
    matrixAttivitaToggleBtn.textContent = only || "Attivita selezionata";
  } else {
    matrixAttivitaToggleBtn.textContent = `${count} attivita`;
  }
  matrixAttivitaToggleBtn.classList.add("active");
  if (matrixAttivitaClearBtn) matrixAttivitaClearBtn.classList.remove("hidden");
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
  updateMatrixQuickCommessaUi();
}

function clearCommessaHighlight(options = {}) {
  if (!matrixGrid) return;
  if (!commessaHighlightId) return;
  const { preserveViewport: keepViewport = false, anchorEl = null } = options;
  const viewportAnchor = resolveMatrixViewportAnchor(anchorEl);
  const bars = matrixGrid.querySelectorAll(".matrix-activity-bar");
  bars.forEach((bar) => {
    bar.classList.remove("commessa-highlight", "commessa-dim");
  });
  clearMatrixResourceCompaction({
    preserveViewport: keepViewport,
    anchorEl: viewportAnchor,
  });
  commessaHighlightId = null;
  if (matrixState.selectedAttivita.size > 0) {
    applyMatrixActivityFilterCompaction();
  }
}

function clearMatrixResourceCompaction(options = {}) {
  if (!matrixGrid) return;
  const { preserveViewport: keepViewport = false, anchorEl = null } = options;
  const runClear = () => {
    matrixGrid.querySelectorAll(".matrix-row-uninvolved").forEach((row) => {
      row.classList.remove("matrix-row-uninvolved");
    });
    matrixGrid.querySelectorAll(".matrix-section-row-uninvolved").forEach((row) => {
      row.classList.remove("matrix-section-row-uninvolved");
    });
  };
  if (!keepViewport) {
    runClear();
    return;
  }
  preserveMatrixViewport(anchorEl, runClear);
}

function applyMatrixResourceCompaction(involvedRisorse = new Set(), options = {}) {
  if (!matrixGrid) return;
  const { preserveViewport: keepViewport = false, anchorEl = null } = options;
  const runCompaction = () => {
    matrixGrid.querySelectorAll(".matrix-row[data-risorsa-id]").forEach((row) => {
      const risorsaId = String(row.dataset.risorsaId || "");
      const involved = involvedRisorse.has(risorsaId);
      row.classList.toggle("matrix-row-uninvolved", !involved);
    });
    const sectionRows = Array.from(matrixGrid.querySelectorAll(".matrix-section-row"));
    sectionRows.forEach((section) => {
      let hasInvolved = false;
      let cursor = section.nextElementSibling;
      while (cursor && !cursor.classList.contains("matrix-section-row")) {
        if (
          cursor.classList &&
          cursor.classList.contains("matrix-row") &&
          cursor.dataset.risorsaId &&
          !cursor.classList.contains("matrix-row-uninvolved")
        ) {
          hasInvolved = true;
          break;
        }
        cursor = cursor.nextElementSibling;
      }
      section.classList.toggle("matrix-section-row-uninvolved", !hasInvolved);
    });
  };
  if (!keepViewport) {
    runCompaction();
    return;
  }
  preserveMatrixViewport(anchorEl, runCompaction);
}

function applyMatrixActivityFilterCompaction() {
  if (!matrixState.selectedAttivita.size) {
    clearMatrixResourceCompaction();
    return;
  }
  const involved = new Set();
  const commessaFilter = getActiveMatrixCommessaFilter();
  matrixState.attivita.forEach((a) => {
    if (!matrixState.selectedAttivita.has(a.titolo)) return;
    if (commessaFilter.active && !commessaFilter.ids.has(String(a.commessa_id || ""))) return;
    if (!a.risorsa_id) return;
    involved.add(String(a.risorsa_id));
  });
  applyMatrixResourceCompaction(involved);
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
  setStatus(nextState === "completata" ? "Activity completed." : "Activity reopened.", "ok");
  const touchedCommessaId = quickMenuAttivita?.commessa_id || null;
  closeMatrixQuickMenu();
  await loadMatrixAttivita();
  await refreshTodoRealtimeSnapshot([touchedCommessaId]);
}

function setMatrixViewLabel() {
  return;
}

function updateMatrixDropEffect(e) {
  if (!e?.dataTransfer) return;
  const isCopy = e.altKey || e.ctrlKey || e.metaKey;
  try {
    e.dataTransfer.dropEffect = isCopy ? "copy" : "move";
  } catch (_err) {
    // Ignore unsupported dropEffect assignments.
  }
}

function updateMatrixDragFollowerPosition(clientX, clientY) {
  const el = matrixState.dragFollowerEl;
  if (!el) return;
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY)) return;
  const rect = el.getBoundingClientRect();
  const offsetX = Number.isFinite(matrixState.dragFollowerOffsetX)
    ? matrixState.dragFollowerOffsetX
    : rect.width / 2;
  const offsetY = Number.isFinite(matrixState.dragFollowerOffsetY)
    ? matrixState.dragFollowerOffsetY
    : rect.height / 2;
  el.style.left = `${clientX - offsetX}px`;
  el.style.top = `${clientY - offsetY}px`;
}

function removeMatrixDragFollower() {
  if (matrixState.dragFollowerEl) {
    matrixState.dragFollowerEl.remove();
    matrixState.dragFollowerEl = null;
  }
  matrixState.dragFollowerOffsetX = null;
  matrixState.dragFollowerOffsetY = null;
}

function clearMatrixDragSourceRestoreTimer() {
  if (!matrixState.dragSourceRestoreTimer) return;
  window.clearTimeout(matrixState.dragSourceRestoreTimer);
  matrixState.dragSourceRestoreTimer = null;
}

function restoreMatrixDragSourceVisual() {
  clearMatrixDragSourceRestoreTimer();
  removeMatrixDragFollower();
  const source = matrixState.dragSourceEl;
  if (source && source.isConnected) {
    source.style.removeProperty("opacity");
    source.style.removeProperty("pointer-events");
    source.classList.remove("dragging");
  }
  matrixState.dragSourceEl = null;
  matrixState.dragDropHandled = false;
}

function markMatrixDragDropHandled() {
  matrixState.dragDropHandled = true;
  const source = matrixState.dragSourceEl;
  if (source && source.isConnected) {
    source.style.opacity = "0";
    source.style.pointerEvents = "none";
  }
  clearMatrixDragSourceRestoreTimer();
  matrixState.dragSourceRestoreTimer = window.setTimeout(() => {
    // Fallback: if no redraw happened (drop error/cancel), make source visible again.
    restoreMatrixDragSourceVisual();
  }, 2600);
}

function createMatrixDragFollower(bar, e) {
  removeMatrixDragFollower();
  restoreMatrixDragSourceVisual();
  if (!bar) return;
  const rect = bar.getBoundingClientRect();
  if (!rect.width || !rect.height) return;
  const follower = bar.cloneNode(true);
  follower.classList.remove("dragging", "resizing");
  follower.classList.add("matrix-drag-follower");
  follower.querySelectorAll(".matrix-resize-handle").forEach((node) => node.remove());
  follower.style.width = `${rect.width}px`;
  follower.style.height = `${rect.height}px`;
  follower.style.left = `${rect.left}px`;
  follower.style.top = `${rect.top}px`;
  document.body.appendChild(follower);
  matrixState.dragFollowerEl = follower;
  matrixState.dragFollowerOffsetX = Number.isFinite(matrixState.dragPointerOffsetX)
    ? matrixState.dragPointerOffsetX
    : rect.width / 2;
  matrixState.dragFollowerOffsetY = Number.isFinite(matrixState.dragPointerOffsetY)
    ? matrixState.dragPointerOffsetY
    : rect.height / 2;
  matrixState.dragSourceEl = bar;
  matrixState.dragDropHandled = false;
  updateMatrixDragFollowerPosition(e?.clientX, e?.clientY);
}

function resetMatrixMovementBaseline() {
  matrixState.customWeeks = null;
  matrixState.view = "three";
  matrixState.date = normalizeMatrixStartDate(new Date());
  matrixState.panVisibleBusinessDays = 0;
  matrixState.panBufferBusinessDays = 0;
  matrixState.panDayWidthPx = 0;
  matrixState.panBaseOffsetPx = 0;
  matrixState.panResidualDays = 0;
  matrixPan = null;
  if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
  if (matrixGrid) {
    matrixGrid.classList.remove("is-panning");
    matrixGrid.classList.remove("is-zooming-preview");
    matrixGrid.classList.remove("is-pan-snapping");
  }
  setMatrixViewLabel();
}

function cancelPendingMatrixInitialRender() {
  if (!matrixInitialRenderRafId) return;
  cancelAnimationFrame(matrixInitialRenderRafId);
  matrixInitialRenderRafId = 0;
}

function scheduleMatrixRenderStabilized() {
  if (matrixState.initialRenderStabilized) {
    renderMatrix();
    renderMatrixReport();
    return;
  }
  cancelPendingMatrixInitialRender();
  const getProbeWidth = () => {
    if (!matrixGrid) return 0;
    const rectW = Number(matrixGrid.getBoundingClientRect?.().width || 0);
    const clientW = Number(matrixGrid.clientWidth || 0);
    return Math.max(rectW, clientW);
  };
  matrixInitialRenderRafId = requestAnimationFrame(() => {
    const w1 = getProbeWidth();
    matrixInitialRenderRafId = requestAnimationFrame(() => {
      const w2 = getProbeWidth();
      const widthStable = w1 > 0 && w2 > 0 && Math.abs(w2 - w1) < 1.5;
      const maxAttempts = 10;
      if (!widthStable && matrixState.initialRenderStabilizeAttempts < maxAttempts) {
        matrixInitialRenderRafId = 0;
        matrixState.initialRenderStabilizeAttempts += 1;
        scheduleMatrixRenderStabilized();
        return;
      }
      matrixInitialRenderRafId = 0;
      renderMatrix();
      renderMatrixReport();
      matrixState.initialRenderStabilized = true;
      matrixState.initialRenderStabilizeAttempts = 0;
    });
  });
}

function applyMatrixMovementControlsState() {
  const navControls = [matrixPrevBtn, matrixNextBtn];
  navControls.forEach((btn) => {
    if (!btn) return;
    btn.disabled = !MATRIX_BUTTON_NAV_ENABLED;
    btn.classList.toggle("hidden", !MATRIX_BUTTON_NAV_ENABLED);
    btn.setAttribute("aria-disabled", !MATRIX_BUTTON_NAV_ENABLED ? "true" : "false");
    btn.title = MATRIX_BUTTON_NAV_ENABLED ? "" : "Navigazione da pulsanti disattivata";
  });
  const zoomControls = [matrixZoomInBtn, matrixZoomOutBtn];
  zoomControls.forEach((btn) => {
    if (!btn) return;
    btn.disabled = !MATRIX_ZOOM_BUTTONS_ENABLED;
    btn.classList.toggle("hidden", !MATRIX_ZOOM_BUTTONS_ENABLED);
    btn.setAttribute("aria-disabled", !MATRIX_ZOOM_BUTTONS_ENABLED ? "true" : "false");
    btn.title = MATRIX_ZOOM_BUTTONS_ENABLED ? "" : "Zoom da pulsanti disattivato";
  });
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
  const commessaFilter = getActiveMatrixCommessaFilter();
  return commessaFilter.active ? Array.from(commessaFilter.ids) : [];
}

function getCommessaYear(commessa) {
  const code = (commessa.codice || "").trim();
  const match = code.match(/^(\d{4})/);
  return match ? match[1] : "";
}

function openMatrixFilter(mode) {
  if (!matrixFilterModal || !matrixFilterSearch || !matrixFilterTitle || !matrixFilterList) return;
  if (mode === "commessa" && isMatrixQuickCommessaFilterActive()) {
    setStatus('Disattiva prima il filtro rapido "numero comm." per usare "Filtra commesse".', "error");
    matrixQuickCommessaFilter?.focus();
    matrixQuickCommessaFilter?.select?.();
    return;
  }
  matrixState.filterMode = mode;
  matrixState.filterDraft = new Set(
    mode === "commessa" ? Array.from(matrixState.selectedCommesse) : Array.from(matrixState.selectedAttivita)
  );
  matrixFilterTitle.textContent = mode === "commessa" ? "Filtra commesse" : "Filtra attivita";
  matrixFilterSearch.value = "";
  if (matrixFilterYear) {
    if (mode === "commessa") {
      const years = Array.from(
        new Set(
          state.commesse
            .map((c) => String(c.anno || getCommessaYear(c) || "").trim())
            .filter((y) => /^\d{4}$/.test(y))
        )
      ).sort((a, b) => b.localeCompare(a));
      matrixFilterYear.innerHTML = "";
      if (!years.includes("2026")) years.unshift("2026");
      years.forEach((y) => {
        const opt = document.createElement("option");
        opt.value = y;
        opt.textContent = y;
        matrixFilterYear.appendChild(opt);
      });
      matrixFilterYear.classList.remove("hidden");
      matrixFilterYear.value = years.includes("2026") ? "2026" : years[0] || "";
    } else {
      matrixFilterYear.classList.add("hidden");
      matrixFilterYear.value = "";
    }
  }
  if (matrixFilterNumero) {
    if (mode === "commessa") {
      matrixFilterNumero.classList.remove("hidden");
      matrixFilterNumero.value = "";
    } else {
      matrixFilterNumero.classList.add("hidden");
      matrixFilterNumero.value = "";
    }
  }
  renderMatrixFilterList();
  matrixFilterModal.classList.remove("hidden");
  requestAnimationFrame(() => {
    if (mode === "commessa" && matrixFilterNumero && !matrixFilterNumero.classList.contains("hidden")) {
      matrixFilterNumero.focus();
      matrixFilterNumero.select?.();
      return;
    }
    matrixFilterSearch.focus();
    matrixFilterSearch.select?.();
  });
}

function closeMatrixFilter() {
  if (!matrixFilterModal) return;
  matrixFilterModal.classList.add("hidden");
}

function renderMatrixFilterSelectedPills() {
  if (!matrixFilterSelectedPills) return;
  if (matrixState.filterMode !== "commessa") {
    matrixFilterSelectedPills.classList.add("hidden");
    matrixFilterSelectedPills.innerHTML = "";
    return;
  }
  const selected = state.commesse
    .filter((c) => matrixState.filterDraft.has(c.id))
    .slice()
    .sort(compareCommesse);
  if (!selected.length) {
    matrixFilterSelectedPills.classList.add("hidden");
    matrixFilterSelectedPills.innerHTML = "";
    return;
  }
  matrixFilterSelectedPills.classList.remove("hidden");
  matrixFilterSelectedPills.innerHTML = "";
  selected.forEach((c) => {
    const pill = document.createElement("button");
    pill.type = "button";
    pill.className = "matrix-filter-pill";
    pill.textContent = c.codice || "commessa";
    pill.title = c.titolo ? `${c.codice} - ${c.titolo}` : c.codice || "commessa";
    pill.addEventListener("click", () => {
      matrixState.filterDraft.delete(c.id);
      renderMatrixFilterList();
    });
    matrixFilterSelectedPills.appendChild(pill);
  });
}

function renderMatrixFilterList() {
  if (!matrixFilterList || !matrixFilterSearch) return;
  const q = matrixFilterSearch.value.trim().toLowerCase();
  matrixFilterList.innerHTML = "";
  renderMatrixFilterSelectedPills();
  if (matrixState.filterMode === "commessa") {
    const yearFilter = matrixFilterYear ? matrixFilterYear.value : "";
    const numeroRaw = matrixFilterNumero ? matrixFilterNumero.value.trim() : "";
    const numeroFilter = /^\d+$/.test(numeroRaw) ? Number.parseInt(numeroRaw, 10) : null;
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
      if (numeroFilter != null) {
        const parsed = parseCommessaCode(c.codice || "");
        const cNum =
          c.numero != null && Number.isFinite(Number(c.numero))
            ? Number(c.numero)
            : parsed.num != null && Number.isFinite(Number(parsed.num))
            ? Number(parsed.num)
            : NaN;
        if (!Number.isFinite(cNum) || cNum !== numeroFilter) return false;
      }
      if (q) {
        const hay = `${c.codice || ""} ${c.titolo || ""} ${c.tipo_macchina || ""}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    const filtered = (yearFilter
      ? options.filter((c) => String(c.anno || getCommessaYear(c) || "") === yearFilter)
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
      btn.className = "filter-pill matrix-filter-commessa-pill";
      const codiceEl = document.createElement("span");
      codiceEl.className = "matrix-filter-commessa-code";
      codiceEl.textContent = c.codice || "-";
      const titoloEl = document.createElement("span");
      titoloEl.className = "matrix-filter-commessa-title";
      titoloEl.textContent = c.titolo || "Senza descrizione";
      const tipoEl = document.createElement("span");
      tipoEl.className = "matrix-filter-commessa-type";
      tipoEl.textContent = getMachineTypeDisplayLabel(c.tipo_macchina || "Altro tipo");
      btn.append(codiceEl, titoloEl, tipoEl);
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

function getMatrixFilteredBoundsFromLoaded() {
  const filterAttivita = matrixState.selectedAttivita.size > 0;
  const commessaFilter = getActiveMatrixCommessaFilter();
  const visible = (matrixState.attivita || []).filter((a) => {
    if (filterAttivita && !matrixState.selectedAttivita.has(a.titolo)) return false;
    if (commessaFilter.active && !commessaFilter.ids.has(String(a.commessa_id || ""))) return false;
    return true;
  });
  if (!visible.length) return null;
  let minStart = null;
  let maxEnd = null;
  visible.forEach((a) => {
    const start = a.data_inizio ? new Date(a.data_inizio) : null;
    const end = a.data_fine ? new Date(a.data_fine) : null;
    if (start && !Number.isNaN(start.getTime())) {
      if (!minStart || start < minStart) minStart = start;
    }
    if (end && !Number.isNaN(end.getTime())) {
      if (!maxEnd || end > maxEnd) maxEnd = end;
    }
  });
  if (!minStart || !maxEnd) return null;
  return { minStart, maxEnd };
}

async function fitMatrixToFilteredActivities() {
  if (!state.session || !matrixState.selectedAttivita.size) return;
  const selectedTitles = Array.from(matrixState.selectedAttivita);
  const commessaFilter = getActiveMatrixCommessaFilter();
  const selectedCommesse = Array.from(commessaFilter.ids);
  let minStart = null;
  let maxEnd = null;
  try {
    let startQuery = supabase
      .from("attivita")
      .select("data_inizio")
      .in("titolo", selectedTitles)
      .order("data_inizio", { ascending: true })
      .limit(1);
    let endQuery = supabase
      .from("attivita")
      .select("data_fine")
      .in("titolo", selectedTitles)
      .order("data_fine", { ascending: false })
      .limit(1);
    if (commessaFilter.active && selectedCommesse.length) {
      startQuery = startQuery.in("commessa_id", selectedCommesse);
      endQuery = endQuery.in("commessa_id", selectedCommesse);
    } else if (commessaFilter.active && !selectedCommesse.length) {
      return;
    }
    const [startRes, endRes] = await Promise.all([
      withTimeout(startQuery, FETCH_TIMEOUT_MS),
      withTimeout(endQuery, FETCH_TIMEOUT_MS),
    ]);
    if (!startRes.error && startRes.data?.length) {
      const d = new Date(startRes.data[0].data_inizio);
      if (!Number.isNaN(d.getTime())) minStart = d;
    }
    if (!endRes.error && endRes.data?.length) {
      const d = new Date(endRes.data[0].data_fine);
      if (!Number.isNaN(d.getTime())) maxEnd = d;
    }
  } catch {
    // fallback below
  }
  if (!minStart || !maxEnd) {
    const fallback = getMatrixFilteredBoundsFromLoaded();
    if (!fallback) return;
    minStart = fallback.minStart;
    maxEnd = fallback.maxEnd;
  }
  const fitStart = normalizeMatrixStartDate(minStart);
  const fitEnd = endOfDay(maxEnd);
  const totalBusinessDays = Math.max(1, businessDaysBetweenInclusive(fitStart, fitEnd));
  const weeksNeeded = Math.max(1, Math.ceil(totalBusinessDays / 5));
  matrixState.customWeeks = clampMatrixWeeks(weeksNeeded);
  syncMatrixViewFromWeeks(matrixState.customWeeks);
  matrixState.date = fitStart;
  if (matrixDate) matrixDate.value = formatDateInput(fitStart);
}

async function applyMatrixFilter() {
  if (matrixState.filterMode === "commessa") {
    setMatrixSelectedCommesse(Array.from(matrixState.filterDraft));
  } else if (matrixState.filterMode === "attivita") {
    matrixState.selectedAttivita = new Set(matrixState.filterDraft);
    updateMatrixAttivitaPickerLabel();
  }
  closeMatrixFilter();
  if (matrixState.filterMode === "attivita") {
    if (MATRIX_AUTO_FIT_ON_ACTIVITY_FILTER && matrixState.selectedAttivita.size > 0) {
      await fitMatrixToFilteredActivities();
    } else if (MATRIX_AUTO_FIT_ON_ACTIVITY_FILTER) {
      matrixState.customWeeks = null;
    }
    await loadMatrixAttivita();
    return;
  }
  renderMatrix();
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
  const commessa = attivita?.commessa_id
    ? state.commesse.find((c) => String(c.id) === String(attivita.commessa_id))
    : null;
  if (activityModalCommessaInfo) {
    if (commessa) {
      const numero = Number.isFinite(Number(commessa.numero)) ? Number(commessa.numero) : commessa.numero || "-";
      const anno = Number.isFinite(Number(commessa.anno)) ? Number(commessa.anno) : commessa.anno || "-";
      const descrizione = String(commessa.titolo || "").trim();
      activityModalCommessaInfo.textContent = descrizione
        ? `Commessa ${numero} / ${anno} - ${descrizione}`
        : `Commessa ${numero} / ${anno}`;
    } else {
      activityModalCommessaInfo.textContent = "Attivita senza commessa";
    }
  }
  const start = new Date(attivita.data_inizio);
  activityDurationInput.value =
    attivita.ore_stimate != null && !Number.isNaN(Number(attivita.ore_stimate))
      ? String(Number(attivita.ore_stimate))
      : "";
  updateActivityQuickHourButtons();
  const repartoName = getAttivitaRepartoName(attivita);
  const currentTodoStatus =
    attivita?.commessa_id && attivita?.titolo ? getTodoStatusFor(attivita.commessa_id, attivita.titolo, repartoName) : "";
  const statoKey = normalizePhaseKey(attivita?.stato || "");
  const isDoneStatus = currentTodoStatus === "fatta" || statoKey === "completata";
  const canChangeActivityTitle = currentTodoStatus === "da_schedulare" && statoKey !== "pianificata" && statoKey !== "completata";
  const changeBlockedReason = isDoneStatus
    ? "Attivita in stato DONE: il titolo non e modificabile."
    : "Il titolo attivita e modificabile solo in stato TO PLAN.";
  if (activityTitleList) {
    activityTitleList.innerHTML = "";
    activityTitleList.title = canChangeActivityTitle ? "" : changeBlockedReason;
    const options = getMatrixAttivitaOptions();
    const allowed = options.filter((name) => isActivityAllowedForRisorsa(name, attivita.risorsa_id));
    const list = options.includes(attivita.titolo)
      ? allowed.includes(attivita.titolo)
        ? allowed
        : [...allowed, attivita.titolo]
      : allowed;
    const orderedList = sortActivitiesByTodoOrder(list, repartoName);
    const repartoActivities = orderedList.filter((name) => !isLavenderActivity(name) && !isAssenteTitle(name));
    const genericActivities = orderedList.filter((name) => isLavenderActivity(name) && !isAssenteTitle(name));
    const assenteActivities = orderedList.filter((name) => isAssenteTitle(name));
    const appendOptionButton = (name, host = activityTitleList) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "filter-pill";
      if (isAssenteTitle(name)) btn.classList.add("assente-pill");
      if (isLavenderActivity(name)) btn.classList.add("lavender-pill");
      btn.textContent = getActivityDisplayLabel(name);
      if (matrixState.editingTitles.has(name)) btn.classList.add("active");
      if (!canChangeActivityTitle) {
        btn.disabled = true;
        btn.title = changeBlockedReason;
      }
      btn.addEventListener("click", () => {
        if (!canChangeActivityTitle) return;
        matrixState.editingTitles = new Set([name]);
        activityTitleList.querySelectorAll(".filter-pill.active").forEach((el) => el.classList.remove("active"));
        btn.classList.add("active");
        if (activityAssenzaWrap && activityAssenzaOre) {
          const hasAssente = isAssenteTitle(name);
          activityAssenzaWrap.classList.toggle("hidden", !hasAssente);
          if (!hasAssente) activityAssenzaOre.value = "";
        }
      });
      host.appendChild(btn);
    };
    repartoActivities.forEach((name) => appendOptionButton(name));
    if (repartoActivities.length && genericActivities.length) {
      const divider = document.createElement("div");
      divider.className = "activity-list-divider";
      divider.setAttribute("aria-hidden", "true");
      activityTitleList.appendChild(divider);
    }
    if (genericActivities.length) {
      const genericGrid = document.createElement("div");
      genericGrid.className = "activity-generic-grid";
      activityTitleList.appendChild(genericGrid);
      genericActivities.forEach((name) => appendOptionButton(name, genericGrid));
    }
    if (assenteActivities.length && (repartoActivities.length || genericActivities.length)) {
      const divider = document.createElement("div");
      divider.className = "activity-list-divider";
      divider.setAttribute("aria-hidden", "true");
      activityTitleList.appendChild(divider);
    }
    assenteActivities.forEach((name) => appendOptionButton(name));
  }
  activityMeta.textContent = commessaLabel
    ? `${attivita.titolo} \u2022 ${commessaLabel} \u2022 ${formatDateLocal(start)}`
    : `${attivita.titolo} \u2022 ${formatDateLocal(start)}`;
  activityModal.classList.remove("hidden");
  activityDurationInput.focus();
}

function updateActivityQuickHourButtons() {
  if (!activityDurationInput) return;
  const raw = String(activityDurationInput.value || "").trim();
  const value = raw === "" ? null : Number(raw);
  const isMatch = (target) => value != null && Number.isFinite(value) && Math.abs(value - target) < 0.0001;
  const isOne = isMatch(1);
  const isFour = isMatch(4);
  const isEight = isMatch(8);
  if (activityHours1Btn) activityHours1Btn.classList.toggle("active", isOne);
  if (activityHours4Btn) activityHours4Btn.classList.toggle("active", isFour);
  if (activityHours8Btn) activityHours8Btn.classList.toggle("active", isEight);
  if (activityHoursQuickButtons) {
    let activeIndex = -1;
    if (isOne) activeIndex = 0;
    else if (isFour) activeIndex = 1;
    else if (isEight) activeIndex = 2;
    activityHoursQuickButtons.style.setProperty("--quick-hours-active-index", String(Math.max(0, activeIndex)));
    activityHoursQuickButtons.style.setProperty("--quick-hours-active-opacity", activeIndex >= 0 ? "1" : "0");
  }
}

function clampActivityModalPosition(left, top) {
  if (!activityModalCard) return { left, top };
  const rect = activityModalCard.getBoundingClientRect();
  const pad = 8;
  const maxLeft = Math.max(pad, window.innerWidth - rect.width - pad);
  const maxTop = Math.max(pad, window.innerHeight - rect.height - pad);
  return {
    left: Math.min(maxLeft, Math.max(pad, left)),
    top: Math.min(maxTop, Math.max(pad, top)),
  };
}

function ensureActivityModalFixedPosition() {
  if (!activityModalCard) return;
  if (activityModalCard.style.position === "fixed") return;
  const rect = activityModalCard.getBoundingClientRect();
  activityModalCard.style.position = "fixed";
  activityModalCard.style.left = `${rect.left}px`;
  activityModalCard.style.top = `${rect.top}px`;
  activityModalCard.style.margin = "0";
}

function resetActivityModalPosition() {
  if (!activityModalCard) return;
  activityModalCard.style.position = "";
  activityModalCard.style.left = "";
  activityModalCard.style.top = "";
  activityModalCard.style.margin = "";
}

function stopActivityModalDrag(pointerId = null) {
  if (!activityModalDrag) return;
  if (pointerId != null && pointerId !== activityModalDrag.pointerId) return;
  const dragPointerId = activityModalDrag.pointerId;
  activityModalDrag = null;
  if (activityModal) activityModal.classList.remove("is-dragging-modal");
  if (activityModalDragHandle && typeof activityModalDragHandle.releasePointerCapture === "function") {
    try {
      activityModalDragHandle.releasePointerCapture(dragPointerId);
    } catch {}
  }
}

function startActivityModalDrag(event) {
  if (!activityModal || !activityModalCard || !activityModalDragHandle) return;
  if (activityModal.classList.contains("hidden")) return;
  if (event.button !== 0) return;
  if (event.target.closest("button, a, input, textarea, select, label")) return;
  ensureActivityModalFixedPosition();
  const rect = activityModalCard.getBoundingClientRect();
  activityModalDrag = {
    pointerId: event.pointerId,
    offsetX: event.clientX - rect.left,
    offsetY: event.clientY - rect.top,
  };
  if (typeof activityModalDragHandle.setPointerCapture === "function") {
    try {
      activityModalDragHandle.setPointerCapture(event.pointerId);
    } catch {}
  }
  if (activityModal) activityModal.classList.add("is-dragging-modal");
}

function moveActivityModalDrag(event) {
  if (!activityModalDrag || !activityModalCard) return;
  if (event.pointerId !== activityModalDrag.pointerId) return;
  const rawLeft = event.clientX - activityModalDrag.offsetX;
  const rawTop = event.clientY - activityModalDrag.offsetY;
  const next = clampActivityModalPosition(rawLeft, rawTop);
  activityModalCard.style.left = `${next.left}px`;
  activityModalCard.style.top = `${next.top}px`;
}

function keepActivityModalInViewport() {
  if (!activityModalCard || activityModalCard.style.position !== "fixed") return;
  const left = Number.parseFloat(activityModalCard.style.left);
  const top = Number.parseFloat(activityModalCard.style.top);
  if (!Number.isFinite(left) || !Number.isFinite(top)) return;
  const next = clampActivityModalPosition(left, top);
  activityModalCard.style.left = `${next.left}px`;
  activityModalCard.style.top = `${next.top}px`;
}

function positionMatrixDualStatusMenu() {
  if (!matrixDualPanelsActive || !todoStatusMenu || !matrixDualStatusHost) return;
  if (todoStatusMenu.parentElement !== matrixDualStatusHost) {
    matrixDualStatusHost.appendChild(todoStatusMenu);
  }
}

function mountTodoStatusMenuInDualHost() {
  if (!todoStatusMenu || !matrixDualStatusHost) return;
  todoStatusMenu.classList.add("matrix-side-by-side");
  if (matrixDualStatusPlaceholder) matrixDualStatusPlaceholder.classList.add("hidden");
  matrixDualStatusHost.appendChild(todoStatusMenu);
}

function unmountTodoStatusMenuFromDualHost(message = "") {
  if (!todoStatusMenu) return;
  todoStatusMenu.classList.remove("matrix-side-by-side");
  todoStatusMenu.style.left = "";
  todoStatusMenu.style.top = "";
  if (todoStatusMenuHomeParent && todoStatusMenu.parentElement !== todoStatusMenuHomeParent) {
    todoStatusMenuHomeParent.appendChild(todoStatusMenu);
  }
  if (matrixDualStatusPlaceholder) {
    matrixDualStatusPlaceholder.textContent = message || "Status disponibile dopo assegnazione attivita.";
    matrixDualStatusPlaceholder.classList.remove("hidden");
  }
}

function openMatrixDualPanels(attivita, anchorEl) {
  if (!attivita || !anchorEl) return;
  matrixDualPanelsActive = true;
  openActivityModal(attivita);
  if (activityModal) activityModal.classList.add("matrix-dual-open");
  if (!attivita.commessa_id || !attivita.titolo) {
    unmountTodoStatusMenuFromDualHost("Status disponibile dopo assegnazione attivita.");
    return;
  }
  mountTodoStatusMenuInDualHost();
  openTodoStatusMenu(anchorEl, 0, 0);
  if (todoStatusMenu && !todoStatusMenu.classList.contains("hidden")) {
    positionMatrixDualStatusMenu();
    return;
  }
  unmountTodoStatusMenuFromDualHost("Status non disponibile per questa attivita.");
}

function closeActivityModal() {
  if (!activityModal) return;
  const shouldCloseStatus = matrixDualPanelsActive;
  stopActivityModalDrag();
  activityModal.classList.add("hidden");
  activityModal.classList.remove("matrix-dual-open");
  activityModal.classList.remove("is-dragging-modal");
  resetActivityModalPosition();
  matrixDualPanelsActive = false;
  unmountTodoStatusMenuFromDualHost();
  if (shouldCloseStatus) closeTodoStatusMenu();
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
  const oreRaw = activityDurationInput ? String(activityDurationInput.value || "").trim() : "";
  const oreStimate = oreRaw === "" ? null : Number(oreRaw);
  if (oreRaw !== "" && (Number.isNaN(oreStimate) || oreStimate < 0)) {
    setStatus("Le ore devono essere un numero valido (>= 0).", "error");
    return;
  }
  const titles = Array.from(matrixState.editingTitles);
  if (!titles.length) {
    setStatus("Select at least one activity.", "error");
    return;
  }
  const firstTitle = titles[0];
  const hasAssente = isAssenteTitle(firstTitle);
  if (!isActivityAllowedForRisorsa(firstTitle, attivita.risorsa_id)) {
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
  const { error: updError } = await supabase
    .from("attivita")
    .update({
      titolo: firstTitle,
      descrizione: note || null,
      ore_stimate: oreStimate,
      ore_assenza: isAssenteTitle(firstTitle) ? assenzaOre || null : null,
      commessa_id: isAssenteTitle(firstTitle) ? null : attivita.commessa_id,
    })
    .eq("id", attivita.id);
  if (updError) {
    setStatus(`Update error: ${updError.message}`, "error");
    return;
  }
  setStatus("Activity updated.", "ok");
  closeActivityModal();
  await loadMatrixAttivita();
  await refreshTodoRealtimeSnapshot([attivita.commessa_id]);
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
    await refreshTodoRealtimeSnapshot([attivita.commessa_id]);
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
    await refreshTodoRealtimeSnapshot([attivita?.commessa_id]);
  });
}

async function commitResize() {
  const ctx = matrixState.resizing;
  if (!ctx) return;
  matrixState.resizing = null;
  clearMatrixResizeTargetCell(ctx);
  if (ctx.originalBar && !ctx.usesOriginalBarForResize) {
    ctx.originalBar.style.opacity = "0";
    ctx.originalBar.style.pointerEvents = "none";
  }
  await animateMatrixResizePreviewToAnchor(ctx);

  try {
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
    let telaioPlannedSynced = true;
    if (isTelaioActivityTitle(ctx.attivita.titolo) && ctx.attivita.commessa_id) {
      telaioPlannedSynced = await syncTelaioPlannedDateFromMatrix(ctx.attivita.commessa_id, newEnd);
      const commessa = findCommessaInState(ctx.attivita.commessa_id);
      await warnTelaioPlannedMismatch(
        ctx.attivita.commessa_id,
        commessa?.data_conferma_consegna_telaio || null,
        { telaioEndDay: startOfDay(newEnd) }
      );
    }
    if (!telaioPlannedSynced) {
      setStatus("Duration aggiornata, ma non ho potuto aggiornare la data telaio pianificata.", "error");
    } else {
      setStatus("Duration updated.", "ok");
    }
    await loadMatrixAttivita();
    await refreshTodoRealtimeSnapshot([ctx.attivita.commessa_id]);
  } finally {
    cleanupMatrixResizeVisual(ctx);
  }
}

function clearMatrixResizeTargetCell(ctx = matrixState.resizing) {
  if (!ctx?.targetCell) return;
  ctx.targetCell.classList.remove("resize-target-col");
  ctx.targetCell = null;
}

function setMatrixResizeTargetCell(ctx, cell) {
  if (!ctx || !cell) return;
  if (ctx.targetCell === cell) return;
  clearMatrixResizeTargetCell(ctx);
  cell.classList.add("resize-target-col");
  ctx.targetCell = cell;
}

function getMatrixResizeAnchoredMetrics(ctx) {
  if (!ctx?.layerRect || !ctx?.startCellRect) return null;
  const targetCell =
    ctx.targetCell ||
    (ctx.rowEl && ctx.targetDayKey ? ctx.rowEl.querySelector(`.matrix-cell[data-day="${ctx.targetDayKey}"]`) : null);
  if (!targetCell) return null;
  const targetRect = targetCell.getBoundingClientRect();
  const left = ctx.startCellRect.left - ctx.layerRect.left;
  const maxRight = ctx.layerRect.right - ctx.layerRect.left;
  const anchoredRight = Math.max(left, Math.min(targetRect.right - ctx.layerRect.left, maxRight));
  const width = Math.max(0, anchoredRight - left);
  return { left, width };
}

function animateMatrixResizePreviewToAnchor(ctx) {
  return new Promise((resolve) => {
    const bar = ctx?.previewBar;
    if (!bar) return resolve();
    const metrics = getMatrixResizeAnchoredMetrics(ctx);
    if (!metrics) return resolve();
    const currentWidth = Number.parseFloat(bar.style.width || "0");
    if (!Number.isFinite(currentWidth)) return resolve();
    if (Math.abs(currentWidth - metrics.width) < 0.5) {
      bar.style.left = `${metrics.left}px`;
      bar.style.width = `${metrics.width}px`;
      return resolve();
    }
    bar.style.left = `${metrics.left}px`;
    bar.style.transition = "width 120ms cubic-bezier(0.22, 1, 0.36, 1)";
    let settled = false;
    const done = () => {
      if (settled) return;
      settled = true;
      bar.removeEventListener("transitionend", done);
      bar.style.transition = "none";
      resolve();
    };
    bar.addEventListener("transitionend", done);
    window.setTimeout(done, 170);
    requestAnimationFrame(() => {
      bar.style.width = `${metrics.width}px`;
    });
  });
}

function cleanupMatrixResizeVisual(ctx) {
  if (!ctx) return;
  clearMatrixResizeTargetCell(ctx);
  if (ctx.previewBar && !ctx.usesOriginalBarForResize) {
    ctx.previewBar.remove();
    ctx.previewBar = null;
  }
  if (ctx.usesOriginalBarForResize && ctx.originalBar) {
    const s = ctx.originalBar.style;
    const prev = ctx.originalBarInitialStyle || {};
    s.position = prev.position || "";
    s.top = prev.top || "";
    s.left = prev.left || "";
    s.width = prev.width || "";
    s.height = prev.height || "";
    s.marginTop = prev.marginTop || "";
    s.gridColumn = prev.gridColumn || "";
    s.zIndex = prev.zIndex || "";
    s.transition = prev.transition || "";
  }
  if (ctx.originalBar) {
    ctx.originalBar.classList.remove("resizing");
    ctx.originalBar.style.removeProperty("opacity");
    ctx.originalBar.style.removeProperty("pointer-events");
  }
}

function resolveMatrixResizeTargetCell(ctx, pointerClientX) {
  if (!ctx) return null;
  const dayCells = Array.isArray(ctx.dayCells) && ctx.dayCells.length
    ? ctx.dayCells
    : Array.from((ctx.rowEl || document).querySelectorAll(".matrix-cell[data-day]"));
  if (!dayCells.length) return null;

  let startIdx = Number.isInteger(ctx.startCellIndex) ? ctx.startCellIndex : 0;
  startIdx = Math.max(0, Math.min(startIdx, dayCells.length - 1));
  let initialIdx = Number.isInteger(ctx.initialTargetIndex) ? ctx.initialTargetIndex : startIdx;
  initialIdx = Math.max(startIdx, Math.min(initialIdx, dayCells.length - 1));

  const grabX = Number.isFinite(ctx.dragStartX) ? ctx.dragStartX : pointerClientX;
  if (Math.abs(pointerClientX - grabX) < 6) {
    return dayCells[initialIdx];
  }

  // Shift to next day only after passing half of the next cell.
  let targetIdx = startIdx;
  for (let i = startIdx + 1; i < dayCells.length; i += 1) {
    const rect = dayCells[i].getBoundingClientRect();
    const halfX = rect.left + rect.width / 2;
    if (pointerClientX >= halfX) targetIdx = i;
    else break;
  }

  return dayCells[targetIdx];
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
  if (progressMeta) progressMeta.textContent = "Seleziona una o piu commesse.";
  if (progressList) progressList.innerHTML = "";
  if (progressResults) progressResults.innerHTML = "";
  progressSelectedIds.clear();
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
  if (!commessa || (Array.isArray(commessa) && !commessa.length)) {
    progressMeta.textContent = "Nessuna commessa selezionata.";
    progressList.innerHTML = "";
    if (progressTimelineDays) progressTimelineDays.innerHTML = "";
    return;
  }
  const selectedCommesse = Array.isArray(commessa) ? commessa : [commessa];
  const selectedCodes = selectedCommesse.map((c) => c.codice || "").filter(Boolean);
  progressMeta.textContent =
    selectedCommesse.length === 1
      ? `${selectedCommesse[0].codice}${selectedCommesse[0].titolo ? " - " + selectedCommesse[0].titolo : ""}`
      : `${selectedCommesse.length} commesse selezionate: ${selectedCodes.join(", ")}`;
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
    const grouped = new Map();
    sorted.forEach((a) => {
      const key = a.progress_commessa_code || "Commessa n/d";
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(a);
    });
    const orderedGroups = selectedCommesse
      .map((c) => c.codice || "Commessa n/d")
      .filter((code) => grouped.has(code))
      .concat(Array.from(grouped.keys()).filter((code) => !selectedCommesse.some((c) => (c.codice || "Commessa n/d") === code)));
    const showGroupTitle = orderedGroups.length > 1;
    let rowCursor = 1;
    orderedGroups.forEach((groupCode) => {
      const groupActivities = grouped.get(groupCode) || [];
      if (!groupActivities.length) return;
      if (showGroupTitle) {
        const header = document.createElement("div");
        header.className = "progress-group-title";
        header.textContent = groupCode;
        header.style.gridColumn = `1 / ${days.length + 1}`;
        header.style.gridRow = `${rowCursor}`;
        progressList.appendChild(header);
        rowCursor += 1;
      }
      const laneEnds = [];
      groupActivities.forEach((a) => {
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
        if (lane === laneEnds.length) laneEnds.push(endIndex);
        else laneEnds[lane] = endIndex;
        item.style.gridColumn = `${startIndex + 1} / ${endIndex + 2}`;
        item.style.gridRow = `${rowCursor + lane}`;
        progressList.appendChild(item);
      });
      rowCursor += Math.max(1, laneEnds.length);
      if (showGroupTitle) rowCursor += 1;
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

const progressSelectedIds = new Set();

async function loadProgressForSelection() {
  const ids = Array.from(progressSelectedIds);
  if (!ids.length) {
    renderProgressList([], []);
    return;
  }
  const selectedCommesse = (state.commesse || []).filter((c) => ids.includes(c.id));
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .in("commessa_id", ids)
    .order("data_inizio", { ascending: true });
  if (error) {
    setStatus(`Progress error: ${error.message}`, "error");
    return;
  }
  const codeById = new Map(selectedCommesse.map((c) => [String(c.id), c.codice || ""]));
  const rows = (data || []).map((row) => ({
    ...row,
    progress_commessa_code: codeById.get(String(row.commessa_id || "")) || "",
  }));
  renderProgressList(selectedCommesse, rows);
}

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
    if (progressSelectedIds.has(c.id)) btn.classList.add("active");
    btn.addEventListener("click", async () => {
      if (progressSelectedIds.has(c.id)) progressSelectedIds.delete(c.id);
      else progressSelectedIds.add(c.id);
      Array.from(progressResults.querySelectorAll(".progress-result")).forEach((el) => {
        el.classList.toggle("active", progressSelectedIds.has(el.dataset.id));
      });
      await loadProgressForSelection();
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
  matrixBootstrapSyncInProgress = true;
  state.session = session;
  matrixState.initialRenderStabilized = false;
  matrixState.initialRenderStabilizeAttempts = 0;
  cancelPendingMatrixInitialRender();
  debugLog(`syncSignedIn start: user=${session.user?.id || "n/a"}`);
  try {
    debugLog("syncSignedIn step: loadProfile");
    const ok = await loadProfile(session.user);
    if (!ok || !state.profile) {
      syncSignedOut({ showStatus: false });
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
    debugLog("syncSignedIn step: loadMachineTypes");
    await loadMachineTypes();
    calendarDate.value = formatDateInput(calendarState.date);
    debugLog("syncSignedIn step: loadAttivita");
    await loadAttivita();
    if (!MATRIX_MOVEMENT_ENABLED) {
      resetMatrixMovementBaseline();
    }
    matrixDate.value = formatDateInput(matrixState.date);
    debugLog("syncSignedIn step: loadMatrixAttivita");
    await loadMatrixAttivita({ deferRender: true });
    debugLog("syncSignedIn step: setupRealtime");
    await setupRealtime();
    matrixBootstrapSyncInProgress = false;
    scheduleMatrixRenderStabilized();
    startTodoVisualRefreshLoop();
    debugLog("syncSignedIn complete");
  } catch (err) {
    debugLog(`syncSignedIn error: ${err?.message || err}`);
    setStatus(`Login sync error: ${err?.message || err}`, "error");
  } finally {
    matrixBootstrapSyncInProgress = false;
    authSyncInProgress = false;
  }
}

function syncSignedOut(options = {}) {
  const showStatus = options.showStatus !== false;
  debugLog("syncSignedOut");
  matrixBootstrapSyncInProgress = false;
  stopTodoVisualRefreshLoop();
  state.session = null;
  state.profile = null;
  state.commesse = [];
  state.reportActivitiesMap = new Map();
  state.todoOverridesMap = new Map();
  state.todoClientFlowMap = new Map();
  state.todoClientFlowLoadedCommesse = new Set();
  state.machineTypes = DEFAULT_MACHINE_TYPES.slice();
  matrixState.initialRenderStabilized = false;
  matrixState.initialRenderStabilizeAttempts = 0;
  cancelPendingMatrixInitialRender();
  authStatus.textContent = "Non autenticato";
  logoutBtn.classList.add("hidden");
  if (setPasswordBtn) setPasswordBtn.classList.add("hidden");
  if (openResourcesBtn) openResourcesBtn.classList.add("hidden");
  if (permissionsLink) permissionsLink.classList.add("hidden");
  if (reportsLink) reportsLink.classList.add("hidden");
  if (authActions) authActions.classList.remove("is-authenticated");
  refreshMachineTypeSelects({
    newType: "",
    newVariant: "standard",
    detailType: "",
    detailVariant: "standard",
  });
  if (authDot) authDot.classList.add("hidden");
  setRoleBadge("");
  setWriteAccess(false);
  commesseList.innerHTML = "";
  if (commesseSeqAlert) {
    commesseSeqAlert.textContent = "";
    commesseSeqAlert.classList.add("hidden");
    commesseSeqAlert.classList.remove("error");
  }
  clearSelection();
  calendarGrid.innerHTML = "";
  matrixGrid.innerHTML = "";
  if (showStatus) setStatus("Session ended.");
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
  refreshDetailTelaioOrderAccess();
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
  refreshDetailTelaioOrderAccess();
  updateDetailFieldCompletionUI();
  setDetailSnapshot();
  if (repartiList) repartiList.innerHTML = "";
  if (commessaDetailModal) commessaDetailModal.classList.add("hidden");
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  return dateStr.split("T")[0];
}

function normalizeMachineTypeValue(value) {
  return String(value || "").trim();
}

function normalizeMachineTypeLookup(value) {
  return normalizeMachineTypeValue(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "");
}

function getMachineTypeDisplayLabel(value) {
  const raw = normalizeMachineTypeValue(value) || "Altro tipo";
  const key = normalizeMachineTypeLookup(raw);
  if (key === "miniboosterswiss" || key === "miniboosterswss") return "MBS";
  return raw;
}

function normalizeMachineVariantValue(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized || "standard";
}

function supportsCustomVariantForMachineType(typeValue) {
  const key = normalizeMachineTypeLookup(typeValue);
  if (!key) return false;
  return CUSTOM_VARIANT_MACHINE_TYPES.some(
    (item) => normalizeMachineTypeLookup(item) === key
  );
}

function sanitizeMachineVariantForType(typeValue, variantValue) {
  const normalizedVariant = normalizeMachineVariantValue(variantValue || "standard");
  if (normalizedVariant === "custom" && !supportsCustomVariantForMachineType(typeValue)) {
    return "standard";
  }
  return normalizedVariant;
}

function collectMachineTypeOptions() {
  const list = [];
  const seen = new Set();
  const add = (raw) => {
    const value = normalizeMachineTypeValue(raw);
    if (!value) return;
    const key = value.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    list.push(value);
  };
  (state.machineTypes || []).forEach(add);
  DEFAULT_MACHINE_TYPES.forEach(add);
  (state.commesse || []).forEach((row) => add(row?.tipo_macchina));
  add("Altro tipo");
  return list;
}

function setMachineTypeSelectOptions(selectEl, options, preferredValue = "", includePlaceholder = false) {
  if (!selectEl) return;
  const currentValue = normalizeMachineTypeValue(selectEl.value);
  const desiredValue = normalizeMachineTypeValue(preferredValue || currentValue);
  const values = Array.isArray(options) ? options.slice() : [];
  if (desiredValue && !values.some((item) => item.toLowerCase() === desiredValue.toLowerCase())) {
    values.push(desiredValue);
  }

  selectEl.innerHTML = "";
  if (includePlaceholder) {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Seleziona tipologia...";
    selectEl.appendChild(placeholder);
  }
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });
  if (desiredValue) {
    selectEl.value = desiredValue;
  } else if (includePlaceholder) {
    selectEl.value = "";
  } else if (values.length) {
    selectEl.value = values[0];
  }
}

function setMachineVariantSelectOptions(selectEl, machineTypeValue = "", preferredValue = "") {
  if (!selectEl) return;
  const currentValue = normalizeMachineVariantValue(selectEl.value || "standard");
  const desiredValue = sanitizeMachineVariantForType(
    machineTypeValue,
    preferredValue || currentValue
  );
  const baseOptions = [{ value: "standard", label: "Standard" }];
  if (supportsCustomVariantForMachineType(machineTypeValue)) {
    baseOptions.push({ value: "custom", label: "Custom" });
  }
  if (!baseOptions.some((item) => item.value === desiredValue)) {
    const label = desiredValue.charAt(0).toUpperCase() + desiredValue.slice(1);
    baseOptions.push({ value: desiredValue, label });
  }
  selectEl.innerHTML = "";
  baseOptions.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    selectEl.appendChild(option);
  });
  selectEl.value = desiredValue;
}

function refreshMachineTypeSelects(options = {}) {
  const typeOptions = collectMachineTypeOptions();
  const nextNewType = options.newType ?? n.tipo_macchina?.value ?? "";
  const nextDetailType = options.detailType ?? d.tipo_macchina?.value ?? "";
  setMachineTypeSelectOptions(
    n.tipo_macchina,
    typeOptions,
    nextNewType,
    true
  );
  setMachineTypeSelectOptions(
    d.tipo_macchina,
    typeOptions,
    nextDetailType,
    true
  );
  setMachineVariantSelectOptions(
    n.variante_macchina,
    nextNewType,
    options.newVariant ?? n.variante_macchina?.value ?? "standard"
  );
  setMachineVariantSelectOptions(
    d.variante_macchina,
    nextDetailType,
    options.detailVariant ?? d.variante_macchina?.value ?? "standard"
  );
}

async function loadMachineTypes() {
  if (!state.session) {
    state.machineTypes = DEFAULT_MACHINE_TYPES.slice();
    refreshMachineTypeSelects();
    return;
  }
  let rows = null;
  let error = null;
  if (!PREFER_REST_ON_RELOAD) {
    try {
      const result = await withTimeout(
        supabase.from("motore_tipologie").select("nome").order("nome", { ascending: true }),
        FETCH_TIMEOUT_MS
      );
      rows = result.data;
      error = result.error;
    } catch (err) {
      error = err;
    }
  }
  if (!rows || error) {
    rows = await fetchTableViaRest("motore_tipologie", "select=nome&order=nome.asc");
  }
  if (!rows || !Array.isArray(rows) || !rows.length) {
    state.machineTypes = DEFAULT_MACHINE_TYPES.slice();
    refreshMachineTypeSelects();
    return;
  }
  state.machineTypes = rows
    .map((row) => normalizeMachineTypeValue(row?.nome))
    .filter(Boolean);
  refreshMachineTypeSelects();
}

function openImportModal() {
  if (!state.canWrite) return;
  importModal.classList.remove("hidden");
  resetImportFeedback();
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
  refreshMachineTypeSelects({
    newType: "",
    newVariant: "standard",
  });
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
      const plannedRaw = c.data_conferma_consegna_telaio || "";
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
  matrixState.date = normalizeMatrixStartDate(date);
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

async function copyTextToClipboard(text) {
  const payload = String(text || "");
  if (!payload) return false;
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(payload);
    return true;
  }
  const temp = document.createElement("textarea");
  temp.value = payload;
  temp.setAttribute("readonly", "");
  temp.style.position = "fixed";
  temp.style.opacity = "0";
  temp.style.pointerEvents = "none";
  document.body.appendChild(temp);
  temp.focus();
  temp.select();
  const ok = document.execCommand("copy");
  temp.remove();
  if (!ok) throw new Error("Copy command not available");
  return true;
}

async function handleCopyImportHeader(text, label) {
  try {
    await copyTextToClipboard(text);
    showToast(`Header copiato: ${label}.`, "ok");
  } catch (err) {
    setStatus(`Copia header non riuscita: ${err?.message || err}`, "error");
    showToast("Copia header non riuscita.", "error");
  }
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
      local: BACKUP_LOCAL_STORAGE_KEYS.reduce((acc, key) => {
        try {
          const value = localStorage.getItem(key);
          if (value != null) acc[key] = value;
        } catch (_err) {}
        return acc;
      }, {}),
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
  const local = payload?.local && typeof payload.local === "object" ? payload.local : {};
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
    const localCount = Object.keys(local).length;
    if (localCount) return `Backup locale trovato (${localCount} chiavi), ma nessuna tabella dati valida.`;
    return "No usable tables found in the backup.";
  }
  const localCount = Object.keys(local).length;
  return `Tables: ${lines.join(", ")}. Total rows: ${totalRows}. Local keys: ${localCount}.`;
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
    const local = backupPayload?.local;
    if (local && typeof local === "object") {
      BACKUP_LOCAL_STORAGE_KEYS.forEach((key) => {
        if (!Object.prototype.hasOwnProperty.call(local, key)) return;
        try {
          const value = local[key];
          if (value == null) localStorage.removeItem(key);
          else localStorage.setItem(key, String(value));
        } catch (_err) {}
      });
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
  resetImportFeedback();
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

function resetImportFeedback() {
  if (importProgressWrap) importProgressWrap.classList.add("hidden");
  if (importProgressBar) importProgressBar.style.width = "0%";
  if (importProgressText) importProgressText.textContent = "0%";
  if (importResult) {
    importResult.classList.add("hidden");
    importResult.classList.remove("ok", "error");
    importResult.textContent = "";
  }
}

function updateImportProgress(percent = 0, note = "") {
  const safe = Math.max(0, Math.min(100, Number(percent) || 0));
  if (importProgressWrap) importProgressWrap.classList.remove("hidden");
  if (importProgressBar) importProgressBar.style.width = `${safe}%`;
  if (importProgressText) {
    importProgressText.textContent = note ? `${safe}% · ${note}` : `${safe}%`;
  }
}

function showImportResult(message, tone = "ok") {
  if (!importResult) return;
  importResult.classList.remove("hidden", "ok", "error");
  importResult.classList.add(tone === "error" ? "error" : "ok");
  importResult.textContent = message;
}

function parseImportRows(text) {
  const cleaned = text.trim();
  if (!cleaned) return { rows: [], errors: [], total: 0 };
  const delimiter = detectDelimiter(cleaned);
  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim() !== "");
  let startIndex = 0;
  if (lines.length) {
    const first = lines[0].toLowerCase();
    if (first.includes("anno") || first.includes("numero") || first.includes("descr") || first.includes("tipo")) {
      startIndex = 1;
    }
  }

  const rows = [];
  const errors = [];
  const expectedTypes = Array.from(
    new Set(
      [...(state.machineTypes || []), ...DEFAULT_MACHINE_TYPES, "Altro tipo"]
        .map((value) => normalizeMachineTypeValue(value))
        .filter(Boolean)
    )
  );
  const expectedTypeMap = new Map(
    expectedTypes.map((typeValue) => [normalizeMachineTypeLookup(typeValue), typeValue])
  );
  const expectedTypesLabel = expectedTypes.join(", ");
  const parseMachineImportCell = (rawValue) => {
    const raw = normalizeMachineTypeValue(rawValue || "");
    if (!raw) {
      return { ok: false, reason: "missing" };
    }
    const hasSpecialTag = /\b(spec|speciale|spaciale|special)\b/i.test(raw);
    const cleanedType = raw
      .replace(/\b(spec|speciale|spaciale|special)\b/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!cleanedType) {
      return { ok: false, reason: "missing" };
    }
    const canonicalType = expectedTypeMap.get(normalizeMachineTypeLookup(cleanedType)) || "";
    if (!canonicalType) {
      return { ok: false, reason: "unknown", value: cleanedType };
    }
    if (hasSpecialTag && !supportsCustomVariantForMachineType(canonicalType)) {
      return { ok: false, reason: "spec_not_supported", value: canonicalType };
    }
    return {
      ok: true,
      type: canonicalType,
      variant: hasSpecialTag ? "custom" : "standard",
    };
  };

  for (let i = startIndex; i < lines.length; i++) {
    const cols = lines[i]
      .split(delimiter)
      .map((c) => c.trim().replace(/^"(.*)"$/, "$1"));

    if (cols.length < 4) {
      errors.push(`Riga ${i + 1}: manca tipo_macchina (4a colonna obbligatoria).`);
      continue;
    }

    const annoRaw = cols[0];
    const numeroRaw = cols[1];
    const descrizione = String(cols[2] || "").trim();
    const descrizioneKey = descrizione
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
    const skipDescrizione = !descrizione || /^(0+|zero|null|nullo)$/.test(descrizioneKey);
    if (skipDescrizione) {
      continue;
    }
    const machineCell = parseMachineImportCell(cols[3] || "");

    if (!/^\d{4}$/.test(annoRaw)) {
      errors.push(`Riga ${i + 1}: anno non valido`);
      continue;
    }
    if (!/^\d+$/.test(numeroRaw)) {
      errors.push(`Riga ${i + 1}: numero non valido`);
      continue;
    }
    if (!machineCell.ok) {
      if (machineCell.reason === "missing") {
        errors.push(`Riga ${i + 1}: tipo_macchina mancante. Tipologie attese: ${expectedTypesLabel}`);
      } else if (machineCell.reason === "spec_not_supported") {
        errors.push(`Riga ${i + 1}: "${machineCell.value}" non supporta variante SPEC/custom.`);
      } else {
        errors.push(
          `Riga ${i + 1}: tipo_macchina "${machineCell.value}" non riconosciuto. Tipologie attese: ${expectedTypesLabel}`
        );
      }
      continue;
    }

    const anno = Number(annoRaw);
    const numero = Number(numeroRaw);
    const codice = `${anno}_${numero}`;
    const titolo = descrizione || null;
    const variante_macchina = machineCell.variant;

    rows.push({
      codice,
      titolo,
      anno,
      numero,
      tipo_macchina: machineCell.type,
      variante_macchina,
    });
  }

  return { rows, errors, total: lines.length - startIndex };
}

function updateImportPreview() {
  resetImportFeedback();
  const { rows, errors, total } = parseImportRows(importTextarea.value);
  const sample = rows
    .slice(0, 3)
    .map(
      (r) =>
        `${r.codice} - ${r.titolo || "Senza titolo"} - ${r.tipo_macchina || "Altro tipo"} (${r.variante_macchina || "standard"})`
    );
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
    if (row.stimato && commessa.data_conferma_consegna_telaio) {
      const existing = formatDate(commessa.data_conferma_consegna_telaio);
      const incoming = formatDateInput(row.stimato);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (ordine telaio stimato)`);
      }
    }
    if (row.effettivo && commessa.data_consegna_telaio_effettiva) {
      const existing = formatDate(commessa.data_consegna_telaio_effettiva);
      const incoming = formatDateInput(row.effettivo);
      if (existing && incoming && existing !== incoming) {
        overwrites.push(`${key} (ordine telaio effettivo)`);
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
    if (
      row.ingresso ||
      row.ordine ||
      row.stimato ||
      row.effettivo ||
      row.prelievo ||
      row.consegna ||
      row.ordinato != null
    ) {
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
  const hasEffettivo = rows.some((r) => r.effettivoRaw);
  const headerCols = [
    "Codice",
    ...(hasIngresso ? ["Ingresso ordine"] : []),
    "Ordine telaio target",
    "Ordine telaio stimato",
    ...(hasPrelievo ? ["Prelievo materiali"] : []),
    ...(hasOrdinato ? ["Ordinato"] : []),
    ...(hasEffettivo ? ["Ordine telaio effettivo"] : []),
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
        ...(hasEffettivo ? [row.effettivoRaw || ""] : []),
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
    let payload = {};
    if (row.ingresso) payload.data_ingresso = formatDateInput(row.ingresso);
    if (row.ordine) payload.data_ordine_telaio = formatDateInput(row.ordine);
    if (row.stimato) payload.data_conferma_consegna_telaio = formatDateInput(row.stimato);
    if (row.effettivo) payload.data_consegna_telaio_effettiva = formatDateInput(row.effettivo);
    if (row.prelievo) payload.data_prelievo = formatDateInput(row.prelievo);
    if (row.ordinato != null) {
      payload.telaio_consegnato = row.ordinato;
      if (!row.ordinato) payload.data_consegna_telaio_effettiva = null;
    }
    if (row.consegna) payload.data_consegna_macchina = formatDateInput(row.consegna);
    const consistency = ensureTelaioOrderPayloadConsistency(commessa.id, payload, { notify: false });
    if (!consistency.ok) {
      skipped.push(`${key} ordinato senza data telaio`);
      return;
    }
    payload = consistency.payload;
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
        if (col.includes("effett")) {
          map.effettivo = index;
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
    let effettivoCell = "";
    let prelievoCell = "";
    let ordinatoCell = "";
    let consegnaCell = "";

    if (headerMap) {
      ingressoCell = headerMap.ingresso != null ? cols[headerMap.ingresso] || "" : "";
      ordineCell = headerMap.ordine != null ? cols[headerMap.ordine] || "" : "";
      stimatoCell = headerMap.stimato != null ? cols[headerMap.stimato] || "" : "";
      effettivoCell = headerMap.effettivo != null ? cols[headerMap.effettivo] || "" : "";
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
        if (hasFlag && cols.length >= 8) {
          effettivoCell = cols[6] || "";
          consegnaCell = cols[7] || "";
        } else {
          consegnaCell = hasFlag ? cols[6] || "" : cols[5] || "";
        }
      } else if (hasIngresso) {
        ordinatoCell = hasFlag ? cols[4] || "" : "";
        if (hasFlag && cols.length >= 7) {
          effettivoCell = cols[5] || "";
          consegnaCell = cols[6] || "";
        } else {
          consegnaCell = hasFlag ? cols[5] || "" : cols[4] || "";
        }
      } else {
        ordinatoCell = hasFlag ? cols[3] || "" : "";
        if (hasFlag && cols.length >= 7) {
          effettivoCell = cols[4] || "";
          consegnaCell = cols[5] || "";
          prelievoCell = cols[6] || "";
        } else if (hasFlag && cols.length >= 6) {
          effettivoCell = cols[4] || "";
          consegnaCell = cols[5] || "";
        } else {
          consegnaCell = hasFlag
            ? cols[4] || ""
            : cols.length >= 4
            ? cols[3] || ""
            : cols[2] || "";
        }
        if (!hasFlag && cols.length >= 6) {
          prelievoCell = cols[5] || "";
        } else if (!hasFlag && cols.length >= 5) {
          prelievoCell = cols[4] || "";
        }
      }
    }
    const ingressoParsed = parseDateCell(ingressoCell);
    const ordineParsed = parseDateCell(ordineCell);
    const stimatoParsed = parseDateCell(stimatoCell);
    const effettivoParsed = parseDateCell(effettivoCell);
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
    if (effettivoCell && effettivoParsed.reason === "invalid") {
      warnings.push(`Riga ${i + 1}: data ordine telaio effettivo non valida`);
    } else if (effettivoCell && effettivoParsed.reason === "empty") {
      warnings.push(`Riga ${i + 1}: data ordine telaio effettivo vuota/zero`);
    } else if (effettivoCell && effettivoParsed.reason === "range") {
      warnings.push(`Riga ${i + 1}: data ordine telaio effettivo fuori range (2000-2100)`);
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
    if (ordinatoParsed.value === true && !effettivoParsed.date) {
      errors.push(`Riga ${i + 1}: ordinato=VERO richiede ordine_telaio_effettivo valido`);
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
      effettivo: effettivoParsed.date,
      prelievo: prelievoParsed.date,
      ordinato: ordinatoParsed.value,
      consegna: consegnaParsed.date,
      ingressoReason: ingressoParsed.reason,
      ordineReason: ordineParsed.reason,
      stimatoReason: stimatoParsed.reason,
      effettivoReason: effettivoParsed.reason,
      prelievoReason: prelievoParsed.reason,
      ordinatoReason: ordinatoParsed.reason,
      consegnaReason: consegnaParsed.reason,
      ingressoRaw: ingressoCell,
      ordineRaw: ordineCell,
      stimatoRaw: stimatoCell,
      effettivoRaw: effettivoCell,
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

function getMatrixDefaultWeeks(view = matrixState.view) {
  if (view === "six") return 6;
  if (view === "three") return 3;
  if (view === "two") return 2;
  return 1;
}

function getMatrixWeeks() {
  const custom = Number(matrixState.customWeeks || 0);
  if (Number.isFinite(custom) && custom > 0) return Math.max(1, Math.round(custom));
  return getMatrixDefaultWeeks(matrixState.view);
}

function clampMatrixWeeks(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return MATRIX_MIN_WEEKS;
  return Math.min(MATRIX_MAX_WEEKS, Math.max(MATRIX_MIN_WEEKS, numeric));
}

function syncMatrixViewFromWeeks(weeks) {
  const normalized = clampMatrixWeeks(weeks);
  if (normalized <= 1) matrixState.view = "week";
  else if (normalized <= 2) matrixState.view = "two";
  else if (normalized <= 3) matrixState.view = "three";
  else matrixState.view = "six";
}

function normalizeMatrixStartDate(date) {
  const candidate = date ? new Date(date) : new Date();
  const safeDate = Number.isNaN(candidate.getTime()) ? new Date() : candidate;
  const day = startOfDay(safeDate);
  return normalizeBusinessDay(day, 1);
}

function getMatrixVisibleBusinessDays() {
  return Math.max(1, Math.round(getMatrixWeeks() * 5));
}

function getMatrixPanBufferBusinessDays(visibleBusinessDays = getMatrixVisibleBusinessDays()) {
  if (!MATRIX_MOVEMENT_ENABLED) return 0;
  const visible = Math.max(1, Math.round(Number(visibleBusinessDays) || 1));
  // Buffer laterale ampio per evitare "muro" visivo durante il pan.
  return Math.max(8, Math.min(35, Math.round(visible * 0.8)));
}

function getMatrixRenderWindow(baseDate) {
  const renderStart = startOfWeek(baseDate);
  const weeks = getMatrixWeeks();
  const days = [];
  for (let i = 0; i < weeks; i += 1) {
    days.push(...weekdayDates(addDays(renderStart, i * 7)));
  }
  const visibleBusinessDays = days.length;
  return {
    days,
    renderStart,
    visibleStart: renderStart,
    visibleBusinessDays,
    bufferBusinessDays: 0,
    totalBusinessDays: visibleBusinessDays,
  };
}

function weekDaysForMatrix(baseDate) {
  return getMatrixRenderWindow(baseDate).days;
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
      let payload = {};
      if (field === "ordine") payload.data_ordine_telaio = formatDateInput(date);
      if (field === "prelievo") payload.data_prelievo = formatDateInput(date);
      if (field === "kit_cavi") payload.data_arrivo_kit_cavi = formatDateInput(date);
      if (field === "consegna") payload.data_consegna_macchina = formatDateInput(date);
      if (field === "ordine_pianificato") payload.data_conferma_consegna_telaio = formatDateInput(date);
      const consistency = ensureTelaioOrderPayloadConsistency(commessaId, payload, { notify: false });
      payload = consistency.payload;
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
        d.data_conferma_consegna_telaio
      ) {
        d.data_conferma_consegna_telaio.value = payload.data_conferma_consegna_telaio || "";
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
        if (!canEditTelaioOrderByRole()) {
          const msg = "Solo admin e responsabile possono modificare lo stato ordine telaio.";
          setStatus(msg, "error");
          showToast(msg, "error");
          return;
        }
        const commessaId = milestonePickerState.commessaId;
        if (!commessaId) return;
        const nextState = !isOrdered;
        if (!nextState) {
          const confirmed = await openConfirmModal(
            "Il telaio risulta ORDINATO. Passare a NON ORDINATO e azzerare la data ordine?",
            null,
            {
              okLabel: "Conferma",
              cancelLabel: "Annulla",
            }
          );
          if (!confirmed) return;
        }
        let payload = { telaio_consegnato: nextState };
        if (nextState && milestonePickerState.selected) {
          payload.data_consegna_telaio_effettiva = formatDateInput(milestonePickerState.selected);
        }
        if (!nextState) {
          payload.data_consegna_telaio_effettiva = null;
        }
        const consistency = ensureTelaioOrderPayloadConsistency(commessaId, payload, {
          focusDateField: true,
          toast: true,
        });
        if (!consistency.ok) return;
        payload = consistency.payload;
        const { error } = await supabase
          .from("commesse")
          .update(payload)
          .eq("id", commessaId);
        if (error) {
          setStatus(`Update error: ${error.message}`, "error");
          return;
        }
        updateCommessaInState(commessaId, payload);
        if (
          state.selected?.id === commessaId &&
          d.data_consegna_telaio_effettiva &&
          Object.prototype.hasOwnProperty.call(payload, "data_consegna_telaio_effettiva")
        ) {
          d.data_consegna_telaio_effettiva.value = payload.data_consegna_telaio_effettiva || "";
        }
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
  let payload = {};
  if (drag.field === "ordine") payload.data_ordine_telaio = formatDateInput(newDate);
  if (drag.field === "kit_cavi") payload.data_arrivo_kit_cavi = formatDateInput(newDate);
  if (drag.field === "prelievo") payload.data_prelievo = formatDateInput(newDate);
  if (drag.field === "consegna") payload.data_consegna_macchina = formatDateInput(newDate);
  if (drag.field === "ordine_pianificato") payload.data_conferma_consegna_telaio = formatDateInput(newDate);
  const consistency = ensureTelaioOrderPayloadConsistency(drag.commessaId, payload, { notify: false });
  payload = consistency.payload;
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
    d.data_conferma_consegna_telaio
  ) {
    d.data_conferma_consegna_telaio.value = payload.data_conferma_consegna_telaio || "";
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
  const weeks = getMatrixWeeks();
  const end = endOfDay(addDays(start, weeks * 7 - 1));
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

function getAttivitaRepartoName(attivita) {
  if (!attivita) return "";
  if (attivita.risorsa_id) return getRisorsaDeptName(attivita.risorsa_id);
  if (attivita.reparto_id != null) {
    const reparto = (state.reparti || []).find((rep) => String(rep.id) === String(attivita.reparto_id));
    return reparto ? reparto.nome : "";
  }
  return "";
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

const TODO_ACTIVITY_PRIMARY_PHASE_KEY = "progettazione termodinamica";
const TODO_CAD_ACTIVITY_ORDER_KEYS = [
  "preliminare",
  "3d",
  "carenatura",
  "telaio",
  "costruttivi 1-5",
  "consegna totale",
];

function sortActivitiesByTodoOrder(names, repartoName = "") {
  const list = Array.isArray(names) ? names.slice() : [];
  const deptKey = normalizeDeptKey(repartoName || "");
  if (deptKey.includes("CAD")) {
    return list.sort((a, b) => {
      const aKey = normalizePhaseKey(a);
      const bKey = normalizePhaseKey(b);
      const aIdx = TODO_CAD_ACTIVITY_ORDER_KEYS.indexOf(aKey);
      const bIdx = TODO_CAD_ACTIVITY_ORDER_KEYS.indexOf(bKey);
      const aRank = aIdx >= 0 ? aIdx : Number.MAX_SAFE_INTEGER;
      const bRank = bIdx >= 0 ? bIdx : Number.MAX_SAFE_INTEGER;
      if (aRank !== bRank) return aRank - bRank;
      return String(a || "").localeCompare(String(b || ""));
    });
  }
  return list.sort((a, b) => {
    const aTarget = normalizePhaseKey(a) === TODO_ACTIVITY_PRIMARY_PHASE_KEY ? 0 : 1;
    const bTarget = normalizePhaseKey(b) === TODO_ACTIVITY_PRIMARY_PHASE_KEY ? 0 : 1;
    if (aTarget !== bTarget) return aTarget - bTarget;
    return String(a || "").localeCompare(String(b || ""));
  });
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

function getActivityDisplayLabel(name) {
  const raw = String(name || "").trim();
  if (!raw) return raw;
  if (!isLavenderActivity(raw) && !isAssenteTitle(raw)) return raw;
  return raw
    .toLocaleLowerCase("it-IT")
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toLocaleUpperCase("it-IT") + part.slice(1))
    .join(" ");
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

async function clearTodoNotNeededOverrideSilently(commessaId, titolo) {
  if (!commessaId || !titolo) return true;
  const { error } = await supabase
    .from("commessa_attivita_override")
    .delete()
    .eq("commessa_id", commessaId)
    .eq("titolo", titolo);
  if (error) return false;
  const key = String(commessaId);
  const list = state.todoOverridesMap.get(key) || new Map();
  list.delete(normalizePhaseKey(titolo));
  state.todoOverridesMap.set(key, list);
  return true;
}

async function syncTelaioPlannedDateFromMatrix(commessaId, endDate) {
  if (!commessaId || !endDate) return true;
  const day = startOfDay(endDate);
  if (Number.isNaN(day.getTime())) return false;
  const nextIso = formatDateInput(day);
  const current = findCommessaInState(commessaId);
  const currentIso = formatIsoDateOnly(current?.data_conferma_consegna_telaio || null);
  if (currentIso === nextIso) return true;

  const { error } = await supabase
    .from("commesse")
    .update({ data_conferma_consegna_telaio: nextIso })
    .eq("id", commessaId);
  if (error) return false;

  updateCommessaInState(commessaId, { data_conferma_consegna_telaio: nextIso });
  if (state.selected?.id === commessaId && d.data_conferma_consegna_telaio) {
    d.data_conferma_consegna_telaio.value = nextIso;
  }
  return true;
}

function animateMatrixMoveGhost(sourceBar, targetCell, durationDays = 1, options = {}) {
  return new Promise((resolve) => {
    if (!targetCell) return resolve();
    const followerGhost =
      options.useExistingFollower !== false && matrixState.dragFollowerEl && matrixState.dragFollowerEl.isConnected
        ? matrixState.dragFollowerEl
        : null;
    const sourceEl = followerGhost || sourceBar;
    if (!sourceEl) return resolve();
    const sourceRect = sourceEl.getBoundingClientRect();
    if (!sourceRect.width || !sourceRect.height) return resolve();

    const targetRow = targetCell.closest(".matrix-row");
    if (!targetRow) return resolve();
    const dayCells = Array.from(targetRow.querySelectorAll(".matrix-cell[data-day]"));
    const startIdx = dayCells.indexOf(targetCell);
    if (startIdx < 0) return resolve();
    const spanDays = Math.max(1, Number(durationDays) || 1);
    const endIdx = Math.min(dayCells.length - 1, startIdx + spanDays - 1);
    const startCell = dayCells[startIdx];
    const endCell = dayCells[endIdx] || startCell;
    if (!startCell || !endCell) return resolve();

    const startCellRect = startCell.getBoundingClientRect();
    const endCellRect = endCell.getBoundingClientRect();
    const targetLayer = targetRow.querySelector(".matrix-bar-layer");
    const targetLayerRect = targetLayer ? targetLayer.getBoundingClientRect() : null;
    const targetLeft = startCellRect.left;
    const targetWidth = Math.max(8, endCellRect.right - startCellRect.left);
    const targetTopBase = targetLayerRect ? targetLayerRect.top : startCellRect.top;
    const targetTop = targetTopBase + 6;

    const dx = Math.abs(targetLeft - sourceRect.left);
    const dy = Math.abs(targetTop - sourceRect.top);
    const dw = Math.abs(targetWidth - sourceRect.width);
    if (dx < 0.5 && dy < 0.5 && dw < 0.5) return resolve();

    let ghost = followerGhost;
    if (!ghost) {
      const dropClientX = Number(options.dropClientX);
      const dropClientY = Number(options.dropClientY);
      const pointerOffsetX = Number(options.pointerOffsetX);
      const pointerOffsetY = Number(options.pointerOffsetY);
      const hasDropPoint = Number.isFinite(dropClientX) && Number.isFinite(dropClientY);
      const hasOffsets = Number.isFinite(pointerOffsetX) && Number.isFinite(pointerOffsetY);
      const startLeft = hasDropPoint
        ? dropClientX - (hasOffsets ? pointerOffsetX : sourceRect.width / 2)
        : sourceRect.left;
      const startTop = hasDropPoint
        ? dropClientY - (hasOffsets ? pointerOffsetY : sourceRect.height / 2)
        : sourceRect.top;
      ghost = sourceBar.cloneNode(true);
      ghost.classList.remove("dragging", "resizing");
      ghost.style.position = "fixed";
      ghost.style.left = `${startLeft}px`;
      ghost.style.top = `${startTop}px`;
      ghost.style.width = `${sourceRect.width}px`;
      ghost.style.height = `${sourceRect.height}px`;
      ghost.style.marginTop = "0";
      ghost.style.zIndex = "2400";
      ghost.style.pointerEvents = "none";
      document.body.appendChild(ghost);
    } else {
      ghost.style.left = `${sourceRect.left}px`;
      ghost.style.top = `${sourceRect.top}px`;
      ghost.style.width = `${sourceRect.width}px`;
      ghost.style.height = `${sourceRect.height}px`;
    }
    ghost.style.transition =
      "left 170ms cubic-bezier(0.22, 1, 0.36, 1), top 170ms cubic-bezier(0.22, 1, 0.36, 1), width 170ms cubic-bezier(0.22, 1, 0.36, 1), opacity 120ms ease";
    ghost.style.willChange = "left, top, width";

    const holdAtTarget = options.holdAtTarget === true;

    let settled = false;
    let removed = false;
    const removeGhost = () => {
      if (removed) return;
      removed = true;
      ghost.removeEventListener("transitionend", done);
      if (matrixState.dragFollowerEl === ghost) {
        removeMatrixDragFollower();
      } else {
        ghost.remove();
      }
    };
    const releaseGhost = ({ immediate = false, anchorEl = null, onAfterRemove = null } = {}) => {
      const runAfterRemove = () => {
        if (typeof onAfterRemove === "function") {
          try {
            onAfterRemove();
          } catch (_err) {
            // Ignore callback errors to keep drag flow stable.
          }
        }
      };
      if (removed) return;
      if (immediate) {
        removeGhost();
        runAfterRemove();
        return;
      }
      const anchorRect = anchorEl?.getBoundingClientRect?.();
      const hasAnchorRect =
        anchorRect &&
        Number.isFinite(anchorRect.left) &&
        Number.isFinite(anchorRect.top) &&
        Number.isFinite(anchorRect.width) &&
        Number.isFinite(anchorRect.height) &&
        anchorRect.width > 0 &&
        anchorRect.height > 0;
      if (hasAnchorRect) {
        ghost.style.transition =
          "left 110ms cubic-bezier(0.22, 1, 0.36, 1), top 110ms cubic-bezier(0.22, 1, 0.36, 1), width 110ms cubic-bezier(0.22, 1, 0.36, 1), height 110ms cubic-bezier(0.22, 1, 0.36, 1), opacity 120ms ease";
        ghost.style.left = `${anchorRect.left}px`;
        ghost.style.top = `${anchorRect.top}px`;
        ghost.style.width = `${anchorRect.width}px`;
        ghost.style.height = `${anchorRect.height}px`;
        window.setTimeout(() => {
          if (removed) return;
          ghost.style.transition = "opacity 120ms ease";
          ghost.style.opacity = "0";
          window.setTimeout(() => {
            removeGhost();
            runAfterRemove();
          }, 150);
        }, 115);
        return;
      }
      ghost.style.transition = "opacity 120ms ease";
      ghost.style.opacity = "0";
      window.setTimeout(() => {
        removeGhost();
        runAfterRemove();
      }, 150);
    };
    const done = () => {
      if (settled) return;
      settled = true;
      if (holdAtTarget) {
        resolve(releaseGhost);
        return;
      }
      removeGhost();
      resolve(null);
    };
    ghost.addEventListener("transitionend", done);
    window.setTimeout(done, 240);

    requestAnimationFrame(() => {
      ghost.style.left = `${targetLeft}px`;
      ghost.style.top = `${targetTop}px`;
      ghost.style.width = `${targetWidth}px`;
    });
  });
}

async function getExistingTelaioAssignments(commessaId) {
  if (!commessaId) return [];
  const commessaKey = String(commessaId);
  const risorsaById = new Map((state.risorse || []).map((r) => [String(r.id), r.nome || `Risorsa ${r.id}`]));
  let rows = [];
  try {
    const { data, error } = await supabase
      .from("attivita")
      .select("id,commessa_id,titolo,risorsa_id,data_inizio,stato")
      .eq("commessa_id", commessaId);
    if (!error && Array.isArray(data)) rows = data;
  } catch {
    // fallback locale
  }
  if (!rows.length) {
    rows = (matrixState.attivita || []).filter((row) => String(row?.commessa_id || "") === commessaKey);
  }
  return rows
    .filter((row) => isTelaioActivityTitle(row?.titolo || ""))
    .filter((row) => normalizePhaseKey(row?.stato || "") !== "annullata")
    .map((row) => ({
      id: row?.id || null,
      risorsaLabel: risorsaById.get(String(row?.risorsa_id || "")) || `Risorsa ${row?.risorsa_id || "n/d"}`,
      start: row?.data_inizio ? startOfDay(new Date(row.data_inizio)) : null,
    }))
    .sort((a, b) => {
      const at = a.start?.getTime?.() || Number.POSITIVE_INFINITY;
      const bt = b.start?.getTime?.() || Number.POSITIVE_INFINITY;
      return at - bt;
    });
}

async function confirmTelaioDuplicateAssignment(commessaId) {
  const existing = await getExistingTelaioAssignments(commessaId);
  if (!existing.length) return true;
  const first = existing[0];
  const dayLabel = first.start ? formatDateDMY(first.start) : "data n/d";
  const extra = existing.length > 1 ? ` (+${existing.length - 1} altre assegnazioni)` : "";
  const confirmed = await openConfirmModal(
    `Attenzione: Telaio gia assegnato a ${first.risorsaLabel} il giorno ${dayLabel}${extra}. Vuoi assegnarlo comunque?`,
    null,
    { okLabel: "Assegna comunque", cancelLabel: "Annulla" }
  );
  if (!confirmed) {
    setStatus("Assegnazione annullata.", "info");
    return false;
  }
  return true;
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
  const shouldForcePlannedTelaio = isTelaioActivityTitle(item.titolo) && Boolean(item.commessa_id);
  let pendingMoveGhostRelease = null;

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
      stato: shouldForcePlannedTelaio ? "pianificata" : item.stato || "pianificata",
      data_inizio: newStart.toISOString(),
      data_fine: newEnd.toISOString(),
    });
    if (error) {
      setStatus(`Copy error: ${error.message}`, "error");
      return;
    }
    const overrideCleared = shouldForcePlannedTelaio
      ? await clearTodoNotNeededOverrideSilently(item.commessa_id, item.titolo)
      : true;
    const telaioPlannedSynced = shouldForcePlannedTelaio
      ? await syncTelaioPlannedDateFromMatrix(item.commessa_id, newEnd)
      : true;
    if (!overrideCleared) {
      setStatus("Activity copied, ma override ToDos non aggiornato.", "error");
    } else if (!telaioPlannedSynced) {
      setStatus("Activity copied, ma data telaio pianificata non aggiornata.", "error");
    } else {
      setStatus("Activity copied.", "ok");
    }
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

    let moveCompleted = false;
    try {
      let sourceBarEl = null;
      if (matrixGrid && item?.id != null) {
        sourceBarEl = Array.from(matrixGrid.querySelectorAll(".matrix-activity-bar[data-activity-id]")).find(
          (el) => String(el.dataset.activityId || "") === String(item.id)
        );
      }

      pendingMoveGhostRelease = await animateMatrixMoveGhost(sourceBarEl, cell, durationDays, {
        dropClientX: e?.clientX,
        dropClientY: e?.clientY,
        pointerOffsetX: matrixState.dragPointerOffsetX,
        pointerOffsetY: matrixState.dragPointerOffsetY,
        useExistingFollower: true,
        holdAtTarget: true,
      });

      const movePayload = {
        risorsa_id: r.id,
        data_inizio: newStart.toISOString(),
        data_fine: newEnd.toISOString(),
      };
      if (shouldForcePlannedTelaio) {
        movePayload.stato = "pianificata";
      }
      const { error } = await supabase.from("attivita").update(movePayload).eq("id", item.id);
      if (error) {
        setStatus(`Move error: ${error.message}`, "error");
        return;
      }
      if (isPhaseKey(item.titolo) && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
        const ok = await shiftDependentPhases(item, deltaDays, dependentToShift);
        if (!ok) return;
      }

      let telaioPlannedSynced = true;
      if (isTelaioActivityTitle(item.titolo) && item.commessa_id) {
        telaioPlannedSynced = await syncTelaioPlannedDateFromMatrix(item.commessa_id, newEnd);
        const commessa = findCommessaInState(item.commessa_id);
        await warnTelaioPlannedMismatch(item.commessa_id, commessa?.data_conferma_consegna_telaio || null, {
          telaioEndDay: startOfDay(newEnd),
        });
      }
      const overrideCleared = shouldForcePlannedTelaio
        ? await clearTodoNotNeededOverrideSilently(item.commessa_id, item.titolo)
        : true;
      if (!overrideCleared) {
        setStatus("Activity moved, ma override ToDos non aggiornato.", "error");
      } else if (!telaioPlannedSynced) {
        setStatus("Activity moved, ma data telaio pianificata non aggiornata.", "error");
      } else {
        setStatus("Activity moved.", "ok");
      }
      moveCompleted = true;
    } finally {
      if (!moveCompleted && pendingMoveGhostRelease) {
        pendingMoveGhostRelease({ immediate: true });
        pendingMoveGhostRelease = null;
      }
    }
  }
  await loadMatrixAttivita();
  if (pendingMoveGhostRelease) {
    let finalBarEl = null;
    if (matrixGrid && item?.id != null) {
      finalBarEl = Array.from(matrixGrid.querySelectorAll(".matrix-activity-bar[data-activity-id]")).find(
        (el) => String(el.dataset.activityId || "") === String(item.id)
      );
    }
    pendingMoveGhostRelease({ anchorEl: finalBarEl || null });
    pendingMoveGhostRelease = null;
  }
  await refreshTodoRealtimeSnapshot([item.commessa_id]);
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

function formatMissingSequence(values, limit = 12) {
  if (!Array.isArray(values) || !values.length) return "";
  const head = values.slice(0, limit);
  const text = head.join(", ");
  if (values.length <= limit) return text;
  return `${text} +${values.length - limit}`;
}

function updateCommesseSequenceAlert() {
  if (!commesseSeqAlert) return;
  const currentYear = new Date().getFullYear();
  const numbers = (state.commesse || [])
    .filter((c) => Number(c.anno) === currentYear)
    .map((c) => Number(c.numero))
    .filter((n) => Number.isInteger(n) && n > 0);
  const sequenceNumbers = numbers.filter((n) => n !== 999);

  commesseSeqAlert.classList.remove("hidden", "error");

  if (!sequenceNumbers.length) {
    commesseSeqAlert.textContent = `Anno ${currentYear}: nessuna commessa.`;
    return;
  }

  const uniqueSorted = Array.from(new Set(sequenceNumbers)).sort((a, b) => a - b);
  const hasBelowTen = uniqueSorted.some((n) => n < 10);
  const highest = uniqueSorted[uniqueSorted.length - 1];
  const present = new Set(uniqueSorted.filter((n) => n >= 10));
  const missing = [];
  for (let n = 10; n <= highest; n += 1) {
    if (!present.has(n)) missing.push(n);
  }

  if (hasBelowTen || missing.length) {
    const issues = [];
    if (hasBelowTen) issues.push("presenti numeri < 10");
    if (missing.length) issues.push(`buchi: ${formatMissingSequence(missing)}`);
    commesseSeqAlert.textContent = `Anno ${currentYear}: numerazione non continua (${issues.join("; ")}).`;
    commesseSeqAlert.classList.add("error");
    return;
  }

  commesseSeqAlert.textContent = `Anno ${currentYear}: numerazione continua da 10 a ${highest}.`;
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
  updateCommesseSequenceAlert();
  renderTelaioOrderAlarms();
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
  if (mode === "asc") {
    todoPriorityBtn.textContent = "Priorità: crescente";
    return;
  }
  if (mode === "desc") {
    todoPriorityBtn.textContent = "Priorità: decrescente";
    return;
  }
  todoPriorityBtn.textContent = "Priorità: OFF";
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
  if (field === "telaio_stimato") return parse(commessa.data_conferma_consegna_telaio);
  if (field === "prelievo") return parse(commessa.data_prelievo);
  if (field === "kit_cavi") return parse(commessa.data_arrivo_kit_cavi);
  if (field === "consegna") {
    return parse(commessa.data_consegna_macchina || commessa.data_consegna_prevista || commessa.data_consegna);
  }
  return null;
}

function getTodoPriorityFieldLabel(field) {
  if (field === "telaio_target") return "Telaio target";
  if (field === "telaio_stimato") return "Telaio pianificato";
  if (field === "prelievo") return "Prelievo";
  if (field === "kit_cavi") return "Kit cavi";
  if (field === "consegna") return "Consegna";
  return "Priorita";
}

function getTodoPriorityDisplay(commessa, field) {
  const value = getTodoPriorityDate(commessa, field);
  if (value == null) return "n/d";
  return formatDateDMY(new Date(value));
}

function updateTodoSyncMatrixFiltersButton() {
  if (!todoSyncMatrixFiltersBtn) return;
  const count = todoSyncedCommessaIds.size;
  todoSyncMatrixFiltersBtn.classList.toggle("active", count > 0);
  if (count > 0) {
    todoSyncMatrixFiltersBtn.dataset.count = String(count);
  } else {
    delete todoSyncMatrixFiltersBtn.dataset.count;
  }
  if (count === 0) {
    todoSyncMatrixFiltersBtn.title = "Quick fill: copia i filtri commessa attivi dalla Matrice risorse";
    return;
  }
  if (count === 1) {
    const only = Array.from(todoSyncedCommessaIds)[0];
    const commessa = (state.commesse || []).find((c) => String(c.id) === String(only));
    const code = commessa?.codice || "1 commessa";
    todoSyncMatrixFiltersBtn.title = `Quick fill attivo: ${code}`;
    return;
  }
  todoSyncMatrixFiltersBtn.title = `Quick fill attivo: ${count} commesse sincronizzate dalla Matrice risorse`;
}

function setTodoSyncedCommessaIds(ids = []) {
  const normalized = Array.from(
    new Set(
      (ids || [])
        .map((id) => String(id || "").trim())
        .filter(Boolean)
    )
  );
  todoSyncedCommessaIds = new Set(normalized);
  updateTodoSyncMatrixFiltersButton();
}

function handleTodoQuickSyncFromMatrixFilters() {
  const matrixCommesse = Array.from(matrixState.selectedCommesse || [])
    .map((id) => String(id || "").trim())
    .filter(Boolean);
  const normalizedMatrix = Array.from(new Set(matrixCommesse));
  const current = Array.from(todoSyncedCommessaIds);
  const sameSelection =
    normalizedMatrix.length === current.length &&
    normalizedMatrix.every((id) => todoSyncedCommessaIds.has(id));
  if (sameSelection && normalizedMatrix.length > 0) {
    setTodoSyncedCommessaIds([]);
    renderMatrixReportFromTodoFilters();
    setStatus("Quick fill To Dos disattivato.", "ok");
    return;
  }
  if (!matrixCommesse.length) {
    if (todoSyncedCommessaIds.size > 0) {
      setTodoSyncedCommessaIds([]);
      renderMatrixReportFromTodoFilters();
      setStatus("Quick fill To Dos disattivato: nessun filtro commessa in Matrice.", "ok");
      return;
    }
    if ((matrixState.selectedAttivita || new Set()).size > 0) {
      setStatus("In Matrice hai solo filtro attivita: quick fill To Dos copia i filtri commessa.", "error");
      return;
    }
    setStatus("Nessun filtro commessa attivo in Matrice risorse.", "error");
    return;
  }
  setTodoSyncedCommessaIds(normalizedMatrix);
  renderMatrixReportFromTodoFilters();
  if (normalizedMatrix.length === 1) {
    setStatus("Quick fill To Dos: filtro commessa sincronizzato dalla Matrice.", "ok");
  } else {
    setStatus(`Quick fill To Dos: sincronizzate ${normalizedMatrix.length} commesse dalla Matrice.`, "ok");
  }
}

function applyTodoFilters(list, columns) {
  const { desc, numVal, year, priorityField, priorityMode, sortMode, showComplete } = getTodoFilters();
  const isComplete = (commessa) => {
    if (!columns.length) return false;
    return columns.every((col) => getTodoStatusFor(commessa.id, col.titolo, col.reparto) === "fatta");
  };
  const isReleasedOrClosed = (commessa) => {
    const global = getTodoGlobalCommessaStatus(commessa).label;
    return global === "Sospesa" || global === "Evasa" || global === "Rilasciata a produzione";
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
      const text = `${c.titolo || ""} ${c.tipo_macchina || ""}`.toLowerCase();
      return text.includes(desc);
    });
  }
  if (todoSyncedCommessaIds.size > 0) {
    commesse = commesse.filter((c) => todoSyncedCommessaIds.has(String(c.id || "")));
  }
  if (!showComplete) {
    commesse = commesse.filter((c) => !isComplete(c) && !isReleasedOrClosed(c));
  }
  if ((priorityMode === "asc" || priorityMode === "desc") && priorityField) {
    const dir = priorityMode === "asc" ? 1 : -1;
    commesse.sort((a, b) => {
      const aa = getTodoPriorityDate(a, priorityField);
      const bb = getTodoPriorityDate(b, priorityField);
      if (aa == null && bb == null) return compareCommesseDesc(a, b);
      if (aa == null) return 1;
      if (bb == null) return -1;
      if (aa !== bb) return (aa - bb) * dir;
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
  refreshMachineTypeSelects({
    detailType: normalizeMachineTypeValue(commessa.tipo_macchina || "Altro tipo"),
    detailVariant: normalizeMachineVariantValue(commessa.variante_macchina || "standard"),
  });
  d.stato.value = commessa.stato || "nuova";
  d.priorita.value = commessa.priorita || "";
  d.data_ingresso.value = formatDate(commessa.data_ingresso);
  if (d.data_ordine_telaio) d.data_ordine_telaio.value = formatDate(commessa.data_ordine_telaio);
  if (d.data_conferma_consegna_telaio) {
    d.data_conferma_consegna_telaio.value = formatDate(commessa.data_conferma_consegna_telaio);
  }
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
  refreshDetailTelaioOrderAccess();
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
  updateDetailFieldCompletionUI();
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
  if (reportsLink) {
    const role = String(data.ruolo || "").trim().toLowerCase();
    if (role === "admin") reportsLink.classList.remove("hidden");
    else reportsLink.classList.add("hidden");
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
  if (matrixBootstrapSyncInProgress) {
    renderResourcesPanel();
    return;
  }
  const canRenderMatrixNow =
    !authSyncInProgress ||
    (Array.isArray(matrixState.attivita) && matrixState.attivita.length > 0);
  if (canRenderMatrixNow) {
    scheduleMatrixRenderStabilized();
  }
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
  setMatrixQuickCommessaFilterRaw(matrixState.quickCommessaFilterRaw, { render: false });
  refreshMachineTypeSelects();
  updateCommesseSequenceAlert();
  renderYearFilter();
  renderReportYearFilter();
  renderTodoYearFilter();
  applyFilters();
  renderMatrixCommesse();
  renderMatrixCommesseYearFilter();
  renderMatrixCommesseColumn();
  renderMatrixReport();
  if (state.selected) {
    const stillThere = state.commesse.find((c) => c.id === state.selected.id);
    if (stillThere) selectCommessa(stillThere.id);
    else clearSelection();
  }
}

function renderMatrixHeader(days, gridTemplateColumns = `160px repeat(${days.length}, minmax(0, 1fr))`) {
  matrixGrid.innerHTML = "";
  const header = document.createElement("div");
  header.className = "matrix-row matrix-header-row";
  header.style.gridTemplateColumns = gridTemplateColumns;

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
    const dayLabel = d.toLocaleDateString("it-IT", { weekday: "short", day: "2-digit", month: "short" });
    const weekLabel = d.getDay() === 1 ? `W${String(getIsoWeekNumber(d)).padStart(2, "0")}` : "";
    cell.innerHTML = `
      <span class="matrix-header-day">${dayLabel}</span>
      <span class="matrix-header-week${weekLabel ? "" : " matrix-header-week-placeholder"}">${weekLabel || "&nbsp;"}</span>
    `;
    header.appendChild(cell);
  });

  matrixGrid.appendChild(header);
  setupMatrixPan(header);
}

function applyMatrixPanPreview(drag, dx) {
  if (!drag?.panTargets?.length) return;
  const baseOffset = Number.isFinite(drag.baseOffsetPx)
    ? drag.baseOffsetPx
    : Number.isFinite(matrixState.panBaseOffsetPx)
    ? matrixState.panBaseOffsetPx
    : 0;
  const safeDx = Number.isFinite(dx) ? dx : 0;
  drag.currentDx = safeDx;
  const totalDx = baseOffset + safeDx;
  drag.panTargets.forEach((el) => {
    el.style.transform = totalDx ? `translateX(${totalDx}px)` : "";
  });
}

function clearMatrixPanPreview(drag) {
  if (!drag) return;
  if (drag.previewRafId) {
    cancelAnimationFrame(drag.previewRafId);
    drag.previewRafId = null;
  }
  if (!drag.panTargets?.length) return;
  const baseOffset = Number.isFinite(matrixState.panBaseOffsetPx)
    ? matrixState.panBaseOffsetPx
    : Number.isFinite(drag.baseOffsetPx)
    ? drag.baseOffsetPx
    : 0;
  const baseTransform = baseOffset ? `translateX(${baseOffset}px)` : "";
  drag.panTargets.forEach((el) => {
    el.style.transform = baseTransform;
    el.style.transition = "";
  });
}

function scheduleMatrixPanPreview(drag, dx) {
  if (!drag) return;
  drag.pendingDx = Number.isFinite(dx) ? dx : 0;
  if (drag.previewRafId) return;
  drag.previewRafId = requestAnimationFrame(() => {
    drag.previewRafId = null;
    applyMatrixPanPreview(drag, drag.pendingDx);
  });
}

function flushMatrixPanPreview(drag) {
  if (!drag) return;
  if (drag.previewRafId) {
    cancelAnimationFrame(drag.previewRafId);
    drag.previewRafId = null;
  }
  if (Number.isFinite(drag.pendingDx)) {
    applyMatrixPanPreview(drag, drag.pendingDx);
  }
}

function applyMatrixStaticPanOffset() {
  if (!matrixGrid) return;
  matrixGrid
    .querySelectorAll(".matrix-cell:not(.matrix-resource), .matrix-activity-bar")
    .forEach((el) => {
      el.style.transform = "";
    });
}

function canStartMatrixPanFromTarget(target, origin = "header") {
  if (!target) return false;
  if (
    target.closest(
      ".matrix-activity-bar, .matrix-resize-handle, button, input, select, textarea, a, label, [contenteditable='true']"
    )
  ) {
    return false;
  }
  if (origin === "grid") {
    if (target.closest(".matrix-resource")) return false;
    return Boolean(
      target.closest(".matrix-row, .matrix-cell, .matrix-bar-layer, .matrix-section-cell, .matrix-section-row")
    );
  }
  return true;
}

function setupMatrixPan(headerRow) {
  if (!headerRow || !matrixGrid) return;
  if (!MATRIX_MOVEMENT_ENABLED) {
    headerRow.onpointerdown = null;
    headerRow.onpointermove = null;
    headerRow.onpointerup = null;
    headerRow.onpointerleave = null;
    headerRow.onpointercancel = null;
    matrixGrid.onpointerdown = null;
    matrixGrid.onpointermove = null;
    matrixGrid.onpointerup = null;
    matrixGrid.onpointerleave = null;
    matrixGrid.onpointercancel = null;
    matrixGrid.classList.remove("is-panning");
    matrixGrid.classList.remove("is-zooming-preview");
    matrixGrid.classList.remove("is-pan-snapping");
    matrixPan = null;
    return;
  }
  const clearPreview = () => {
    if (!matrixPan?.panTargets?.length) return;
    matrixPan.panTargets.forEach((el) => {
      el.style.transform = "";
    });
  };
  headerRow.onpointerdown = (e) => {
    if (matrixPan) return;
    if (e.button !== 0) return;
    if (!canStartMatrixPanFromTarget(e.target, "header")) return;
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
      pointerId: e.pointerId,
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
    if (matrixPan.pointerId != null && e.pointerId != null && e.pointerId !== matrixPan.pointerId) return;
    const dx = e.clientX - matrixPan.startX;
    const dy = e.clientY - matrixPan.startY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) matrixPan.moved = true;
    if (!matrixPan.mode) {
      if (Math.abs(dy) > 3 && Math.abs(dy) >= Math.abs(dx)) matrixPan.mode = "zoom";
      else if (Math.abs(dx) > 3) matrixPan.mode = "pan";
    }
    if (matrixPan.mode === "pan") {
      matrixPan.panTargets.forEach((el) => {
        el.style.transform = dx ? `translateX(${dx}px)` : "";
      });
    }
  };
  headerRow.onpointerup = async (e) => {
    if (!matrixPan) return;
    if (matrixPan.pointerId != null && e.pointerId != null && e.pointerId !== matrixPan.pointerId) return;
    headerRow.releasePointerCapture?.(e.pointerId);
    const dx = e.clientX - matrixPan.startX;
    const dy = e.clientY - matrixPan.startY;
    const dayWidth = matrixPan.dayWidth;
    const shiftDays = dayWidth ? Math.round(-dx / dayWidth) : 0;
    clearPreview();
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
      matrixState.customWeeks = null;
      matrixState.panResidualDays = 0;
      matrixState.date = startOfWeek(matrixState.date);
      if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
      setMatrixViewLabel();
      await loadMatrixAttivita();
      return;
    }
    if (!shouldMove) return;
    matrixState.panResidualDays = 0;
    matrixState.date = addDays(matrixState.date, shiftDays);
    if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
    await loadMatrixAttivita();
  };
  headerRow.onpointercancel = (e) => {
    if (!matrixPan) return;
    if (matrixPan.pointerId != null && e?.pointerId != null && e.pointerId !== matrixPan.pointerId) return;
    headerRow.releasePointerCapture?.(matrixPan.pointerId ?? e?.pointerId);
    clearPreview();
    matrixGrid.classList.remove("is-panning");
    matrixGrid.classList.remove("is-zooming-preview");
    matrixGrid.classList.remove("is-pan-snapping");
    matrixPan = null;
  };
  headerRow.onpointerleave = () => {
    if (!matrixPan) return;
    clearPreview();
    matrixGrid.classList.remove("is-panning");
    matrixPan = null;
  };
  matrixGrid.onpointerdown = null;
  matrixGrid.onpointermove = null;
  matrixGrid.onpointerup = null;
  matrixGrid.onpointerleave = null;
  matrixGrid.onpointercancel = null;
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

function renderMatrixCommesseYearFilter() {
  if (!matrixCommessaYear) return;
  const currentYear = String(new Date().getFullYear());
  const years = Array.from(
    new Set(
      (state.commesse || [])
        .map((c) => Number(c.anno))
        .filter((y) => Number.isFinite(y) && y > 0)
    )
  ).sort((a, b) => b - a);
  if (!years.includes(Number(currentYear))) years.unshift(Number(currentYear));
  const prev = matrixCommessaYear.value;
  matrixCommessaYear.innerHTML = "";
  years.forEach((year) => {
    const opt = document.createElement("option");
    opt.value = String(year);
    opt.textContent = String(year);
    matrixCommessaYear.appendChild(opt);
  });
  matrixCommessaYear.value = prev && years.includes(Number(prev)) ? prev : currentYear;
}

function renderMatrixCommesseColumn() {
  if (!matrixCommesseList) return;
  const q = (matrixCommessaSearch.value || "").trim().toLowerCase();
  const year = matrixCommessaYear ? Number(matrixCommessaYear.value || 0) : 0;
  const items = state.commesse.filter((c) => {
    if (year && Number(c.anno) !== year) return false;
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

function getActivityEstimatedHours(attivita) {
  const est = Number(attivita?.ore_stimate);
  if (Number.isFinite(est) && est > 0) return est;
  const assenza = Number(attivita?.ore_assenza);
  if (Number.isFinite(assenza) && assenza > 0) return assenza;
  return 0;
}

function getActivityEstimatedHoursPerDay(attivita) {
  const total = getActivityEstimatedHours(attivita);
  if (!total) return 0;
  const start = attivita?.data_inizio ? startOfDay(new Date(attivita.data_inizio)) : null;
  const end = attivita?.data_fine ? startOfDay(new Date(attivita.data_fine)) : null;
  if (!start || !end || Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return total;
  const businessDays = Math.max(1, businessDaysBetweenInclusive(start, end));
  return total / businessDays;
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
  const windowSpec = getMatrixRenderWindow(matrixState.date);
  const days = windowSpec.days;
  const visibleBusinessDays = windowSpec.visibleBusinessDays;
  const bufferBusinessDays = windowSpec.bufferBusinessDays;
  const resourceColWidth = 160;
  const gridRect = matrixGrid.getBoundingClientRect();
  const wrapRect = matrixWrap ? matrixWrap.getBoundingClientRect() : null;
  const fallbackWidth = Math.max(320, (wrapRect ? wrapRect.width - 252 : 0));
  const viewportWidth = Math.max(320, gridRect.width || matrixGrid.clientWidth || fallbackWidth);
  const availableDaysWidth = Math.max(160, viewportWidth - resourceColWidth);
  const dayWidthPx = Math.max(20, availableDaysWidth / Math.max(1, visibleBusinessDays));
  const gridTemplateColumns = `160px repeat(${days.length}, ${dayWidthPx}px)`;
  const residualDays = 0;
  matrixState.panResidualDays = 0;
  matrixState.panVisibleBusinessDays = visibleBusinessDays;
  matrixState.panBufferBusinessDays = bufferBusinessDays;
  matrixState.panDayWidthPx = dayWidthPx;
  matrixState.panBaseOffsetPx = (-bufferBusinessDays + residualDays) * dayWidthPx;
  renderMatrixHeader(days, gridTemplateColumns);

  if (!matrixState.risorse.length) {
    const empty = document.createElement("div");
    empty.className = "matrix-empty";
    empty.textContent = "Nessuna risorsa.";
    matrixGrid.appendChild(empty);
    applyMatrixStaticPanOffset();
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
    sectionRow.style.gridTemplateColumns = gridTemplateColumns;
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
      row.dataset.risorsaId = String(r.id);
      row.style.gridTemplateColumns = gridTemplateColumns;
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
    const commessaFilter = getActiveMatrixCommessaFilter();
    const visibleActivities = rowActivities.filter((a) => {
      if (filterAttivita && !matrixState.selectedAttivita.has(a.titolo)) return false;
      if (commessaFilter.active && !commessaFilter.ids.has(String(a.commessa_id || ""))) return false;
      return true;
    });

    const occupiedDays = new Set();
    const dayHours = new Map();
    visibleActivities.forEach((a) => {
      const start = startOfDay(new Date(a.data_inizio));
      const end = startOfDay(new Date(a.data_fine));
      const hoursPerDay = getActivityEstimatedHoursPerDay(a);
      let current = new Date(start);
      while (current <= end) {
        const key = formatDateLocal(current);
        if (dayIndex.has(key)) {
          occupiedDays.add(key);
          if (hoursPerDay > 0) dayHours.set(key, (dayHours.get(key) || 0) + hoursPerDay);
        }
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
      const totalHours = dayHours.get(dayKey) || 0;
      if (totalHours > 9) {
        cell.classList.add("overbooked-hours");
        cell.title = `${totalHours.toFixed(1)}h pianificate`;
      }

      cell.addEventListener("dragover", (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        updateMatrixDropEffect(e);
        updateMatrixDragFollowerPosition(e.clientX, e.clientY);
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
        markMatrixDragDropHandled();
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
        const includesTelaio = Boolean(commessaId) && attivitaList.some((title) => isTelaioActivityTitle(title));
        let telaioPlannedSynced = true;
        if (commessaId && attivitaList.some((title) => isTelaioActivityTitle(title))) {
          const ok = await confirmTelaioDuplicateAssignment(commessaId);
          if (!ok) return;
        }
        const { error } = await supabase.from("attivita").insert(rows);
        if (error) {
          setStatus(`Assignment error: ${error.message}`, "error");
          return;
        }
        if (includesTelaio) {
          telaioPlannedSynced = await syncTelaioPlannedDateFromMatrix(commessaId, endAt);
        }
        if (!telaioPlannedSynced) {
          setStatus("Activities assigned, ma data telaio pianificata non aggiornata.", "error");
        } else {
          setStatus(rows.length > 1 ? "Activities assigned." : "Activity assigned.", "ok");
        }
        await loadMatrixAttivita();
        await refreshTodoRealtimeSnapshot([commessaId]);
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
      bars.push({ a, laneIndex, startIdx, endIdx, startKey, endKey });
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
      applyMatrixBarPastClass(bar, b.a);
      applyMatrixBarClientFlowClasses(bar, b.a, bar.dataset.reparto || "");
      applyMatrixBarTelaioOrderClass(bar, b.a);
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
        const barRect = bar.getBoundingClientRect();
        const hasPointer = Number.isFinite(e.clientX) && Number.isFinite(e.clientY);
        matrixState.dragPointerOffsetX = hasPointer ? e.clientX - barRect.left : barRect.width / 2;
        matrixState.dragPointerOffsetY = hasPointer ? e.clientY - barRect.top : barRect.height / 2;
        if (e.dataTransfer) {
          e.dataTransfer.effectAllowed = "copyMove";
          try {
            e.dataTransfer.setDragImage(matrixTransparentDragImage, 0, 0);
          } catch (_err) {
            // Ignore browsers that block custom drag image.
          }
        }
        createMatrixDragFollower(bar, e);
        if (matrixGrid) matrixGrid.classList.add("matrix-dragging");
        bar.classList.add("dragging");
        e.dataTransfer.setData("text/plain", b.a.id);
      });
      bar.addEventListener("dragend", () => {
        matrixState.draggingId = null;
        matrixState.dragPointerOffsetX = null;
        matrixState.dragPointerOffsetY = null;
        if (!matrixState.dragDropHandled) {
          removeMatrixDragFollower();
          restoreMatrixDragSourceVisual();
        }
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
          if (b.a.commessa_id) {
            applyCommessaHighlight(b.a.commessa_id, {
              preserveViewport: true,
              anchorEl: bar,
            });
            commessaLongPressSuppress = true;
          } else {
            clearCommessaHighlight({
              preserveViewport: true,
              anchorEl: bar,
            });
          }
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
        updateMatrixDropEffect(e);
        updateMatrixDragFollowerPosition(e.clientX, e.clientY);
      });
      bar.addEventListener("drop", async (e) => {
        if (!state.canWrite) return;
        e.preventDefault();
        e.stopPropagation();
        const elements = document.elementsFromPoint(e.clientX, e.clientY);
        const cell = elements.find((el) => el.classList && el.classList.contains("matrix-cell"));
        if (cell) {
          markMatrixDragDropHandled();
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
        clearCommessaHighlight({
          preserveViewport: true,
          anchorEl: bar,
        });
        if (!b.a.titolo) {
          openActivityModal(b.a);
          return;
        }
        if (b.a.commessa_id) {
          mergeReportActivitiesFromMatrix(b.a.commessa_id);
        }
        openMatrixDualPanels(b.a, bar);
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
        const layerRect = layerEl ? layerEl.getBoundingClientRect() : null;
        const barRect = bar.getBoundingClientRect();
        const originalBarInitialStyle = {
          position: bar.style.position || "",
          top: bar.style.top || "",
          left: bar.style.left || "",
          width: bar.style.width || "",
          height: bar.style.height || "",
          marginTop: bar.style.marginTop || "",
          gridColumn: bar.style.gridColumn || "",
          zIndex: bar.style.zIndex || "",
          transition: bar.style.transition || "",
        };
        if (layerRect) {
          bar.style.position = "absolute";
          bar.style.top = `${barRect.top - layerRect.top}px`;
          bar.style.left = `${barRect.left - layerRect.left}px`;
          bar.style.width = `${barRect.width}px`;
          bar.style.height = `${barRect.height}px`;
          bar.style.marginTop = "0";
          bar.style.gridColumn = "auto";
          bar.style.zIndex = "5";
          bar.style.transition = "none";
        }
        const startDate = startOfDay(new Date(b.a.data_inizio));
        const endDate = new Date(b.a.data_fine);
        const dayCells = rowEl ? Array.from(rowEl.querySelectorAll(".matrix-cell[data-day]")) : [];
        const startCell = rowEl ? rowEl.querySelector(`.matrix-cell[data-day="${b.startKey}"]`) : null;
        const endCell = rowEl ? rowEl.querySelector(`.matrix-cell[data-day="${b.endKey}"]`) : null;
        const startCellIndex = startCell ? dayCells.indexOf(startCell) : -1;
        const endCellIndex = endCell ? dayCells.indexOf(endCell) : -1;
        const startCellRect = startCell ? startCell.getBoundingClientRect() : null;
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
          targetDayKey: b.endKey || b.startKey,
          startIdx: b.startIdx,
          rowEl,
          layerEl,
          dayCells,
          startCellIndex: startCellIndex >= 0 ? startCellIndex : b.startIdx,
          initialTargetIndex: endCellIndex >= 0 ? endCellIndex : b.endIdx,
          dragStartX: e.clientX,
          layerRect,
          startCellRect,
          previewBar: bar,
          usesOriginalBarForResize: true,
          originalBarInitialStyle,
          originalBar: bar,
          targetCell: null,
        };
        if (endCell) setMatrixResizeTargetCell(matrixState.resizing, endCell);
        else if (startCell) setMatrixResizeTargetCell(matrixState.resizing, startCell);
        bar.classList.add("resizing");
      });
      bar.appendChild(handle);
      barLayer.appendChild(bar);
    });

      matrixGrid.appendChild(row);
    });
  });
  applyMatrixStaticPanOffset();
  if (commessaHighlightId) {
    applyCommessaHighlight(commessaHighlightId);
  } else if (matrixState.selectedAttivita.size > 0) {
    applyMatrixActivityFilterCompaction();
  } else {
    clearMatrixResourceCompaction();
  }
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
  matrixReportRange.textContent = "";

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
  const getReportActivityDate = (entry) => {
    if (!entry) return null;
    const raw = entry.end || entry.start || null;
    if (!raw) return null;
    const parsed = raw instanceof Date ? new Date(raw.getTime()) : new Date(raw);
    if (Number.isNaN(parsed.getTime())) return null;
    const year = parsed.getFullYear();
    if (year < 2000 || year > 2100) return null;
    return parsed;
  };
  const pickLatestActivityByDate = (entries) => {
    let picked = null;
    let pickedTime = -Infinity;
    (entries || []).forEach((entry) => {
      const date = getReportActivityDate(entry);
      if (!date) return;
      const t = date.getTime();
      if (t <= pickedTime) return;
      picked = { entry, date };
      pickedTime = t;
    });
    return picked;
  };
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
    const telaioProgrammataRaw = c.data_conferma_consegna_telaio || "";
    const telaioProgrammataDate = telaioProgrammataRaw ? new Date(telaioProgrammataRaw) : null;
    const telaioRichiestaRaw = c.data_ordine_telaio || "";
    const telaioRichiestaDate = telaioRichiestaRaw ? new Date(telaioRichiestaRaw) : null;
    const telaioOrdinatoRaw = c.data_consegna_telaio_effettiva || "";
    const telaioOrdinatoDate = telaioOrdinatoRaw ? new Date(telaioOrdinatoRaw) : null;
    const kitRaw = c.data_arrivo_kit_cavi || "";
    const kitDate = kitRaw ? new Date(kitRaw) : null;
    const prelievoRaw = c.data_prelievo || "";
    const prelievoDate = prelievoRaw ? new Date(prelievoRaw) : null;
    const isDateOk = (d) => d && !Number.isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
    const safeTarget = isDateOk(targetDate) ? targetDate : null;
    const safeTelaioProgrammata = isDateOk(telaioProgrammataDate) ? telaioProgrammataDate : null;
    const safeTelaioRichiesta = isDateOk(telaioRichiestaDate) ? telaioRichiestaDate : null;
    const safeTelaioOrdinato = isDateOk(telaioOrdinatoDate) ? telaioOrdinatoDate : null;
    const telaioOrdinatoFlag = Boolean(c.telaio_consegnato);
    const safeKit = isDateOk(kitDate) ? kitDate : null;
    const safePrelievo = isDateOk(prelievoDate) ? prelievoDate : null;
    const thermoEntries = getTodoEntriesForCell(c.id, TODO_ACTIVITY_PRIMARY_PHASE_KEY).filter(
      (entry) => normalizePhaseKey(entry.stato) !== "annullata"
    );
    const thermoDone = pickLatestActivityByDate(
      thermoEntries.filter((entry) => normalizePhaseKey(entry.stato) === "completata")
    );
    const thermoPlanned = thermoDone
      ? null
      : pickLatestActivityByDate(thermoEntries.filter((entry) => normalizePhaseKey(entry.stato) === "pianificata"));
    const thermoFineDate = thermoDone?.date || thermoPlanned?.date || null;
    const thermoFineInfo = thermoDone
      ? { label: "done", cls: "is-ok" }
      : thermoPlanned
      ? { label: "planned", cls: "is-planned" }
      : { label: "n/d", cls: "" };
    const telaioProgrammataInfo = getBusinessDiffInfo(today, safeTelaioProgrammata);
    const telaioRichiestaInfo = getBusinessDiffInfo(today, safeTelaioRichiesta);
    const telaioOrdinatoInfo = telaioOrdinatoFlag
      ? safeTelaioOrdinato
        ? { label: "ordinato", cls: "is-ok" }
        : { label: "manca data", cls: "is-late" }
      : { label: "non ordinato", cls: "" };
    const missingPrelievo = !safePrelievo;
    const missingKit = !safeKit;
    const missingOrdine = !safeTelaioRichiesta || missingPrelievo;
    const missingTarget = !safeTarget;
    const codeInfo = parseCommessaCode(c.codice || "");
    const numeroLabel = c.numero ?? codeInfo.num ?? c.codice ?? "-";
    const annoLabel = c.anno ?? codeInfo.year ?? "-";
    const descrizioneLabel = c.titolo || "Senza titolo";
    const machineTypeRaw = String(c.tipo_macchina || "Altro tipo").trim() || "Altro tipo";
    const machineTypeLabel = getMachineTypeDisplayLabel(machineTypeRaw);

    const item = document.createElement("div");
    item.className = "report-item";
    const scheduleHtml = buildReportScheduleHtml(c.id);
    item.innerHTML = `
      <div class="report-main">
        <button class="report-title report-commessa-link" type="button" data-commessa-id="${c.id}">
          <span class="report-title-parts">
            <span class="report-title-num">${escapeHtml(String(numeroLabel))}</span>
            <span class="report-title-year">${escapeHtml(String(annoLabel))}</span>
            <span class="report-title-desc" title="${escapeHtml(descrizioneLabel)}">${escapeHtml(descrizioneLabel)}</span>
          </span>
          <span class="report-machine-pill" title="${escapeHtml(machineTypeRaw)}">${escapeHtml(machineTypeLabel)}</span>
        </button>
        <div class="report-meta">Stato: ${c.stato || "-"} - Cliente: ${c.cliente || "-"}</div>
        ${scheduleHtml}
        ${
          missingOrdine || missingTarget
            ? `<div class="report-missing">
                ${!safeTelaioRichiesta ? `<span class="missing-pill">Manca data ordine telaio</span>` : ""}
                ${missingKit ? `<span class="missing-pill">Manca data arrivo kit cavi</span>` : ""}
                ${missingPrelievo ? `<span class="missing-pill">Manca data prelievo materiali</span>` : ""}
                ${missingTarget ? `<span class="missing-pill">Manca data consegna macchina</span>` : ""}
              </div>`
            : ""
        }
      </div>
      <div class="report-milestone ${thermoFineInfo.cls}">
        <div class="report-label">Fine progettazione termodinamica</div>
        <div class="report-value">${thermoFineDate ? formatDateDMY(thermoFineDate) : "-"}</div>
        <div class="report-sub">${thermoFineInfo.label}</div>
      </div>
      <div class="report-milestone ${telaioProgrammataInfo.cls}">
        <div class="report-label">Data telaio programmata</div>
        <div class="report-value">${safeTelaioProgrammata ? formatDateDMY(safeTelaioProgrammata) : "-"}</div>
        <div class="report-sub">${telaioProgrammataInfo.label}</div>
      </div>
      <div class="report-milestone ${telaioOrdinatoInfo.cls}">
        <div class="report-label">Data telaio ordinato</div>
        <div class="report-value">${telaioOrdinatoFlag && safeTelaioOrdinato ? formatDateDMY(safeTelaioOrdinato) : "-"}</div>
        <div class="report-sub">${telaioOrdinatoInfo.label}</div>
      </div>
      <div class="report-milestone ${telaioRichiestaInfo.cls}">
        <div class="report-label">Data telaio richiesta da produzione</div>
        <div class="report-value">${safeTelaioRichiesta ? formatDateDMY(safeTelaioRichiesta) : "-"}</div>
        <div class="report-sub">${telaioRichiestaInfo.label}</div>
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
const TODO_TELAIO_ORDER_TRACKED_TITLES = new Set(["telaio"]);
const NOTIFICATION_COMMESSA_DEADLINES = [
  { field: "data_ordine_telaio", label: "Telaio target" },
  { field: "data_arrivo_kit_cavi", label: "Arrivo kit cavi" },
  { field: "data_prelievo", label: "Prelievo materiali" },
  { field: "data_consegna_macchina", label: "Consegna macchina" },
];
const NOTIFICATION_PHASE_CLASS_KEYS = new Set(["preliminare", "3d"]);

function isTodoClientFlowTracked(titolo, repartoName = "") {
  const titleKey = normalizePhaseKey(titolo || "");
  const deptKey = normalizeDeptKey(repartoName || "");
  return TODO_CLIENT_FLOW_TRACKED_TITLES.has(titleKey) && deptKey.includes("CAD");
}

function isTodoTelaioOrderTracked(titolo, repartoName = "") {
  const titleKey = normalizePhaseKey(titolo || "");
  const deptKey = normalizeDeptKey(repartoName || "");
  return TODO_TELAIO_ORDER_TRACKED_TITLES.has(titleKey) && deptKey.includes("CAD");
}

function getTodoClientFlowSectionTitle(titolo = "") {
  const cleanTitle = String(titolo || "").trim();
  return cleanTitle ? `Cliente ${cleanTitle}` : "Cliente";
}

function getTodoTelaioOrderSectionTitle(titolo = "") {
  const cleanTitle = String(titolo || "").trim();
  return cleanTitle ? `Ordine ${cleanTitle}` : "Ordine telaio";
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

function applyMatrixBarTelaioOrderClass(bar, activity) {
  if (!bar) return;
  bar.classList.remove("is-telaio-ordered");
  if (!activity) return;
  const titleKey = normalizePhaseKey(activity.titolo || bar.dataset.attivita || "");
  if (titleKey !== "telaio") return;
  const commessaId = activity.commessa_id ?? activity.commessaId ?? bar.dataset.commessaId ?? "";
  if (!commessaId) return;
  const commessa = findCommessaInState(commessaId);
  const isOrdered =
    Boolean(commessa?.telaio_consegnato) && Boolean(parseIsoDateOnly(commessa?.data_consegna_telaio_effettiva || null));
  if (isOrdered) bar.classList.add("is-telaio-ordered");
}

function applyMatrixBarPastClass(bar, activity) {
  if (!bar) return;
  bar.classList.remove("is-past-task");
  if (!activity) return;
  const statoKey = normalizePhaseKey(activity.stato || (bar.classList.contains("is-done") ? "completata" : ""));
  if (statoKey !== "completata") return;
  const endRaw = activity.data_fine || null;
  if (!endRaw) return;
  const endDate = startOfDay(new Date(endRaw));
  if (Number.isNaN(endDate.getTime())) return;
  const today = startOfDay(new Date());
  if (endDate.getTime() < today.getTime()) bar.classList.add("is-past-task");
}

function refreshMatrixTelaioOrderIndicators(commessaId = null) {
  if (!matrixGrid) return;
  const targetId = commessaId != null ? String(commessaId) : "";
  const byId = new Map((matrixState.attivita || []).map((entry) => [String(entry.id), entry]));
  matrixGrid.querySelectorAll(".matrix-activity-bar[data-activity-id]").forEach((bar) => {
    const barCommessaId = String(bar.dataset.commessaId || "");
    if (targetId && barCommessaId !== targetId) return;
    const entry = byId.get(String(bar.dataset.activityId || ""));
    const fallback = entry || {
      commessa_id: barCommessaId,
      titolo: bar.dataset.attivita || "",
    };
    applyMatrixBarTelaioOrderClass(bar, fallback);
  });
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
    applyMatrixBarTelaioOrderClass(bar, activity);
  });
}

function getTodoClientFlowEditBlockReason(commessaId, titolo, repartoName = "") {
  if (!isTodoClientFlowTracked(titolo, repartoName)) return "Tracking cliente non disponibile per questa attivita.";
  if (!state.canWrite) return "Permessi insufficienti.";
  const role = getTodoRole();
  if (role === "admin" || role === "planner") return "";
  if (role === "responsabile") {
    const userDept = getProfileDeptKey();
    const targetDept = normalizeDeptKey(repartoName || "");
    if (!userDept || !targetDept || userDept !== targetDept) {
      return "Responsabile: puoi modificare solo il tracking del tuo reparto.";
    }
    return "";
  }
  if (role === "operatore") {
    const entries = getTodoEntriesForCell(commessaId, titolo, repartoName);
    const editableEntries = getTodoEditableEntries(entries, repartoName);
    if (!editableEntries.length) {
      return "Operatore: puoi modificare il tracking cliente solo sulle task assegnate a te.";
    }
    return "";
  }
  return "Permessi insufficienti.";
}

function getTodoTelaioOrderEditBlockReason(commessaId, titolo, repartoName = "") {
  if (!isTodoTelaioOrderTracked(titolo, repartoName)) return "Data ordine non disponibile per questa attivita.";
  if (!state.canWrite) return "Permessi insufficienti.";
  const role = getTodoRole();
  if (role !== "admin" && role !== "responsabile") {
    return "Solo admin o responsabile possono registrare la data ordine telaio.";
  }
  if (role === "responsabile") {
    const userDept = getProfileDeptKey();
    const targetDept = normalizeDeptKey(repartoName || "");
    if (!userDept || !targetDept || userDept !== targetDept) {
      return "Responsabile: puoi modificare solo il tracking del tuo reparto.";
    }
  }
  const entries = getTodoEntriesForCell(commessaId, titolo, repartoName);
  if (!entries.length) return "Nessuna attivita valida trovata.";
  return "";
}

function getTodoTelaioOrderedDate(commessaId) {
  if (!commessaId) return null;
  const key = String(commessaId);
  const commessa =
    (state.commesse || []).find((c) => String(c.id) === key) ||
    (todoLastRendered || []).find((c) => String(c.id) === key) ||
    null;
  return parseIsoDateOnly(commessa?.data_consegna_telaio_effettiva || null);
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

const TODO_GLOBAL_STATUS_ORDER = ["In corso", "Rilasciata a produzione", "Sospesa", "Evasa"];
const TODO_GLOBAL_STATUS_TO_DB = {
  "In corso": "in_corso",
  "Rilasciata a produzione": "rilasciata_produzione",
  Sospesa: "sospesa",
  Evasa: "chiusa",
};
const TODO_GLOBAL_RELEASE_DB_STATUS = "rilasciata_produzione";

function isCommessaActivityCompletedByPhaseKey(commessaId, phaseKey) {
  if (!commessaId || !phaseKey) return false;
  const idKey = String(commessaId);
  const phase = normalizePhaseKey(phaseKey);
  const reportEntries = state.reportActivitiesMap.get(idKey) || [];
  if (
    reportEntries.some(
      (entry) =>
        normalizePhaseKey(entry.titolo || "") === phase && normalizePhaseKey(entry.stato || "") === "completata"
    )
  ) {
    return true;
  }
  return (matrixState.attivita || []).some(
    (entry) =>
      String(entry.commessa_id || "") === idKey &&
      normalizePhaseKey(entry.titolo || "") === phase &&
      normalizePhaseKey(entry.stato || "") === "completata"
  );
}

function getCommessaReleaseReadiness(commessa, overrides = {}) {
  const telaioOrdered = Object.prototype.hasOwnProperty.call(overrides, "telaioOrdered")
    ? Boolean(overrides.telaioOrdered)
    : Boolean(commessa?.telaio_consegnato);
  const telaioOrderedDateRaw = Object.prototype.hasOwnProperty.call(overrides, "telaioOrderedDate")
    ? overrides.telaioOrderedDate
    : commessa?.data_consegna_telaio_effettiva || null;
  const telaioOrderedDate = parseIsoDateOnly(telaioOrderedDateRaw || null);
  const termoDone = isCommessaActivityCompletedByPhaseKey(commessa?.id, TODO_ACTIVITY_PRIMARY_PHASE_KEY);
  const missing = [];
  if (!telaioOrdered || !telaioOrderedDate) missing.push("Telaio ordinato con data");
  if (!termoDone) missing.push("Progettazione termodinamica in DONE");
  return {
    ok: missing.length === 0,
    missing,
  };
}

function getCommessaReleaseReadinessMessage(commessa, overrides = {}) {
  const readiness = getCommessaReleaseReadiness(commessa, overrides);
  if (readiness.ok) return "";
  return `Per impostare \"Rilasciata a produzione\" servono: ${readiness.missing.join(" + ")}.`;
}

function canEditTodoGlobalStatus() {
  const role = getTodoRole();
  return role === "admin" || role === "responsabile" || role === "planner";
}

function getTodoGlobalCommessaStatus(commessa) {
  const raw = normalizePhaseKey(commessa?.stato || "");
  if (raw === "nuova" || raw.includes("nuov")) return { label: "Nuova", cls: "todo-global-nuova" };
  if (raw === "in_corso") return { label: "In corso", cls: "todo-global-in-corso" };
  if (raw === "rilasciata_produzione") return { label: "Rilasciata a produzione", cls: "todo-global-rilasciata" };
  if (raw === "sospesa") return { label: "Sospesa", cls: "todo-global-sospesa" };
  if (raw === "chiusa") return { label: "Evasa", cls: "todo-global-evasa" };
  if (raw.includes("evas")) return { label: "Evasa", cls: "todo-global-evasa" };
  if (raw.includes("rilasci")) return { label: "Rilasciata a produzione", cls: "todo-global-rilasciata" };
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
  const currentDbValue = normalizePhaseKey(commessa?.stato || "");
  if (dbValue === TODO_GLOBAL_RELEASE_DB_STATUS && currentDbValue !== TODO_GLOBAL_RELEASE_DB_STATUS) {
    const msg = getCommessaReleaseReadinessMessage(commessa);
    if (msg) {
      setStatus(msg, "error");
      showToast(msg, "error");
      return false;
    }
  }
  if (dbValue === "sospesa" && currentDbValue !== "sospesa") {
    const confirmed = await openConfirmModal(
      "Confermi di mettere la commessa in stato SOSPESA?",
      null,
      { okLabel: "Sospendi", cancelLabel: "Annulla" }
    );
    if (!confirmed) return false;
  }
  const { error } = await supabase.from("commesse").update({ stato: dbValue }).eq("id", commessa.id);
  if (error) {
    const msg = formatDbErrorMessage(error);
    setStatus(`Update error: ${msg}`, "error");
    showToast(msg, "error");
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
  const orderedGroups = groups
    .map((group) => ({
      ...group,
      attivita: sortActivitiesByTodoOrder(group?.attivita || [], group?.reparto || ""),
    }))
    .sort((a, b) => {
      const aHas = (a.attivita || []).some((name) => normalizePhaseKey(name) === TODO_ACTIVITY_PRIMARY_PHASE_KEY);
      const bHas = (b.attivita || []).some((name) => normalizePhaseKey(name) === TODO_ACTIVITY_PRIMARY_PHASE_KEY);
      if (aHas !== bHas) return aHas ? -1 : 1;
      return String(a.reparto || "").localeCompare(String(b.reparto || ""));
    });

  const columns = [];
  orderedGroups.forEach((g) => {
    g.attivita.forEach((name) => {
      columns.push({ reparto: g.reparto, titolo: name });
    });
  });
  const todoFilters = getTodoFilters();
  const showPriorityColumn =
    (todoFilters.priorityMode === "asc" || todoFilters.priorityMode === "desc") &&
    Boolean(todoFilters.priorityField);
  const priorityField = showPriorityColumn ? todoFilters.priorityField : "";
  const priorityLabel = showPriorityColumn ? getTodoPriorityFieldLabel(priorityField) : "";
  const activityStartCol = showPriorityColumn ? 5 : 4;

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

  const priorityColWidth = showPriorityColumn ? "140px " : "";
  const gridCols = `240px 156px 184px ${priorityColWidth}repeat(${columns.length}, minmax(120px, 1fr))`;
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

  const machineBlank = document.createElement("div");
  machineBlank.className = "todo-cell todo-head";
  machineBlank.style.gridRow = "1";
  machineBlank.style.gridColumn = "2";
  machineBlank.textContent = "";
  headerTable.appendChild(machineBlank);

  const statusBlank = document.createElement("div");
  statusBlank.className = "todo-cell todo-head";
  statusBlank.style.gridRow = "1";
  statusBlank.style.gridColumn = "3";
  statusBlank.textContent = "";
  headerTable.appendChild(statusBlank);

  if (showPriorityColumn) {
    const priorityBlank = document.createElement("div");
    priorityBlank.className = "todo-cell todo-head";
    priorityBlank.style.gridRow = "1";
    priorityBlank.style.gridColumn = "4";
    priorityBlank.textContent = "";
    headerTable.appendChild(priorityBlank);
  }

  let colStart = activityStartCol;
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

  const machineHeader = document.createElement("div");
  machineHeader.className = "todo-cell todo-head todo-machine-head";
  machineHeader.style.gridRow = "2";
  machineHeader.style.gridColumn = "2";
  machineHeader.textContent = "Tipo macchina";
  headerTable.appendChild(machineHeader);

  const statusHeader = document.createElement("div");
  statusHeader.className = "todo-cell todo-head";
  statusHeader.style.gridRow = "2";
  statusHeader.style.gridColumn = "3";
  statusHeader.textContent = "Stato";
  headerTable.appendChild(statusHeader);

  if (showPriorityColumn) {
    const priorityHeader = document.createElement("div");
    priorityHeader.className = "todo-cell todo-head todo-priority-head";
    priorityHeader.style.gridRow = "2";
    priorityHeader.style.gridColumn = "4";
    priorityHeader.textContent = `Priorita: ${priorityLabel}`;
    headerTable.appendChild(priorityHeader);
  }

  colStart = activityStartCol;
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

    const machineCell = document.createElement("div");
    machineCell.className = "todo-cell todo-machine-cell";
    machineCell.style.gridRow = String(rowIndex);
    machineCell.style.gridColumn = "2";
    machineCell.dataset.todoRow = rowKey;
    const machineTypeRaw = String(commessa.tipo_macchina || "Altro tipo").trim() || "Altro tipo";
    const machineTypeLabel = getMachineTypeDisplayLabel(machineTypeRaw);
    machineCell.innerHTML = `<span class="todo-machine-value" title="${escapeHtml(machineTypeRaw)}">${escapeHtml(machineTypeLabel)}</span>`;
    bodyTable.appendChild(machineCell);

    const globalStatus = getTodoGlobalCommessaStatus(commessa);
    const statusCell = document.createElement("div");
    statusCell.className = "todo-cell todo-commessa-status";
    statusCell.style.gridRow = String(rowIndex);
    statusCell.style.gridColumn = "3";
    statusCell.dataset.todoRow = rowKey;
    if (canEditTodoGlobalStatus()) statusCell.classList.add("todo-commessa-status-clickable");
    statusCell.innerHTML = `<span class="todo-global-pill ${globalStatus.cls}">${globalStatus.label}</span>`;
    statusCell.addEventListener("click", async () => {
      const pill = statusCell.querySelector(".todo-global-pill");
      const rect = (pill || statusCell).getBoundingClientRect();
      openTodoGlobalStatusMenu(statusCell, commessa, rect.left + rect.width / 2, rect.bottom);
    });
    bodyTable.appendChild(statusCell);

    if (showPriorityColumn) {
      const priorityCell = document.createElement("div");
      priorityCell.className = "todo-cell todo-priority-cell";
      priorityCell.style.gridRow = String(rowIndex);
      priorityCell.style.gridColumn = "4";
      priorityCell.dataset.todoRow = rowKey;
      const value = getTodoPriorityDisplay(commessa, priorityField);
      const isEmpty = value === "n/d";
      priorityCell.innerHTML = `<span class="todo-priority-value${isEmpty ? " is-empty" : ""}">${value}</span>`;
      bodyTable.appendChild(priorityCell);
    }

    colStart = activityStartCol;
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

function formatSectionToolbarTodayLabel(date = new Date()) {
  const safeDate = date instanceof Date ? date : new Date(date);
  const dayLabel = safeDate.toLocaleDateString("it-IT", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const week = getIsoWeekNumber(safeDate);
  return `${dayLabel} - Week ${week}`;
}

function updateSectionToolbarDate() {
  if (!sectionToolbarDate) return;
  const now = new Date();
  const key = `${now.getFullYear()}-${now.getMonth()}-${now.getDate()}`;
  if (key === sectionToolbarDateKey && sectionToolbarDate.textContent) return;
  sectionToolbarDate.textContent = formatSectionToolbarTodayLabel(now);
  sectionToolbarDateKey = key;
}

function startSectionToolbarDateClock() {
  updateSectionToolbarDate();
  if (sectionToolbarDateTimer) return;
  sectionToolbarDateTimer = window.setInterval(updateSectionToolbarDate, 60 * 1000);
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
  const confirmSubtitle = panel.querySelector('.todo-status-date-confirm .todo-status-date-subtitle');
  if (!title || !dualWrap || !confirmWrap) return;

  const renderCalendar = (field, monthInput, selectedInput) => {
    const root = panel.querySelector(`.todo-status-date-calendar[data-field="${field}"]`);
    if (!root) return;
    const label = root.querySelector(".todo-status-date-label");
    const daysHost = root.querySelector(".todo-status-date-days");
    if (!label || !daysHost) return;
    const selected = selectedInput ? startOfDay(selectedInput) : null;
    const baseDate = selected || startOfDay(new Date());
    const month = monthInput
      ? new Date(monthInput.getFullYear(), monthInput.getMonth(), 1)
      : new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
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
      if (selected && cursor.getTime() === selected.getTime()) day.classList.add("selected");
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
  const isTelaioOrdered = todoStatusDatePickerState.action === "telaio_ordered";
  title.textContent = isTelaioOrdered
    ? "Data ordine telaio"
    : isSilenceDecision
    ? "Data decisione silenzio assenso"
    : "Data conferma cliente";
  if (applyBtn) {
    applyBtn.textContent = isTelaioOrdered
      ? "Conferma ordine"
      : isSilenceDecision
      ? "Conferma silenzio assenso"
      : "Conferma data";
    const telaioCommessa =
      isTelaioOrdered && todoMenuTarget?.commessaId
        ? (state.commesse || []).find((c) => String(c.id) === String(todoMenuTarget.commessaId)) || null
        : null;
    const telaioIsAlreadyOrdered =
      isTelaioOrdered &&
      Boolean(telaioCommessa?.telaio_consegnato) &&
      Boolean(getTodoTelaioOrderedDate(todoMenuTarget?.commessaId));
    applyBtn.classList.toggle("todo-order-confirm-glow", isTelaioOrdered && !telaioIsAlreadyOrdered);
    applyBtn.classList.toggle("is-warning", isTelaioOrdered && !telaioIsAlreadyOrdered);
    applyBtn.classList.toggle("is-success", isTelaioOrdered && telaioIsAlreadyOrdered);
  }
  if (confirmSubtitle) {
    confirmSubtitle.textContent = isTelaioOrdered ? "Ordine" : "Conferma";
  }
  dualWrap.classList.add("hidden");
  confirmWrap.classList.remove("hidden");
  renderCalendar("confirm", todoStatusDatePickerState.confirmMonth, todoStatusDatePickerState.confirmSelected);
}

function openTodoStatusDatePicker(action) {
  if (!todoStatusMenu || !todoMenuTarget) return;
  const panel = todoStatusMenu.querySelector(".todo-status-date-panel");
  if (!panel) return;
  const now = startOfDay(new Date());
  todoStatusDatePickerState.action = action;
  if (action === "client_sent") {
    const existing = getTodoClientFlow(todoMenuTarget.commessaId, todoMenuTarget.titolo, todoMenuTarget.reparto);
    const sent = parseIsoDateOnly(existing?.sentAt) || now;
    const due = parseIsoDateOnly(existing?.dueAt) || getDefaultClientFlowDueDate(sent) || now;
    todoStatusDatePickerState.sentSelected = sent;
    todoStatusDatePickerState.dueSelected = due;
    todoStatusDatePickerState.sentMonth = new Date(sent.getFullYear(), sent.getMonth(), 1);
    todoStatusDatePickerState.dueMonth = new Date(due.getFullYear(), due.getMonth(), 1);
  } else {
    const confirmed = action === "telaio_ordered"
      ? getTodoTelaioOrderedDate(todoMenuTarget.commessaId) || null
      : parseIsoDateOnly(getTodoClientFlow(todoMenuTarget.commessaId, todoMenuTarget.titolo, todoMenuTarget.reparto)?.confirmedAt) ||
        now;
    todoStatusDatePickerState.confirmSelected = confirmed;
    const base = confirmed || now;
    todoStatusDatePickerState.confirmMonth = new Date(base.getFullYear(), base.getMonth(), 1);
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
    if (todoStatusDatePickerState.action === "telaio_ordered" && !todoStatusDatePickerState.confirmSelected) {
      setStatus("Seleziona una data ordine telaio.", "error");
      return false;
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
  const isTelaioOrderTracked = isTodoTelaioOrderTracked(titolo, reparto);
  const canAccessTelaioOrder = isTelaioOrderTracked && currentStatus === "fatta";
  const clientFlow = isClientTracked ? getTodoClientFlow(commessaId, titolo, reparto) : null;
  const clientFlowResolvedStatus = isClientTracked ? getTodoClientFlowResolvedStatus(clientFlow, startOfDay(new Date())) : "";
  const clientFlowBlockReason = isClientTracked ? getTodoClientFlowEditBlockReason(commessaId, titolo, reparto) : "";
  const telaioOrderBlockReason = isTelaioOrderTracked
    ? getTodoTelaioOrderEditBlockReason(commessaId, titolo, reparto)
    : "";
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
  const telaioActions = [];
  const telaioOrderedDate = isTelaioOrderTracked ? getTodoTelaioOrderedDate(commessaId) : null;
  const telaioOrderedOk = Boolean(commessa?.telaio_consegnato) && Boolean(telaioOrderedDate);
  const telaioOrderedLabel = telaioOrderedOk
    ? `TELAIO ORDINATO (${formatDateDMY(telaioOrderedDate)})`
    : "TELAIO: DATA ORDINE...";
  if (canAccessTelaioOrder) {
    telaioActions.push({ action: "telaio_ordered", label: telaioOrderedLabel });
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
    if (!blockReason && def.action === "telaio_ordered" && telaioOrderBlockReason) {
      blockReason = telaioOrderBlockReason;
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
      if (def.action === "telaio_ordered") {
        button.classList.add(telaioOrderedOk ? "is-success" : "is-warning");
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
  const telaioSectionTitle = getTodoTelaioOrderSectionTitle(titolo);
  if (canAccessTelaioOrder) {
    appendSection(telaioSectionTitle, telaioActions);
  } else if (isTelaioOrderTracked) {
    appendInfoSection(telaioSectionTitle, "Data ordine disponibile solo con stato DONE.");
  }
  if (clientActions.length || telaioActions.length) {
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
  if (matrixDualPanelsActive) {
    positionMatrixDualStatusMenu();
  } else {
    positionTodoStatusMenuCentered();
    requestAnimationFrame(() => positionTodoStatusMenuCentered());
  }
}

function repositionTodoStatusMenuToAnchor() {
  if (!todoStatusMenu || todoStatusMenu.classList.contains("hidden")) return;
  if (matrixDualPanelsActive) {
    positionMatrixDualStatusMenu();
    return;
  }
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
  if (matrixDualPanelsActive) {
    if (todoFocusOverlay) todoFocusOverlay.classList.add("hidden");
    document.body.classList.remove("todo-menu-open");
    return;
  }
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
  unmountTodoStatusMenuFromDualHost();
  if (!matrixDualPanelsActive && activityModal && !activityModal.classList.contains("hidden")) {
    activityModal.classList.remove("matrix-dual-open");
  }
  if (!matrixDualPanelsActive) {
    matrixDualPanelsActive = false;
  }
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
    button.textContent = label;
    if (label === "Rilasciata a produzione") {
      const msg = getCommessaReleaseReadinessMessage(commessa);
      if (msg) {
        button.classList.add("is-blocked");
        button.disabled = true;
        button.title = msg;
      }
    }
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
  const blockReason = getTodoClientFlowEditBlockReason(target.commessaId, target.titolo, target.reparto || "");
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

async function applyTodoTelaioOrderAction(action, target, options = {}) {
  if (action !== "telaio_ordered" || !target?.commessaId) return false;
  const activityStatus = getTodoStatusFor(target.commessaId, target.titolo, target.reparto || "");
  if (activityStatus !== "fatta") {
    setStatus("Data ordine telaio disponibile solo quando l'attivita e DONE.", "error");
    return false;
  }
  const blockReason = getTodoTelaioOrderEditBlockReason(target.commessaId, target.titolo, target.reparto || "");
  if (blockReason) {
    setStatus(blockReason, "error");
    return false;
  }
  const orderedAt = parseIsoDateOnly(options.selectedDate) || startOfDay(new Date());
  const payload = {
    telaio_consegnato: true,
    data_consegna_telaio_effettiva: formatIsoDateOnly(orderedAt),
  };
  const { error } = await supabase.from("commesse").update(payload).eq("id", target.commessaId);
  if (error) {
    setStatus(`Update error: ${error.message}`, "error");
    return false;
  }
  updateCommessaInState(target.commessaId, payload);
  if (state.selected && state.selected.id === target.commessaId) {
    if (d.data_consegna_telaio_effettiva) {
      d.data_consegna_telaio_effettiva.value = payload.data_consegna_telaio_effettiva || "";
    }
    if (d.telaio_ordinato) {
      setTelaioOrdinatoButton(d.telaio_ordinato, true, false, d.data_consegna_telaio_effettiva || null);
    }
  }
  applyFilters();
  renderMatrixReport();
  setStatus("Data ordine telaio aggiornata.", "ok");
  return true;
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
    applyMatrixBarPastClass(bar, fallback);
    applyMatrixBarClientFlowClasses(bar, fallback, bar.dataset.reparto || "");
    applyMatrixBarTelaioOrderClass(bar, fallback);
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
  if (action === "telaio_ordered") {
    return applyTodoTelaioOrderAction(action, target, options);
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

function clampReportGanttRowHeight(value) {
  const raw = Number.parseInt(String(value ?? ""), 10);
  const normalized = Number.isFinite(raw) ? raw : REPORT_GANTT_ROW_HEIGHT_DEFAULT;
  return Math.min(REPORT_GANTT_ROW_HEIGHT_MAX, Math.max(REPORT_GANTT_ROW_HEIGHT_MIN, normalized));
}

function getReportGanttRowHeight() {
  if (state.reportGanttRowHeight != null) {
    return clampReportGanttRowHeight(state.reportGanttRowHeight);
  }
  let stored = null;
  try {
    stored = localStorage.getItem(REPORT_GANTT_ROW_HEIGHT_STORAGE_KEY);
  } catch {}
  const height = clampReportGanttRowHeight(stored);
  state.reportGanttRowHeight = height;
  return height;
}

function persistReportGanttRowHeight(height) {
  try {
    localStorage.setItem(REPORT_GANTT_ROW_HEIGHT_STORAGE_KEY, String(height));
  } catch {}
}

function applyReportGanttRowHeightToDom(height) {
  if (!matrixReportGrid) return;
  const rowHeight = clampReportGanttRowHeight(height);
  matrixReportGrid.style.setProperty("--report-gantt-row-height", `${rowHeight}px`);
  const rowTracks = Array.from(matrixReportGrid.querySelectorAll(".report-gantt-row .report-gantt-track"));
  rowTracks.forEach((track) => {
    const rawContent = Number.parseFloat(track.dataset.contentHeight || "");
    const contentHeight = Number.isFinite(rawContent) ? rawContent : 28;
    track.style.height = `${Math.max(rowHeight, contentHeight)}px`;
  });
}

function setReportGanttRowHeight(height, { persist = false } = {}) {
  const rowHeight = clampReportGanttRowHeight(height);
  state.reportGanttRowHeight = rowHeight;
  applyReportGanttRowHeightToDom(rowHeight);
  if (persist) persistReportGanttRowHeight(rowHeight);
  return rowHeight;
}

function stopReportGanttRowResize(pointerId = null) {
  if (!reportGanttRowResize) return;
  if (pointerId != null && reportGanttRowResize.pointerId !== pointerId) return;
  const { handle } = reportGanttRowResize;
  if (handle) {
    handle.classList.remove("is-dragging");
    try {
      handle.releasePointerCapture?.(reportGanttRowResize.pointerId);
    } catch {}
  }
  document.body.classList.remove("is-report-gantt-row-resizing");
  const current = clampReportGanttRowHeight(state.reportGanttRowHeight ?? REPORT_GANTT_ROW_HEIGHT_DEFAULT);
  state.reportGanttRowHeight = current;
  persistReportGanttRowHeight(current);
  reportGanttRowResize = null;
}

function getReportGanttScrollHost() {
  if (!matrixReportGrid) return null;
  return matrixReportGrid;
}

function setupReportGanttRowResizeHandle(handle) {
  if (!handle) return;
  handle.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const startHeight = getReportGanttRowHeight();
    const scrollHost = getReportGanttScrollHost();
    const hostRect = scrollHost ? scrollHost.getBoundingClientRect() : null;
    const pointerOffsetInHost = hostRect ? event.clientY - hostRect.top : 0;
    reportGanttRowResize = {
      pointerId: event.pointerId,
      startY: event.clientY,
      startHeight,
      handle,
      scrollHost,
      startScrollTop: scrollHost ? scrollHost.scrollTop : 0,
      pointerOffsetInHost,
    };
    handle.classList.add("is-dragging");
    document.body.classList.add("is-report-gantt-row-resizing");
    handle.setPointerCapture?.(event.pointerId);
  });
  handle.addEventListener("pointermove", (event) => {
    if (!reportGanttRowResize || reportGanttRowResize.pointerId !== event.pointerId) return;
    event.preventDefault();
    const dy = event.clientY - reportGanttRowResize.startY;
    const next = reportGanttRowResize.startHeight + dy;
    const nextHeight = setReportGanttRowHeight(next, { persist: false });
    const baseHeight = Math.max(1, reportGanttRowResize.startHeight || REPORT_GANTT_ROW_HEIGHT_DEFAULT);
    const ratio = nextHeight / baseHeight;
    const host = reportGanttRowResize.scrollHost;
    if (host && Number.isFinite(ratio)) {
      const anchor = (reportGanttRowResize.startScrollTop || 0) + (reportGanttRowResize.pointerOffsetInHost || 0);
      const nextScrollTop = anchor * ratio - (reportGanttRowResize.pointerOffsetInHost || 0);
      const maxScrollTop = Math.max(0, host.scrollHeight - host.clientHeight);
      host.scrollTop = Math.max(0, Math.min(maxScrollTop, nextScrollTop));
    }
  });
  handle.addEventListener("pointerup", (event) => {
    if (!reportGanttRowResize || reportGanttRowResize.pointerId !== event.pointerId) return;
    event.preventDefault();
    stopReportGanttRowResize(event.pointerId);
  });
  handle.addEventListener("pointercancel", (event) => {
    if (!reportGanttRowResize || reportGanttRowResize.pointerId !== event.pointerId) return;
    stopReportGanttRowResize(event.pointerId);
  });
  handle.addEventListener("lostpointercapture", () => {
    if (!reportGanttRowResize) return;
    stopReportGanttRowResize();
  });
}

function renderReportGantt(commesse, today) {
  if (!matrixReportGrid) return;
  matrixReportGrid.classList.remove("report-list");
  matrixReportGrid.classList.add("report-gantt");
  matrixReportGrid.innerHTML = "";
  const ganttRowHeight = getReportGanttRowHeight();
  matrixReportGrid.style.setProperty("--report-gantt-row-height", `${ganttRowHeight}px`);

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
    const raw = c.data_conferma_consegna_telaio || "";
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
  const headerLabel = header.querySelector(".report-gantt-label");
  if (headerLabel) {
    const headerResizeHandle = document.createElement("button");
    headerResizeHandle.type = "button";
    headerResizeHandle.className = "report-gantt-label-resize";
    headerResizeHandle.title = "Trascina su/giu per aumentare o ridurre l'altezza delle righe";
    headerResizeHandle.setAttribute("aria-label", "Ridimensiona altezza righe gantt");
    headerLabel.appendChild(headerResizeHandle);
    setupReportGanttRowResizeHandle(headerResizeHandle);
  }
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
    const rowResizeHandle = document.createElement("button");
    rowResizeHandle.type = "button";
    rowResizeHandle.className = "report-gantt-label-resize";
    rowResizeHandle.title = "Trascina su/giu per aumentare o ridurre l'altezza delle righe";
    rowResizeHandle.setAttribute("aria-label", "Ridimensiona altezza righe gantt");
    label.appendChild(rowResizeHandle);
    setupReportGanttRowResizeHandle(rowResizeHandle);
    const track = document.createElement("div");
    track.className = "report-gantt-track report-gantt-track-pan";
    track.dataset.rangeStart = formatDateLocal(rangeStart);
    track.dataset.totalDays = String(totalDays);
    track.dataset.commessaId = String(c.id);
    track.dataset.contentHeight = "28";
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
          ? `Ordine telaio (pianificato): ${formatDateDMY(ordineDay)}`
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
        effMarker.title = `Ordine telaio (pianificato): ${formatDateDMY(ordinePianificatoDay)}${
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
        kitMarker.innerHTML = '<span class="report-gantt-kit-glyph" aria-hidden="true">\uD83D\uDD0C</span>';
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
        const lanesBottomBase = 14;
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
          bar.style.bottom = `${lanesBottomBase + laneIndex * laneHeight}px`;
          bar.style.height = `${barHeight}px`;
          bar.title = `${dept.label}: ${formatDateDMY(range.start)} \u2192 ${formatDateDMY(range.end)}`;
          bar.addEventListener("click", (e) => {
            e.stopPropagation();
            showReportDeptMenu(e.clientX, e.clientY, dept.label, schedule, dept.key, c.id);
          });
          trackContent.appendChild(bar);
          laneIndex += 1;
        });
        const neededHeight = Math.max(28, lanesBottomBase + laneIndex * laneHeight + barHeight + 8);
        track.dataset.contentHeight = String(neededHeight);
      }
    }

    const contentHeightRaw = Number.parseFloat(track.dataset.contentHeight || "");
    const contentHeight = Number.isFinite(contentHeightRaw) ? contentHeightRaw : 28;
    track.style.height = `${Math.max(ganttRowHeight, contentHeight)}px`;

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
      anchorRatio: Math.min(1, Math.max(0, (e.clientX - rect.left) / Math.max(1, rect.width))),
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
    const shiftDaysFloat = reportPanZoom.mode === "pan" ? -dx * daysPerPx : 0;
    const zoomStepsFloat = reportPanZoom.mode === "zoom" ? dy / 18 : 0;
    let newDays = reportPanZoom.baseDays + zoomStepsFloat * 8;
    newDays = Math.min(REPORT_MAX_DAYS, Math.max(REPORT_MIN_DAYS, newDays));
    const anchorRatio = Number.isFinite(reportPanZoom.anchorRatio) ? reportPanZoom.anchorRatio : 0.5;
    const oldAnchorOffset = anchorRatio * Math.max(0, reportPanZoom.baseDays - 1);
    const newAnchorOffset = anchorRatio * Math.max(0, newDays - 1);
    const newStartOffset = oldAnchorOffset - newAnchorOffset + shiftDaysFloat;
    reportPanZoom.previewStartOffset = newStartOffset;
    reportPanZoom.previewDays = newDays;
    applyReportHeaderPreview(reportPanZoom, reportPanZoom.previewStartOffset, reportPanZoom.previewDays);
  };
  track.onpointerup = (e) => {
    if (!reportPanZoom) return;
    track.classList.remove("is-dragging");
    track.releasePointerCapture?.(e.pointerId);
    const drag = reportPanZoom;
    clearReportPreview(drag);
    reportPanZoom = null;
    if (!drag.moved) return;
    if (Number.isFinite(drag.previewStartOffset) || Number.isFinite(drag.previewDays)) {
      const startOffset = Number.isFinite(drag.previewStartOffset) ? drag.previewStartOffset : 0;
      let nextDays = Math.round(drag.previewDays || drag.baseDays);
      nextDays = Math.min(REPORT_MAX_DAYS, Math.max(REPORT_MIN_DAYS, nextDays));
      let nextStart = addBusinessDays(drag.baseStart, Math.round(startOffset));
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

function applyReportHeaderPreview(drag, previewStartOffset, previewDays) {
  if (!drag || !drag.previewTargets || !Number.isFinite(previewDays)) return;
  const baseDays = drag.baseDays;
  const width = Math.max(1, drag.rect.width);
  const shiftDays = Number.isFinite(previewStartOffset) ? previewStartOffset : 0;
  const scale = baseDays / Math.max(1, previewDays);
  const shiftPx = -scale * (shiftDays / Math.max(1, baseDays)) * width;
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

async function loadMatrixAttivita(options = {}) {
  if (!state.session) return;
  debugLog("loadMatrixAttivita query");
  const deferRender = Boolean(options?.deferRender) || matrixBootstrapSyncInProgress;
  if (matrixState.customWeeks != null) {
    const parsedWeeks = Number(matrixState.customWeeks);
    if (Number.isFinite(parsedWeeks) && parsedWeeks > 0) {
      matrixState.customWeeks = clampMatrixWeeks(parsedWeeks);
    } else {
      matrixState.customWeeks = null;
    }
  }
  const rawDate = matrixState.date ? new Date(matrixState.date) : new Date();
  matrixState.date = Number.isNaN(rawDate.getTime()) ? new Date() : rawDate;
  if (matrixDate) matrixDate.value = formatDateInput(matrixState.date);
  const { start, end } = getMatrixRange();
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
  if (!deferRender) {
    scheduleMatrixRenderStabilized();
  }
  removeMatrixDragFollower();
  restoreMatrixDragSourceVisual();
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

function buildTodoRealtimeSnapshotBase(targetCommessaIds = null) {
  if (Array.isArray(targetCommessaIds) && targetCommessaIds.length) {
    const ids = Array.from(
      new Set(
        targetCommessaIds
          .map((id) => String(id || "").trim())
          .filter(Boolean)
      )
    );
    if (!ids.length) return [];
    const byId = new Map();
    const register = (commessa) => {
      if (!commessa || !commessa.id) return;
      byId.set(String(commessa.id), commessa);
    };
    (state.commesse || []).forEach(register);
    (state.reportFilteredCommesse || []).forEach(register);
    (todoLastRendered || []).forEach(register);
    return ids.slice(0, REPORT_MAX_ITEMS).map((id) => byId.get(id) || { id });
  }
  const source = todoLastRendered && todoLastRendered.length ? todoLastRendered : (state.commesse || []);
  return source.slice(0, REPORT_MAX_ITEMS);
}

async function refreshTodoRealtimeSnapshot(targetCommessaIds = null) {
  if (!state.session) return;
  const base = buildTodoRealtimeSnapshotBase(targetCommessaIds);
  if (!base.length) return;
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
    const role = String(state.profile?.ruolo || "").trim().toLowerCase();
    const machineType = normalizeMachineTypeValue(
      d.tipo_macchina?.value || state.selected?.tipo_macchina || ""
    );
    const machineVariant = sanitizeMachineVariantForType(
      machineType,
      d.variante_macchina?.value || state.selected?.variante_macchina || "standard"
    );
    if (role !== "planner" && !machineType) {
      setStatus("Tipo macchina obbligatorio.", "error");
      if (d.tipo_macchina) d.tipo_macchina.focus();
      return;
    }
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
    if (d.stato.value === TODO_GLOBAL_RELEASE_DB_STATUS) {
      const msg = getCommessaReleaseReadinessMessage(state.selected, {
        telaioOrdered: telaioConsegnato,
        telaioOrderedDate: telaioEffettivo,
      });
      if (msg) {
        setStatus(msg, "error");
        showToast(msg, "error");
        return;
      }
    }
    const consegnaValue = d.data_consegna.value || null;
    const payload = {
      codice: normalized.codice,
      anno: normalized.anno,
      numero: normalized.numero,
      titolo: d.titolo.value.trim() || null,
      cliente: d.cliente.value.trim() || null,
      tipo_macchina: machineType || "Altro tipo",
      variante_macchina: machineVariant,
      stato: d.stato.value,
      priorita: d.priorita.value || null,
      data_ingresso: d.data_ingresso.value || null,
      data_ordine_telaio: d.data_ordine_telaio ? d.data_ordine_telaio.value || null : null,
      data_conferma_consegna_telaio: d.data_conferma_consegna_telaio
        ? d.data_conferma_consegna_telaio.value || null
        : null,
      data_arrivo_kit_cavi: d.data_arrivo_kit_cavi ? d.data_arrivo_kit_cavi.value || null : null,
      data_prelievo: d.data_prelievo_materiali ? d.data_prelievo_materiali.value || null : null,
      telaio_consegnato: telaioConsegnato,
      data_consegna_telaio_effettiva: telaioEffettivo,
      data_consegna_prevista: consegnaValue,
      data_consegna_macchina: consegnaValue,
      note_generali: d.note.value.trim() || null,
    };
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
        const msg = formatDbErrorMessage(error);
        setStatus(`Update error: ${msg}`, "error");
        showToast(msg, "error");
        return;
      }
      updateCommessaInState(state.selected.id, payload);
      if (Object.prototype.hasOwnProperty.call(payload, "data_conferma_consegna_telaio")) {
        await warnTelaioPlannedMismatch(state.selected.id, payload.data_conferma_consegna_telaio);
      }
    }
    applyFilters();
    renderMatrixReport();
    setDetailSnapshot();
    setStatus("Commessa updated.", "ok");
    if (closeOnSuccess) {
      closeCommessaDetailModal();
    }
    await loadCommesse();
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
  const role = String(state.profile?.ruolo || "").trim().toLowerCase();
  const machineType = normalizeMachineTypeValue(n.tipo_macchina?.value || "");
  const machineVariant = sanitizeMachineVariantForType(
    machineType,
    n.variante_macchina?.value || "standard"
  );
  if (!machineType) {
    setStatus("Seleziona una tipologia macchina.", "error");
    if (n.tipo_macchina) n.tipo_macchina.focus();
    return;
  }
  const plannerTriedTelaioOrdered =
    role === "planner" &&
    (Boolean(n.telaio_consegnato?.checked) ||
      Boolean((n.data_consegna_telaio_effettiva ? n.data_consegna_telaio_effettiva.value : "").trim()));
  if (plannerTriedTelaioOrdered) {
    const msg = "Planner: in creazione non puoi impostare Stato telaio e Data telaio ordinato.";
    setStatus(msg, "error");
    showToast(msg, "error");
    if (n.data_consegna_telaio_effettiva) n.data_consegna_telaio_effettiva.focus();
    return;
  }
  const telaioEffettivo =
    n.data_consegna_telaio_effettiva ? n.data_consegna_telaio_effettiva.value || null : null;
  const telaioConsegnato = Boolean(n.telaio_consegnato?.checked);
  if (telaioConsegnato && !telaioEffettivo) {
    setStatus("Set a planned frame order date before marking as ordered.", "error");
    return;
  }
  if (n.stato.value === TODO_GLOBAL_RELEASE_DB_STATUS) {
    setStatus(
      'In creazione non puoi impostare "Rilasciata a produzione": servono Telaio ordinato con data e Progettazione termodinamica DONE.',
      "error"
    );
    return;
  }
  const consegnaValue = n.data_consegna.value || null;
  const payload = {
    codice: normalized.codice,
    anno: normalized.anno,
    numero: normalized.numero,
    titolo: n.titolo.value.trim(),
    cliente: n.cliente.value.trim() || null,
    tipo_macchina: machineType,
    variante_macchina: machineVariant,
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
    const msg = formatDbErrorMessage(error);
    setStatus(`Create error: ${msg}`, "error");
    showToast(msg, "error");
    return;
  }
  newForm.reset();
  refreshMachineTypeSelects({
    newType: "",
    newVariant: "standard",
  });
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
  resetImportFeedback();
  const { rows, errors } = parseImportRows(importTextarea.value);
  if (!rows.length) {
    const msg = "No valid rows to import.";
    setStatus(msg, "error");
    showImportResult(msg, "error");
    return;
  }
  if (errors.length) {
    const msg = `Import blocked: fix the errors (e.g. ${errors[0]}).`;
    setStatus(msg, "error");
    showImportResult(msg, "error");
    return;
  }

  if (importCommitBtn) {
    importCommitBtn.disabled = true;
    importCommitBtn.textContent = "Import 0%";
  }

  try {
    const chunks = chunkArray(rows, IMPORT_CHUNK_SIZE);
    let imported = 0;
    updateImportProgress(2, "Preparazione...");

    for (let i = 0; i < chunks.length; i += 1) {
      const chunk = chunks[i];
      const startPct = Math.round((imported / rows.length) * 100);
      updateImportProgress(startPct, `Blocco ${i + 1}/${chunks.length}`);

      const { error } = await supabase
        .from("commesse")
        .upsert(chunk, { onConflict: "codice" })
        .select("codice");

      if (error) {
        const msg = `Import error: ${formatDbErrorMessage(error)}`;
        setStatus(msg, "error");
        showImportResult(msg, "error");
        importPreview.textContent = msg;
        console.error("Import error", error);
        return;
      }

      imported += chunk.length;
      const pct = Math.round((imported / rows.length) * 100);
      updateImportProgress(pct, `${imported}/${rows.length} righe`);
      if (importCommitBtn) {
        importCommitBtn.textContent = `Import ${pct}%`;
      }
    }

    updateImportProgress(100, "Completato");
    const okMsg = `Import completed: ${rows.length} rows.`;
    setStatus(okMsg, "ok");
    showImportResult(okMsg, "ok");
    showToast(okMsg, "ok");
    await loadCommesse();
    window.setTimeout(() => {
      importTextarea.value = "";
      closeImportModal();
      updateImportPreview();
    }, 900);
  } catch (err) {
    const msg = `Import error: ${formatDbErrorMessage(err)}`;
    setStatus(msg, "error");
    showImportResult(msg, "error");
  } finally {
    if (importCommitBtn) {
      importCommitBtn.disabled = false;
      importCommitBtn.textContent = "Importa";
    }
  }
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
  const includesTelaio = selected.some((title) => isTelaioActivityTitle(title));

  if (commessaId && selected.some((title) => isTelaioActivityTitle(title))) {
    const ok = await confirmTelaioDuplicateAssignment(commessaId);
    if (!ok) return;
  }

  const { error } = await supabase.from("attivita").insert(rows);
  if (error) {
    setStatus(`Assignment error: ${error.message}`, "error");
    return;
  }
  let telaioPlannedSynced = true;
  if (includesTelaio) {
    telaioPlannedSynced = await syncTelaioPlannedDateFromMatrix(commessaId, endAt);
  }
  if (!telaioPlannedSynced) {
    setStatus("Activities assigned, ma data telaio pianificata non aggiornata.", "error");
  } else {
    setStatus("Activities assigned.", "ok");
  }
  closeAssignModal();
  await loadMatrixAttivita();
  await refreshTodoRealtimeSnapshot([commessaId]);
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
if (todoSyncMatrixFiltersBtn) {
  updateTodoSyncMatrixFiltersButton();
  todoSyncMatrixFiltersBtn.addEventListener("click", handleTodoQuickSyncFromMatrixFilters);
}
if (todoYearFilter) todoYearFilter.addEventListener("change", renderMatrixReportFromTodoFilters);
if (todoPriorityField) todoPriorityField.addEventListener("change", renderMatrixReportFromTodoFilters);
if (todoPriorityBtn) {
  renderTodoPriorityButtonLabel();
  todoPriorityBtn.addEventListener("click", () => {
    const mode = todoPriorityBtn.dataset.mode || "off";
    const next = mode === "off" ? "asc" : mode === "asc" ? "desc" : "off";
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
  d.telaio_ordinato.addEventListener("click", async () => {
    if (!canEditTelaioOrderByRole()) {
      const msg = "Solo admin e responsabile possono modificare lo stato ordine telaio.";
      setStatus(msg, "error");
      showToast(msg, "error");
      return;
    }
    const isOrdered = d.telaio_ordinato.dataset.ordered === "true";
    const plannedDate = d.data_consegna_telaio_effettiva ? d.data_consegna_telaio_effettiva.value || "" : "";
    if (isOrdered) {
      const confirmed = await openConfirmModal(
        "Il telaio risulta ORDINATO. Passare a NON ORDINATO e azzerare la data ordine?",
        null,
        {
          okLabel: "Conferma",
          cancelLabel: "Annulla",
        }
      );
      if (!confirmed) return;
      if (d.data_consegna_telaio_effettiva) {
        d.data_consegna_telaio_effettiva.value = "";
      }
      setTelaioOrdinatoButton(
        d.telaio_ordinato,
        false,
        true,
        d.data_consegna_telaio_effettiva || null
      );
      updateDetailDirty();
      return;
    }
    if (!isOrdered && !plannedDate) {
      if (d.data_consegna_telaio_effettiva) {
        d.data_consegna_telaio_effettiva.classList.add("is-missing-required");
        setTimeout(() => d.data_consegna_telaio_effettiva.classList.remove("is-missing-required"), 1600);
        d.data_consegna_telaio_effettiva.focus();
        if (typeof d.data_consegna_telaio_effettiva.showPicker === "function") {
          try {
            d.data_consegna_telaio_effettiva.showPicker();
          } catch (_err) {
            // Ignore unsupported showPicker errors.
          }
        }
      }
      const msg = "Seleziona prima una data ordine telaio, poi conferma lo stato ORDINATO.";
      setStatus(msg, "error");
      showToast(msg, "error");
      return;
    }
    const plannedDateObj = parseIsoDateOnly(plannedDate);
    const plannedDateLabel = plannedDateObj ? formatDateDMY(plannedDateObj) : plannedDate;
    const confirmOrdered = await openConfirmModal(
      `Confermare TELAIO ORDINATO con data ${plannedDateLabel}?`,
      null,
      {
        okLabel: "Conferma ordine",
        cancelLabel: "Annulla",
      }
    );
    if (!confirmOrdered) return;
    setTelaioOrdinatoButton(d.telaio_ordinato, true, true, d.data_consegna_telaio_effettiva || null);
    updateDetailDirty();
  });
}
if (d.data_consegna_telaio_effettiva) {
  d.data_consegna_telaio_effettiva.addEventListener("input", () => {
    if (!d.telaio_ordinato || d.telaio_ordinato.dataset.ordered !== "true") return;
    setTelaioOrdinatoButton(d.telaio_ordinato, true, false, d.data_consegna_telaio_effettiva || null);
  });
  d.data_consegna_telaio_effettiva.addEventListener("change", () => {
    if (!d.telaio_ordinato || d.telaio_ordinato.dataset.ordered !== "true") return;
    setTelaioOrdinatoButton(d.telaio_ordinato, true, false, d.data_consegna_telaio_effettiva || null);
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
if (n.tipo_macchina) {
  n.tipo_macchina.addEventListener("change", () => {
    setMachineVariantSelectOptions(
      n.variante_macchina,
      n.tipo_macchina.value,
      n.variante_macchina?.value || "standard"
    );
  });
}
if (d.tipo_macchina) {
  d.tipo_macchina.addEventListener("change", () => {
    setMachineVariantSelectOptions(
      d.variante_macchina,
      d.tipo_macchina.value,
      d.variante_macchina?.value || "standard"
    );
    updateDetailDirty();
  });
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
if (commessePanel) {
  const toggle = () => {
    commessePanel.classList.toggle("collapsed");
    const isCollapsed = commessePanel.classList.contains("collapsed");
    if (commesseToggleBtn) {
      commesseToggleBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
    }
  };
  if (commesseToggleBtn) commesseToggleBtn.addEventListener("click", toggle);
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
if (openNotificationsBtn) {
  openNotificationsBtn.addEventListener("click", openNotificationsModal);
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
if (closeNotificationsBtn) {
  closeNotificationsBtn.addEventListener("click", closeNotificationsModal);
}
if (refreshNotificationsBtn) {
  refreshNotificationsBtn.addEventListener("click", () => {
    void refreshNotificationsPanel();
  });
}
if (notificationsOnlyMineBtn) {
  notificationsOnlyMineBtn.addEventListener("click", () => {
    notificationsOnlyMine = !notificationsOnlyMine;
    updateNotificationsOnlyMineButton();
    renderNotificationsFromCache();
  });
}
if (notificationsModal) {
  notificationsModal.addEventListener("click", (e) => {
    if (e.target === notificationsModal) closeNotificationsModal();
  });
}
if (progressYearCurrent) {
  progressYearCurrent.addEventListener("click", () => {
    progressYearCurrent.classList.add("active");
    if (progressYearPrev) progressYearPrev.classList.remove("active");
    renderProgressResults();
  });
}
if (progressYearPrev) {
  progressYearPrev.addEventListener("click", () => {
    progressYearPrev.classList.add("active");
    if (progressYearCurrent) progressYearCurrent.classList.remove("active");
    renderProgressResults();
  });
}
  if (progressSearch) {
    progressSearch.addEventListener("input", () => {
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
if (copyImportHeaderBtn) {
  copyImportHeaderBtn.addEventListener("click", () => {
    void handleCopyImportHeader(IMPORT_HEADER_BASE, "Importa da Excel");
  });
}
if (copyImportDatesHeaderBtn) {
  copyImportDatesHeaderBtn.addEventListener("click", () => {
    void handleCopyImportHeader(IMPORT_HEADER_DATES, "Importa date commesse");
  });
}
if (copyImportProductionHeaderBtn) {
  copyImportProductionHeaderBtn.addEventListener("click", () => {
    void handleCopyImportHeader(IMPORT_HEADER_PRODUCTION, "Importa date produzione");
  });
}
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
    closeNotificationsModal();
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
  matrixState.panResidualDays = 0;
  matrixState.date = new Date(e.target.value);
  await loadMatrixAttivita();
});
matrixPrevBtn.addEventListener("click", async () => {
  matrixState.panResidualDays = 0;
  matrixState.date = addDays(matrixState.date, -7);
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixNextBtn.addEventListener("click", async () => {
  matrixState.panResidualDays = 0;
  matrixState.date = addDays(matrixState.date, 7);
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixTodayBtn.addEventListener("click", async () => {
  matrixState.panResidualDays = 0;
  matrixState.date = new Date();
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixZoomInBtn.addEventListener("click", async () => {
  if (!MATRIX_ZOOM_BUTTONS_ENABLED) return;
  matrixState.customWeeks = null;
  if (matrixState.view === "six") {
    matrixState.view = "three";
  } else if (matrixState.view === "three") {
    matrixState.view = "two";
  } else if (matrixState.view === "two") {
    matrixState.view = "week";
  } else {
    matrixState.view = "week";
  }
  matrixState.panResidualDays = 0;
  matrixState.date = startOfWeek(matrixState.date);
  matrixDate.value = formatDateInput(matrixState.date);
  setMatrixViewLabel();
  await loadMatrixAttivita();
});
matrixZoomOutBtn.addEventListener("click", async () => {
  if (!MATRIX_ZOOM_BUTTONS_ENABLED) return;
  matrixState.customWeeks = null;
  if (matrixState.view === "week") {
    matrixState.view = "two";
  } else if (matrixState.view === "two") {
    matrixState.view = "three";
  } else if (matrixState.view === "three") {
    matrixState.view = "six";
  }
  matrixState.panResidualDays = 0;
  matrixState.date = startOfWeek(matrixState.date);
  matrixDate.value = formatDateInput(matrixState.date);
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
    updateMatrixDropEffect(e);
    updateMatrixDragFollowerPosition(e.clientX, e.clientY);
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
    markMatrixDragDropHandled();
    matrixTrash.classList.remove("is-dragover");
    if (!dragId) return;
    await deleteActivityById(dragId);
  });
}
if (matrixCommessaPickerToggleBtn) {
  matrixCommessaPickerToggleBtn.addEventListener("click", () => {
    if (isMatrixQuickCommessaFilterActive()) {
      setStatus('Disattiva il filtro rapido "numero comm." per usare "Filtra commesse".', "error");
      matrixQuickCommessaFilter?.focus();
      matrixQuickCommessaFilter?.select?.();
      return;
    }
    openMatrixFilter("commessa");
  });
}
if (matrixCommessaClearBtn) {
  matrixCommessaClearBtn.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    setMatrixSelectedCommesse([]);
    renderMatrix();
    setStatus("Filtro commesse rimosso.", "ok");
  });
}
if (matrixAttivitaToggleBtn) {
  matrixAttivitaToggleBtn.addEventListener("click", () => openMatrixFilter("attivita"));
}
if (matrixAttivitaClearBtn) {
  matrixAttivitaClearBtn.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    matrixState.selectedAttivita = new Set();
    updateMatrixAttivitaPickerLabel();
    await loadMatrixAttivita();
    setStatus("Filtro attivita rimosso.", "ok");
  });
}
matrixToggleListBtn.addEventListener("click", () => {
  if (!matrixCommessePanel) return;
  matrixCommessePanel.classList.toggle("collapsed");
  matrixToggleListBtn.textContent = matrixCommessePanel.classList.contains("collapsed")
    ? "Espandi lista"
    : "Comprimi lista";
});
matrixCommessaSearch.addEventListener("input", renderMatrixCommesseColumn);
if (matrixCommessaYear) {
  matrixCommessaYear.addEventListener("change", renderMatrixCommesseColumn);
}
if (matrixQuickCommessaFilter) {
  matrixQuickCommessaFilter.addEventListener("input", (event) => {
    const raw = event?.target?.value ?? "";
    setMatrixQuickCommessaFilterRaw(raw);
  });
}
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
if (matrixFilterNumero) {
  matrixFilterNumero.addEventListener("input", renderMatrixFilterList);
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
    if (matrixFilterNumero) matrixFilterNumero.value = "";
    if (matrixFilterYear) {
      matrixFilterYear.value = matrixState.filterMode === "commessa" ? "2026" : "";
    }
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
if (activityModalDragHandle) {
  activityModalDragHandle.addEventListener("pointerdown", startActivityModalDrag);
  activityModalDragHandle.addEventListener("pointermove", moveActivityModalDrag);
  activityModalDragHandle.addEventListener("pointerup", (event) => stopActivityModalDrag(event.pointerId));
  activityModalDragHandle.addEventListener("pointercancel", (event) => stopActivityModalDrag(event.pointerId));
}
if (activitySaveBtn) {
  activitySaveBtn.addEventListener("click", saveActivityDuration);
}
if (activityHours1Btn && activityDurationInput) {
  activityHours1Btn.addEventListener("click", () => {
    activityDurationInput.value = "1";
    updateActivityQuickHourButtons();
    activityDurationInput.focus();
  });
}
if (activityHours4Btn && activityDurationInput) {
  activityHours4Btn.addEventListener("click", () => {
    activityDurationInput.value = "4";
    updateActivityQuickHourButtons();
    activityDurationInput.focus();
  });
}
if (activityHours8Btn && activityDurationInput) {
  activityHours8Btn.addEventListener("click", () => {
    activityDurationInput.value = "8";
    updateActivityQuickHourButtons();
    activityDurationInput.focus();
  });
}
if (activityDurationInput) {
  activityDurationInput.addEventListener("input", updateActivityQuickHourButtons);
  activityDurationInput.addEventListener("change", updateActivityQuickHourButtons);
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
    let payload = {};
    if (field === "ordine") payload.data_ordine_telaio = null;
    if (field === "kit_cavi") payload.data_arrivo_kit_cavi = null;
    if (field === "prelievo") payload.data_prelievo = null;
    if (field === "consegna") payload.data_consegna_macchina = null;
    if (field === "ordine_pianificato") payload.data_conferma_consegna_telaio = null;
    const consistency = ensureTelaioOrderPayloadConsistency(commessaId, payload, {
      focusDateField: true,
      toast: true,
    });
    if (!consistency.ok) {
      milestonePickerState.saving = false;
      return;
    }
    payload = consistency.payload;
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
  const ctx = matrixState.resizing;
  if (!ctx) return;
  const targetCell = resolveMatrixResizeTargetCell(ctx, e.clientX);
  if (!targetCell || !targetCell.dataset?.day) return;
  const dayKey = targetCell.dataset.day;
  if (dayKey < ctx.startDayKey) return;
  ctx.targetDayKey = dayKey;
  setMatrixResizeTargetCell(ctx, targetCell);
  if (ctx.previewBar && ctx.layerRect && ctx.startCellRect) {
    const left = ctx.startCellRect.left - ctx.layerRect.left;
    const maxRight = ctx.layerRect.right - ctx.layerRect.left;
    const mouseX = e.clientX - ctx.layerRect.left;
    const clampedX = Math.max(left, Math.min(mouseX, maxRight));
    const width = Math.max(0, clampedX - left);
    ctx.previewBar.style.left = `${left}px`;
    ctx.previewBar.style.width = `${width}px`;
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
  if (!matrixState.dragDropHandled) {
    removeMatrixDragFollower();
    restoreMatrixDragSourceVisual();
  }
  document.querySelectorAll(".matrix-activity-bar.dragging").forEach((el) => el.classList.remove("dragging"));
  closeMatrixQuickMenu();
  if (matrixTrash) matrixTrash.classList.remove("is-dragover");
});
document.addEventListener("drop", () => {
  if (matrixGrid) matrixGrid.classList.remove("matrix-dragging");
  matrixState.draggingId = null;
  if (!matrixState.dragDropHandled) {
    removeMatrixDragFollower();
    restoreMatrixDragSourceVisual();
  }
  document.querySelectorAll(".matrix-activity-bar.dragging").forEach((el) => el.classList.remove("dragging"));
  closeMatrixQuickMenu();
  if (matrixTrash) matrixTrash.classList.remove("is-dragover");
});
document.addEventListener("dragover", (e) => {
  if (!matrixState.dragFollowerEl) return;
  updateMatrixDragFollowerPosition(e.clientX, e.clientY);
});
sectionLinks.forEach((btn) => {
  btn.addEventListener("click", () => {
    const targetId = btn.dataset.target;
    if (!targetId) return;
    const ensureSectionExpanded = (panelId) => {
      const panel = document.getElementById(panelId);
      if (!panel) return null;
      if (!panel.classList.contains("collapsed")) return panel;
      panel.classList.remove("collapsed");
      if (panelId === "commessePanel" && commesseToggleBtn) {
        commesseToggleBtn.setAttribute("aria-expanded", "true");
      }
      if (panelId === "reportPanel" || panelId === "todoSection") {
        renderMatrixReport();
      }
      return panel;
    };
    const target = ensureSectionExpanded(targetId);
    if (!target) return;
    setActiveSectionLink(targetId);
    window.requestAnimationFrame(() => {
      const rect = target.getBoundingClientRect();
      const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
      const top = rect.top + window.scrollY - toolbarOffset - getSectionToolbarGap();
      window.scrollTo({ top: Math.max(0, top), left: 0, behavior: "smooth" });
    });
  });
});
const sectionSpyLinks = sectionLinks.filter((btn) => {
  const targetId = btn.dataset.target;
  return Boolean(targetId && document.getElementById(targetId));
});
let sectionSpyRafId = 0;
const setActiveSectionLink = (targetId) => {
  sectionSpyLinks.forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.target === targetId);
  });
};
const getActiveSectionTargetId = () => {
  if (!sectionSpyLinks.length) return null;
  const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
  const probeY = toolbarOffset + getSectionToolbarGap() + 10;
  let current = null;
  let next = null;
  sectionSpyLinks.forEach((btn) => {
    const targetId = btn.dataset.target;
    const target = targetId ? document.getElementById(targetId) : null;
    if (!target) return;
    const rect = target.getBoundingClientRect();
    if (rect.top <= probeY) current = targetId;
    else if (!next) next = targetId;
  });
  return current || next || sectionSpyLinks[sectionSpyLinks.length - 1].dataset.target || null;
};
const updateActiveSectionLinkFromScroll = () => {
  const targetId = getActiveSectionTargetId();
  if (targetId) setActiveSectionLink(targetId);
};
const scheduleSectionSpyUpdate = () => {
  if (sectionSpyRafId) return;
  sectionSpyRafId = window.requestAnimationFrame(() => {
    sectionSpyRafId = 0;
    updateActiveSectionLinkFromScroll();
  });
};
const updateSectionToolbarOffset = () => {
  const toolbarOffset = sectionToolbar ? sectionToolbar.getBoundingClientRect().height : 0;
  document.documentElement.style.setProperty("--section-toolbar-h", `${toolbarOffset}px`);
  updateTodoPanelHeaderOffset();
  scheduleSectionSpyUpdate();
};
updateSectionToolbarOffset();
startSectionToolbarDateClock();
window.addEventListener("resize", updateSectionToolbarOffset);
window.addEventListener("scroll", scheduleSectionSpyUpdate, { passive: true });
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
      const prevTarget = todoMenuTarget ? { ...todoMenuTarget } : null;
      const prevManual = todoStatusMenuManualPosition;
      const prevLeft = todoStatusMenu?.style.left || "";
      const prevTop = todoStatusMenu?.style.top || "";
      const shouldClose = await handleTodoStatusDatePickerAction(action, target);
      if (shouldClose) {
        const targetCell = prevTarget
          ? findTodoCellByTarget(prevTarget.commessaId, prevTarget.titolo, prevTarget.reparto || "")
          : null;
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
      }
      return;
    }
    if (target.dataset.blocked === "1") {
      setStatus(target.dataset.blockReason || "Azione non disponibile.", "error");
      return;
    }
    if (action === "client_sent" || action === "client_confirmed" || action === "client_silence" || action === "telaio_ordered") {
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
  positionMatrixDualStatusMenu();
  keepActivityModalInViewport();
});
window.addEventListener("focus", refreshTodoVisualizationNow);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshTodoVisualizationNow();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeNotificationsModal();
    if (activityModal && !activityModal.classList.contains("hidden")) {
      e.preventDefault();
      closeTodoStatusMenu();
      closeActivityModal();
      return;
    }
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
    const insideDualModal = matrixDualPanelsActive && Boolean(e.target.closest("#activityModal"));
    if (!e.target.closest("#todoStatusMenu") && !insideDualModal) {
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
    clearCommessaHighlight({ preserveViewport: true });
  }
});

setMatrixColorMode("commessa");
setMatrixAutoShift(false);
resetMatrixMovementBaseline();
applyMatrixMovementControlsState();

supabase.auth.onAuthStateChange(async (event, session) => {
  debugLog(`auth event: ${event} session=${session ? "yes" : "no"}`);
  if (session) {
    const sameSessionToken =
      Boolean(session.access_token) &&
      Boolean(state.session?.access_token) &&
      session.access_token === state.session.access_token;
    const sameUser =
      Boolean(session.user?.id) &&
      Boolean(state.profile?.id) &&
      String(session.user.id) === String(state.profile.id);
    if ((event === "INITIAL_SESSION" || event === "SIGNED_IN") && sameSessionToken && sameUser) {
      debugLog(`auth event ignored (duplicate ${event})`);
      return;
    }
    await syncSignedIn(session);
  } else {
    const hadSession = Boolean(state.session);
    if (event === "SIGNED_OUT") {
      syncSignedOut({ showStatus: hadSession });
      return;
    }
    if (!hadSession) return;
    syncSignedOut({ showStatus: true });
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
