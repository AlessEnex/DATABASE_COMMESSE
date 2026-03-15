const baseRuleTemplate = [
  { title: "CAD layout generale", dept: "CAD", hours: null, depends: "inizio" },
  { title: "Disegni costruttivi", dept: "CAD", hours: null, depends: "layout" },
  { title: "Distinte base", dept: "CAD", hours: null, depends: "layout" },
  { title: "Progetto termodinamico", dept: "Termodinamici", hours: null, depends: "layout" },
  { title: "Schema elettrico", dept: "Elettrici", hours: null, depends: "layout" },
];

const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";
const THEME_KEY = "commesse_theme";

const motoreSupabase = window.supabase
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      auth: { persistSession: true, autoRefreshToken: true },
    })
  : null;

const gate = document.getElementById("motoreAccessGate");
const gateTitle = document.getElementById("motoreGateTitle");
const gateText = document.getElementById("motoreGateText");
const ruleWarnings = document.getElementById("motoreRuleWarnings");
const motoreValidationSummary = document.getElementById("motoreValidationSummary");
const motoreValidationList = document.getElementById("motoreValidationList");
const motoreSandboxState = document.getElementById("motoreSandboxState");
const motoreSandboxToggleBtn = document.getElementById("motoreSandboxToggleBtn");
const motoreSandboxHint = document.getElementById("motoreSandboxHint");
const motoreThemeToggleBtn = document.getElementById("motoreThemeToggleBtn");

let activeMachineId = "";
let templateLoaded = false;
let metaLoaded = false;
let rulesLoadToken = 0;
let motoreUserId = null;
let motoreValidationItemsCache = [];
const motoreTipologie = new Map();
const motoreVariantIds = new Map();

function showGate(title, text) {
  if (gateTitle) gateTitle.textContent = title;
  if (gateText) gateText.textContent = text;
  document.body.classList.add("motore-locked");
  if (gate) gate.classList.remove("hidden");
}

function applyMotoreTheme(theme) {
  const isDark = theme === "dark";
  document.body.classList.toggle("theme-dark", isDark);
  if (!motoreThemeToggleBtn) return;
  motoreThemeToggleBtn.textContent = isDark ? "Dark" : "Light";
  motoreThemeToggleBtn.title = isDark ? "Tema notte" : "Tema giorno";
  motoreThemeToggleBtn.setAttribute("aria-label", isDark ? "Tema notte" : "Tema giorno");
}

function initMotoreTheme() {
  const stored = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = stored || (prefersDark ? "dark" : "light");
  applyMotoreTheme(theme);
}

function toggleMotoreTheme() {
  const isDark = document.body.classList.contains("theme-dark");
  const next = isDark ? "light" : "dark";
  localStorage.setItem(THEME_KEY, next);
  applyMotoreTheme(next);
}

function hideGate() {
  document.body.classList.remove("motore-locked");
  if (gate) gate.classList.add("hidden");
}

async function checkMotoreAccess() {
  motoreUserId = null;
  if (!motoreSupabase) {
    showGate("Accesso non disponibile", "Supabase non disponibile.");
    return;
  }
  const { data: sessionData } = await motoreSupabase.auth.getSession();
  const session = sessionData?.session || null;
  if (!session) {
    showGate("Accesso richiesto", "Effettua il login nella app per continuare.");
    return;
  }
  const { data: userData } = await motoreSupabase.auth.getUser();
  const user = userData?.user || session.user;
  if (!user) {
    showGate("Accesso richiesto", "Effettua il login nella app per continuare.");
    return;
  }

  let profile = null;
  try {
    const result = await motoreSupabase.from("utenti").select("ruolo, email").eq("id", user.id).single();
    if (result.error) throw result.error;
    profile = result.data;
  } catch {
    try {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/utenti?id=eq.${user.id}&select=ruolo,email`,
        {
          headers: {
            apikey: SUPABASE_ANON_KEY,
            Authorization: `Bearer ${session.access_token}`,
          },
        }
      );
      if (res.ok) {
        const data = await res.json();
        profile = Array.isArray(data) ? data[0] : null;
      }
    } catch {
      profile = null;
    }
  }

  if (!profile) {
    showGate("Accesso negato", "Profilo non trovato o non autorizzato.");
    return;
  }
  const role = String(profile.ruolo || "").trim().toLowerCase();
  if (role !== "admin") {
    const shown = profile.ruolo ? `Ruolo attuale: ${profile.ruolo}` : "Ruolo non impostato.";
    showGate("Accesso negato", `Solo gli admin possono accedere al motore. ${shown}`);
    return;
  }

  motoreUserId = user.id;
  hideGate();
  await initMotoreTemplate();
}

const machineData = [
  { id: "tago", name: "TAGO", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  { id: "minibooster-swiss", name: "MINIBOOSTER SWISS", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  { id: "drava", name: "DRAVA", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  {
    id: "senna",
    name: "SENNA",
    variants: [
      {
        id: "standard",
        label: "Standard",
        rules: [
          { title: "CAD layout generale", dept: "CAD", hours: 32, depends: "inizio" },
          { title: "Distinte base", dept: "CAD", hours: 16, depends: "layout" },
          { title: "Progetto termodinamico", dept: "Termodinamici", hours: 32, depends: "layout" },
          { title: "Schema elettrico", dept: "Elettrici", hours: 24, depends: "layout" },
        ],
      },
    ],
  },
  { id: "senna-xs", name: "SENNA-XS", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  {
    id: "neva",
    name: "NEVA",
    variants: [
      {
        id: "standard",
        label: "Standard",
        rules: [
          { title: "CAD layout generale", dept: "CAD", hours: 16, depends: "inizio" },
          { title: "Disegni costruttivi", dept: "CAD", hours: 24, depends: "layout" },
          { title: "Progetto termodinamico", dept: "Termodinamici", hours: 16, depends: "layout" },
          { title: "Schema elettrico", dept: "Elettrici", hours: 16, depends: "layout" },
        ],
      },
      { id: "custom", label: "Custom", rules: [] },
    ],
  },
  {
    id: "elba",
    name: "ELBA",
    variants: [
      {
        id: "standard",
        label: "Standard",
        rules: [
          { title: "CAD layout generale", dept: "CAD", hours: 24, depends: "inizio" },
          { title: "Disegni costruttivi", dept: "CAD", hours: 32, depends: "layout" },
          { title: "Progetto termodinamico", dept: "Termodinamici", hours: 24, depends: "layout" },
          { title: "Schema elettrico", dept: "Elettrici", hours: 16, depends: "layout" },
        ],
      },
      { id: "custom", label: "Custom", rules: [] },
    ],
  },
  {
    id: "lt-unit",
    name: "LT-UNIT",
    variants: [
      { id: "standard", label: "Standard", rules: [] },
      { id: "custom", label: "Custom", rules: [] },
    ],
  },
  { id: "ah", name: "AH", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  { id: "gh", name: "GH", variants: [{ id: "standard", label: "Standard", rules: [] }] },
  {
    id: "yukon",
    name: "YUKON",
    variants: [
      { id: "standard", label: "Standard", rules: [] },
      { id: "custom", label: "Custom", rules: [] },
    ],
  },
];
const motoreSeedDeptByTitle = new Map();
machineData.forEach((machine) => {
  (machine.variants || []).forEach((variant) => {
    (variant.rules || []).forEach((rule) => {
      const titleKey = String(rule?.title || "").trim().toLowerCase();
      const dept = String(rule?.dept || "").trim();
      if (!titleKey || !dept) return;
      if (!motoreSeedDeptByTitle.has(titleKey)) motoreSeedDeptByTitle.set(titleKey, dept);
    });
  });
});

const machineList = document.getElementById("motoreMachineList");
const ruleTitle = document.getElementById("motoreRuleTitle");
const ruleList = document.getElementById("motoreRuleList");
const motoreTimelineViewBtn = document.getElementById("motoreTimelineViewBtn");
const motoreTimelineDefaultView = document.getElementById("motoreTimelineDefaultView");
const motoreTimelineCompareView = document.getElementById("motoreTimelineCompareView");
const motoreCompareActivityFilter = document.getElementById("motoreCompareActivityFilter");
const motoreCompareTableWrap = document.getElementById("motoreCompareTableWrap");
const machineSelect = document.getElementById("motoreMachineSelect");
const motoreSaveBtn = document.getElementById("motoreSaveBtn");
const motoreSimOrderCode = document.getElementById("motoreSimOrderCode");
const motoreSimFrameOrderDate = document.getElementById("motoreSimFrameOrderDate");
const motoreSimDeliveryDate = document.getElementById("motoreSimDeliveryDate");
const motoreSimRunBtn = document.getElementById("motoreSimRunBtn");
const motoreSimResetBtn = document.getElementById("motoreSimResetBtn");
const motoreSimMeta = document.getElementById("motoreSimMeta");
const motoreSimApplyBtn = document.getElementById("motoreSimApplyBtn");
const motoreSimKpis = document.getElementById("motoreSimKpis");
const motoreSimPlanList = document.getElementById("motoreSimPlanList");
const motoreSimIssues = document.getElementById("motoreSimIssues");
const motoreActivityCatalog = document.getElementById("motoreActivityCatalog");
const flowTableBody = document.getElementById("motoreFlowTableBody");
const skillTableBody = document.getElementById("motoreSkillTableBody");
const skillSummary = document.getElementById("motoreSkillSummary");
const motoreSynopticMeta = document.getElementById("motoreSynopticMeta");
const motoreSynopticFlow = document.getElementById("motoreSynopticFlow");
const motoreSynopticResources = document.getElementById("motoreSynopticResources");
const priorityInputs = {
  due: document.getElementById("motorePriorityDue"),
  delay: document.getElementById("motorePriorityDelay"),
  strategic: document.getElementById("motorePriorityStrategic"),
  complexity: document.getElementById("motorePriorityComplexity"),
};
const plannerRules = {
  hardSkills: document.getElementById("motoreRuleHardSkills"),
  continuity: document.getElementById("motoreRulePreferContinuity"),
  balance: document.getElementById("motoreRuleBalanceLoad"),
  notes: document.getElementById("motorePlannerNotes"),
};
const motoreParams = {
  bufferDays: document.getElementById("motoreParamBufferDays"),
  basePriority: document.getElementById("motoreParamBasePriority"),
  blockOverlap: document.getElementById("motoreParamBlockOverlap"),
};
const motoreTransferSource = document.getElementById("motoreTransferSource");
const motoreTransferInfo = document.getElementById("motoreTransferInfo");
const motoreTransferApplyBtn = document.getElementById("motoreTransferApplyBtn");
const motoreTransferAllBtn = document.getElementById("motoreTransferAllBtn");
const motoreTransferNoneBtn = document.getElementById("motoreTransferNoneBtn");
const motoreTransferSectionInputs = Array.from(
  document.querySelectorAll('input[data-transfer-section]')
);

const MOTORE_STRUCTURE_STORAGE_KEY = "motore_canvas_structure_v1";
const MOTORE_FLOW_LAYOUT_STORAGE_KEY = "motore_flow_layout_v1";
const MOTORE_SANDBOX_MODE_KEY = "motore_sandbox_mode_v1";
const defaultMachineParams = {
  bufferDays: 2,
  basePriority: "Media",
  blockOverlap: true,
};
const defaultPriorityWeights = {
  due: 40,
  delay: 30,
  strategic: 20,
  complexity: 10,
};
const defaultPlannerRules = {
  hardSkills: true,
  continuity: true,
  balance: true,
  notes: "",
};
const motoreActivityPresetGroups = [
  {
    dept: "CAD",
    titles: ["Preliminare", "3D", "Carenatura", "Telaio", "Costruttivi", "Consegna totale"],
  },
  {
    dept: "Termodinamici",
    titles: ["Progettazione termodinamica"],
  },
  {
    dept: "Elettrici",
    titles: ["Progettazione elettrica", "Ordine Kit Cavi"],
  },
  {
    dept: "Generali",
    titles: ["PED", "Segnatura elettrica", "OP Ordini"],
  },
];
const motoreGenericActivityTitles = motoreActivityPresetGroups
  .find((group) => group.dept === "Generali")
  ?.titles.slice() || [];
const motoreExcludedActivityTitles = new Set(
  [
    ...motoreGenericActivityTitles,
    "ORDINI MATERIALE",
    "ASSENTE",
    "ALTRO",
    "KICKOFF COMMESSA",
  ].map((t) => String(t).toLowerCase())
);
const motoreHourPresets = [
  { label: "15m", hours: 0.25 },
  { label: "30m", hours: 0.5 },
  { label: "1h", hours: 1 },
  { label: "4h", hours: 4 },
  { label: "8h", hours: 8 },
  { label: "2gg", hours: 16 },
  { label: "3gg", hours: 24 },
  { label: "4gg", hours: 32 },
  { label: "custom...", hours: null },
];
const MOTORE_HOURS_CUSTOM_INDEX = motoreHourPresets.length - 1;
const MOTORE_HOURS_SCALE_LABEL = motoreHourPresets.map((item) => item.label).join(" · ");

let motoreResources = [];
let motoreStructure = loadMotoreStructure();
let motoreFlowLayout = loadMotoreFlowLayout();
let motoreSandboxMode = loadMotoreSandboxMode();
let motoreLastSimulationResult = null;
let motoreTimelineCompareMode = false;
let motoreAllRulesLoaded = false;
let motoreAllRulesLoading = false;
const motoreActivityDeptByTitle = new Map();
let motoreActivityCatalogGroups = [];

function formatMotoreDeptLabel(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Generale";
  const key = raw.toLowerCase();
  if (key === "cad") return "CAD";
  if (key.startsWith("termodinam")) return "Termodinamico";
  if (key.startsWith("elettr")) return "Elettrico";
  if (key === "tutti" || key === "generale") return "Generale";
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

function getSkillActivityPresetGroups() {
  return motoreActivityPresetGroups.map((group) => ({
    dept: group.dept,
    titles: (group.titles || []).slice(),
  }));
}

function inferMotoreDeptFromTitle(title) {
  const rawTitle = String(title || "").trim();
  if (!rawTitle) return "Generale";
  const key = rawTitle.toLowerCase();
  const seedDept = motoreSeedDeptByTitle.get(key);
  if (seedDept) return formatMotoreDeptLabel(seedDept);

  const normalized = key
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
  if (/(^| )cad( |$)|layout|disegn|distint|assonometr|modell|3d/.test(normalized)) return "CAD";
  if (/termodinam|frigor|ciclo|gas|compressor|evapor|condens/.test(normalized)) return "Termodinamico";
  if (/elettr|schema|quadro|cavi|cablag|plc/.test(normalized)) return "Elettrico";
  if (/ped|segnatura|op ordini|ordini materiale/.test(normalized)) return "Generali";
  return "Generale";
}

function normalizeMotoreBasePriority(value) {
  const raw = String(value || "").trim();
  if (raw === "Alta" || raw === "Critica") return raw;
  return "Media";
}

function normalizeMachineParams(params) {
  const source = params && typeof params === "object" ? params : {};
  const bufferRaw = Number(source.bufferDays);
  return {
    bufferDays: Number.isFinite(bufferRaw) ? Math.max(0, Math.round(bufferRaw)) : defaultMachineParams.bufferDays,
    basePriority: normalizeMotoreBasePriority(source.basePriority),
    blockOverlap: Boolean(source.blockOverlap),
  };
}

function cloneMotoreRules(rules) {
  return (Array.isArray(rules) ? rules : []).map((rule) => ({ ...rule }));
}

function buildCatalogGroupsFromRuleRows(rows) {
  const grouped = new Map();
  (rows || []).forEach((row) => {
    const title = String(row?.title || "").trim();
    if (!title) return;
    const inputDept = formatMotoreDeptLabel(row?.dept || "Generale");
    const dept = inputDept === "Generale" || inputDept === "Senza reparto"
      ? inferMotoreDeptFromTitle(title)
      : inputDept;
    if (!grouped.has(dept)) grouped.set(dept, new Map());
    const map = grouped.get(dept);
    const key = title.toLowerCase();
    if (!map.has(key)) map.set(key, title);
  });
  return Array.from(grouped.entries())
    .map(([dept, map]) => ({
      dept,
      titles: Array.from(map.values()).sort((a, b) => a.localeCompare(b, "it")),
    }))
    .filter((group) => group.titles.length > 0)
    .sort((a, b) => a.dept.localeCompare(b.dept, "it"));
}

function renderActivityCatalog() {
  if (!motoreActivityCatalog) return;
  const groups = getSkillActivityPresetGroups();
  if (!groups.length) {
    motoreActivityCatalog.innerHTML = `<div class="motore-catalog-empty">Nessuna attivita caricata.</div>`;
    return;
  }
  const groupMarkup = groups
    .map((group) => {
      const titles = (group.titles || []);
      if (!titles.length) return "";
      const items = titles.map((title) => `<li>${escapeHtml(title)}</li>`).join("");
      return `
        <div class="motore-catalog-group${group.dept === "Generali" ? " motore-catalog-group-generic" : ""}">
          <div class="motore-catalog-group-title">${escapeHtml(group.dept)}</div>
          <ul class="motore-catalog-list">${items}</ul>
        </div>
      `;
    })
    .filter(Boolean)
    .join("");
  motoreActivityCatalog.innerHTML = groupMarkup;
}

function loadMotoreStructure() {
  try {
    const raw = localStorage.getItem(MOTORE_STRUCTURE_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function loadMotoreFlowLayout() {
  try {
    const raw = localStorage.getItem(MOTORE_FLOW_LAYOUT_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function loadMotoreSandboxMode() {
  try {
    const raw = localStorage.getItem(MOTORE_SANDBOX_MODE_KEY);
    if (raw === null) return true;
    return raw !== "false";
  } catch {
    return true;
  }
}

function saveMotoreStructure() {
  try {
    localStorage.setItem(MOTORE_STRUCTURE_STORAGE_KEY, JSON.stringify(motoreStructure));
  } catch {
    // Ignore localStorage write failures.
  }
}

function saveMotoreFlowLayout() {
  try {
    localStorage.setItem(MOTORE_FLOW_LAYOUT_STORAGE_KEY, JSON.stringify(motoreFlowLayout));
  } catch {
    // Ignore localStorage write failures.
  }
}

function saveMotoreSandboxMode() {
  try {
    localStorage.setItem(MOTORE_SANDBOX_MODE_KEY, motoreSandboxMode ? "true" : "false");
  } catch {
    // Ignore localStorage write failures.
  }
}

function refreshMotoreSandboxUI() {
  if (motoreSandboxState) {
    motoreSandboxState.textContent = motoreSandboxMode ? "Sandbox protetta" : "Live";
    motoreSandboxState.classList.toggle("is-live", !motoreSandboxMode);
  }
  if (motoreSandboxToggleBtn) {
    motoreSandboxToggleBtn.textContent = motoreSandboxMode ? "Passa a live" : "Torna sandbox";
    motoreSandboxToggleBtn.title = motoreSandboxMode
      ? "Attiva scrittura reale su Supabase"
      : "Blocca scrittura reale e torna in simulazione protetta";
  }
  if (motoreSandboxHint) {
    motoreSandboxHint.textContent = motoreSandboxMode
      ? "Sandbox attiva: simulazione protetta, nessuna scrittura su piano reale."
      : "Modalita live: il pulsante Salva bozza scrive su Supabase.";
  }
  if (motoreSaveBtn) {
    motoreSaveBtn.textContent = motoreSandboxMode ? "Salva locale" : "Salva bozza";
  }
  document.body.classList.toggle("motore-live-mode", !motoreSandboxMode);
}

function setMotoreSandboxMode(nextMode) {
  motoreSandboxMode = Boolean(nextMode);
  saveMotoreSandboxMode();
  refreshMotoreSandboxUI();
}

function toggleMotoreSandboxMode() {
  if (motoreSandboxMode) {
    if (!motoreSupabase) {
      window.alert("Supabase non disponibile: non puoi passare in modalita live.");
      return;
    }
    const goLive = window.confirm(
      "Passare in modalita LIVE? Da ora il salvataggio scrivera su Supabase."
    );
    if (!goLive) return;
    setMotoreSandboxMode(false);
    return;
  }
  setMotoreSandboxMode(true);
}

function getMotoreFlowCardOffset(machineId, activityKey) {
  const machineKey = String(machineId || "").trim();
  const nodeKey = String(activityKey || "").trim();
  if (!machineKey || !nodeKey) return { x: 0, y: 0 };
  const machineMap = motoreFlowLayout?.[machineKey];
  const entry = machineMap?.[nodeKey];
  const x = Number(entry?.x);
  const y = Number(entry?.y);
  return {
    x: Number.isFinite(x) ? x : 0,
    y: Number.isFinite(y) ? y : 0,
  };
}

function setMotoreFlowCardOffset(machineId, activityKey, x, y) {
  const machineKey = String(machineId || "").trim();
  const nodeKey = String(activityKey || "").trim();
  if (!machineKey || !nodeKey) return;
  if (!motoreFlowLayout[machineKey] || typeof motoreFlowLayout[machineKey] !== "object") {
    motoreFlowLayout[machineKey] = {};
  }
  const safeX = Number.isFinite(Number(x)) ? Number(x) : 0;
  const safeY = Number.isFinite(Number(y)) ? Number(y) : 0;
  motoreFlowLayout[machineKey][nodeKey] = { x: safeX, y: safeY };
}

function ensureMachineStructure(machineId) {
  if (!machineId) return null;
  if (!motoreStructure[machineId]) {
    motoreStructure[machineId] = {
      params: { ...defaultMachineParams },
      priorities: { ...defaultPriorityWeights },
      flow: {},
      skills: {},
      rules: { ...defaultPlannerRules },
    };
  }
  const machine = motoreStructure[machineId];
  machine.params = normalizeMachineParams({
    ...defaultMachineParams,
    ...(machine.params || {}),
  });
  machine.priorities = { ...defaultPriorityWeights, ...(machine.priorities || {}) };
  machine.rules = { ...defaultPlannerRules, ...(machine.rules || {}) };
  if (!machine.flow || typeof machine.flow !== "object") machine.flow = {};
  if (!machine.skills || typeof machine.skills !== "object") machine.skills = {};
  return machine;
}

function getMachineRuleKeySet(machine) {
  const keys = new Set();
  (machine?.rules || []).forEach((rule) => {
    if (isMotoreGenericActivityTitle(rule?.title || "")) return;
    const key = normalizeMotoreActivityKey(rule?.title || "");
    if (key) keys.add(key);
  });
  return keys;
}

function pruneOrphanFlowEntries(machine, machineStruct) {
  const flow = machineStruct?.flow;
  if (!flow || typeof flow !== "object") return [];
  const ruleKeys = getMachineRuleKeySet(machine);
  if (!ruleKeys.size) return [];
  const removed = [];
  Object.keys(flow).forEach((flowTitle) => {
    const key = normalizeMotoreActivityKey(flowTitle);
    if (!key || !ruleKeys.has(key)) {
      removed.push(flowTitle);
      delete flow[flowTitle];
    }
  });
  return removed;
}

async function loadPlannerCanvasForActive(machine) {
  if (!motoreSupabase || !machine) return;
  await loadMotoreMeta();
  const variantId = getVariantDbId(machine);
  if (!variantId) return;

  let configRows = null;
  let flowRows = null;
  let skillRows = null;
  try {
    const [cfgRes, flowRes, skillRes] = await Promise.all([
      motoreSupabase.from("motore_canvas_config").select("*").eq("variante_id", variantId).limit(1),
      motoreSupabase.from("motore_canvas_flow").select("*").eq("variante_id", variantId),
      motoreSupabase.from("motore_canvas_skill").select("*").eq("variante_id", variantId),
    ]);
    configRows = cfgRes.data || [];
    flowRows = flowRes.data || [];
    skillRows = skillRes.data || [];
  } catch {
    return;
  }

  const hasRemoteData = Boolean((configRows && configRows.length) || (flowRows && flowRows.length) || (skillRows && skillRows.length));
  if (!hasRemoteData) return;

  const machineStruct = ensureMachineStructure(machine.id);
  if (!machineStruct) return;

  const cfg = configRows[0] || null;
  if (cfg) {
    machineStruct.priorities = {
      due: Number(cfg.priority_due ?? defaultPriorityWeights.due),
      delay: Number(cfg.priority_delay ?? defaultPriorityWeights.delay),
      strategic: Number(cfg.priority_strategic ?? defaultPriorityWeights.strategic),
      complexity: Number(cfg.priority_complexity ?? defaultPriorityWeights.complexity),
    };
    machineStruct.rules = {
      hardSkills: Boolean(cfg.rule_hard_skills ?? defaultPlannerRules.hardSkills),
      continuity: Boolean(cfg.rule_prefer_continuity ?? defaultPlannerRules.continuity),
      balance: Boolean(cfg.rule_balance_load ?? defaultPlannerRules.balance),
      notes: String(cfg.planner_notes || ""),
    };
  }

  if (flowRows && flowRows.length) {
    machineStruct.flow = {};
    flowRows.forEach((row) => {
      const key = String(row.attivita_titolo || "").trim();
      if (!key) return;
      machineStruct.flow[key] = {
        depends: String(row.depends_on || ""),
        constraint: String(row.constraint_level || "none"),
        parallel: Boolean(row.allow_parallel),
        notes: String(row.notes || ""),
      };
    });
  }

  if (skillRows && skillRows.length) {
    machineStruct.skills = {};
    skillRows.forEach((row) => {
      const key = String(row.risorsa_id || "").trim();
      if (!key) return;
      machineStruct.skills[key] = {
        enabled: Boolean(row.enabled),
        level: String(row.skill_level || "medio"),
        activities: String(row.allowed_activities || ""),
        maxHours: Number(row.max_hours_per_day ?? 8),
      };
    });
  }

  pruneOrphanFlowEntries(machine, machineStruct);

  saveMotoreStructure();
  renderPlannerCanvas(machine);
}

async function savePlannerCanvasForActive(machine) {
  if (motoreSandboxMode) {
    saveMotoreStructure();
    return { ok: true, sandbox: true };
  }
  if (!motoreSupabase || !machine) return { ok: false, error: "Supabase not available." };
  await loadMotoreMeta();
  const variantId = getVariantDbId(machine);
  if (!variantId) return { ok: false, error: "Variant not found." };
  const machineStruct = ensureMachineStructure(machine.id);
  if (!machineStruct) return { ok: false, error: "Machine structure not found." };
  const staleFlowTitles = pruneOrphanFlowEntries(machine, machineStruct);

  const configPayload = {
    variante_id: variantId,
    priority_due: Number(machineStruct.priorities?.due ?? defaultPriorityWeights.due),
    priority_delay: Number(machineStruct.priorities?.delay ?? defaultPriorityWeights.delay),
    priority_strategic: Number(machineStruct.priorities?.strategic ?? defaultPriorityWeights.strategic),
    priority_complexity: Number(machineStruct.priorities?.complexity ?? defaultPriorityWeights.complexity),
    rule_hard_skills: Boolean(machineStruct.rules?.hardSkills),
    rule_prefer_continuity: Boolean(machineStruct.rules?.continuity),
    rule_balance_load: Boolean(machineStruct.rules?.balance),
    planner_notes: String(machineStruct.rules?.notes || ""),
    updated_by: motoreUserId || null,
  };

  const flowPayload = (machine.rules || []).map((rule) => {
    const key = String(rule.title || "").trim();
    const current = machineStruct.flow?.[key] || {
      depends: rule.depends === "inizio" ? "" : String(rule.depends || ""),
      constraint: getDefaultFlowConstraint(rule),
      parallel: false,
      notes: "",
    };
    return {
      variante_id: variantId,
      attivita_titolo: key,
      depends_on: String(current.depends || ""),
      constraint_level: String(current.constraint || "none"),
      allow_parallel: Boolean(current.parallel),
      notes: String(current.notes || ""),
      updated_by: motoreUserId || null,
    };
  });

  const skillPayload = Object.entries(machineStruct.skills || {})
    .map(([risorsaId, value]) => {
      const parsedRisorsaId = Number.parseInt(String(risorsaId), 10);
      if (!Number.isFinite(parsedRisorsaId)) return null;
      return {
        variante_id: variantId,
        risorsa_id: parsedRisorsaId,
        enabled: Boolean(value?.enabled),
        skill_level: String(value?.level || "medio"),
        allowed_activities: String(value?.activities || ""),
        max_hours_per_day: Number(value?.maxHours ?? 8),
        updated_by: motoreUserId || null,
      };
    })
    .filter(Boolean);

  const cfgResult = await motoreSupabase
    .from("motore_canvas_config")
    .upsert(configPayload, { onConflict: "variante_id" });
  if (cfgResult.error) return { ok: false, error: cfgResult.error.message };

  if (flowPayload.length) {
    const flowResult = await motoreSupabase
      .from("motore_canvas_flow")
      .upsert(flowPayload, { onConflict: "variante_id,attivita_titolo" });
    if (flowResult.error) return { ok: false, error: flowResult.error.message };
  }

  if (staleFlowTitles.length) {
    const uniqueStale = Array.from(
      new Set(staleFlowTitles.map((title) => String(title || "").trim()).filter(Boolean))
    );
    if (uniqueStale.length) {
      const staleDeleteResult = await motoreSupabase
        .from("motore_canvas_flow")
        .delete()
        .eq("variante_id", variantId)
        .in("attivita_titolo", uniqueStale);
      if (staleDeleteResult.error) return { ok: false, error: staleDeleteResult.error.message };
    }
  }

  if (skillPayload.length) {
    const skillResult = await motoreSupabase
      .from("motore_canvas_skill")
      .upsert(skillPayload, { onConflict: "variante_id,risorsa_id" });
    if (skillResult.error) return { ok: false, error: skillResult.error.message };
  }

  return { ok: true };
}

async function loadActivityTemplate() {
  if (!motoreSupabase) {
    motoreActivityCatalogGroups = buildCatalogGroupsFromRuleRows(
      baseRuleTemplate.map((rule) => ({ title: rule.title, dept: rule.dept }))
    );
    renderActivityCatalog();
    return baseRuleTemplate.map((r) => r.title);
  }
  motoreActivityDeptByTitle.clear();
  let activityRows = null;
  let repartoRows = null;
  try {
    const [activityRes, repartoRes] = await Promise.all([
      motoreSupabase
        .from("attivita")
        .select("titolo, reparto_id")
        .range(0, 9999),
      motoreSupabase
        .from("reparti")
        .select("id, nome"),
    ]);
    activityRows = activityRes.data || null;
    repartoRows = repartoRes.error ? [] : repartoRes.data || [];
  } catch {
    activityRows = null;
    repartoRows = [];
  }

  const repartoMap = new Map(
    (repartoRows || []).map((r) => [String(r.id), String(r.nome || "").trim() || "Generale"])
  );
  const titlesMap = new Map();
  const deptCountByTitle = new Map();
  const catalogRows = [];
  (activityRows || []).forEach((row) => {
    const title = String(row?.titolo || "").trim();
    if (!title) return;
    const key = title.toLowerCase();
    if (motoreExcludedActivityTitles.has(key)) return;
    const dbDept = repartoMap.get(String(row?.reparto_id ?? "")) || "Generale";
    const repartoName = formatMotoreDeptLabel(dbDept) === "Generale"
      ? inferMotoreDeptFromTitle(title)
      : formatMotoreDeptLabel(dbDept);
    catalogRows.push({ title, dept: repartoName });
    if (!deptCountByTitle.has(key)) deptCountByTitle.set(key, new Map());
    const deptCounter = deptCountByTitle.get(key);
    deptCounter.set(repartoName, (deptCounter.get(repartoName) || 0) + 1);
    if (!titlesMap.has(key)) titlesMap.set(key, title);
  });
  motoreActivityCatalogGroups = buildCatalogGroupsFromRuleRows(catalogRows);
  if (!motoreActivityCatalogGroups.length) {
    motoreActivityCatalogGroups = buildCatalogGroupsFromRuleRows(
      baseRuleTemplate.map((rule) => ({ title: rule.title, dept: rule.dept }))
    );
  }
  renderActivityCatalog();

  deptCountByTitle.forEach((counter, titleKey) => {
    const bestDept = Array.from(counter.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "it"))[0]?.[0];
    if (bestDept) motoreActivityDeptByTitle.set(titleKey, bestDept);
  });

  const titles = Array.from(titlesMap.values()).sort((a, b) => a.localeCompare(b, "it"));
  if (!titles.length) {
    return baseRuleTemplate.map((r) => r.title);
  }
  return titles;
}

async function loadMotoreMeta() {
  if (!motoreSupabase || metaLoaded) return;
  try {
    const [tipRes, varRes] = await Promise.all([
      motoreSupabase.from("motore_tipologie").select("id, nome"),
      motoreSupabase.from("motore_varianti").select("id, tipologia_id, nome"),
    ]);
    const tipRows = tipRes.data || [];
    const varRows = varRes.data || [];
    tipRows.forEach((row) => {
      const name = String(row.nome || "").trim();
      if (!name) return;
      motoreTipologie.set(name.toUpperCase(), row.id);
    });
    varRows.forEach((row) => {
      const tipName = Array.from(motoreTipologie.entries()).find(([, id]) => id === row.tipologia_id)?.[0];
      const varName = String(row.nome || "").trim().toLowerCase();
      if (!tipName || !varName) return;
      motoreVariantIds.set(`${tipName}|${varName}`, row.id);
    });
    metaLoaded = true;
  } catch {
    metaLoaded = false;
  }
}

async function loadMotoreResources() {
  if (!motoreSupabase) {
    motoreResources = [];
    return;
  }
  try {
    // Primary path: current schema (risorse has reparto_id).
    let risorseRows = null;
    let repartoRows = null;
    const [risorseRes, repartiRes] = await Promise.all([
      motoreSupabase.from("risorse").select("id, nome, reparto_id, attiva").order("nome", { ascending: true }),
      motoreSupabase.from("reparti").select("id, nome"),
    ]);
    if (risorseRes.error) throw risorseRes.error;
    risorseRows = risorseRes.data || [];
    repartoRows = repartiRes.error ? [] : repartiRes.data || [];

    const repartoMap = new Map(
      (repartoRows || []).map((r) => [String(r.id), String(r.nome || "").trim()])
    );
    motoreResources = (risorseRows || [])
      .filter((row) => row && row.attiva !== false)
      .map((row) => ({
        id: row.id,
        nome: row.nome,
        reparto_id: row.reparto_id ?? null,
        attiva: row.attiva !== false,
        reparto: repartoMap.get(String(row.reparto_id ?? "")) || "",
      }));
  } catch {
    // Compatibility fallback for legacy schema variants.
    try {
      const legacyRes = await motoreSupabase
        .from("risorse")
        .select("id, nome, attiva")
        .order("nome", { ascending: true });
      if (legacyRes.error) throw legacyRes.error;
      motoreResources = (legacyRes.data || [])
        .filter((row) => row && row.attiva !== false)
        .map((row) => ({
          id: row.id,
          nome: row.nome,
          attiva: row.attiva !== false,
          reparto: "",
        }));
    } catch {
      motoreResources = [];
    }
  }
}

function getVariantDbId(machine) {
  const tipName = String(machine.machineName || machine.name || "").trim().toUpperCase();
  const varName = String(machine.variantLabel || "standard").trim().toLowerCase();
  return motoreVariantIds.get(`${tipName}|${varName}`) || null;
}

async function loadRulesForActive(machine) {
  if (!motoreSupabase || !machine) return;
  await loadMotoreMeta();
  const variantId = getVariantDbId(machine);
  if (!variantId) {
    renderRuleWarnings(machine);
    return;
  }
  const token = ++rulesLoadToken;
  let rows = null;
  try {
    const result = await motoreSupabase
      .from("motore_regole")
      .select("attivita_titolo, ore")
      .eq("variante_id", variantId);
    rows = result.data || null;
  } catch {
    rows = null;
  }
  if (token !== rulesLoadToken) return;
  if (rows) {
    const map = new Map(
      rows.map((r) => [String(r.attivita_titolo || "").toLowerCase(), r.ore])
    );
    const nextRules = machine.rules.map((rule) => {
      const key = String(rule.title || "").toLowerCase();
      if (map.has(key)) {
        return { ...rule, hours: map.get(key) };
      }
      return rule;
    });
    machine.rules = nextRules;
    if (machine.variantRef) {
      machine.variantRef.rules = nextRules;
    }
  }
  renderRuleList(machine);
}

async function saveRulesForActive(machine) {
  if (motoreSandboxMode) {
    saveMotoreStructure();
    return { ok: true, sandbox: true };
  }
  if (!motoreSupabase || !machine) return { ok: false, error: "Supabase not available." };
  await loadMotoreMeta();
  const variantId = getVariantDbId(machine);
  if (!variantId) return { ok: false, error: "Variant not found." };
  const payload = (machine.rules || []).map((rule) => ({
    variante_id: variantId,
    attivita_titolo: rule.title,
    ore: Number.isFinite(Number(rule.hours)) ? Number(rule.hours) : null,
  }));
  const result = await motoreSupabase
    .from("motore_regole")
    .upsert(payload, { onConflict: "variante_id,attivita_titolo" });
  if (result.error) {
    return { ok: false, error: result.error.message };
  }
  return { ok: true };
}

function applyTemplateToMachines(titles) {
  machineData.forEach((machine) => {
    machine.variants.forEach((variant) => {
      const existing = new Map(
        (variant.rules || []).map((r) => [String(r.title || "").toLowerCase(), r])
      );
      const nextRules = titles.map((title) => {
        const key = title.toLowerCase();
        const mappedDept = motoreActivityDeptByTitle.get(key) || "Generale";
        const rule = existing.get(key);
        if (rule) {
          const existingDept = String(rule.dept || "").trim();
          return {
            ...rule,
            title,
            dept: existingDept && existingDept.toLowerCase() !== "generale"
              ? existingDept
              : mappedDept,
          };
        }
        return { title, dept: mappedDept, hours: null, depends: "da definire" };
      });
      variant.rules = nextRules;
    });
  });
}

async function initMotoreTemplate() {
  if (templateLoaded) return;
  const [titles] = await Promise.all([loadActivityTemplate(), loadMotoreResources()]);
  applyTemplateToMachines(titles);
  templateLoaded = true;
  if (activeMachineId) {
    setActiveMachine(activeMachineId);
  } else {
    const first = getMachineVariants()[0]?.id || "";
    if (first) setActiveMachine(first);
  }
  renderMachineSelect(activeMachineId || getMachineVariants()[0]?.id);
}

function getMachineVariants() {
  const items = [];
  machineData.forEach((machine) => {
    machine.variants.forEach((variant) => {
      items.push({
        id: `${machine.id}:${variant.id}`,
        machineId: machine.id,
        machineName: machine.name,
        variantId: variant.id,
        variantLabel: variant.label,
        rules: variant.rules || [],
        machineRef: machine,
        variantRef: variant,
      });
    });
  });
  return items;
}

function getMachineVariantById(id) {
  const list = getMachineVariants();
  return list.find((entry) => entry.id === id) || list[0];
}

function renderMachineList(activeId) {
  if (!machineList) return;
  machineList.innerHTML = "";
  const active = getMachineVariantById(activeId);
  machineData.forEach((m) => {
    const item = document.createElement("div");
    const isActive = active && active.machineId === m.id;
    item.className = `motore-list-item${isActive ? " active" : ""}`;
    item.innerHTML = `
      <div class="motore-machine-name">${m.name}</div>
      ${m.variants.length > 1 ? `<div class="motore-variant-chips"></div>` : ""}
    `;
    item.addEventListener("click", () => setActiveMachine(`${m.id}:${m.variants[0].id}`));
    if (m.variants.length > 1) {
      const chips = item.querySelector(".motore-variant-chips");
      m.variants.forEach((variant) => {
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = `motore-variant-chip${
          active && active.machineId === m.id && active.variantId === variant.id ? " active" : ""
        }`;
        chip.textContent = variant.label;
        chip.addEventListener("click", (e) => {
          e.stopPropagation();
          setActiveMachine(`${m.id}:${variant.id}`);
        });
        chips.appendChild(chip);
      });
    }
    machineList.appendChild(item);
  });
}

function renderRuleList(machine) {
  if (!ruleList) return;
  ruleList.innerHTML = "";
  machine.rules.forEach((r) => {
    const presetIndex = getMotoreHourPresetIndex(r.hours);
    const card = document.createElement("div");
    card.className = "motore-rule-card";
    card.innerHTML = `
      <div class="motore-rule-handle"></div>
      <div class="motore-rule-meta">
        <div class="motore-rule-title">${r.title}</div>
        <div class="motore-rule-sub">${r.dept} · dipende da: ${r.depends}</div>
      </div>
      <div class="motore-rule-time-control">
        <label class="motore-rule-duration">
          <input
            class="motore-hours-input"
            type="number"
            min="0"
            step="0.5"
            value="${r.hours ?? ""}"
            title="Ore standard stimate per l'attivita ${escapeHtml(String(r.title || ""))}. Usa 0 se non necessaria."
          />
          <span>h</span>
        </label>
        <div
          class="motore-hours-slider-wrap"
          title="Preset rapidi durata: 15m, 30m, 1h, 4h, 8h, 2gg, 3gg, 4gg o Custom."
        >
          <input
            class="motore-hours-slider"
            type="range"
            min="0"
            max="${MOTORE_HOURS_CUSTOM_INDEX}"
            step="1"
            value="${presetIndex}"
          />
          <span class="motore-hours-slider-label">${motoreHourPresets[presetIndex]?.label || "custom..."}</span>
        </div>
        <div class="motore-hours-slider-scale">${MOTORE_HOURS_SCALE_LABEL}</div>
      </div>
    `;
    const input = card.querySelector(".motore-hours-input");
    const slider = card.querySelector(".motore-hours-slider");
    const sliderLabel = card.querySelector(".motore-hours-slider-label");
    const syncPresetUI = () => {
      const index = getMotoreHourPresetIndex(r.hours);
      if (slider) slider.value = String(index);
      if (sliderLabel) sliderLabel.textContent = motoreHourPresets[index]?.label || "custom...";
      if (input) {
        input.classList.toggle("motore-hours-input-custom", index === MOTORE_HOURS_CUSTOM_INDEX);
      }
    };
    if (input) {
      input.addEventListener("input", () => {
        const value = Number(input.value);
        r.hours = Number.isFinite(value) ? value : null;
        syncPresetUI();
        renderRuleWarnings(machine);
        refreshMotoreValidation(machine);
      });
    }
    if (slider) {
      const applySlider = () => {
        const index = Math.max(0, Math.min(MOTORE_HOURS_CUSTOM_INDEX, Number.parseInt(slider.value, 10) || 0));
        const preset = motoreHourPresets[index] || motoreHourPresets[MOTORE_HOURS_CUSTOM_INDEX];
        if (sliderLabel) sliderLabel.textContent = preset.label;
        if (preset.hours === null) {
          const value = Number(input?.value);
          r.hours = Number.isFinite(value) ? value : null;
          if (input) input.focus();
        } else {
          r.hours = preset.hours;
          if (input) input.value = String(preset.hours);
        }
        syncPresetUI();
        renderRuleWarnings(machine);
        refreshMotoreValidation(machine);
      };
      slider.addEventListener("input", applySlider);
      slider.addEventListener("change", applySlider);
    }
    syncPresetUI();
    ruleList.appendChild(card);
  });
  renderRuleWarnings(machine);
  refreshMotoreValidation(machine);
  if (motoreTimelineCompareMode) {
    renderTimelineComparison(machine);
  }
}

function getMotoreHourPresetIndex(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return MOTORE_HOURS_CUSTOM_INDEX;
  const exactIndex = motoreHourPresets.findIndex(
    (preset) => preset.hours !== null && Math.abs(preset.hours - numeric) < 0.001
  );
  return exactIndex >= 0 ? exactIndex : MOTORE_HOURS_CUSTOM_INDEX;
}

function renderMachineSelect(activeId) {
  if (!machineSelect) return;
  machineSelect.innerHTML = "";
  const items = getMachineVariants();
  const grouped = new Map();
  items.forEach((entry) => {
    const list = grouped.get(entry.machineName) || [];
    list.push(entry);
    grouped.set(entry.machineName, list);
  });
  Array.from(grouped.entries()).forEach(([name, entries]) => {
    entries.forEach((entry) => {
      const opt = document.createElement("option");
      opt.value = entry.id;
      opt.textContent = entries.length > 1 ? `${name} — ${entry.variantLabel}` : name;
      if (entry.id === activeId) opt.selected = true;
      machineSelect.appendChild(opt);
    });
  });
  machineSelect.onchange = (e) => setActiveMachine(e.target.value);
}

function getMotoreVariantHoursMap(machine) {
  const map = new Map();
  (machine?.rules || []).forEach((rule) => {
    if (isMotoreGenericActivityTitle(rule?.title || "")) return;
    const key = normalizeMotoreActivityKey(rule?.title || "");
    if (!key) return;
    const hours = Number(rule?.hours);
    map.set(key, Number.isFinite(hours) ? hours : null);
  });
  return map;
}

async function loadRulesForAllVariants() {
  if (!motoreSupabase || motoreAllRulesLoaded || motoreAllRulesLoading) return;
  motoreAllRulesLoading = true;
  try {
    await loadMotoreMeta();
    const variantIds = Array.from(new Set(Array.from(motoreVariantIds.values()).filter(Boolean)));
    if (!variantIds.length) {
      motoreAllRulesLoaded = true;
      return;
    }
    const res = await motoreSupabase
      .from("motore_regole")
      .select("variante_id, attivita_titolo, ore")
      .in("variante_id", variantIds)
      .range(0, 60000);
    if (res.error) return;
    const byVariantId = new Map();
    (res.data || []).forEach((row) => {
      const variantId = String(row?.variante_id || "").trim();
      if (!variantId) return;
      if (!byVariantId.has(variantId)) byVariantId.set(variantId, []);
      byVariantId.get(variantId).push(row);
    });
    getMachineVariants().forEach((entry) => {
      if (entry.id === activeMachineId) return;
      const variantId = getVariantDbId(entry);
      if (!variantId) return;
      const rows = byVariantId.get(String(variantId)) || [];
      if (!rows.length) return;
      const dbHours = new Map();
      rows.forEach((row) => {
        const key = normalizeMotoreActivityKey(row?.attivita_titolo || "");
        if (!key) return;
        const hours = Number(row?.ore);
        dbHours.set(key, Number.isFinite(hours) ? hours : null);
      });
      const nextRules = (entry.rules || []).map((rule) => {
        const key = normalizeMotoreActivityKey(rule?.title || "");
        if (!key || !dbHours.has(key)) return rule;
        return { ...rule, hours: dbHours.get(key) };
      });
      entry.rules = nextRules;
      if (entry.variantRef) entry.variantRef.rules = nextRules;
    });
    motoreAllRulesLoaded = true;
  } finally {
    motoreAllRulesLoading = false;
  }
}

function renderTimelineComparison(machine) {
  if (!motoreCompareTableWrap) return;
  const active = machine || getMachineVariantById(activeMachineId);
  if (!active) return;
  const filterKey = normalizeMotoreTextKey(motoreCompareActivityFilter?.value || "");
  const baseRows = (active.rules || [])
    .filter((rule) => !isMotoreGenericActivityTitle(rule?.title || ""))
    .map((rule) => {
      const title = getCanonicalMotoreActivityTitle(rule?.title || "");
      return {
        title,
        key: normalizeMotoreActivityKey(title),
      };
    })
    .filter((row) => Boolean(row.key))
    .filter((row) => !filterKey || normalizeMotoreTextKey(row.title).includes(filterKey));

  if (!baseRows.length) {
    motoreCompareTableWrap.innerHTML = `<div class="motore-compare-empty">Nessuna attivita da confrontare.</div>`;
    return;
  }

  const variants = getMachineVariants().slice().sort((a, b) => {
    return getMotoreMachineVariantLabel(a).localeCompare(getMotoreMachineVariantLabel(b), "it");
  });
  const hoursByVariant = new Map(variants.map((entry) => [entry.id, getMotoreVariantHoursMap(entry)]));

  const headCols = variants
    .map((entry) => {
      const cls = entry.id === activeMachineId ? " is-active-col" : "";
      return `<th class="${cls.trim()}">${escapeHtml(getMotoreMachineVariantLabel(entry))}</th>`;
    })
    .join("");
  const bodyRows = baseRows
    .map((row) => {
      const cells = variants
        .map((entry) => {
          const cls = entry.id === activeMachineId ? " is-active-col" : "";
          const hoursMap = hoursByVariant.get(entry.id) || new Map();
          const value = hoursMap.get(row.key);
          const text = Number.isFinite(value) ? `${Number(value).toFixed(1)}h` : "-";
          return `<td class="${cls.trim()}">${escapeHtml(text)}</td>`;
        })
        .join("");
      return `<tr><td>${escapeHtml(row.title)}</td>${cells}</tr>`;
    })
    .join("");

  motoreCompareTableWrap.innerHTML = `
    <table class="motore-compare-table">
      <thead>
        <tr>
          <th>Attivita</th>
          ${headCols}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  `;
}

function setMotoreTimelineCompareMode(enabled) {
  motoreTimelineCompareMode = Boolean(enabled);
  if (motoreTimelineDefaultView) {
    motoreTimelineDefaultView.classList.toggle("hidden", motoreTimelineCompareMode);
  }
  if (motoreTimelineCompareView) {
    motoreTimelineCompareView.classList.toggle("hidden", !motoreTimelineCompareMode);
  }
  if (motoreTimelineViewBtn) {
    motoreTimelineViewBtn.textContent = motoreTimelineCompareMode ? "Torna timeline" : "Vista confronto ore";
  }
  if (motoreTimelineCompareMode) {
    renderTimelineComparison(getMachineVariantById(activeMachineId));
  }
}

function bindMotoreTimelineCompare() {
  if (motoreCompareActivityFilter) {
    motoreCompareActivityFilter.addEventListener("input", () => {
      if (!motoreTimelineCompareMode) return;
      renderTimelineComparison(getMachineVariantById(activeMachineId));
    });
  }
  if (motoreTimelineViewBtn) {
    motoreTimelineViewBtn.addEventListener("click", async () => {
      const nextMode = !motoreTimelineCompareMode;
      setMotoreTimelineCompareMode(nextMode);
      if (!nextMode) return;
      motoreTimelineViewBtn.disabled = true;
      try {
        await loadRulesForAllVariants();
        renderTimelineComparison(getMachineVariantById(activeMachineId));
      } finally {
        motoreTimelineViewBtn.disabled = false;
      }
    });
  }
}

function getMotoreMachineVariantLabel(machine) {
  if (!machine) return "";
  const variant = String(machine.variantLabel || "").trim();
  if (!variant || variant.toLowerCase() === "standard") return `${machine.machineName} — Standard`;
  return `${machine.machineName} — ${variant}`;
}

function renderMotoreTransferPanel(targetMachineId) {
  if (!motoreTransferSource) return;
  const activeId = String(targetMachineId || activeMachineId || "").trim();
  const previous = String(motoreTransferSource.value || "").trim();
  const items = getMachineVariants()
    .filter((entry) => String(entry.id) !== activeId)
    .sort((a, b) => getMotoreMachineVariantLabel(a).localeCompare(getMotoreMachineVariantLabel(b), "it"));
  motoreTransferSource.innerHTML = "";
  items.forEach((entry) => {
    const opt = document.createElement("option");
    opt.value = entry.id;
    opt.textContent = getMotoreMachineVariantLabel(entry);
    motoreTransferSource.appendChild(opt);
  });
  if (items.length) {
    const fallback = items[0].id;
    const nextValue = items.some((entry) => entry.id === previous) ? previous : fallback;
    motoreTransferSource.value = nextValue;
  }
  if (motoreTransferApplyBtn) {
    motoreTransferApplyBtn.disabled = !items.length;
    motoreTransferApplyBtn.title = items.length
      ? "Copia le sezioni selezionate nella tipologia corrente"
      : "Nessuna altra tipologia disponibile";
  }
  if (motoreTransferInfo) {
    const active = getMachineVariantById(activeId);
    const targetLabel = getMotoreMachineVariantLabel(active);
    motoreTransferInfo.textContent = items.length
      ? `Target corrente: ${targetLabel}`
      : `Solo una tipologia disponibile: ${targetLabel}`;
  }
}

function getSelectedMotoreTransferSections() {
  return motoreTransferSectionInputs
    .filter((input) => input && input.checked)
    .map((input) => String(input.dataset.transferSection || "").trim())
    .filter(Boolean);
}

function setMotoreTransferInfo(message) {
  if (!motoreTransferInfo) return;
  motoreTransferInfo.textContent = String(message || "");
}

function applyMotoreTransfer() {
  const target = getMachineVariantById(activeMachineId);
  const sourceId = String(motoreTransferSource?.value || "").trim();
  const source = getMachineVariantById(sourceId);
  if (!target || !sourceId || !source || source.id === target.id) {
    setMotoreTransferInfo("Seleziona una tipologia sorgente diversa da quella corrente.");
    return;
  }
  const sections = getSelectedMotoreTransferSections();
  if (!sections.length) {
    setMotoreTransferInfo("Seleziona almeno una sezione da trasferire.");
    return;
  }
  const sourceStruct = ensureMachineStructure(source.id);
  const targetStruct = ensureMachineStructure(target.id);
  if (!sourceStruct || !targetStruct) {
    setMotoreTransferInfo("Impossibile leggere struttura sorgente o target.");
    return;
  }

  const summary = sections.join(", ");
  const ok = window.confirm(
    `Copiare [${summary}] da ${getMotoreMachineVariantLabel(source)} a ${getMotoreMachineVariantLabel(target)}?`
  );
  if (!ok) return;

  const copiedLabels = [];
  if (sections.includes("timeline")) {
    const nextRules = cloneMotoreRules(source.rules || []);
    target.rules = nextRules;
    if (target.variantRef) target.variantRef.rules = cloneMotoreRules(nextRules);
    copiedLabels.push("Timeline logica");
  }
  if (sections.includes("params")) {
    targetStruct.params = normalizeMachineParams({
      ...defaultMachineParams,
      ...(sourceStruct.params || {}),
    });
    copiedLabels.push("Parametri");
  }
  if (sections.includes("priorities")) {
    targetStruct.priorities = { ...defaultPriorityWeights, ...(sourceStruct.priorities || {}) };
    copiedLabels.push("Priorita backlog");
  }
  if (sections.includes("flow")) {
    targetStruct.flow = JSON.parse(JSON.stringify(sourceStruct.flow || {}));
    copiedLabels.push("Dipendenze e flusso");
  }
  if (sections.includes("skills")) {
    targetStruct.skills = JSON.parse(JSON.stringify(sourceStruct.skills || {}));
    copiedLabels.push("Skills e capacita");
  }
  if (sections.includes("assignmentRules")) {
    targetStruct.rules = { ...defaultPlannerRules, ...(sourceStruct.rules || {}) };
    copiedLabels.push("Regole di assegnazione");
  }
  if (sections.includes("timeline") && !sections.includes("flow")) {
    pruneOrphanFlowEntries(target, targetStruct);
  }

  saveMotoreStructure();
  renderRuleList(target);
  renderPlannerCanvas(target);
  renderMachineList(target.id);
  renderMachineSelect(target.id);
  renderMotoreTransferPanel(target.id);
  setMotoreTransferInfo(`Copiato: ${copiedLabels.join(", ")}.`);
}

function bindMotoreTransferPanel() {
  if (motoreTransferAllBtn) {
    motoreTransferAllBtn.addEventListener("click", () => {
      motoreTransferSectionInputs.forEach((input) => {
        input.checked = true;
      });
    });
  }
  if (motoreTransferNoneBtn) {
    motoreTransferNoneBtn.addEventListener("click", () => {
      motoreTransferSectionInputs.forEach((input) => {
        input.checked = false;
      });
    });
  }
  if (motoreTransferApplyBtn) {
    motoreTransferApplyBtn.addEventListener("click", applyMotoreTransfer);
  }
}

function renderRuleWarnings(machine) {
  if (!ruleWarnings) return;
  const missing = (machine.rules || []).filter((r) => {
    if (r.hours === 0) return false;
    return !Number.isFinite(Number(r.hours));
  });
  if (!missing.length) {
    ruleWarnings.classList.add("hidden");
    ruleWarnings.innerHTML = "";
    return;
  }
  const typeLabel = machine.variantLabel && machine.variantLabel.toLowerCase() !== "standard"
    ? `${machine.machineName} ${machine.variantLabel}`
    : machine.machineName;
  const typeUpper = typeLabel.toUpperCase();
  ruleWarnings.innerHTML = missing
    .map(
      (r) =>
        `<div class="motore-rule-warning">L'attivit&#224; ${escapeHtml(String(r.title))} non ha ore assegnate nella TIPOLOGIA MACCHINA ${escapeHtml(typeUpper)}</div>`
    )
    .join("");
  ruleWarnings.classList.remove("hidden");
}

function getMachineDisplayLabel(machine) {
  if (!machine) return "";
  return machine.variantLabel && machine.variantLabel.toLowerCase() !== "standard"
    ? `${machine.machineName} ${machine.variantLabel}`
    : machine.machineName;
}

function collectMotoreValidation(machine, machineStruct) {
  const items = [];
  let checks = 0;
  const rules = Array.isArray(machine?.rules) ? machine.rules : [];
  const ruleByKey = new Map();
  const ruleTitleByKey = new Map();
  rules.forEach((rule) => {
    const title = String(rule?.title || "").trim();
    if (!title) return;
    if (isMotoreGenericActivityTitle(title)) return;
    const key = normalizeMotoreActivityKey(title);
    if (!key) return;
    ruleByKey.set(key, rule);
    if (!ruleTitleByKey.has(key)) ruleTitleByKey.set(key, title);
  });

  checks += 1;
  if (!rules.length) {
    items.push({ severity: "error", text: "Nessuna attivita configurata per questa tipologia." });
  }

  const priorities = machineStruct?.priorities || {};
  const due = Number(priorities.due ?? defaultPriorityWeights.due);
  const delay = Number(priorities.delay ?? defaultPriorityWeights.delay);
  const strategic = Number(priorities.strategic ?? defaultPriorityWeights.strategic);
  const complexity = Number(priorities.complexity ?? defaultPriorityWeights.complexity);
  const priorityTotal = [due, delay, strategic, complexity]
    .map((v) => (Number.isFinite(v) ? v : 0))
    .reduce((acc, value) => acc + value, 0);
  checks += 1;
  if (Math.round(priorityTotal) !== 100) {
    items.push({
      severity: "warn",
      text: `La somma dei pesi priorita e ${priorityTotal} (consigliato 100).`,
    });
  }

  const flow = machineStruct?.flow && typeof machineStruct.flow === "object" ? machineStruct.flow : {};
  Object.keys(flow).forEach((flowTitle) => {
    const key = normalizeMotoreActivityKey(flowTitle);
    if (!key || ruleByKey.has(key)) return;
    checks += 1;
    items.push({
      severity: "warn",
      text: `Flusso presente per "${String(flowTitle)}" ma attivita non trovata nella timeline.`,
    });
  });

  rules.forEach((rule) => {
    const title = String(rule?.title || "").trim();
    if (!title) return;
    if (isMotoreGenericActivityTitle(title)) return;
    const key = normalizeMotoreActivityKey(title);
    if (!key) return;
    const numericHours = Number(rule?.hours);
    checks += 1;
    if (!Number.isFinite(numericHours)) {
      items.push({ severity: "error", text: `L'attivita "${title}" non ha ore valorizzate.` });
    } else if (numericHours < 0) {
      items.push({ severity: "error", text: `L'attivita "${title}" ha ore negative (${numericHours}).` });
    } else if (numericHours === 0) {
      items.push({ severity: "warn", text: `L'attivita "${title}" ha durata 0h.` });
    }

    checks += 1;
    const flowRow = flow[title];
    if (!flowRow) {
      items.push({ severity: "warn", text: `Flusso non compilato per l'attivita "${title}".` });
      return;
    }

    const rawDepends = parseAllowedActivitiesList(String(flowRow.depends || ""));
    const dependsList = rawDepends
      .map((item) => normalizeMotoreActivityKey(item))
      .filter(Boolean);
    const uniqueDepends = Array.from(new Set(dependsList));
    const constraint = String(flowRow.constraint || "none").trim().toLowerCase();
    checks += 1;
    if (uniqueDepends.includes(key)) {
      items.push({
        severity: "error",
        text: `L'attivita "${title}" non puo dipendere da se stessa.`,
      });
    }
    const invalidDepends = uniqueDepends.filter((depKey) => depKey !== key && !ruleByKey.has(depKey));
    if (invalidDepends.length) {
      const invalidLabels = rawDepends.filter((item) => invalidDepends.includes(normalizeMotoreActivityKey(item)));
      const shown = invalidLabels.length ? invalidLabels : invalidDepends;
      items.push({
        severity: "warn",
        text: `Dipendenza non trovata: "${title}" dipende da "${shown.join(", ")}".`,
      });
    }

    checks += 1;
    if (!uniqueDepends.length && constraint !== "none") {
      items.push({
        severity: "warn",
        text: `L'attivita "${title}" ha vincolo "${constraint}" ma senza predecessore.`,
      });
    }
    checks += 1;
    if (uniqueDepends.length && constraint === "none") {
      items.push({
        severity: "warn",
        text: `L'attivita "${title}" ha predecessore ma vincolo impostato su "none".`,
      });
    }
  });

  const skills = machineStruct?.skills && typeof machineStruct.skills === "object" ? machineStruct.skills : {};
  const enabledRows = Object.entries(skills).filter(([, row]) => Boolean(row?.enabled));
  checks += 1;
  if (!enabledRows.length) {
    items.push({ severity: "error", text: "Nessuna risorsa abilitata nella sezione Skill e capacita." });
  }

  const presetGroups = getSkillActivityPresetGroups();
  const presetKeys = new Set(
    presetGroups.flatMap((group) => (group?.titles || []).map((title) => normalizeMotoreActivityKey(title)))
  );
  const resourceById = new Map((motoreResources || []).map((res) => [String(res.id || ""), res]));
  const coverageByRule = new Map(
    Array.from(ruleByKey.keys()).map((ruleKey) => [ruleKey, 0])
  );

  enabledRows.forEach(([resourceId, row]) => {
    const res = resourceById.get(String(resourceId)) || null;
    const resName = String(res?.nome || `Risorsa ${resourceId}`);
    const allowedGroups = filterActivityGroupsForResource(
      presetGroups,
      String(res?.reparto || "")
    );
    const allowedKeys = new Set(
      allowedGroups.flatMap((group) => (group?.titles || []).map((title) => normalizeMotoreActivityKey(title)))
    );
    const selected = parseAllowedActivitiesList(row?.activities || "");
    const selectedUnique = Array.from(new Set(selected.map((item) => String(item || "").trim()).filter(Boolean)));

    checks += 1;
    const maxHours = Number(row?.maxHours);
    if (!Number.isFinite(maxHours) || maxHours <= 0) {
      items.push({
        severity: "error",
        text: `La risorsa "${resName}" ha Max h/giorno non valido (${String(row?.maxHours ?? "")}).`,
      });
    }

    checks += 1;
    if (!selectedUnique.length) {
      items.push({
        severity: "warn",
        text: `La risorsa "${resName}" e abilitata ma senza attivita consentite.`,
      });
    }

    selectedUnique.forEach((activityTitle) => {
      const activityKey = normalizeMotoreActivityKey(activityTitle);
      if (!activityKey) return;
      checks += 1;
      if (!presetKeys.has(activityKey)) {
        items.push({
          severity: "warn",
          text: `La risorsa "${resName}" contiene attivita non a catalogo: "${activityTitle}".`,
        });
      }
      checks += 1;
      if (presetKeys.has(activityKey) && allowedKeys.size && !allowedKeys.has(activityKey)) {
        items.push({
          severity: "warn",
          text: `La risorsa "${resName}" ha attivita fuori reparto: "${activityTitle}".`,
        });
      }
      if (coverageByRule.has(activityKey)) {
        coverageByRule.set(activityKey, Number(coverageByRule.get(activityKey) || 0) + 1);
      }
    });
  });

  coverageByRule.forEach((count, key) => {
    checks += 1;
    if (count > 0) return;
    items.push({
      severity: "error",
      text: `Nessuna risorsa abilitata/coperta per l'attivita "${ruleTitleByKey.get(key) || key}".`,
    });
  });

  const severityOrder = { error: 0, warn: 1, ok: 2 };
  items.sort((a, b) => (severityOrder[a.severity] ?? 9) - (severityOrder[b.severity] ?? 9));
  return { items, checks };
}

function renderMotoreValidation(machine, machineStruct) {
  if (!motoreValidationSummary || !motoreValidationList || !machine || !machineStruct) return;
  const report = collectMotoreValidation(machine, machineStruct);
  motoreValidationItemsCache = Array.isArray(report.items) ? report.items.slice() : [];
  const errors = report.items.filter((item) => item.severity === "error").length;
  const warns = report.items.filter((item) => item.severity === "warn").length;
  const machineLabel = getMachineDisplayLabel(machine);
  motoreValidationSummary.textContent = `${machineLabel}: ${errors} errori, ${warns} warning su ${report.checks} verifiche.`;

  if (!report.items.length) {
    motoreValidationList.innerHTML = `
      <div class="motore-validation-item is-ok">
        <span class="motore-validation-badge">OK</span>
        <span class="motore-validation-text">Nessuna incongruenza rilevata. Configurazione pronta per i test di pianificazione.</span>
      </div>
    `;
    return;
  }

  motoreValidationList.innerHTML = report.items
    .map((item, index) => {
      const badge = item.severity === "error" ? "Errore" : item.severity === "warn" ? "Warning" : "OK";
      const rowClass = item.severity === "error" ? "is-error" : item.severity === "warn" ? "is-warn" : "is-ok";
      return `
        <div class="motore-validation-item ${rowClass}">
          <span class="motore-validation-badge">${badge}</span>
          <span class="motore-validation-text">${escapeHtml(String(item.text || ""))}</span>
          <button
            type="button"
            class="motore-copy-btn"
            data-copy-validation="${index}"
            title="Copia negli appunti"
            aria-label="Copia negli appunti"
          >
            <span class="motore-copy-icon" aria-hidden="true"></span>
          </button>
        </div>
      `;
    })
    .join("");
}

function getFlowRowForRule(machineStruct, ruleTitle) {
  const flow = machineStruct?.flow && typeof machineStruct.flow === "object" ? machineStruct.flow : {};
  if (flow[ruleTitle]) return flow[ruleTitle];
  const targetKey = normalizeMotoreActivityKey(ruleTitle);
  if (!targetKey) return null;
  const entry = Object.entries(flow).find(([title]) => normalizeMotoreActivityKey(title) === targetKey);
  return entry?.[1] || null;
}

function buildMotoreFlowMapNodes(planningRules, coverage) {
  const nodeMap = new Map();
  const deptOrder = ["cad", "termodinamico", "elettrico", "generali"];
  const toHoursLabel = (hoursValue) => (Number.isFinite(hoursValue) ? `${hoursValue}h` : "--");

  planningRules.forEach((rule) => {
    const dependsRawList = parseAllowedActivitiesList(String(rule.flow?.depends || ""));
    const dependsKeys = Array.from(
      new Set(
        dependsRawList
          .map((value) => normalizeMotoreActivityKey(value))
          .filter((depKey) => depKey && depKey !== rule.key)
      )
    );
    nodeMap.set(rule.key, {
      key: rule.key,
      title: rule.title,
      dept: String(rule.dept || "").trim() || "Generale",
      hours: Number(rule.hours),
      hoursLabel: toHoursLabel(Number(rule.hours)),
      constraint: String(rule.flow?.constraint || "none").trim().toLowerCase(),
      dependsRawList,
      dependsKeys,
      dependsTitles: [],
      coverageCount: Number(coverage.get(rule.key)?.count || 0),
      coverageNames: (coverage.get(rule.key)?.names || []).slice(),
      level: 0,
    });
  });

  nodeMap.forEach((node) => {
    const validDeps = node.dependsKeys.filter((depKey) => depKey && nodeMap.has(depKey) && depKey !== node.key);
    node.dependsKeys = validDeps;
    node.dependsTitles = validDeps.map((depKey) => String(nodeMap.get(depKey)?.title || depKey));
  });

  const memo = new Map();
  const visiting = new Set();
  const computeLevel = (key) => {
    if (memo.has(key)) return memo.get(key);
    if (visiting.has(key)) return 0;
    const node = nodeMap.get(key);
    if (!node) return 0;
    visiting.add(key);
    let level = 0;
    if (node.dependsKeys.length) {
      level = node.dependsKeys.reduce((acc, depKey) => Math.max(acc, computeLevel(depKey) + 1), 0);
    }
    visiting.delete(key);
    memo.set(key, level);
    return level;
  };
  nodeMap.forEach((node) => {
    node.level = computeLevel(node.key);
  });

  const nodes = Array.from(nodeMap.values()).sort((a, b) => {
    if (a.level !== b.level) return a.level - b.level;
    const depA = normalizeDeptKey(a.dept);
    const depB = normalizeDeptKey(b.dept);
    const idxA = deptOrder.indexOf(depA);
    const idxB = deptOrder.indexOf(depB);
    if (idxA !== idxB) {
      if (idxA === -1) return 1;
      if (idxB === -1) return -1;
      return idxA - idxB;
    }
    return a.title.localeCompare(b.title, "it");
  });
  const maxLevel = nodes.reduce((max, node) => Math.max(max, Number(node.level || 0)), 0);
  return { nodes, maxLevel };
}

function drawMotoreFlowMapLinks(container) {
  const stage = container?.querySelector(".motore-flow-map-stage");
  const svg = stage?.querySelector(".motore-flow-map-links");
  if (!stage || !svg) return;
  const stageRect = stage.getBoundingClientRect();
  const width = Math.max(1, Math.ceil(stage.scrollWidth || stageRect.width));
  const height = Math.max(1, Math.ceil(stage.scrollHeight || stageRect.height));
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.style.width = `${width}px`;
  svg.style.height = `${height}px`;

  const cards = new Map(
    Array.from(stage.querySelectorAll(".motore-flow-map-card[data-key]"))
      .map((el) => [String(el.getAttribute("data-key") || ""), el])
      .filter(([key]) => Boolean(key))
  );

  const edges = [];
  cards.forEach((cardEl) => {
    const targetKey = String(cardEl.getAttribute("data-key") || "");
    const dependsKeys = parseAllowedActivitiesList(String(cardEl.getAttribute("data-depends") || ""))
      .map((value) => normalizeMotoreActivityKey(value))
      .filter((depKey) => depKey && depKey !== targetKey);
    if (!targetKey || !dependsKeys.length) return;
    dependsKeys.forEach((dependsKey) => {
      if (!cards.has(dependsKey)) return;
      const sourceEl = cards.get(dependsKey);
      if (!sourceEl) return;
      edges.push({ targetKey, dependsKey, targetEl: cardEl, sourceEl });
    });
  });

  const outCount = new Map();
  const inCount = new Map();
  edges.forEach((edge) => {
    outCount.set(edge.dependsKey, Number(outCount.get(edge.dependsKey) || 0) + 1);
    inCount.set(edge.targetKey, Number(inCount.get(edge.targetKey) || 0) + 1);
  });

  const outCursor = new Map();
  const inCursor = new Map();
  const slotOffset = (index, total, step = 5) => {
    if (total <= 1) return 0;
    const center = (total - 1) / 2;
    return (index - center) * step;
  };

  const paths = [];
  edges.forEach((edge) => {
    const { targetEl: cardEl, sourceEl, targetKey, dependsKey } = edge;
    const sourceRect = sourceEl.getBoundingClientRect();
    const targetRect = cardEl.getBoundingClientRect();

    const sourceTotal = Number(outCount.get(dependsKey) || 1);
    const sourceIndex = Number(outCursor.get(dependsKey) || 0);
    outCursor.set(dependsKey, sourceIndex + 1);
    const targetTotal = Number(inCount.get(targetKey) || 1);
    const targetIndex = Number(inCursor.get(targetKey) || 0);
    inCursor.set(targetKey, targetIndex + 1);

    const sourceYOffset = slotOffset(sourceIndex, sourceTotal, 5);
    const targetYOffset = slotOffset(targetIndex, targetTotal, 5);

    const x1 = sourceRect.right - stageRect.left + stage.scrollLeft + 2;
    const y1 = sourceRect.top + sourceRect.height / 2 - stageRect.top + stage.scrollTop + sourceYOffset;
    const x2 = targetRect.left - stageRect.left + stage.scrollLeft - 2;
    const y2 = targetRect.top + targetRect.height / 2 - stageRect.top + stage.scrollTop + targetYOffset;
    const isBackLink = x2 <= x1 + 6;
    const curve = Math.max(42, Math.abs(x2 - x1) * 0.5);
    const c1x = isBackLink ? x1 + 52 : x1 + curve;
    const c2x = isBackLink ? x2 - 52 : x2 - curve;
    const constraint = String(cardEl.getAttribute("data-constraint") || "none");
    const linkClass = constraint === "soft"
      ? "motore-flow-link is-soft"
      : constraint === "none"
        ? "motore-flow-link is-none"
        : "motore-flow-link is-hard";

    const d = `M ${x1.toFixed(1)} ${y1.toFixed(1)} C ${c1x.toFixed(1)} ${y1.toFixed(1)}, ${c2x.toFixed(1)} ${y2.toFixed(1)}, ${x2.toFixed(1)} ${y2.toFixed(1)}`;
    paths.push(`
      <path class="${linkClass}" d="${d}" pathLength="100" marker-end="url(#motoreFlowArrow)" />
      <path class="${linkClass} motore-flow-link-pulse" d="${d}" pathLength="100" />
    `);
  });

  svg.innerHTML = `
    <defs>
      <marker id="motoreFlowArrow" markerWidth="8" markerHeight="8" refX="6.2" refY="4" orient="auto" markerUnits="strokeWidth">
        <path d="M 0 0 L 8 4 L 0 8 z" fill="currentColor" />
      </marker>
    </defs>
    ${paths.join("")}
  `;
}

function bindMotoreFlowMapDrag() {
  if (!motoreSynopticFlow) return;
  let dragState = null;
  let redrawFrame = 0;

  const scheduleRedraw = () => {
    if (redrawFrame) return;
    redrawFrame = window.requestAnimationFrame(() => {
      redrawFrame = 0;
      drawMotoreFlowMapLinks(motoreSynopticFlow);
    });
  };

  const finishDrag = (event) => {
    if (!dragState) return;
    if (event && dragState.pointerId !== null && event.pointerId !== dragState.pointerId) return;
    const { cardEl, machineId, activityKey, x, y } = dragState;
    if (cardEl) {
      cardEl.classList.remove("is-dragging");
      cardEl.style.willChange = "";
      cardEl.dataset.offsetX = String(x);
      cardEl.dataset.offsetY = String(y);
      cardEl.style.transform = `translate(${x}px, ${y}px)`;
    }
    if (machineId && activityKey) {
      setMotoreFlowCardOffset(machineId, activityKey, x, y);
      saveMotoreFlowLayout();
    }
    dragState = null;
    scheduleRedraw();
  };

  motoreSynopticFlow.addEventListener("pointerdown", (event) => {
    const cardEl = event.target.closest(".motore-flow-map-card[data-key]");
    if (!cardEl || !motoreSynopticFlow.contains(cardEl)) return;
    if (event.button !== 0) return;
    const stageEl = cardEl.closest(".motore-flow-map-stage");
    const machineId = String(stageEl?.dataset.machineId || activeMachineId || "").trim();
    const activityKey = String(cardEl.dataset.key || "").trim();
    if (!machineId || !activityKey) return;
    event.preventDefault();

    const startX = event.clientX;
    const startY = event.clientY;
    const currentX = Number(cardEl.dataset.offsetX);
    const currentY = Number(cardEl.dataset.offsetY);
    dragState = {
      pointerId: typeof event.pointerId === "number" ? event.pointerId : null,
      cardEl,
      machineId,
      activityKey,
      startX,
      startY,
      baseX: Number.isFinite(currentX) ? currentX : 0,
      baseY: Number.isFinite(currentY) ? currentY : 0,
      x: Number.isFinite(currentX) ? currentX : 0,
      y: Number.isFinite(currentY) ? currentY : 0,
    };
    cardEl.classList.add("is-dragging");
    cardEl.style.willChange = "transform";
    scheduleRedraw();
  });

  window.addEventListener("pointermove", (event) => {
    if (!dragState) return;
    if (dragState.pointerId !== null && event.pointerId !== dragState.pointerId) return;
    const dx = event.clientX - dragState.startX;
    const dy = event.clientY - dragState.startY;
    dragState.x = dragState.baseX + dx;
    dragState.y = dragState.baseY + dy;
    dragState.cardEl.style.transform = `translate(${dragState.x}px, ${dragState.y}px)`;
    dragState.cardEl.dataset.offsetX = String(dragState.x);
    dragState.cardEl.dataset.offsetY = String(dragState.y);
    scheduleRedraw();
  });

  window.addEventListener("pointerup", finishDrag);
  window.addEventListener("pointercancel", finishDrag);
}

function renderMotoreSynoptic(machine, machineStruct) {
  if (!motoreSynopticMeta || !motoreSynopticFlow || !motoreSynopticResources || !machine || !machineStruct) return;
  const planningRules = (machine.rules || [])
    .filter((rule) => !isMotoreGenericActivityTitle(rule?.title || ""))
    .map((rule) => ({
      title: getCanonicalMotoreActivityTitle(rule.title || ""),
      key: normalizeMotoreActivityKey(rule.title || ""),
      dept: String(rule.dept || "").trim() || "Generale",
      hours: Number(rule.hours),
      flow: getFlowRowForRule(machineStruct, rule.title || ""),
    }))
    .filter((rule) => Boolean(rule.key));

  if (!planningRules.length) {
    motoreSynopticMeta.innerHTML = `<div class="motore-synoptic-empty">Nessuna attivita tecnica disponibile per il sinottico.</div>`;
    motoreSynopticFlow.innerHTML = "";
    motoreSynopticResources.innerHTML = "";
    return;
  }

  const coverage = new Map(planningRules.map((rule) => [rule.key, { count: 0, names: [] }]));
  const resourceById = new Map((motoreResources || []).map((res) => [String(res.id || ""), res]));
  const enabledRows = Object.entries(machineStruct.skills || {}).filter(([, row]) => Boolean(row?.enabled));
  enabledRows.forEach(([resourceId, row]) => {
    const resName = String(resourceById.get(resourceId)?.nome || `Risorsa ${resourceId}`);
    const selected = parseAllowedActivitiesList(row?.activities || "")
      .map((title) => normalizeMotoreActivityKey(title))
      .filter(Boolean);
    const selectedSet = new Set(selected);
    coverage.forEach((entry, key) => {
      if (!selectedSet.has(key)) return;
      entry.count += 1;
      entry.names.push(resName);
    });
  });

  const totalActivities = planningRules.length;
  const coveredActivities = planningRules.filter((rule) => Number(coverage.get(rule.key)?.count || 0) > 0).length;
  const weakActivities = planningRules.filter((rule) => Number(coverage.get(rule.key)?.count || 0) <= 1).length;
  const totalCapacity = enabledRows.reduce((sum, [, row]) => {
    const value = Number(row?.maxHours);
    return sum + (Number.isFinite(value) && value > 0 ? value : 0);
  }, 0);
  const cycleHours = planningRules.reduce((sum, rule) => {
    const value = Number(rule.hours);
    return sum + (Number.isFinite(value) && value > 0 ? value : 0);
  }, 0);
  const cycleDays = totalCapacity > 0 ? cycleHours / totalCapacity : 0;
  const coveragePct = totalActivities > 0 ? Math.round((coveredActivities / totalActivities) * 100) : 0;

  motoreSynopticMeta.innerHTML = `
    <div class="motore-synoptic-kpi">
      <div class="motore-synoptic-kpi-label">Copertura attivita</div>
      <div class="motore-synoptic-kpi-value">${coveredActivities}/${totalActivities}</div>
      <div class="motore-synoptic-kpi-note">${coveragePct}% coperte</div>
    </div>
    <div class="motore-synoptic-kpi">
      <div class="motore-synoptic-kpi-label">Rischio colli</div>
      <div class="motore-synoptic-kpi-value">${weakActivities}</div>
      <div class="motore-synoptic-kpi-note">attivita con <=1 risorsa</div>
    </div>
    <div class="motore-synoptic-kpi">
      <div class="motore-synoptic-kpi-label">Risorse abilitate</div>
      <div class="motore-synoptic-kpi-value">${enabledRows.length}</div>
      <div class="motore-synoptic-kpi-note">su ${motoreResources.length || 0} disponibili</div>
    </div>
    <div class="motore-synoptic-kpi">
      <div class="motore-synoptic-kpi-label">Durata teorica</div>
      <div class="motore-synoptic-kpi-value">${cycleDays > 0 ? cycleDays.toFixed(1) : "0.0"}g</div>
      <div class="motore-synoptic-kpi-note">${Math.round(cycleHours)}h / ${Math.round(totalCapacity)}h-giorno</div>
    </div>
  `;

  const flowModel = buildMotoreFlowMapNodes(planningRules, coverage);
  const columns = Array.from({ length: flowModel.maxLevel + 1 }, (_, level) => {
    const levelNodes = flowModel.nodes.filter((node) => node.level === level);
    const cards = levelNodes.length
      ? levelNodes
        .map((node) => {
          const offset = getMotoreFlowCardOffset(machine.id, node.key);
          const riskClass = node.coverageCount === 0 ? "is-risk-high" : node.coverageCount === 1 ? "is-risk-mid" : "is-risk-low";
          const dependsLabel = node.dependsTitles.length ? `da ${node.dependsTitles.join(" + ")}` : "start";
          const coverageTitle = node.coverageNames.length
            ? `Coperto da: ${node.coverageNames.join(", ")}`
            : "Nessuna risorsa copre questa attivita";
          return `
            <div
              class="motore-flow-map-card"
              data-key="${escapeHtml(node.key)}"
              data-depends="${escapeHtml(node.dependsKeys.join(", "))}"
              data-constraint="${escapeHtml(node.constraint || "none")}"
              data-offset-x="${offset.x}"
              data-offset-y="${offset.y}"
              style="transform: translate(${offset.x}px, ${offset.y}px);"
            >
              <div class="motore-flow-map-card-name">${escapeHtml(node.title)}</div>
              <div class="motore-flow-map-card-sub">${escapeHtml(node.dept)} · ${escapeHtml(node.hoursLabel)} · ${escapeHtml(dependsLabel)}</div>
              <div class="motore-flow-map-card-foot">
                <span class="motore-synoptic-chip ${riskClass}" title="${escapeHtml(coverageTitle)}">${node.coverageCount} ris.</span>
                <span class="motore-flow-map-constraint is-${escapeHtml(node.constraint || "none")}">${escapeHtml(node.constraint || "none")}</span>
              </div>
            </div>
          `;
        })
        .join("")
      : `<div class="motore-flow-map-empty">-</div>`;
    return `
      <div class="motore-flow-map-col">
        <div class="motore-flow-map-col-head">${level === 0 ? "Start" : `Step ${level}`}</div>
        <div class="motore-flow-map-col-body">${cards}</div>
      </div>
    `;
  }).join("");

  motoreSynopticFlow.innerHTML = `
    <div class="motore-flow-map-wrapper">
      <div class="motore-flow-map-legend">Mappa temporale: rettangoli = attivita, frecce = dipendenze, chip = copertura risorse.</div>
      <div class="motore-flow-map-stage" data-machine-id="${escapeHtml(machine.id)}">
        <svg class="motore-flow-map-links" aria-hidden="true"></svg>
        <div class="motore-flow-map-cols" style="--motore-flow-cols: ${flowModel.maxLevel + 1};">
          ${columns}
        </div>
      </div>
    </div>
  `;
  window.requestAnimationFrame(() => drawMotoreFlowMapLinks(motoreSynopticFlow));

  if (!enabledRows.length) {
    motoreSynopticResources.innerHTML = `<div class="motore-synoptic-empty">Nessuna risorsa abilitata per questa tipologia.</div>`;
    return;
  }

  const deptColumns = [
    { key: "cad", label: "CAD" },
    { key: "termodinamico", label: "Termodinamici" },
    { key: "elettrico", label: "Elettrici" },
  ];
  const groupedResources = new Map(
    deptColumns.map((col) => [col.key, []])
  );
  enabledRows.forEach(([resourceId, row]) => {
    const res = resourceById.get(String(resourceId)) || null;
    const deptKey = normalizeDeptKey(String(res?.reparto || ""));
    if (!groupedResources.has(deptKey)) return;
    const selectedKeys = parseAllowedActivitiesList(row?.activities || "")
      .map((title) => normalizeMotoreActivityKey(title))
      .filter(Boolean);
    const coveredCount = new Set(selectedKeys.filter((key) => coverage.has(key))).size;
    groupedResources.get(deptKey).push({
      name: String(res?.nome || `Risorsa ${resourceId}`),
      level: String(row?.level || "medio"),
      maxHours: Number(row?.maxHours),
      coveredCount,
    });
  });

  motoreSynopticResources.innerHTML = deptColumns
    .map((column) => {
      const entries = (groupedResources.get(column.key) || []).slice();
      const rows = entries
        .sort((a, b) => a.name.localeCompare(b.name, "it"))
        .map((entry) => {
          const maxHours = Number.isFinite(entry.maxHours) ? entry.maxHours : 0;
          return `
            <div class="motore-synoptic-res-row">
              <div class="motore-synoptic-res-main">
                <div class="motore-synoptic-res-name">${escapeHtml(entry.name)}</div>
                <div class="motore-synoptic-res-sub">${entry.coveredCount} att. coperte · ${maxHours}h/giorno</div>
              </div>
              <span class="motore-synoptic-level">${escapeHtml(entry.level)}</span>
            </div>
          `;
        })
        .join("") || `<div class="motore-synoptic-empty">Nessuna risorsa abilitata</div>`;
      return `
        <div class="motore-synoptic-res-group">
          <div class="motore-synoptic-res-group-title">${escapeHtml(column.label)}</div>
          ${rows}
        </div>
      `;
    })
    .join("");
}

async function copyTextToClipboard(text) {
  const value = String(text || "");
  if (!value) return false;
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(value);
      return true;
    } catch {
      // Fallback below.
    }
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = value;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return Boolean(ok);
  } catch {
    return false;
  }
}

function bindMotoreValidationCopy() {
  if (!motoreValidationList) return;
  motoreValidationList.addEventListener("click", async (event) => {
    const button = event.target.closest(".motore-copy-btn[data-copy-validation]");
    if (!button || !motoreValidationList.contains(button)) return;
    const idx = Number.parseInt(String(button.getAttribute("data-copy-validation") || ""), 10);
    if (!Number.isFinite(idx) || idx < 0) return;
    const text = String(motoreValidationItemsCache[idx]?.text || "");
    if (!text) return;
    const copied = await copyTextToClipboard(text);
    button.classList.toggle("is-copied", copied);
    button.setAttribute("title", copied ? "Copiato" : "Copia non riuscita");
    button.setAttribute("aria-label", copied ? "Copiato" : "Copia non riuscita");
    window.setTimeout(() => {
      button.classList.remove("is-copied");
      button.setAttribute("title", "Copia negli appunti");
      button.setAttribute("aria-label", "Copia negli appunti");
    }, 1200);
  });
}

function refreshMotoreValidation(machine, machineStruct) {
  const activeMachine = machine || getMachineVariantById(activeMachineId);
  if (!activeMachine) return;
  const struct = machineStruct || ensureMachineStructure(activeMachine.id);
  if (!struct) return;
  renderMotoreValidation(activeMachine, struct);
  renderMotoreSynoptic(activeMachine, struct);
}

function renderLegacyParamsCanvas(machineStruct) {
  if (!machineStruct) return;
  const params = normalizeMachineParams(machineStruct.params);
  machineStruct.params = params;
  if (motoreParams.bufferDays) {
    motoreParams.bufferDays.value = String(params.bufferDays);
    motoreParams.bufferDays.oninput = () => {
      const value = Number(motoreParams.bufferDays.value);
      machineStruct.params.bufferDays = Number.isFinite(value) ? Math.max(0, Math.round(value)) : defaultMachineParams.bufferDays;
      saveMotoreStructure();
    };
  }
  if (motoreParams.basePriority) {
    motoreParams.basePriority.value = normalizeMotoreBasePriority(params.basePriority);
    motoreParams.basePriority.onchange = () => {
      machineStruct.params.basePriority = normalizeMotoreBasePriority(motoreParams.basePriority.value);
      saveMotoreStructure();
    };
  }
  if (motoreParams.blockOverlap) {
    motoreParams.blockOverlap.checked = Boolean(params.blockOverlap);
    motoreParams.blockOverlap.onchange = () => {
      machineStruct.params.blockOverlap = Boolean(motoreParams.blockOverlap.checked);
      saveMotoreStructure();
    };
  }
}

function renderPriorityCanvas(machine, machineStruct) {
  if (!machine || !machineStruct) return;
  Object.entries(priorityInputs).forEach(([key, input]) => {
    if (!input) return;
    input.value = String(Number(machineStruct.priorities?.[key] ?? defaultPriorityWeights[key] ?? 0));
    input.oninput = () => {
      const value = Number(input.value);
      machineStruct.priorities[key] = Number.isFinite(value) ? Math.max(0, value) : 0;
      saveMotoreStructure();
      refreshMotoreValidation(machine, machineStruct);
    };
  });
}

function getDefaultFlowConstraint(rule) {
  const dep = String(rule?.depends || "").trim().toLowerCase();
  if (!dep || dep === "inizio" || dep === "da definire") return "none";
  return "hard";
}

function getCanonicalMotoreActivityTitle(value) {
  const raw = String(value || "").trim();
  const key = normalizeMotoreActivityKey(raw);
  if (!key) return raw;
  for (const group of motoreActivityPresetGroups) {
    for (const title of group?.titles || []) {
      if (normalizeMotoreActivityKey(title) === key) return title;
    }
  }
  return raw;
}

function getFlowDependencyGroups(machine) {
  const groups = new Map();
  (machine?.rules || []).forEach((rule) => {
    const rawTitle = String(rule?.title || "").trim();
    if (!rawTitle) return;
    if (isMotoreGenericActivityTitle(rawTitle)) return;
    const canonicalTitle = getCanonicalMotoreActivityTitle(rawTitle);
    const titleKey = normalizeMotoreActivityKey(canonicalTitle);
    if (!titleKey) return;
    const rawDept = String(rule?.dept || "").trim();
    const deptLabel = rawDept || formatMotoreDeptLabel(inferMotoreDeptFromTitle(canonicalTitle));
    if (!groups.has(deptLabel)) groups.set(deptLabel, new Map());
    const titles = groups.get(deptLabel);
    if (!titles.has(titleKey)) titles.set(titleKey, canonicalTitle);
  });

  const deptOrder = ["cad", "termodinamico", "elettrico", "generali"];
  return Array.from(groups.entries())
    .map(([dept, titlesMap]) => ({
      dept,
      titles: Array.from(titlesMap.values()).sort((a, b) => a.localeCompare(b, "it")),
    }))
    .sort((a, b) => {
      const keyA = normalizeDeptKey(a.dept);
      const keyB = normalizeDeptKey(b.dept);
      const idxA = deptOrder.indexOf(keyA);
      const idxB = deptOrder.indexOf(keyB);
      if (idxA !== idxB) {
        if (idxA === -1) return 1;
        if (idxB === -1) return -1;
        return idxA - idxB;
      }
      return String(a.dept || "").localeCompare(String(b.dept || ""), "it");
    });
}

function buildFlowDependencyGroupsForRule(groups, ruleTitle) {
  const selfKey = normalizeMotoreActivityKey(ruleTitle);
  return (groups || [])
    .map((group) => ({
      dept: group?.dept || "",
      titles: (group?.titles || []).filter((title) => normalizeMotoreActivityKey(title) !== selfKey),
    }))
    .filter((group) => group.titles.length > 0);
}

function normalizeFlowDependencyTitles(rawValue, dependencyGroups, ruleTitle) {
  const selfKey = normalizeMotoreActivityKey(ruleTitle);
  const titleByKey = new Map();
  (dependencyGroups || []).forEach((group) => {
    (group?.titles || []).forEach((title) => {
      const key = normalizeMotoreActivityKey(title);
      if (!key || key === selfKey || titleByKey.has(key)) return;
      titleByKey.set(key, title);
    });
  });
  const selectedRaw = parseAllowedActivitiesList(rawValue);
  const selectedTitles = [];
  const seen = new Set();
  selectedRaw.forEach((item) => {
    const key = normalizeMotoreActivityKey(item);
    if (!key || key === "da definire" || key === "inizio" || key === selfKey || seen.has(key)) return;
    seen.add(key);
    selectedTitles.push(titleByKey.get(key) || String(item || "").trim());
  });
  return selectedTitles;
}

function formatFlowDependencySummary(selectedTitles) {
  const list = Array.isArray(selectedTitles) ? selectedTitles.filter(Boolean) : [];
  if (!list.length) return "Nessuna dipendenza";
  if (list.length <= 2) return list.join(" + ");
  return `${list.length} predecessori`;
}

function buildFlowDependencyChecklistMarkup(groups, selectedTitles) {
  const selectedSet = new Set((selectedTitles || []).map((title) => normalizeMotoreActivityKey(title)));
  const chunks = [];
  (groups || []).forEach((group) => {
    const options = (group?.titles || [])
      .map((title) => {
        const key = normalizeMotoreActivityKey(title);
        const checked = selectedSet.has(key);
        return `
          <label class="motore-activity-option">
            <input
              data-field="dependsList"
              type="checkbox"
              value="${escapeHtml(title)}"${checked ? " checked" : ""}
            />
            <span>${escapeHtml(title)}</span>
          </label>
        `;
      })
      .join("");
    if (!options) return;
    chunks.push(`
      <div class="motore-activities-group">
        <div class="motore-activities-group-title">${escapeHtml(group.dept)}</div>
        <div class="motore-activities-group-options">${options}</div>
      </div>
    `);
  });
  if (!chunks.length) {
    return `<div class="motore-activities-empty">Nessuna dipendenza disponibile</div>`;
  }
  return chunks.join("");
}

function renderFlowCanvas(machine, machineStruct) {
  if (!flowTableBody || !machine || !machineStruct) return;
  flowTableBody.innerHTML = "";
  const dependencyGroups = getFlowDependencyGroups(machine);
  (machine.rules || []).forEach((rule) => {
    if (isMotoreGenericActivityTitle(rule?.title || "")) return;
    const key = String(rule.title || "").trim();
    if (!key) return;
    const current = machineStruct.flow[key] || {
      depends: rule.depends === "inizio" ? "" : String(rule.depends || ""),
      constraint: getDefaultFlowConstraint(rule),
      parallel: false,
      notes: "",
    };
    const rowDependencyGroups = buildFlowDependencyGroupsForRule(dependencyGroups, key);
    const normalizedDepends = normalizeFlowDependencyTitles(current.depends, rowDependencyGroups, key);
    const dependencySummary = formatFlowDependencySummary(normalizedDepends);
    const dependencyChecklist = buildFlowDependencyChecklistMarkup(rowDependencyGroups, normalizedDepends);
    current.depends = normalizedDepends.join(", ");
    machineStruct.flow[key] = current;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(key)}</td>
      <td>
        <details class="motore-activities-picker motore-depends-picker">
          <summary class="motore-activities-summary" title="Seleziona uno o piu predecessori per ${escapeHtml(key)}.">
            <span class="motore-activities-summary-text">${escapeHtml(dependencySummary)}</span>
          </summary>
          <div class="motore-activities-menu">
            <div class="motore-activities-actions">
              <button type="button" class="ghost motore-activities-action" data-action="depends-none">Nessuna</button>
            </div>
            ${dependencyChecklist}
          </div>
        </details>
      </td>
      <td>
        <select
          data-field="constraint"
          title="Livello vincolo per ${escapeHtml(key)}: hard obbligatorio, soft preferito, none libero."
        >
          <option value="hard"${current.constraint === "hard" ? " selected" : ""}>Rigido</option>
          <option value="soft"${current.constraint === "soft" ? " selected" : ""}>Flessibile</option>
          <option value="none"${current.constraint === "none" ? " selected" : ""}>Nessuno</option>
        </select>
      </td>
      <td>
        <input
          data-field="parallel"
          type="checkbox"${current.parallel ? " checked" : ""}
          title="Se attivo, ${escapeHtml(key)} puo essere pianificata in parallelo con altre attivita compatibili."
        />
      </td>
      <td>
        <input
          data-field="notes"
          type="text"
          value="${escapeHtml(String(current.notes || ""))}"
          placeholder="Eccezioni su questo flusso"
          title="Note operative o eccezioni specifiche per ${escapeHtml(key)}."
        />
      </td>
    `;

    const applyDependencySelection = () => {
      const selected = Array.from(tr.querySelectorAll('input[data-field="dependsList"]:checked'))
        .map((option) => String(option.value || "").trim())
        .filter(Boolean);
      const normalized = normalizeFlowDependencyTitles(selected.join(", "), rowDependencyGroups, key);
      current.depends = normalized.join(", ");
      const summaryEl = tr.querySelector(".motore-activities-summary-text");
      if (summaryEl) summaryEl.textContent = formatFlowDependencySummary(normalized);
      saveMotoreStructure();
      refreshMotoreValidation(machine, machineStruct);
    };

    tr.querySelectorAll("input, select").forEach((el) => {
      el.addEventListener("input", () => {
        const field = el.dataset.field;
        if (!field) return;
        if (field === "parallel") {
          current.parallel = Boolean(el.checked);
        } else if (field === "dependsList") {
          applyDependencySelection();
          return;
        } else {
          current[field] = el.value;
        }
        saveMotoreStructure();
        refreshMotoreValidation(machine, machineStruct);
      });
      el.addEventListener("change", () => {
        const field = el.dataset.field;
        if (!field) return;
        if (field === "parallel") {
          current.parallel = Boolean(el.checked);
        } else if (field === "dependsList") {
          applyDependencySelection();
          return;
        } else {
          current[field] = el.value;
        }
        saveMotoreStructure();
        refreshMotoreValidation(machine, machineStruct);
      });
    });

    const noneBtn = tr.querySelector('[data-action="depends-none"]');
    if (noneBtn) {
      noneBtn.addEventListener("click", () => {
        tr.querySelectorAll('input[data-field="dependsList"]').forEach((checkbox) => {
          checkbox.checked = false;
        });
        applyDependencySelection();
      });
    }

    flowTableBody.appendChild(tr);
  });
}

function parseAllowedActivitiesList(value) {
  const items = [];
  const seen = new Set();
  String(value || "")
    .split(/[,;\n]+/)
    .forEach((raw) => {
      const item = String(raw || "").trim();
      if (!item) return;
      const key = item.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      items.push(item);
    });
  return items;
}

function isMotoreGenericActivityTitle(value) {
  const key = normalizeMotoreActivityKey(value);
  if (!key) return false;
  return motoreGenericActivityTitles.some((title) => normalizeMotoreActivityKey(title) === key);
}

function normalizeMotoreActivityKey(value) {
  const normalized = String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
  if (!normalized) return "";
  if (normalized.startsWith("costruttivi")) return "costruttivi";
  if (normalized.includes("targhett") || normalized.includes("segnatura elettrica")) return "segnatura elettrica";
  if (normalized === "ped" || normalized.startsWith("ped ")) return "ped";
  if (normalized.includes("op ordini") || normalized.includes("ordini materiale")) return "op ordini";
  if (normalized.includes("progetto termodinamico")) return "progettazione termodinamica";
  if (normalized.includes("schema elettrico")) return "progettazione elettrica";
  return normalized;
}

function parseMotoreDateOnly(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const date = new Date(raw);
  if (!Number.isFinite(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatMotoreDateIso(date) {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatMotoreDateLabel(date) {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) return "--";
  return date.toLocaleDateString("it-IT");
}

function getMotoreWeekStart(value) {
  const date = value instanceof Date ? new Date(value) : parseMotoreDateOnly(value);
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  const day = (date.getDay() + 6) % 7;
  date.setDate(date.getDate() - day);
  return date;
}

function getMotoreIsoWeekInfo(value) {
  const date = value instanceof Date ? new Date(value) : parseMotoreDateOnly(value);
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  const weekDate = new Date(date);
  weekDate.setDate(weekDate.getDate() + 3 - ((weekDate.getDay() + 6) % 7));
  const weekYear = weekDate.getFullYear();
  const firstThursday = new Date(weekYear, 0, 4);
  firstThursday.setDate(firstThursday.getDate() + 3 - ((firstThursday.getDay() + 6) % 7));
  const weekNo = 1 + Math.round((weekDate - firstThursday) / (7 * 24 * 60 * 60 * 1000));
  return {
    year: weekYear,
    week: weekNo,
  };
}

function getMotoreWeekIndex(baseWeekStart, value) {
  if (!(baseWeekStart instanceof Date) || !Number.isFinite(baseWeekStart.getTime())) return null;
  const weekStart = getMotoreWeekStart(value);
  if (!weekStart) return null;
  const diff = weekStart.getTime() - baseWeekStart.getTime();
  return Math.floor(diff / (7 * 24 * 60 * 60 * 1000));
}

function formatMotoreWeekLabel(value) {
  const weekStart = getMotoreWeekStart(value);
  const info = getMotoreIsoWeekInfo(weekStart);
  if (!weekStart || !info) return "--";
  return `W${info.week} ${info.year}`;
}

function normalizeMotoreTextKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function formatMotoreCommessaCode(commessa) {
  const year = String(commessa?.anno || "").trim();
  const rawNum = String(commessa?.numero ?? "").trim();
  if (!year && !rawNum) return "";
  const num = /^\d+$/.test(rawNum) ? String(Number.parseInt(rawNum, 10)).padStart(3, "0") : rawNum;
  return year && num ? `${year}_${num}` : `${year}${num}`;
}

function matchMotoreCommessaByCode(commessa, codeFilter) {
  const filter = String(codeFilter || "").trim();
  if (!filter) return true;
  const normalizedFilter = normalizeMotoreTextKey(filter).replace(/\s+/g, "");
  const code = formatMotoreCommessaCode(commessa);
  const codeAlt = `${commessa?.anno || ""}${commessa?.numero || ""}`;
  const bag = [code, codeAlt, commessa?.numero, commessa?.titolo]
    .map((item) => normalizeMotoreTextKey(item).replace(/\s+/g, ""))
    .filter(Boolean);
  return bag.some((item) => item.includes(normalizedFilter));
}

function matchMotoreCommessaMachine(commessa, machine) {
  if (!machine) return true;
  const machineName = normalizeMotoreTextKey(machine.machineName || machine.name || "");
  if (!machineName) return true;
  const machineVariant = normalizeMotoreTextKey(machine.variantLabel || "standard") || "standard";
  const commessaType = normalizeMotoreTextKey(commessa?.tipo_macchina || "");
  if (commessaType) {
    const commessaVariant = normalizeMotoreTextKey(commessa?.variante_macchina || "standard") || "standard";
    return commessaType === machineName && commessaVariant === machineVariant;
  }
  const legacyTitle = normalizeMotoreTextKey(commessa?.titolo || "");
  if (!legacyTitle) return true;
  const nameMatch = legacyTitle.includes(machineName);
  if (!nameMatch) return false;
  if (machineVariant === "standard") return true;
  return legacyTitle.includes(machineVariant);
}

function rankMotoreCommessaPriority(commessa) {
  const key = normalizeMotoreTextKey(commessa?.priorita || "");
  if (key.includes("urgent") || key.includes("urgente") || key.includes("critica")) return 4;
  if (key.includes("alta") || key.includes("high")) return 3;
  if (key.includes("media") || key.includes("medio") || key.includes("normal")) return 2;
  if (key.includes("bassa") || key.includes("low")) return 1;
  return 2;
}

function getMotoreCommessaDueDate(commessa, codeFilter, deliveryOverride) {
  if (codeFilter && deliveryOverride && matchMotoreCommessaByCode(commessa, codeFilter)) {
    const overrideDate = parseMotoreDateOnly(deliveryOverride);
    if (overrideDate) return overrideDate;
  }
  return (
    parseMotoreDateOnly(commessa?.data_consegna_macchina) ||
    parseMotoreDateOnly(commessa?.data_consegna_prevista) ||
    parseMotoreDateOnly(commessa?.data_consegna) ||
    parseMotoreDateOnly(commessa?.data_ordine_telaio) ||
    parseMotoreDateOnly(commessa?.data_ingresso) ||
    null
  );
}

function getMotoreRulePlanningRows(machine, machineStruct) {
  const rows = (machine?.rules || [])
    .filter((rule) => !isMotoreGenericActivityTitle(rule?.title || ""))
    .map((rule) => {
      const title = getCanonicalMotoreActivityTitle(rule?.title || "");
      const key = normalizeMotoreActivityKey(title);
      const hours = Number(rule?.hours);
      const flow = getFlowRowForRule(machineStruct, title) || {};
      const depends = parseAllowedActivitiesList(String(flow?.depends || ""))
        .map((dep) => normalizeMotoreActivityKey(dep))
        .filter((dep) => dep && dep !== key);
      return {
        title,
        key,
        dept: String(rule?.dept || "").trim() || "Generale",
        hours: Number.isFinite(hours) && hours > 0 ? hours : 0,
        depends,
        constraint: String(flow?.constraint || "none").trim().toLowerCase(),
        parallel: Boolean(flow?.parallel),
      };
    })
    .filter((row) => Boolean(row.key));

  const rowByKey = new Map(rows.map((row) => [row.key, row]));
  rows.forEach((row) => {
    row.depends = Array.from(new Set(row.depends.filter((key) => rowByKey.has(key))));
  });

  const memo = new Map();
  const visiting = new Set();
  const getLevel = (key) => {
    if (memo.has(key)) return memo.get(key);
    if (visiting.has(key)) return 0;
    const row = rowByKey.get(key);
    if (!row || !row.depends.length) {
      memo.set(key, 0);
      return 0;
    }
    visiting.add(key);
    const level = row.depends.reduce((max, dep) => Math.max(max, getLevel(dep) + 1), 0);
    visiting.delete(key);
    memo.set(key, level);
    return level;
  };
  rows.forEach((row) => {
    row.level = getLevel(row.key);
  });

  rows.sort((a, b) => {
    if ((a.level || 0) !== (b.level || 0)) return (a.level || 0) - (b.level || 0);
    return a.title.localeCompare(b.title, "it");
  });
  return rows;
}

function buildMotoreCoverageMap(machineStruct, resources, planningRows) {
  const rowsByRule = new Map((planningRows || []).map((row) => [row.key, []]));
  const enabledResources = [];
  const skillRows = machineStruct?.skills && typeof machineStruct.skills === "object" ? machineStruct.skills : {};
  const resourceById = new Map((resources || []).map((res) => [String(res.id || ""), res]));

  Object.entries(skillRows).forEach(([resourceId, row]) => {
    if (!row?.enabled) return;
    const resource = resourceById.get(String(resourceId));
    if (!resource) return;
    const maxHours = Number(row?.maxHours);
    const dailyHours = Number.isFinite(maxHours) && maxHours > 0 ? maxHours : 0;
    const allowed = new Set(
      parseAllowedActivitiesList(row?.activities || "")
        .map((title) => normalizeMotoreActivityKey(title))
        .filter(Boolean)
    );
    const packed = {
      id: String(resource.id),
      nome: String(resource.nome || `Risorsa ${resource.id}`),
      reparto: String(resource.reparto || "").trim() || "Senza reparto",
      dailyHours,
      weeklyHours: dailyHours * 5,
      level: String(row?.level || "medio"),
      allowed,
    };
    enabledResources.push(packed);
    rowsByRule.forEach((list, key) => {
      if (!packed.allowed.has(key)) return;
      list.push(packed);
    });
  });

  return { enabledResources, rowsByRule };
}

function renderMotoreSimulation(result) {
  if (!motoreSimKpis || !motoreSimPlanList || !motoreSimIssues || !motoreSimMeta) return;
  motoreLastSimulationResult = result || null;
  const kpis = result?.kpis || {};
  motoreSimKpis.innerHTML = `
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Commesse analizzate</span>
      <span class="motore-kpi-value">${Number(kpis.commesse || 0)}</span>
    </div>
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Task pianificati</span>
      <span class="motore-kpi-value">${Number(kpis.assigned || 0)} / ${Number(kpis.totalTasks || 0)}</span>
    </div>
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Ore assegnate</span>
      <span class="motore-kpi-value">${Number(kpis.hoursAssigned || 0).toFixed(1)}h</span>
    </div>
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Saturazione stimata</span>
      <span class="motore-kpi-value">${Number(kpis.saturationPct || 0)}%</span>
    </div>
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Task bloccati</span>
      <span class="motore-kpi-value">${Number(kpis.blocked || 0)}</span>
    </div>
    <div class="motore-output-item kpi">
      <span class="motore-kpi-label">Ritardi previsti</span>
      <span class="motore-kpi-value">${Number(kpis.delays || 0)}</span>
    </div>
  `;

  const plan = Array.isArray(result?.plan) ? result.plan : [];
  if (!plan.length) {
    motoreSimPlanList.innerHTML = `<div class="motore-output-item">Nessuna assegnazione proposta.</div>`;
  } else {
    motoreSimPlanList.innerHTML = plan
      .map((item) => {
        const tone = item.delay ? " warn" : " ok";
        return `
          <div class="motore-output-item${tone}">
            ${escapeHtml(String(item.weekLabel || "--"))} · ${escapeHtml(String(item.resourceName || "-"))} · ${escapeHtml(String(item.commessaCode || "-"))}
            <br />
            ${escapeHtml(String(item.activityTitle || "-"))} · ${Number(item.hours || 0).toFixed(1)}h
          </div>
        `;
      })
      .join("");
  }

  const issues = Array.isArray(result?.issues) ? result.issues : [];
  if (!issues.length) {
    motoreSimIssues.innerHTML = `<div class="motore-output-item ok">Nessun blocco critico rilevato.</div>`;
  } else {
    motoreSimIssues.innerHTML = issues
      .map((item) => {
        const severity = String(item?.severity || "warn").toLowerCase();
        const klass = severity === "error" ? "error" : severity === "ok" ? "ok" : "warn";
        return `<div class="motore-output-item ${klass}">${escapeHtml(String(item?.text || ""))}</div>`;
      })
      .join("");
  }

  const runAt = result?.runAt instanceof Date ? result.runAt : new Date();
  motoreSimMeta.textContent = `Ultima simulazione: ${runAt.toLocaleString("it-IT")} · Orizzonte ${Number(kpis.horizonWeeks || 0)} settimane`;
}

function resetMotoreSimulationOutput(clearInputs = false) {
  if (clearInputs) {
    if (motoreSimOrderCode) motoreSimOrderCode.value = "";
    if (motoreSimFrameOrderDate) motoreSimFrameOrderDate.value = "";
    if (motoreSimDeliveryDate) motoreSimDeliveryDate.value = "";
  }
  if (motoreSimKpis) {
    motoreSimKpis.innerHTML = `<div class="motore-output-item">KPI disponibili dopo la prima simulazione.</div>`;
  }
  if (motoreSimPlanList) {
    motoreSimPlanList.innerHTML = `<div class="motore-output-item">Nessuna simulazione eseguita.</div>`;
  }
  if (motoreSimIssues) {
    motoreSimIssues.innerHTML = `<div class="motore-output-item">Nessun blocco analizzato.</div>`;
  }
  if (motoreSimMeta) {
    motoreSimMeta.textContent = "Simulazione protetta: usa backlog reale in sola lettura e non scrive su matrice.";
  }
  if (motoreSimApplyBtn) {
    motoreSimApplyBtn.disabled = true;
    motoreSimApplyBtn.title = "Disponibile in step successivo: applicazione controllata al piano reale.";
  }
  motoreLastSimulationResult = null;
}

async function loadMotoreSimulationData(machine, options = {}) {
  const codeFilter = String(options.codeFilter || "").trim();
  const maxCommesse = Number.isFinite(Number(options.maxCommesse)) ? Number(options.maxCommesse) : 80;
  const commesseRes = await motoreSupabase
    .from("commesse")
    .select(
      "id, anno, numero, titolo, cliente, tipo_macchina, variante_macchina, priorita, stato, data_ingresso, data_ordine_telaio, data_consegna_macchina, data_consegna_prevista, data_consegna"
    )
    .range(0, 5000);
  if (commesseRes.error) {
    return { ok: false, error: commesseRes.error.message };
  }
  const rows = Array.isArray(commesseRes.data) ? commesseRes.data : [];
  const filtered = rows
    .filter((row) => {
      const stato = normalizeMotoreTextKey(row?.stato || "");
      if (stato.includes("chius") || stato.includes("archiviat") || stato.includes("cancell")) return false;
      if (!matchMotoreCommessaByCode(row, codeFilter)) return false;
      return true;
    })
    .filter((row) => matchMotoreCommessaMachine(row, machine));
  const scoped = filtered
    .sort((a, b) => {
      const da = getMotoreCommessaDueDate(a, "", "")?.getTime() || Number.MAX_SAFE_INTEGER;
      const db = getMotoreCommessaDueDate(b, "", "")?.getTime() || Number.MAX_SAFE_INTEGER;
      if (da !== db) return da - db;
      return rankMotoreCommessaPriority(b) - rankMotoreCommessaPriority(a);
    })
    .slice(0, maxCommesse);

  const commessaIds = scoped.map((row) => row.id).filter(Boolean);
  let attivitaRows = [];
  if (commessaIds.length) {
    const attivitaRes = await motoreSupabase
      .from("attivita")
      .select("id, commessa_id, titolo, stato, risorsa_id, data_inizio, ore_stimate")
      .in("commessa_id", commessaIds)
      .range(0, 30000);
    if (attivitaRes.error) {
      return { ok: false, error: attivitaRes.error.message };
    }
    attivitaRows = Array.isArray(attivitaRes.data) ? attivitaRes.data : [];
  }
  return {
    ok: true,
    commesse: scoped,
    attivita: attivitaRows,
    totalCommesseScanned: rows.length,
    totalCommesseSelected: scoped.length,
  };
}

async function runMotoreSimulation() {
  if (!motoreSupabase) {
    renderMotoreSimulation({
      runAt: new Date(),
      kpis: { commesse: 0, totalTasks: 0, assigned: 0, blocked: 0, hoursAssigned: 0, saturationPct: 0, delays: 0, horizonWeeks: 0 },
      plan: [],
      issues: [{ severity: "error", text: "Supabase non disponibile: impossibile leggere backlog reale." }],
    });
    return;
  }
  const machine = getMachineVariantById(activeMachineId);
  if (!machine) return;
  const machineStruct = ensureMachineStructure(machine.id);
  if (!machineStruct) return;
  const planningRows = getMotoreRulePlanningRows(machine, machineStruct).filter((row) => row.hours > 0);
  if (!planningRows.length) {
    renderMotoreSimulation({
      runAt: new Date(),
      kpis: { commesse: 0, totalTasks: 0, assigned: 0, blocked: 0, hoursAssigned: 0, saturationPct: 0, delays: 0, horizonWeeks: 0 },
      plan: [],
      issues: [{ severity: "error", text: "Nessuna attivita tecnica con ore > 0: completa prima la timeline." }],
    });
    return;
  }

  const coverage = buildMotoreCoverageMap(machineStruct, motoreResources, planningRows);
  const issues = [];
  const coverageMissing = planningRows.filter((row) => !(coverage.rowsByRule.get(row.key) || []).length);
  coverageMissing.forEach((row) => {
    issues.push({
      severity: "error",
      text: `Skill mancanti: nessuna risorsa abilitata per "${row.title}".`,
    });
  });

  const codeFilter = String(motoreSimOrderCode?.value || "").trim();
  const frameOverride = String(motoreSimFrameOrderDate?.value || "").trim();
  const deliveryOverride = String(motoreSimDeliveryDate?.value || "").trim();
  const dataResult = await loadMotoreSimulationData(machine, { codeFilter, maxCommesse: 80 });
  if (!dataResult.ok) {
    renderMotoreSimulation({
      runAt: new Date(),
      kpis: { commesse: 0, totalTasks: 0, assigned: 0, blocked: 0, hoursAssigned: 0, saturationPct: 0, delays: 0, horizonWeeks: 0 },
      plan: [],
      issues: [{ severity: "error", text: `Errore lettura backlog: ${dataResult.error}` }],
    });
    return;
  }

  const validationReport = collectMotoreValidation(machine, machineStruct);
  const validationErrors = validationReport.items.filter((item) => item.severity === "error").length;
  if (validationErrors > 0) {
    issues.push({
      severity: "warn",
      text: `Configurazione con ${validationErrors} errori nel controllo coerenza: la simulazione potrebbe essere parziale.`,
    });
  }

  const todayWeekStart = getMotoreWeekStart(new Date());
  const horizonWeeks = 12;
  const horizonEnd = new Date(todayWeekStart);
  horizonEnd.setDate(horizonEnd.getDate() + horizonWeeks * 7 - 1);

  const resourceCapacityMap = new Map();
  let totalCapacityHours = 0;
  coverage.enabledResources.forEach((res) => {
    const weeks = Array.from({ length: horizonWeeks }, () => res.weeklyHours);
    totalCapacityHours += res.weeklyHours * horizonWeeks;
    resourceCapacityMap.set(res.id, weeks);
  });

  (dataResult.attivita || []).forEach((row) => {
    const rid = String(row?.risorsa_id || "");
    const slots = resourceCapacityMap.get(rid);
    if (!slots) return;
    const idx = getMotoreWeekIndex(todayWeekStart, row?.data_inizio);
    if (idx == null || idx < 0 || idx >= horizonWeeks) return;
    const rawHours = Number(row?.ore_stimate);
    const hours = Number.isFinite(rawHours) && rawHours > 0 ? rawHours : 8;
    slots[idx] = Math.max(0, slots[idx] - hours);
  });

  const activityByCommessa = new Map();
  (dataResult.attivita || []).forEach((row) => {
    const cid = String(row?.commessa_id || "");
    if (!cid) return;
    if (!activityByCommessa.has(cid)) activityByCommessa.set(cid, []);
    activityByCommessa.get(cid).push(row);
  });

  const sortedCommesse = (dataResult.commesse || [])
    .slice()
    .sort((a, b) => {
      const da = getMotoreCommessaDueDate(a, codeFilter, deliveryOverride)?.getTime() || Number.MAX_SAFE_INTEGER;
      const db = getMotoreCommessaDueDate(b, codeFilter, deliveryOverride)?.getTime() || Number.MAX_SAFE_INTEGER;
      if (da !== db) return da - db;
      return rankMotoreCommessaPriority(b) - rankMotoreCommessaPriority(a);
    });

  if (Number(dataResult.totalCommesseScanned || 0) > Number(dataResult.totalCommesseSelected || 0)) {
    issues.push({
      severity: "warn",
      text: `Simulazione su campione: ${dataResult.totalCommesseSelected} commesse su ${dataResult.totalCommesseScanned} totali (filtro/limite attivo).`,
    });
  }

  const plan = [];
  const taskQueue = [];
  let backlogTasks = 0;
  let alreadyPlannedOrDone = 0;
  const commessaTaskState = new Map();

  sortedCommesse.forEach((commessa) => {
    const commessaId = String(commessa.id || "");
    const dueDate = getMotoreCommessaDueDate(commessa, codeFilter, deliveryOverride);
    const dueIndexRaw = dueDate ? getMotoreWeekIndex(todayWeekStart, dueDate) : null;
    const dueWeekIndex = dueIndexRaw == null ? horizonWeeks - 1 : Math.max(0, dueIndexRaw);
    const priority = rankMotoreCommessaPriority(commessa);
    const code = formatMotoreCommessaCode(commessa) || String(commessaId);
    const commessaRows = activityByCommessa.get(commessaId) || [];
    const statusMap = new Map();
    planningRows.forEach((rule) => {
      const matches = commessaRows.filter((row) => normalizeMotoreActivityKey(row?.titolo || "") === rule.key);
      const doneRow = matches.find((row) => normalizeMotoreTextKey(row?.stato || "").includes("complet"));
      const plannedRow = matches.find((row) => normalizeMotoreTextKey(row?.stato || "").includes("pianificat"));
      if (doneRow) {
        const weekIdx = getMotoreWeekIndex(todayWeekStart, doneRow.data_inizio);
        statusMap.set(rule.key, { state: "done", week: weekIdx == null ? 0 : Math.max(0, weekIdx) });
        alreadyPlannedOrDone += 1;
        return;
      }
      if (plannedRow) {
        const weekIdx = getMotoreWeekIndex(todayWeekStart, plannedRow.data_inizio);
        statusMap.set(rule.key, { state: "planned", week: weekIdx == null ? 0 : Math.max(0, weekIdx) });
        alreadyPlannedOrDone += 1;
        return;
      }
      statusMap.set(rule.key, { state: "todo", week: null });
      backlogTasks += 1;
      taskQueue.push({
        commessaId,
        commessaCode: code,
        commessaTitle: String(commessa?.titolo || ""),
        ruleKey: rule.key,
        activityTitle: rule.title,
        dept: rule.dept,
        hours: rule.hours,
        level: Number(rule.level || 0),
        depends: rule.depends.slice(),
        constraint: rule.constraint || "none",
        parallel: Boolean(rule.parallel),
        dueDate,
        dueWeekIndex: rule.key === "telaio" && frameOverride
          ? Math.max(0, getMotoreWeekIndex(todayWeekStart, frameOverride) ?? dueWeekIndex)
          : dueWeekIndex,
        priority,
      });
    });
    commessaTaskState.set(commessaId, statusMap);
  });

  taskQueue.sort((a, b) => {
    if (a.dueWeekIndex !== b.dueWeekIndex) return a.dueWeekIndex - b.dueWeekIndex;
    if (a.priority !== b.priority) return b.priority - a.priority;
    if (a.level !== b.level) return a.level - b.level;
    return a.activityTitle.localeCompare(b.activityTitle, "it");
  });

  const scheduledMap = new Map();
  const lastResourceByCommessa = new Map();
  let assignedHours = 0;
  let blocked = 0;
  let delays = 0;

  taskQueue.forEach((task) => {
    const stateMap = commessaTaskState.get(task.commessaId) || new Map();
    let earliestWeek = 0;
    let dependencyBlocked = null;
    task.depends.forEach((depKey) => {
      const depState = stateMap.get(depKey);
      if (!depState) return;
      if (depState.state === "todo" && depState.week == null && !scheduledMap.has(`${task.commessaId}|${depKey}`)) {
        dependencyBlocked = depKey;
        return;
      }
      const depWeek = depState.week != null ? depState.week : Number(scheduledMap.get(`${task.commessaId}|${depKey}`) ?? 0);
      if (task.constraint !== "none") {
        const releaseWeek = depWeek + (task.parallel ? 0 : 1);
        earliestWeek = Math.max(earliestWeek, releaseWeek);
      } else {
        earliestWeek = Math.max(earliestWeek, depWeek);
      }
    });
    if (dependencyBlocked) {
      blocked += 1;
      issues.push({
        severity: "error",
        text: `${task.commessaCode} · ${task.activityTitle}: predecessore non pianificato (${dependencyBlocked}).`,
      });
      return;
    }

    const candidates = (coverage.rowsByRule.get(task.ruleKey) || []).filter((res) => res.weeklyHours > 0);
    if (!candidates.length) {
      blocked += 1;
      issues.push({
        severity: "error",
        text: `${task.commessaCode} · ${task.activityTitle}: nessuna risorsa compatibile.`,
      });
      return;
    }

    let picked = null;
    for (let week = Math.max(0, earliestWeek); week < horizonWeeks; week += 1) {
      const feasible = candidates
        .map((res) => {
          const weeks = resourceCapacityMap.get(res.id) || [];
          const remaining = Number(weeks[week] || 0);
          return { res, remaining, week };
        })
        .filter((entry) => entry.remaining >= task.hours);
      if (!feasible.length) continue;
      feasible.sort((a, b) => {
        const continuityOn = Boolean(machineStruct?.rules?.continuity);
        const balancerOn = Boolean(machineStruct?.rules?.balance);
        const lastA = lastResourceByCommessa.get(task.commessaId) === a.res.id ? 1 : 0;
        const lastB = lastResourceByCommessa.get(task.commessaId) === b.res.id ? 1 : 0;
        if (continuityOn && lastA !== lastB) return lastB - lastA;
        if (balancerOn && a.remaining !== b.remaining) return b.remaining - a.remaining;
        return a.res.nome.localeCompare(b.res.nome, "it");
      });
      picked = feasible[0];
      break;
    }

    if (!picked) {
      blocked += 1;
      issues.push({
        severity: "warn",
        text: `${task.commessaCode} · ${task.activityTitle}: capacita insufficiente entro ${horizonWeeks} settimane.`,
      });
      return;
    }

    const slots = resourceCapacityMap.get(picked.res.id) || [];
    slots[picked.week] = Math.max(0, Number(slots[picked.week] || 0) - task.hours);
    resourceCapacityMap.set(picked.res.id, slots);
    stateMap.set(task.ruleKey, { state: "planned", week: picked.week });
    scheduledMap.set(`${task.commessaId}|${task.ruleKey}`, picked.week);
    lastResourceByCommessa.set(task.commessaId, picked.res.id);
    assignedHours += task.hours;
    const weekDate = new Date(todayWeekStart);
    weekDate.setDate(weekDate.getDate() + picked.week * 7);
    const isDelay = picked.week > task.dueWeekIndex;
    if (isDelay) delays += 1;
    plan.push({
      weekIndex: picked.week,
      weekLabel: formatMotoreWeekLabel(weekDate),
      resourceId: picked.res.id,
      resourceName: picked.res.nome,
      commessaCode: task.commessaCode,
      commessaTitle: task.commessaTitle,
      activityTitle: task.activityTitle,
      hours: task.hours,
      delay: isDelay,
    });
  });

  plan.sort((a, b) => {
    if (a.weekIndex !== b.weekIndex) return a.weekIndex - b.weekIndex;
    return a.resourceName.localeCompare(b.resourceName, "it");
  });

  const saturationPct = totalCapacityHours > 0 ? Math.round((assignedHours / totalCapacityHours) * 100) : 0;
  if (!issues.length && backlogTasks > 0) {
    issues.push({ severity: "ok", text: "Backlog elaborato senza blocchi critici nel perimetro simulato." });
  }

  renderMotoreSimulation({
    runAt: new Date(),
    kpis: {
      commesse: sortedCommesse.length,
      totalTasks: backlogTasks,
      assigned: plan.length,
      blocked,
      delays,
      hoursAssigned: assignedHours,
      saturationPct,
      horizonWeeks,
      alreadyPlannedOrDone,
    },
    plan,
    issues,
  });
}

function bindMotoreSimulation() {
  resetMotoreSimulationOutput(false);
  if (motoreSimApplyBtn) {
    motoreSimApplyBtn.addEventListener("click", () => {
      window.alert("Step successivo: applicazione controllata al piano reale.");
    });
  }
  if (motoreSimResetBtn) {
    motoreSimResetBtn.addEventListener("click", () => {
      resetMotoreSimulationOutput(true);
    });
  }
  if (motoreSimRunBtn) {
    motoreSimRunBtn.addEventListener("click", async () => {
      const original = motoreSimRunBtn.textContent;
      motoreSimRunBtn.disabled = true;
      motoreSimRunBtn.textContent = "Simulazione...";
      try {
        await runMotoreSimulation();
      } finally {
        motoreSimRunBtn.textContent = original;
        motoreSimRunBtn.disabled = false;
      }
    });
  }
}

function getMachineActivityGroups(machine) {
  const grouped = new Map();
  (machine?.rules || []).forEach((rule) => {
    const title = String(rule?.title || "").trim();
    if (!title) return;
    const dept = String(rule?.dept || "Generale").trim() || "Generale";
    if (!grouped.has(dept)) grouped.set(dept, new Map());
    const map = grouped.get(dept);
    const key = title.toLowerCase();
    if (!map.has(key)) map.set(key, title);
  });

  return Array.from(grouped.entries())
    .map(([dept, map]) => ({
      dept,
      titles: Array.from(map.values()).sort((a, b) => a.localeCompare(b, "it")),
    }))
    .sort((a, b) => a.dept.localeCompare(b.dept, "it"));
}

function normalizeDeptKey(value) {
  const raw = String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "");
  if (!raw) return "";
  if (raw.includes("cad")) return "cad";
  if (raw.includes("termodinam")) return "termodinamico";
  if (raw.includes("elettr")) return "elettrico";
  if (raw.includes("generic") || raw.endsWith("generali")) return "generali";
  if (raw.includes("tutti") || raw === "generale" || raw === "general" || raw.includes("all")) return "tutti";
  if (raw.includes("senzareparto")) return "senzareparto";
  return raw;
}

function filterActivityGroupsForResource(groups, resourceDept) {
  const resourceKey = normalizeDeptKey(resourceDept);
  if (!resourceKey || resourceKey === "senzareparto" || resourceKey === "tutti") {
    return groups || [];
  }
  const filtered = (groups || []).filter((group) => {
    const groupKey = normalizeDeptKey(group?.dept || "");
    if (!groupKey) return false;
    return groupKey === resourceKey || groupKey === "generali";
  });
  return filtered.length ? filtered : groups || [];
}

function formatAllowedActivitiesSummary(selectedItems) {
  const list = Array.isArray(selectedItems) ? selectedItems.filter(Boolean) : [];
  if (!list.length) return "Nessuna attivita selezionata";
  if (list.length === 1) return list[0];
  return `${list.length} attivita selezionate`;
}

function buildActivityChecklistMarkup(groups, selectedItems) {
  const selectedSet = new Set((selectedItems || []).map((item) => String(item).toLowerCase()));
  const matched = new Set();
  const chunks = [];

  (groups || []).forEach((group) => {
    if (!group?.titles?.length) return;
    const options = group.titles
      .map((title) => {
        const key = String(title).toLowerCase();
        const isSelected = selectedSet.has(key);
        if (isSelected) matched.add(key);
        return `
          <label class="motore-activity-option">
            <input
              data-field="activities"
              type="checkbox"
              value="${escapeHtml(title)}"${isSelected ? " checked" : ""}
            />
            <span>${escapeHtml(title)}</span>
          </label>
        `;
      })
      .join("");
    if (!options) return;
    chunks.push(`
      <div class="motore-activities-group">
        <div class="motore-activities-group-title">${escapeHtml(group.dept)}</div>
        <div class="motore-activities-group-options">${options}</div>
      </div>
    `);
  });

  const customItems = (selectedItems || []).filter((item) => !matched.has(String(item).toLowerCase()));
  if (customItems.length) {
    const options = customItems
      .map(
        (title) => `
          <label class="motore-activity-option">
            <input
              data-field="activities"
              type="checkbox"
              value="${escapeHtml(title)}"
              checked
            />
            <span>${escapeHtml(title)}</span>
          </label>
        `
      )
      .join("");
    chunks.push(`
      <div class="motore-activities-group">
        <div class="motore-activities-group-title">Personalizzate</div>
        <div class="motore-activities-group-options">${options}</div>
      </div>
    `);
  }

  if (!chunks.length) {
    return `<div class="motore-activities-empty">Nessuna attivita disponibile</div>`;
  }
  return chunks.join("");
}

function renderSkillsCanvas(machine, machineStruct) {
  if (!skillTableBody || !skillSummary || !machine || !machineStruct) return;
  skillTableBody.innerHTML = "";
  const resources = Array.isArray(motoreResources) ? motoreResources : [];
  const activityGroups = getSkillActivityPresetGroups();
  if (!resources.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6">Nessuna risorsa disponibile. Verifica tabella risorse nella app principale.</td>`;
    skillTableBody.appendChild(tr);
    skillSummary.textContent = "Nessuna risorsa caricata.";
    refreshMotoreValidation(machine, machineStruct);
    return;
  }

  const groupedResources = resources
    .map((res) => ({
      ...res,
      repartoLabel: String(res?.reparto || "").trim() || "Senza reparto",
    }))
    .sort((a, b) => {
      const deptA = a.repartoLabel.toLowerCase() === "senza reparto" ? "zzz_senza_reparto" : a.repartoLabel;
      const deptB = b.repartoLabel.toLowerCase() === "senza reparto" ? "zzz_senza_reparto" : b.repartoLabel;
      const byDept = deptA.localeCompare(deptB, "it");
      if (byDept !== 0) return byDept;
      return String(a.nome || "").localeCompare(String(b.nome || ""), "it");
    });

  let enabledCount = 0;
  let lastDept = "";
  groupedResources.forEach((res) => {
    if (res.repartoLabel !== lastDept) {
      lastDept = res.repartoLabel;
      const groupTr = document.createElement("tr");
      groupTr.className = "motore-skill-group-row";
      groupTr.innerHTML = `<td colspan="6">${escapeHtml(lastDept)}</td>`;
      skillTableBody.appendChild(groupTr);
    }

    const resId = String(res.id || "");
    if (!resId) return;
    const current = machineStruct.skills[resId] || {
      enabled: false,
      level: "medio",
      activities: "",
      maxHours: 8,
    };
    const selectedActivities = parseAllowedActivitiesList(current.activities);
    const visibleActivityGroups = filterActivityGroupsForResource(activityGroups, res.repartoLabel);
    const activitiesChecklist = buildActivityChecklistMarkup(visibleActivityGroups, selectedActivities);
    const activitiesSummary = formatAllowedActivitiesSummary(selectedActivities);
    machineStruct.skills[resId] = current;
    if (current.enabled) enabledCount += 1;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(res.nome || "-"))}</td>
      <td>${escapeHtml(res.repartoLabel)}</td>
      <td>
        <input
          data-field="enabled"
          type="checkbox"${current.enabled ? " checked" : ""}
          title="Se attivo, la risorsa ${escapeHtml(String(res.nome || ""))} e assegnabile a questa tipologia."
        />
      </td>
      <td>
        <select
          data-field="level"
          title="Livello competenza della risorsa ${escapeHtml(String(res.nome || ""))} sulla tipologia selezionata."
        >
          <option value="base"${current.level === "base" ? " selected" : ""}>Base</option>
          <option value="medio"${current.level === "medio" ? " selected" : ""}>Medio</option>
          <option value="esperto"${current.level === "esperto" ? " selected" : ""}>Esperto</option>
        </select>
      </td>
      <td>
        <details class="motore-activities-picker">
          <summary class="motore-activities-summary" title="Seleziona attivita consentite per questa risorsa.">
            <span class="motore-activities-summary-text">${escapeHtml(activitiesSummary)}</span>
          </summary>
          <div class="motore-activities-menu">
            <div class="motore-activities-actions">
              <button type="button" class="ghost motore-activities-action" data-action="activities-all">Abilita tutte</button>
              <button type="button" class="ghost motore-activities-action" data-action="activities-none">Nessuna</button>
            </div>
            ${activitiesChecklist}
          </div>
        </details>
      </td>
      <td>
        <input
          data-field="maxHours"
          type="number"
          min="0"
          max="24"
          step="0.5"
          value="${escapeHtml(String(current.maxHours ?? 8))}"
          title="Massimo ore/giorno assegnabili a ${escapeHtml(String(res.nome || ""))} su questa tipologia."
        />
      </td>
    `;

    const applyActivitiesSelection = () => {
      const selected = Array.from(tr.querySelectorAll('input[data-field="activities"]:checked'))
        .map((option) => String(option.value || "").trim())
        .filter(Boolean);
      current.activities = selected.join(", ");
      const summaryEl = tr.querySelector(".motore-activities-summary-text");
      if (summaryEl) summaryEl.textContent = formatAllowedActivitiesSummary(selected);
      saveMotoreStructure();
      const enabledNow = Object.values(machineStruct.skills || {}).filter((v) => Boolean(v?.enabled)).length;
      skillSummary.textContent = `${enabledNow} risorse abilitate su ${resources.length} per ${machine.machineName}.`;
      refreshMotoreValidation(machine, machineStruct);
    };

    tr.querySelectorAll("input, select").forEach((el) => {
      const apply = () => {
        const field = el.dataset.field;
        if (!field) return;
        if (field === "enabled") {
          current.enabled = Boolean(el.checked);
        } else if (field === "maxHours") {
          const num = Number(el.value);
          current.maxHours = Number.isFinite(num) ? Math.max(0, num) : 0;
        } else if (field === "activities") {
          applyActivitiesSelection();
          return;
        } else {
          current[field] = el.value;
        }
        saveMotoreStructure();
        const enabledNow = Object.values(machineStruct.skills || {}).filter((v) => Boolean(v?.enabled)).length;
        skillSummary.textContent = `${enabledNow} risorse abilitate su ${resources.length} per ${machine.machineName}.`;
        refreshMotoreValidation(machine, machineStruct);
      };
      el.addEventListener("input", apply);
      el.addEventListener("change", apply);
    });

    const allBtn = tr.querySelector('[data-action="activities-all"]');
    const noneBtn = tr.querySelector('[data-action="activities-none"]');
    const setActivitiesChecked = (checked) => {
      tr.querySelectorAll('input[data-field="activities"]').forEach((checkbox) => {
        checkbox.checked = Boolean(checked);
      });
      applyActivitiesSelection();
    };
    if (allBtn) {
      allBtn.addEventListener("click", () => setActivitiesChecked(true));
    }
    if (noneBtn) {
      noneBtn.addEventListener("click", () => setActivitiesChecked(false));
    }

    skillTableBody.appendChild(tr);
  });

  skillSummary.textContent = `${enabledCount} risorse abilitate su ${resources.length} per ${machine.machineName}.`;
  refreshMotoreValidation(machine, machineStruct);
}

function renderPlannerRulesCanvas(machineStruct) {
  if (!machineStruct) return;
  if (plannerRules.hardSkills) {
    plannerRules.hardSkills.checked = Boolean(machineStruct.rules.hardSkills);
    plannerRules.hardSkills.onchange = () => {
      machineStruct.rules.hardSkills = Boolean(plannerRules.hardSkills.checked);
      saveMotoreStructure();
    };
  }
  if (plannerRules.continuity) {
    plannerRules.continuity.checked = Boolean(machineStruct.rules.continuity);
    plannerRules.continuity.onchange = () => {
      machineStruct.rules.continuity = Boolean(plannerRules.continuity.checked);
      saveMotoreStructure();
    };
  }
  if (plannerRules.balance) {
    plannerRules.balance.checked = Boolean(machineStruct.rules.balance);
    plannerRules.balance.onchange = () => {
      machineStruct.rules.balance = Boolean(plannerRules.balance.checked);
      saveMotoreStructure();
    };
  }
  if (plannerRules.notes) {
    plannerRules.notes.value = String(machineStruct.rules.notes || "");
    plannerRules.notes.oninput = () => {
      machineStruct.rules.notes = plannerRules.notes.value;
      saveMotoreStructure();
    };
  }
}

function renderPlannerCanvas(machine) {
  if (!machine) return;
  const machineStruct = ensureMachineStructure(machine.id);
  if (!machineStruct) return;
  const removedFlowTitles = pruneOrphanFlowEntries(machine, machineStruct);
  if (removedFlowTitles.length) saveMotoreStructure();
  renderLegacyParamsCanvas(machineStruct);
  renderPriorityCanvas(machine, machineStruct);
  renderFlowCanvas(machine, machineStruct);
  renderSkillsCanvas(machine, machineStruct);
  renderPlannerRulesCanvas(machineStruct);
  refreshMotoreValidation(machine, machineStruct);
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function setActiveMachine(id) {
  const machine = getMachineVariantById(id);
  if (!machine) return;
  activeMachineId = machine.id;
  const title = machine.variantLabel && machine.variantLabel.toLowerCase() !== "standard"
    ? `${machine.machineName} · ${machine.variantLabel}`
    : machine.machineName;
  if (ruleTitle) ruleTitle.textContent = `Regola: ${title}`;
  renderRuleList(machine);
  renderPlannerCanvas(machine);
  renderMachineList(machine.id);
  if (machineSelect && machineSelect.value !== machine.id) {
    machineSelect.value = machine.id;
  }
  renderMotoreTransferPanel(machine.id);
  loadRulesForActive(machine);
  loadPlannerCanvasForActive(machine);
}

const firstId = getMachineVariants()[0]?.id || "";
if (firstId) setActiveMachine(firstId);
renderMachineSelect(firstId);

initMotoreTheme();
refreshMotoreSandboxUI();
if (motoreThemeToggleBtn) {
  motoreThemeToggleBtn.addEventListener("click", toggleMotoreTheme);
}
if (motoreSandboxToggleBtn) {
  motoreSandboxToggleBtn.addEventListener("click", toggleMotoreSandboxMode);
}
bindMotoreValidationCopy();
bindMotoreFlowMapDrag();
bindMotoreSimulation();
bindMotoreTransferPanel();
bindMotoreTimelineCompare();
setMotoreTimelineCompareMode(false);
window.addEventListener("resize", () => {
  window.requestAnimationFrame(() => drawMotoreFlowMapLinks(motoreSynopticFlow));
});

checkMotoreAccess();
if (motoreSupabase) {
  motoreSupabase.auth.onAuthStateChange(() => {
    checkMotoreAccess();
  });
}

if (motoreSaveBtn) {
  motoreSaveBtn.addEventListener("click", async () => {
    const machine = getMachineVariantById(activeMachineId);
    if (!machine) return;
    motoreSaveBtn.disabled = true;
    const original = motoreSaveBtn.textContent;
    if (motoreSandboxMode) {
      saveMotoreStructure();
      motoreSaveBtn.textContent = "Salvato locale";
      setTimeout(() => {
        motoreSaveBtn.textContent = original;
        motoreSaveBtn.disabled = false;
      }, 1000);
      return;
    }
    motoreSaveBtn.textContent = "Saving...";
    const rulesResult = await saveRulesForActive(machine);
    if (!rulesResult.ok) {
      motoreSaveBtn.textContent = "Errore";
      setTimeout(() => {
        motoreSaveBtn.textContent = original;
        motoreSaveBtn.disabled = false;
      }, 1200);
      return;
    }
    const canvasResult = await savePlannerCanvasForActive(machine);
    if (!canvasResult.ok) {
      motoreSaveBtn.textContent = "Errore";
      setTimeout(() => {
        motoreSaveBtn.textContent = original;
        motoreSaveBtn.disabled = false;
      }, 1200);
      return;
    }
    motoreSaveBtn.textContent = "Salvato";
    setTimeout(() => {
      motoreSaveBtn.textContent = original;
      motoreSaveBtn.disabled = false;
    }, 1200);
  });
}
