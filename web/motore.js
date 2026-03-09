const baseRuleTemplate = [
  { title: "CAD layout generale", dept: "CAD", hours: null, depends: "inizio" },
  { title: "Disegni costruttivi", dept: "CAD", hours: null, depends: "layout" },
  { title: "Distinte base", dept: "CAD", hours: null, depends: "layout" },
  { title: "Progetto termodinamico", dept: "Termodinamici", hours: null, depends: "layout" },
  { title: "Schema elettrico", dept: "Elettrici", hours: null, depends: "layout" },
];

const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

const motoreSupabase = window.supabase
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      auth: { persistSession: true, autoRefreshToken: true },
    })
  : null;

const gate = document.getElementById("motoreAccessGate");
const gateTitle = document.getElementById("motoreGateTitle");
const gateText = document.getElementById("motoreGateText");
const ruleWarnings = document.getElementById("motoreRuleWarnings");

let activeMachineId = "";
let templateLoaded = false;
let metaLoaded = false;
let rulesLoadToken = 0;
const motoreTipologie = new Map();
const motoreVariantIds = new Map();

function showGate(title, text) {
  if (gateTitle) gateTitle.textContent = title;
  if (gateText) gateText.textContent = text;
  document.body.classList.add("motore-locked");
  if (gate) gate.classList.remove("hidden");
}

function hideGate() {
  document.body.classList.remove("motore-locked");
  if (gate) gate.classList.add("hidden");
}

async function checkMotoreAccess() {
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

const machineList = document.getElementById("motoreMachineList");
const ruleTitle = document.getElementById("motoreRuleTitle");
const ruleList = document.getElementById("motoreRuleList");
const machineSelect = document.getElementById("motoreMachineSelect");
const motoreSaveBtn = document.getElementById("motoreSaveBtn");

async function loadActivityTemplate() {
  if (!motoreSupabase) return baseRuleTemplate.map((r) => r.title);
  let rows = null;
  try {
    const result = await motoreSupabase
      .from("attivita")
      .select("titolo")
      .range(0, 9999);
    rows = result.data || null;
  } catch {
    rows = null;
  }

  const excluded = new Set(
    ["ASSENTE", "ALTRO", "SEGNATURA ELETTRICA", "PED", "ORDINI MATERIALE", "KICKOFF COMMESSA"].map((t) =>
      t.toLowerCase()
    )
  );
  const titlesMap = new Map();
  (rows || []).forEach((row) => {
    const title = String(row?.titolo || "").trim();
    if (!title) return;
    if (excluded.has(title.toLowerCase())) return;
    const key = title.toLowerCase();
    if (!titlesMap.has(key)) titlesMap.set(key, title);
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
        const rule = existing.get(key);
        if (rule) {
          return { ...rule, title };
        }
        return { title, dept: "Generale", hours: null, depends: "da definire" };
      });
      variant.rules = nextRules;
    });
  });
}

async function initMotoreTemplate() {
  if (templateLoaded) return;
  const titles = await loadActivityTemplate();
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
    const card = document.createElement("div");
    card.className = "motore-rule-card";
    card.innerHTML = `
      <div class="motore-rule-handle"></div>
      <div class="motore-rule-meta">
        <div class="motore-rule-title">${r.title}</div>
        <div class="motore-rule-sub">${r.dept} · dipende da: ${r.depends}</div>
      </div>
      <label class="motore-rule-duration">
        <input class="motore-hours-input" type="number" min="0" step="0.5" value="${r.hours ?? ""}" />
        <span>h</span>
      </label>
    `;
    const input = card.querySelector(".motore-hours-input");
    if (input) {
      input.addEventListener("input", () => {
        const value = Number(input.value);
        r.hours = Number.isFinite(value) ? value : null;
        renderRuleWarnings(machine);
      });
    }
    ruleList.appendChild(card);
  });
  renderRuleWarnings(machine);
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
  renderMachineList(machine.id);
  if (machineSelect && machineSelect.value !== machine.id) {
    machineSelect.value = machine.id;
  }
  loadRulesForActive(machine);
}

const firstId = getMachineVariants()[0]?.id || "";
if (firstId) setActiveMachine(firstId);
renderMachineSelect(firstId);

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
    motoreSaveBtn.textContent = "Saving...";
    const result = await saveRulesForActive(machine);
    if (!result.ok) {
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
