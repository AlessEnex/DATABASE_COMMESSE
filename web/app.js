(() => {
  if (window.__commesseAppLoaded) return;
  window.__commesseAppLoaded = true;

const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const state = {
  session: null,
  profile: null,
  reparti: [],
  commesse: [],
  risorse: [],
  selected: null,
  canWrite: false,
  realtime: null,
};

const el = (id) => document.getElementById(id);

const authStatus = el("authStatus");
const roleBadge = el("roleBadge");
const statusMsg = el("statusMsg");

const emailInput = el("emailInput");
const loginBtn = el("loginBtn");
const logoutBtn = el("logoutBtn");

const commesseList = el("commesseList");
const searchInput = el("searchInput");
const yearFilter = el("yearFilter");
const statusFilter = el("statusFilter");
const openImportBtn = el("openImportBtn");

const detailForm = el("detailForm");
const selectedCode = el("selectedCode");
const updateBtn = el("updateBtn");
const clearSelectionBtn = el("clearSelectionBtn");

const newForm = el("newForm");
const newRepartiChecks = el("newRepartiChecks");

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
const resourcesList = el("resourcesList");
const resourceForm = el("resourceForm");
const resourceName = el("resourceName");
const resourceReparto = el("resourceReparto");
const resourceActive = el("resourceActive");

const importModal = el("importModal");
const importTextarea = el("importTextarea");
const importPreview = el("importPreview");
const importCommitBtn = el("importCommitBtn");
const closeImportBtn = el("closeImportBtn");
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
const matrixReportGrid = el("matrixReportGrid");
const matrixReportRange = el("matrixReportRange");
const matrixReportSortBtn = el("matrixReportSortBtn");

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

const d = {
  codice: el("d_codice"),
  titolo: el("d_titolo"),
  cliente: el("d_cliente"),
  stato: el("d_stato"),
  priorita: el("d_priorita"),
  data_ingresso: el("d_data_ingresso"),
  data_consegna: el("d_data_consegna"),
  note: el("d_note"),
};

const n = {
  codice: el("n_codice"),
  titolo: el("n_titolo"),
  cliente: el("n_cliente"),
  stato: el("n_stato"),
  priorita: el("n_priorita"),
  data_ingresso: el("n_data_ingresso"),
  data_consegna: el("n_data_consegna"),
  note: el("n_note"),
};

const calendarState = {
  view: "week",
  date: new Date(),
  attivita: [],
};

const matrixState = {
  date: new Date(),
  attivita: [],
  risorse: [],
  draggingId: null,
  view: "week",
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
  const hue = seed % 360;
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
    matrixState.selectedCommesse = new Set(matrixState.filterDraft);
    if (matrixCommessa) {
      const only = matrixState.selectedCommesse.size === 1 ? Array.from(matrixState.selectedCommesse)[0] : "";
      matrixCommessa.value = only;
    }
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
    const isAltro = attivita.titolo && attivita.titolo.toLowerCase() === "altro";
    activityAltroNoteWrap.classList.toggle("hidden", !isAltro);
    activityAltroNote.value = isAltro ? attivita.descrizione || "" : "";
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
        if (activityAltroNote && activityAltroNoteWrap) {
          const hasAltro = Array.from(matrixState.editingTitles).some((t) => t.toLowerCase() === "altro");
          activityAltroNoteWrap.classList.toggle("hidden", !hasAltro);
          if (!hasAltro) activityAltroNote.value = "";
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
    ? `${attivita.titolo} • ${commessaLabel} • ${formatDateLocal(start)}`
    : `${attivita.titolo} • ${formatDateLocal(start)}`;
  activityModal.classList.remove("hidden");
  activityDurationInput.focus();
}

function closeActivityModal() {
  if (!activityModal) return;
  activityModal.classList.add("hidden");
  matrixState.editingAttivita = null;
}

async function saveActivityDuration() {
  const attivita = matrixState.editingAttivita;
  if (!attivita) return;
  const days = Math.max(1, Number(activityDurationInput.value || 1));
  const titles = Array.from(matrixState.editingTitles);
  if (!titles.length) {
    setStatus("Seleziona almeno una attivita.", "error");
    return;
  }
  const hasAssente = titles.some((t) => isAssenteTitle(t));
  if (hasAssente && titles.length > 1) {
    setStatus("ASSENTE va assegnata da sola.", "error");
    return;
  }
  const notAllowed = titles.filter((t) => !isActivityAllowedForRisorsa(t, attivita.risorsa_id));
  if (notAllowed.length) {
    setStatus("Alcune attivita non sono disponibili per questo reparto.", "error");
    return;
  }
  const altroNote = activityAltroNote ? activityAltroNote.value.trim() : "";
  const assenzaOreRaw = activityAssenzaOre ? activityAssenzaOre.value : "";
  const assenzaOre = Number(assenzaOreRaw);
  if (hasAssente && (!assenzaOreRaw || Number.isNaN(assenzaOre) || assenzaOre <= 0)) {
    setStatus("Inserisci le ore di assenza.", "error");
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
      descrizione: firstTitle.toLowerCase() === "altro" ? altroNote || null : null,
      ore_assenza: isAssenteTitle(firstTitle) ? assenzaOre || null : null,
      commessa_id: isAssenteTitle(firstTitle) ? null : attivita.commessa_id,
    })
    .eq("id", attivita.id);
  if (updError) {
    setStatus(`Errore aggiornamento: ${updError.message}`, "error");
    return;
  }
  if (titles.length > 1) {
    const rows = titles.slice(1).map((title) => ({
      commessa_id: isAssenteTitle(title) ? null : attivita.commessa_id,
      titolo: title,
      descrizione: title.toLowerCase() === "altro" ? altroNote || null : null,
      ore_assenza: isAssenteTitle(title) ? assenzaOre || null : null,
      risorsa_id: attivita.risorsa_id,
      data_inizio: attivita.data_inizio,
      data_fine: newEnd.toISOString(),
      stato: attivita.stato,
      reparto_id: attivita.reparto_id,
    }));
    const { error: insError } = await supabase.from("attivita").insert(rows);
    if (insError) {
      setStatus(`Errore creazione: ${insError.message}`, "error");
      return;
    }
  }
  if (keyForChain && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
    const ok = await shiftDependentPhases(keyForChain, deltaDays, dependentToShift);
    if (!ok) return;
  }
  setStatus(titles.length > 1 ? "Attivita aggiornate." : "Attivita aggiornata.", "ok");
  closeActivityModal();
  await loadMatrixAttivita();
}

async function deleteActivity() {
  const attivita = matrixState.editingAttivita;
  if (!attivita) return;
  const ok = confirm("Eliminare questa attivita?");
  if (!ok) return;
  const { error } = await supabase.from("attivita").delete().eq("id", attivita.id);
  if (error) {
    setStatus(`Errore rimozione: ${error.message}`, "error");
    return;
  }
  setStatus("Attivita rimossa.", "ok");
  closeActivityModal();
  await loadMatrixAttivita();
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
    setStatus(`Errore aggiornamento: ${error.message}`, "error");
    return;
  }
  if (isPhaseKey(ctx.attivita.titolo) && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
    const ok = await shiftDependentPhases(ctx.attivita, deltaDays, dependentToShift);
    if (!ok) return;
  }
  setStatus("Durata aggiornata.", "ok");
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
}

function closeResourcesModal() {
  if (!resourcesModal) return;
  resourcesModal.classList.add("hidden");
}

function openProgressModal() {
  if (!progressModal) return;
  progressModal.classList.remove("hidden");
  if (progressMeta) progressMeta.textContent = "Seleziona una commessa.";
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
        setStatus(`Errore avanzamento: ${error.message}`, "error");
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
  if (statusTimer) clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    statusMsg.classList.remove("ok", "error");
  }, 4000);
}

function setAuthLoading(isLoading) {
  loginBtn.disabled = isLoading;
  emailInput.disabled = isLoading;
  loginBtn.textContent = isLoading ? "Invio..." : "Invia link";
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

function setWriteAccess(canWrite) {
  state.canWrite = canWrite;
  const disabled = !canWrite;
  const inputs = document.querySelectorAll(
    "#detailForm input, #detailForm select, #detailForm textarea, #newForm input, #newForm select, #newForm textarea, #addRepartoForm select, #resourceForm input, #resourceForm select, #resourceForm button, #resourcesList input, #resourcesList select, #resourcesList button"
  );
  inputs.forEach((input) => (input.disabled = disabled));
  updateBtn.disabled = disabled;
  addRepartoForm.querySelector("button").disabled = disabled;
  newForm.querySelector("button").disabled = disabled;
  importCommitBtn.disabled = disabled;
  openImportBtn.disabled = disabled;
  if (matrixCommessa) matrixCommessa.disabled = disabled;
  if (matrixAttivita) matrixAttivita.disabled = disabled;
}

async function signIn() {
  const email = emailInput.value.trim();
  if (!email) {
    setStatus("Inserisci una email valida.", "error");
    return;
  }
  setAuthLoading(true);
  const { error } = await supabase.auth.signInWithOtp({
    email,
    options: {
      emailRedirectTo: window.location.href,
    },
  });
  if (error) {
    setAuthLoading(false);
    setStatus(`Errore login: ${error.message}`, "error");
    return;
  }
  setAuthLoading(false);
  setStatus("Ti ho inviato un link di accesso via email.", "ok");
}

async function signOut() {
  await supabase.auth.signOut();
}

function clearSelection() {
  state.selected = null;
  selectedCode.textContent = "Nessuna selezionata";
  detailForm.reset();
  repartiList.innerHTML = "";
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

function closeImportModal() {
  importModal.classList.add("hidden");
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
  importPreview.textContent =
    `Righe riconosciute: ${rows.length}/${total}. ` +
    `Esempi: ${sample.join(" - ") || "-"}. ${errorText}`;
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

function formatDateHumanNoYear(date) {
  return date.toLocaleDateString("it-IT", {
    weekday: "long",
    day: "numeric",
    month: "long",
  });
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
  if (view === "three") {
    const prev = addDays(start, -7);
    const next = addDays(start, 7);
    return [
      ...weekdayDates(prev),
      ...weekdayDates(start),
      ...weekdayDates(next),
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
  const s = startOfDay(start);
  const e = startOfDay(end);
  if (s.getTime() === e.getTime()) return 0;
  const step = s < e ? 1 : -1;
  let diff = 0;
  const current = new Date(s);
  while (current.getTime() !== e.getTime()) {
    current.setDate(current.getDate() + step);
    if (!isWeekend(current)) diff += step;
  }
  return diff;
}

function getMatrixRange() {
  const start = startOfWeek(matrixState.date);
  const end = matrixState.view === "six"
    ? endOfDay(addDays(start, 41))
    : matrixState.view === "three"
    ? endOfDay(addDays(start, 20))
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
  TERMODINAMICI: ["Progettazione termodinamica"],
  ELETTRICI: ["Progettazione elettrica", "Ordine KIT cavi"],
};

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

function daysBetweenInclusive(start, end) {
  const s = startOfDay(start);
  const e = startOfDay(end);
  const diff = e.getTime() - s.getTime();
  return Math.max(1, Math.floor(diff / 86400000) + 1);
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
    setStatus(`Errore verifica assenze: ${error.message}`, "error");
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
    setStatus(`Errore shift: ${error.message}`, "error");
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
      setStatus("Impossibile trovare uno spazio libero per lo shift.", "error");
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
    setStatus(`Errore shift: ${firstError.error.message}`, "error");
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
    setStatus(`Errore shift: ${error.message}`, "error");
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
    setStatus(`Errore shift: ${firstError.error.message}`, "error");
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
    setStatus(`Errore verifica overlap: ${error.message}`, "error");
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
    setStatus(`Errore dipendenze: ${error.message}`, "error");
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
      setStatus(`Errore dipendenze: ${updError.message}`, "error");
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
  const d = dateFromKey(dayKey);
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
      setStatus(`Errore assenza: ${error.message}`, "error");
      return;
    }
    setStatus("Assenza creata.", "ok");
    openActivityModal(data);
    await loadMatrixAttivita();
    return;
  }
  const commessaId = e.dataTransfer.getData("application/x-commessa-id");
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
            setStatus("Questa attivita non e disponibile per questo reparto.", "error");
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
      setStatus(`Errore copia: ${error.message}`, "error");
      return;
    }
    setStatus("Attivita copiata.", "ok");
        } else {
          if (!isActivityAllowedForRisorsa(item.titolo, r.id)) {
            setStatus("Questa attivita non e disponibile per questo reparto.", "error");
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
    const { error } = await supabase
      .from("attivita")
      .update({
        risorsa_id: r.id,
        data_inizio: newStart.toISOString(),
        data_fine: newEnd.toISOString(),
      })
      .eq("id", item.id);
    if (error) {
      setStatus(`Errore spostamento: ${error.message}`, "error");
      return;
    }
          if (isPhaseKey(item.titolo) && deltaDays !== 0 && dependentToShift && dependentToShift.length) {
            const ok = await shiftDependentPhases(item, deltaDays, dependentToShift);
            if (!ok) return;
          }
    setStatus("Attivita spostata.", "ok");
  }
  await loadMatrixAttivita();
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

function compareCommesse(a, b) {
  const pa = parseCommessaCode(a.codice);
  const pb = parseCommessaCode(b.codice);
  if (pa.year != null && pb.year != null) {
    if (pa.year !== pb.year) return pa.year - pb.year;
    if (pa.num != null && pb.num != null && pa.num !== pb.num) return pa.num - pb.num;
  }
  return (a.codice || "").localeCompare(b.codice || "");
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
  list.slice().sort(compareCommesse).forEach((c) => {
    const item = document.createElement("div");
    item.className = `item${state.selected && state.selected.id === c.id ? " selected" : ""}`;
    item.dataset.id = c.id;
    item.innerHTML = `
      <div class="item-title">${c.codice} - ${c.titolo || "Senza titolo"}</div>
      <div class="item-meta">
        <span>${c.cliente || "Cliente n/d"}</span>
        <span>Stato: ${c.stato}</span>
        <span>Ingresso: ${formatDate(c.data_ingresso) || "-"}</span>
      </div>
    `;
    item.addEventListener("click", () => selectCommessa(c.id));
    commesseList.appendChild(item);
  });
}

function renderReparti(commessa) {
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
        setStatus(`Errore reparto: ${error.message}`, "error");
        return;
      }
  setStatus("Reparto aggiornato.", "ok");
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
        setStatus("Inserisci il nome della risorsa.", "error");
        return;
      }
      const repartoId = reparto.value || null;
      const { error } = await supabase
        .from("risorse")
        .update({ nome, reparto_id: repartoId, attiva: activeInput.checked })
        .eq("id", r.id);
      if (error) {
        setStatus(`Errore risorsa: ${error.message}`, "error");
        return;
      }
      setStatus("Risorsa aggiornata.", "ok");
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
        setStatus(`Errore eliminazione: ${error.message}`, "error");
        return;
      }
      setStatus("Risorsa eliminata.", "ok");
      await loadRisorse();
    });

    row.appendChild(name);
    row.appendChild(reparto);
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
  const q = searchInput.value.trim().toLowerCase();
  const year = yearFilter.value;
  const st = statusFilter.value;
  const filtered = state.commesse.filter((c) => {
    if (st && c.stato !== st) return false;
    if (year && String(c.anno || "") !== year) return false;
    if (!q) return true;
    const hay = `${c.codice} ${c.titolo || ""} ${c.cliente || ""}`.toLowerCase();
    return hay.includes(q);
  });
  renderCommesse(filtered);
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

function selectCommessa(id) {
  const commessa = state.commesse.find((c) => c.id === id);
  if (!commessa) return;
  state.selected = commessa;
  selectedCode.textContent = commessa.codice;

  d.codice.value = commessa.codice || "";
  d.titolo.value = commessa.titolo || "";
  d.cliente.value = commessa.cliente || "";
  d.stato.value = commessa.stato || "nuova";
  d.priorita.value = commessa.priorita || "";
  d.data_ingresso.value = formatDate(commessa.data_ingresso);
  d.data_consegna.value = formatDate(commessa.data_consegna_prevista);
  d.note.value = commessa.note_generali || "";

  renderReparti(commessa);
  renderCommesse(state.commesse);
}

async function loadProfile() {
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return;

  const { data, error } = await supabase.from("utenti").select("*").eq("id", user.id).single();
  if (error) {
    setStatus("Accesso negato: non sei in whitelist o profilo non creato.", "error");
    return;
  }
  state.profile = data;
  setRoleBadge(data.ruolo);
  setWriteAccess(data.ruolo !== "viewer");
  authStatus.textContent = `Connesso come ${data.email}`;
  logoutBtn.classList.remove("hidden");
}

async function loadReparti() {
  const { data, error } = await supabase.from("reparti").select("*").order("nome");
  if (error) {
    setStatus(`Errore reparti: ${error.message}`, "error");
    return;
  }
  state.reparti = data || [];
  renderRepartiChecks();
  renderAddRepartoSelect();
  renderResourceRepartoSelect(resourceReparto);
  renderResourcesPanel();
}

async function loadRisorse() {
  const { data, error } = await supabase.from("risorse").select("*").order("nome");
  if (error) {
    setStatus(`Errore risorse: ${error.message}`, "error");
    return;
  }
  state.risorse = data || [];
  matrixState.risorse = (data || []).filter((r) => r.attiva);
  renderMatrix();
  renderMatrixReport();
  renderResourcesPanel();
}

async function loadCommesse() {
  const { data, error } = await supabase.from("v_commesse").select("*").order("created_at", {
    ascending: false,
  });
  if (error) {
    setStatus(`Errore commesse: ${error.message}`, "error");
    return;
  }
  state.commesse = data || [];
  renderYearFilter();
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
  empty.className = "matrix-cell matrix-header";
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
    const hay = `${c.codice} ${c.titolo || ""}`.toLowerCase();
    return hay.includes(q);
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
        ${isCollapsed ? "▸" : "▾"}
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
        const empty = document.createElement("div");
        empty.className = "matrix-empty";
        empty.textContent = "+";
        cell.appendChild(empty);
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
          setStatus("Seleziona almeno una attivita.", "error");
          return;
        }
        const hasAssente = attivitaList.some((t) => isAssenteTitle(t));
        if (hasAssente && attivitaList.length > 1) {
          setStatus("ASSENTE va assegnata da sola.", "error");
          return;
        }
        if (!hasAssente && commesseSel.length === 0) {
          setStatus("Seleziona una commessa.", "error");
          return;
        }
        if (!hasAssente && commesseSel.length > 1) {
          setStatus("Seleziona una sola commessa per assegnare.", "error");
          return;
        }
        const notAllowed = attivitaList.filter((t) => !isActivityAllowedForRisorsa(t, r.id));
        if (notAllowed.length) {
          setStatus("Alcune attivita non sono disponibili per questo reparto.", "error");
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
          setStatus(`Errore assegnazione: ${error.message}`, "error");
          return;
        }
        setStatus(rows.length > 1 ? "Attivita assegnate." : "Attivita assegnata.", "ok");
        await loadMatrixAttivita();
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
      if (matrixState.colorMode !== "none") {
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
        if (matrixState.resizing) {
          e.preventDefault();
          return;
        }
        cancelCommessaHighlightTimer();
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
        commessaHighlightTimer = setTimeout(() => {
          commessaHighlightTimer = null;
          if (matrixState.draggingId) return;
          if (!b.a.commessa_id) return;
          applyCommessaHighlight(String(b.a.commessa_id));
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
        clearCommessaHighlight();
        openActivityModal(b.a);
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
  const { start, end } = getMatrixRange();
  matrixReportRange.textContent = `Periodo: ${formatDateLocal(start)} → ${formatDateLocal(end)}`;

  const activities = (matrixState.attivita || []).filter((a) => a.commessa_id && overlapsRange(a, start, end));
  if (!activities.length) {
    matrixReportGrid.innerHTML = `<div class="matrix-empty">Nessuna attivita nel periodo.</div>`;
    return;
  }

  const commessaIds = Array.from(new Set(activities.map((a) => a.commessa_id)));
  let commesse = state.commesse.filter((c) => commessaIds.includes(c.id));
  if (matrixState.reportSort === "asc") {
    commesse = commesse.slice().sort(compareCommesse);
  }
  const attivitaOrder = [
    "Preliminare",
    "3D",
    "Telaio",
    "Carenatura",
    "Costruttivi 1-5",
    "Consegna totale",
    "Altro",
  ];
  const attivitaTypes = attivitaOrder.filter((t) => getMatrixAttivitaOptions().includes(t));

  matrixReportGrid.innerHTML = "";
  matrixReportGrid.style.gridTemplateColumns = `260px repeat(${attivitaTypes.length}, minmax(140px, 1fr))`;

  const header = document.createElement("div");
  header.className = "report-cell report-header";
  header.textContent = "Commessa";
  matrixReportGrid.appendChild(header);
  attivitaTypes.forEach((t) => {
    const h = document.createElement("div");
    h.className = "report-cell report-header";
    h.textContent = t;
    matrixReportGrid.appendChild(h);
  });

  const risorseById = new Map((state.risorse || []).map((r) => [r.id, r.nome]));

  commesse.forEach((c) => {
    const commessaCell = document.createElement("div");
    commessaCell.className = "report-cell report-commessa";
    commessaCell.innerHTML = `<div class="item-title">${c.codice}</div><div class="item-meta">${c.titolo || ""}</div>`;
    matrixReportGrid.appendChild(commessaCell);

    attivitaTypes.forEach((type) => {
      const cell = document.createElement("div");
      cell.className = "report-cell";
      const matches = activities.filter((a) => a.commessa_id === c.id && (a.titolo || "") === type);
      if (!matches.length) {
        cell.classList.add("report-empty");
        cell.textContent = "—";
      } else {
        const resSet = new Map();
        matches.forEach((a) => {
          if (a.risorsa_id == null) return;
          const name = risorseById.get(a.risorsa_id) || `Risorsa ${a.risorsa_id}`;
          const dateLabel = `${formatDateLocal(new Date(a.data_inizio))} → ${formatDateLocal(new Date(a.data_fine))}`;
          resSet.set(`${name} (${dateLabel})`, name);
        });
        const badges = document.createElement("div");
        badges.className = "report-badges";
        Array.from(resSet.entries()).forEach(([label, name]) => {
          const badge = document.createElement("div");
          badge.className = "report-badge";
          badge.textContent = label;
          badge.title = name;
          badges.appendChild(badge);
        });
        cell.appendChild(badges);
      }
      matrixReportGrid.appendChild(cell);
    });
  });
}

async function loadMatrixAttivita() {
  if (!state.session) return;
  const start = startOfWeek(matrixState.date);
  const end = matrixState.view === "six"
    ? endOfDay(addDays(start, 41))
    : matrixState.view === "three"
    ? endOfDay(addDays(start, 20))
    : endOfDay(addDays(start, 6));
  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .lte("data_inizio", end.toISOString())
    .gte("data_fine", start.toISOString())
    .order("data_inizio", { ascending: true });
  if (error) {
    setStatus(`Errore matrice: ${error.message}`, "error");
    return;
  }
  matrixState.attivita = data || [];
  renderMatrix();
  renderMatrixReport();
}

async function loadAttivita() {
  if (!state.session) return;
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

  const { data, error } = await supabase
    .from("attivita")
    .select("*")
    .lte("data_inizio", end.toISOString())
    .gte("data_fine", start.toISOString())
    .order("data_inizio", { ascending: true });

  if (error) {
    setStatus(`Errore attivita: ${error.message}`, "error");
    return;
  }
  calendarState.attivita = data || [];
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

async function setupRealtime() {
  if (state.realtime) return;
  state.realtime = supabase
    .channel("realtime-commesse")
    .on("postgres_changes", { event: "*", schema: "public", table: "commesse" }, loadCommesse)
    .on("postgres_changes", { event: "*", schema: "public", table: "commesse_reparti" }, loadCommesse)
    .on("postgres_changes", { event: "*", schema: "public", table: "attivita" }, loadAttivita)
    .on("postgres_changes", { event: "*", schema: "public", table: "attivita" }, loadMatrixAttivita)
    .subscribe();
}

async function handleUpdate(e) {
  e.preventDefault();
  if (!state.selected) {
    setStatus("Seleziona una commessa.", "error");
    return;
  }
  const payload = {
    codice: d.codice.value.trim(),
    titolo: d.titolo.value.trim() || null,
    cliente: d.cliente.value.trim() || null,
    stato: d.stato.value,
    priorita: d.priorita.value || null,
    data_ingresso: d.data_ingresso.value || null,
    data_consegna_prevista: d.data_consegna.value || null,
    note_generali: d.note.value.trim() || null,
  };
  const { error } = await supabase.from("commesse").update(payload).eq("id", state.selected.id);
  if (error) {
    setStatus(`Errore update: ${error.message}`, "error");
    return;
  }
  setStatus("Commessa aggiornata.", "ok");
  await loadCommesse();
}

async function handleCreate(e) {
  e.preventDefault();
  const payload = {
    codice: n.codice.value.trim(),
    titolo: n.titolo.value.trim() || null,
    cliente: n.cliente.value.trim() || null,
    stato: n.stato.value,
    priorita: n.priorita.value || null,
    data_ingresso: n.data_ingresso.value || null,
    data_consegna_prevista: n.data_consegna.value || null,
    note_generali: n.note.value.trim() || null,
  };
  const { data, error } = await supabase.from("commesse").insert(payload).select().single();
  if (error) {
    setStatus(`Errore create: ${error.message}`, "error");
    return;
  }
  const selectedReparti = Array.from(newRepartiChecks.querySelectorAll("input:checked")).map(
    (i) => i.value
  );
  if (selectedReparti.length) {
    const rows = selectedReparti.map((id) => ({
      commessa_id: data.id,
      reparto_id: Number(id),
      stato: "da_fare",
    }));
    const { error: repErr } = await supabase.from("commesse_reparti").insert(rows);
    if (repErr) {
      setStatus(`Commessa creata, errore reparti: ${repErr.message}`, "error");
    }
  }
  newForm.reset();
  setStatus("Commessa creata.", "ok");
  await loadCommesse();
  selectCommessa(data.id);
}

async function handleAddReparto(e) {
  e.preventDefault();
  if (!state.selected) {
    setStatus("Seleziona una commessa.", "error");
    return;
  }
  const repartoId = Number(addRepartoSelect.value);
  const { error } = await supabase.from("commesse_reparti").insert({
    commessa_id: state.selected.id,
    reparto_id: repartoId,
    stato: "da_fare",
  });
  if (error) {
    setStatus(`Errore aggiunta reparto: ${error.message}`, "error");
    return;
  }
  setStatus("Reparto aggiunto.", "ok");
  await loadCommesse();
  selectCommessa(state.selected.id);
}

async function handleImport() {
  if (!state.canWrite) return;
  const { rows, errors } = parseImportRows(importTextarea.value);
  if (!rows.length) {
    setStatus("Nessuna riga valida da importare.", "error");
    return;
  }
  if (errors.length) {
    setStatus(`Import bloccato: correggi gli errori (es: ${errors[0]}).`, "error");
    return;
  }

  const { error } = await supabase
    .from("commesse")
    .upsert(rows, { onConflict: "codice" })
    .select("codice");

  if (error) {
    setStatus(`Errore import: ${error.message}`, "error");
    importPreview.textContent = `Errore import: ${error.message}`;
    console.error("Import error", error);
    return;
  }
  setStatus(`Import completato: ${rows.length} righe.`, "ok");
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
      setStatus("Errore: lista attivita non disponibile.", "error");
      return;
    }
    const selected = Array.from(assignActivities.querySelectorAll(".filter-pill.active")).map((i) => i.textContent.trim());
    const altroNote = assignAltroNote ? assignAltroNote.value.trim() : "";
    const assenzaOreRaw = assignAssenzaOre ? assignAssenzaOre.value : "";
    const assenzaOre = Number(assenzaOreRaw);
    if (selected.some((t) => isAssenteTitle(t)) && (!assenzaOreRaw || Number.isNaN(assenzaOre) || assenzaOre <= 0)) {
      setStatus("Inserisci le ore di assenza.", "error");
      return;
    }
    if (!selected.length) {
      setStatus("Seleziona almeno una attivita.", "error");
      return;
    }
    if (assignConfirmBtn) {
      assignConfirmBtn.disabled = true;
      assignConfirmBtn.textContent = "Assegno...";
    }
  const { commessaId, risorsaId, day } = matrixState.pendingDrop;
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
    setStatus(`Errore assegnazione: ${error.message}`, "error");
    return;
  }
  setStatus("Attivita assegnate.", "ok");
  closeAssignModal();
  await loadMatrixAttivita();
  } catch (err) {
    setStatus(`Errore assegnazione: ${err.message || err}`, "error");
  } finally {
    if (assignConfirmBtn) {
      assignConfirmBtn.disabled = false;
      assignConfirmBtn.textContent = "Assegna";
    }
  }
}

async function init() {
  if (SUPABASE_URL.startsWith("INSERISCI")) {
    setStatus("Configura SUPABASE_URL e SUPABASE_ANON_KEY in web/app.js.");
  }

  closeImportModal();

  const { data } = await supabase.auth.getSession();
  state.session = data.session;

  if (state.session) {
    await loadProfile();
    await loadReparti();
    await loadRisorse();
    await loadCommesse();
    calendarDate.value = formatDateInput(calendarState.date);
    await loadAttivita();
    matrixDate.value = formatDateInput(matrixState.date);
    await loadMatrixAttivita();
    await setupRealtime();
  } else {
    setStatus("Effettua il login con email whitelisted.");
  }
}

loginBtn.addEventListener("click", signIn);
logoutBtn.addEventListener("click", signOut);
searchInput.addEventListener("input", applyFilters);
yearFilter.addEventListener("change", applyFilters);
statusFilter.addEventListener("change", applyFilters);
detailForm.addEventListener("submit", handleUpdate);
newForm.addEventListener("submit", handleCreate);
addRepartoForm.addEventListener("submit", handleAddReparto);
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
if (closeProgressBtn) {
  closeProgressBtn.addEventListener("click", closeProgressModal);
}
if (progressModal) {
  progressModal.addEventListener("click", (e) => {
    if (e.target === progressModal) closeProgressModal();
  });
}
if (progressYearCurrent) {
  progressYearCurrent.addEventListener("click", () => {
    progressYearCurrent.classList.add("active");
    if (progressYearPrev) progressYearPrev.classList.remove("active");
    progressSelectedId = null;
    if (progressMeta) progressMeta.textContent = "Seleziona una commessa.";
    if (progressList) progressList.innerHTML = "";
    renderProgressResults();
  });
}
if (progressYearPrev) {
  progressYearPrev.addEventListener("click", () => {
    progressYearPrev.classList.add("active");
    if (progressYearCurrent) progressYearCurrent.classList.remove("active");
    progressSelectedId = null;
    if (progressMeta) progressMeta.textContent = "Seleziona una commessa.";
    if (progressList) progressList.innerHTML = "";
    renderProgressResults();
  });
}
  if (progressSearch) {
    progressSearch.addEventListener("input", () => {
      progressSelectedId = null;
      if (progressMeta) progressMeta.textContent = "Seleziona una commessa.";
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
      setStatus("Inserisci il nome della risorsa.", "error");
      return;
    }
    const repartoId = resourceReparto.value;
    if (!repartoId) {
      setStatus("Seleziona un reparto.", "error");
      return;
    }
    const payload = {
      nome,
      reparto_id: repartoId,
      attiva: resourceActive.checked,
    };
    const { error } = await supabase.from("risorse").insert(payload);
    if (error) {
      setStatus(`Errore creazione risorsa: ${error.message}`, "error");
      return;
    }
    resourceName.value = "";
    resourceReparto.value = "";
    resourceActive.checked = true;
    setStatus("Risorsa aggiunta.", "ok");
    await loadRisorse();
  });
}
clearSelectionBtn.addEventListener("click", clearSelection);
openImportBtn.addEventListener("click", openImportModal);
closeImportBtn.addEventListener("click", closeImportModal);
importTextarea.addEventListener("input", updateImportPreview);
importCommitBtn.addEventListener("click", handleImport);
importModal.addEventListener("click", (e) => {
  if (e.target === importModal) closeImportModal();
});
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeImportModal();
    closeAssignModal();
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
  const step = matrixState.view === "six" ? 42 : matrixState.view === "three" ? 21 : 7;
  matrixState.date = addDays(matrixState.date, -step);
  matrixDate.value = formatDateInput(matrixState.date);
  await loadMatrixAttivita();
});
matrixNextBtn.addEventListener("click", async () => {
  const step = matrixState.view === "six" ? 42 : matrixState.view === "three" ? 21 : 7;
  matrixState.date = addDays(matrixState.date, step);
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
  } else {
    matrixState.view = "week";
  }
  setMatrixViewLabel();
  await loadMatrixAttivita();
});
matrixZoomOutBtn.addEventListener("click", async () => {
  if (matrixState.view === "week") {
    matrixState.view = "three";
  } else if (matrixState.view === "three") {
    matrixState.view = "six";
  }
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
if (activityCloseBtn) {
  activityCloseBtn.addEventListener("click", closeActivityModal);
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
});
document.addEventListener("drop", () => {
  if (matrixGrid) matrixGrid.classList.remove("matrix-dragging");
  matrixState.draggingId = null;
  document.querySelectorAll(".matrix-activity-bar.dragging").forEach((el) => el.classList.remove("dragging"));
});
document.addEventListener("click", (e) => {
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

supabase.auth.onAuthStateChange(async (_event, session) => {
  state.session = session;
  if (session) {
    await loadProfile();
    await loadReparti();
    await loadRisorse();
    await loadCommesse();
    calendarDate.value = formatDateInput(calendarState.date);
    await loadAttivita();
    matrixDate.value = formatDateInput(matrixState.date);
    await loadMatrixAttivita();
    await setupRealtime();
  } else {
    state.profile = null;
    state.commesse = [];
    authStatus.textContent = "Non autenticato";
    logoutBtn.classList.add("hidden");
    setRoleBadge("");
    setWriteAccess(false);
    commesseList.innerHTML = "";
    clearSelection();
    calendarGrid.innerHTML = "";
    matrixGrid.innerHTML = "";
    setStatus("Sessione terminata.");
  }
});

init();

})();

