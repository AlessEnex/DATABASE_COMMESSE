(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

  const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: { persistSession: true, autoRefreshToken: true },
  });

  const pasteInput = document.getElementById("reportsPasteInput");
  const importBtn = document.getElementById("reportsImportBtn");
  const refreshBtn = document.getElementById("reportsRefreshBtn");
  const yearFilter = document.getElementById("reportsYearFilter");
  const importStatus = document.getElementById("reportsImportStatus");
  const weekList = document.getElementById("reportsWeekList");
  const weekDetail = document.getElementById("reportsWeekDetail");
  const chartCanvas = document.getElementById("reportsChart");
  const targetInput = document.getElementById("reportsTargetInput");

  const kpiAvgWeek = document.getElementById("kpiAvgWeek");
  const kpiMachines = document.getElementById("kpiMachines");
  const kpiAvgMachine = document.getElementById("kpiAvgMachine");
  const kpiTotalEuro = document.getElementById("kpiTotalEuro");

  const gate = document.getElementById("reportsGate");
  const gateTitle = document.getElementById("reportsGateTitle");
  const gateText = document.getElementById("reportsGateText");

  const TARGET_STORAGE_KEY = "reports_weekly_target_eur";
  const DEFAULT_WEEKLY_TARGET_EUR = 1_000_000;

  let chart = null;
  let currentProfile = null;
  let rawWeeklyItems = [];
  let lastWeekDetails = new Map();

  const moneyFmt = new Intl.NumberFormat("it-IT", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  const numberFmt = new Intl.NumberFormat("it-IT", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  });

  function setImportStatus(message, tone = "") {
    if (!importStatus) return;
    importStatus.textContent = message;
    importStatus.classList.remove("ok", "error");
    if (tone) importStatus.classList.add(tone);
  }

  function showGate(title, text) {
    if (gateTitle) gateTitle.textContent = title;
    if (gateText) gateText.textContent = text;
    if (gate) gate.classList.remove("hidden");
  }

  function hideGate() {
    if (gate) gate.classList.add("hidden");
  }

  function normalizeCode(value) {
    return String(value || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "");
  }

  function parseMoney(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const raw = String(value || "").trim();
    if (!raw) return null;
    let text = raw.replace(/\s+/g, "").replace(/EUR|euro|\u20AC/gi, "");
    const hasComma = text.includes(",");
    const hasDot = text.includes(".");

    if (hasComma && hasDot) {
      const lastComma = text.lastIndexOf(",");
      const lastDot = text.lastIndexOf(".");
      if (lastComma > lastDot) {
        text = text.replace(/\./g, "").replace(/,/g, ".");
      } else {
        text = text.replace(/,/g, "");
      }
    } else if (hasComma && !hasDot) {
      if (/^-?\d{1,3}(,\d{3})+$/.test(text)) text = text.replace(/,/g, "");
      else text = text.replace(/,/g, ".");
    } else if (!hasComma && hasDot) {
      if (/^-?\d{1,3}(\.\d{3})+$/.test(text)) text = text.replace(/\./g, "");
    }

    text = text.replace(/'/g, "");
    const parsed = Number.parseFloat(text);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function parseIsoDateOnly(value) {
    const text = String(value || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
    const d = new Date(`${text}T00:00:00`);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function getIsoWeekInfo(dateInput) {
    const d = new Date(Date.UTC(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const week = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
    const year = d.getUTCFullYear();
    const weekPadded = String(week).padStart(2, "0");
    return {
      key: `${year}-W${weekPadded}`,
      label: `${year} - W${weekPadded}`,
      year,
      sortValue: year * 100 + week,
    };
  }

  function getWeeklyTarget() {
    const stored = parseMoney(localStorage.getItem(TARGET_STORAGE_KEY));
    if (stored != null && stored >= 0) return stored;
    return DEFAULT_WEEKLY_TARGET_EUR;
  }

  function setWeeklyTarget(value) {
    if (!Number.isFinite(value) || value < 0) return;
    localStorage.setItem(TARGET_STORAGE_KEY, String(value));
  }

  function renderTargetInputValue() {
    if (!targetInput) return;
    targetInput.value = String(Math.round(getWeeklyTarget()));
  }

  function splitPastedLine(line) {
    if (line.includes("\t")) return line.split("\t");
    if (line.includes(";")) return line.split(";");
    if (line.includes(",")) return line.split(",");
    return [line];
  }

  function parsePastedRows(rawText) {
    const lines = String(rawText || "")
      .replace(/\r/g, "")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);
    if (!lines.length) throw new Error("Nessuna riga incollata.");

    let startIndex = 0;
    const first = lines[0].toLowerCase();
    if (first.includes("codice") && first.includes("imponibile")) startIndex = 1;

    const byCode = new Map();
    let invalidRows = 0;
    for (let i = startIndex; i < lines.length; i += 1) {
      const cols = splitPastedLine(lines[i]).map((c) => String(c || "").trim());
      if (cols.length < 2) {
        invalidRows += 1;
        continue;
      }
      const code = normalizeCode(cols[0]);
      if (!code) continue;
      const imponibile = parseMoney(cols[1]);
      if (imponibile == null) {
        invalidRows += 1;
        continue;
      }
      byCode.set(code, imponibile);
    }

    return {
      entries: Array.from(byCode.entries()).map(([code, imponibile]) => ({ code, imponibile })),
      invalidRows,
    };
  }

  async function ensureAdminAccess() {
    showGate("Verifica accesso", "Controllo sessione e ruolo...");
    const { data: sessionData } = await supabase.auth.getSession();
    const session = sessionData?.session || null;
    if (!session) {
      window.location.replace("./");
      return false;
    }
    const { data: userData } = await supabase.auth.getUser();
    const user = userData?.user || session.user;
    if (!user) {
      window.location.replace("./");
      return false;
    }
    const { data: profile, error } = await supabase
      .from("utenti")
      .select("id, ruolo, email")
      .eq("id", user.id)
      .single();
    if (error || !profile) {
      window.location.replace("./");
      return false;
    }
    if (String(profile.ruolo || "").trim().toLowerCase() !== "admin") {
      window.location.replace("./");
      return false;
    }
    currentProfile = profile;
    hideGate();
    return true;
  }

  async function fetchExistingCommesseMap() {
    const { data, error } = await supabase.from("commesse").select("id, codice");
    if (error) throw error;
    const map = new Map();
    (data || []).forEach((row) => {
      const key = normalizeCode(row.codice);
      if (!key) return;
      map.set(key, row.id);
    });
    return map;
  }

  async function importImponibiliFromPaste(rawText) {
    if (!String(rawText || "").trim()) {
      setImportStatus("Incolla prima i dati da Excel.", "error");
      return;
    }
    importBtn.disabled = true;
    setImportStatus("Parsing dati incollati in corso...");
    try {
      const parsed = parsePastedRows(rawText);
      if (!parsed.entries.length) {
        setImportStatus("Nessuna riga valida da importare.", "error");
        return;
      }

      const commesseMap = await fetchExistingCommesseMap();
      const payloadByCommessa = new Map();
      let skippedMissingCode = 0;

      parsed.entries.forEach((entry) => {
        const commessaId = commesseMap.get(entry.code);
        if (!commessaId) {
          skippedMissingCode += 1;
          return;
        }
        payloadByCommessa.set(commessaId, {
          commessa_id: commessaId,
          imponibile: entry.imponibile,
          updated_by: currentProfile?.id || null,
          updated_at: new Date().toISOString(),
        });
      });

      const payload = Array.from(payloadByCommessa.values());
      if (!payload.length) {
        setImportStatus("Nessun codice macchina esistente trovato nel file.", "error");
        return;
      }

      const chunkSize = 500;
      for (let i = 0; i < payload.length; i += chunkSize) {
        const chunk = payload.slice(i, i + chunkSize);
        const { error } = await supabase.from("commessa_imponibili").upsert(chunk, { onConflict: "commessa_id" });
        if (error) throw error;
      }

      const msg =
        `Import completato: ${payload.length} codici aggiornati` +
        `, ${skippedMissingCode} codici non trovati` +
        (parsed.invalidRows ? `, ${parsed.invalidRows} righe con imponibile non valido` : "") +
        ".";
      setImportStatus(msg, "ok");
      await loadWeeklyChart();
    } catch (err) {
      setImportStatus(`Errore import: ${err?.message || err}`, "error");
    } finally {
      importBtn.disabled = false;
    }
  }

  function populateYearFilter() {
    if (!yearFilter) return;
    const years = new Set([2026]);
    rawWeeklyItems.forEach((item) => {
      if (Number.isFinite(item?.year)) years.add(item.year);
    });
    const options = Array.from(years).sort((a, b) => b - a);
    const current = yearFilter.value || "2026";
    yearFilter.innerHTML = "";

    const allOpt = document.createElement("option");
    allOpt.value = "all";
    allOpt.textContent = "Tutti";
    yearFilter.appendChild(allOpt);

    options.forEach((year) => {
      const opt = document.createElement("option");
      opt.value = String(year);
      opt.textContent = String(year);
      yearFilter.appendChild(opt);
    });

    if (Array.from(yearFilter.options).some((o) => o.value === current)) yearFilter.value = current;
    else if (Array.from(yearFilter.options).some((o) => o.value === "2026")) yearFilter.value = "2026";
    else yearFilter.value = "all";
  }

  function renderWeekList(weekRows) {
    if (!weekList) return;
    weekList.innerHTML = "";
    if (!weekRows.length) {
      const li = document.createElement("li");
      li.className = "reports-week-item";
      li.textContent = "Nessun dato disponibile per il grafico.";
      weekList.appendChild(li);
      if (weekDetail) {
        weekDetail.innerHTML = `<div class="reports-week-detail-title">Nessuna settimana disponibile.</div>`;
      }
      return;
    }

    weekRows.forEach((row) => {
      const li = document.createElement("li");
      li.className = "reports-week-item";
      li.dataset.weekKey = row.key;
      const totalWeek = Number(row.orderedTotal || 0) + Number(row.plannedTotal || 0);
      li.innerHTML = `<span>${row.label}</span><strong>${moneyFmt.format(totalWeek)}</strong>`;
      weekList.appendChild(li);
    });

    renderWeekDetail(weekRows[0].key);
    const first = weekList.querySelector(".reports-week-item");
    if (first) first.classList.add("active");
  }

  function renderWeekDetail(weekKey) {
    if (!weekDetail) return;
    const bucket = lastWeekDetails.get(String(weekKey || "")) || { ordered: [], planned: [] };
    const orderedRows = (bucket.ordered || []).slice();
    const plannedRows = (bucket.planned || []).slice();

    if (!orderedRows.length && !plannedRows.length) {
      weekDetail.innerHTML = `<div class="reports-week-detail-title">Nessuna commessa trovata per la settimana selezionata.</div>`;
      return;
    }

    orderedRows.sort((a, b) => String(a.codice || "").localeCompare(String(b.codice || ""), "it"));
    plannedRows.sort((a, b) => String(a.codice || "").localeCompare(String(b.codice || ""), "it"));

    weekDetail.innerHTML = "";
    const title = document.createElement("div");
    title.className = "reports-week-detail-title";
    title.textContent = `Commesse incluse (ordinati ${orderedRows.length}, pianificati ${plannedRows.length})`;
    weekDetail.appendChild(title);

    const renderSection = (sectionTitle, rows) => {
      if (!rows.length) return;
      const head = document.createElement("div");
      head.className = "reports-week-detail-title";
      head.textContent = sectionTitle;
      weekDetail.appendChild(head);
      rows.forEach((row) => {
        const line = document.createElement("div");
        line.className = "reports-week-detail-row";
        const left = document.createElement("span");
        left.textContent = row.codice || "(codice n/d)";
        const right = document.createElement("strong");
        right.textContent = moneyFmt.format(Number(row.imponibile || 0));
        line.appendChild(left);
        line.appendChild(right);
        weekDetail.appendChild(line);
      });
    };

    renderSection("Telai ordinati", orderedRows);
    renderSection("Telai pianificati (non ordinati)", plannedRows);
  }

  function renderChart(weekRows) {
    if (!chartCanvas) return;
    const labels = weekRows.map((row) => row.label);
    const orderedValues = weekRows.map((row) => Number(row.orderedTotal || 0));
    const plannedValues = weekRows.map((row) => Number(row.plannedTotal || 0));
    const targetSeries = labels.map(() => getWeeklyTarget());

    if (chart) chart.destroy();
    chart = new window.Chart(chartCanvas, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Telai ordinati",
            data: orderedValues,
            backgroundColor: "rgba(75, 192, 132, 0.55)",
            borderColor: "rgba(46, 160, 95, 1)",
            borderWidth: 1.5,
            borderRadius: 8,
            stack: "totale",
            categoryPercentage: 0.82,
            barPercentage: 0.9,
            maxBarThickness: 44,
          },
          {
            label: "Telai pianificati (non ordinati)",
            data: plannedValues,
            backgroundColor: "rgba(102, 187, 255, 0.55)",
            borderColor: "rgba(51, 153, 255, 1)",
            borderWidth: 1.5,
            borderRadius: 8,
            stack: "totale",
            categoryPercentage: 0.82,
            barPercentage: 0.9,
            maxBarThickness: 44,
          },
          {
            type: "line",
            label: "Target settimanale",
            data: targetSeries,
            borderColor: "rgba(230, 57, 70, 0.95)",
            borderWidth: 2,
            borderDash: [8, 6],
            pointRadius: 0,
            pointHoverRadius: 0,
            tension: 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 4, right: 4, bottom: 0, left: 0 },
        },
        scales: {
          x: {
            stacked: true,
            ticks: {
              maxRotation: 0,
              autoSkipPadding: 10,
            },
          },
          y: {
            beginAtZero: true,
            stacked: true,
            ticks: {
              callback: (value) => moneyFmt.format(Number(value || 0)),
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            labels: {
              boxWidth: 12,
              boxHeight: 12,
              padding: 10,
            },
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${moneyFmt.format(Number(ctx.raw || 0))}`,
            },
          },
        },
      },
    });
  }

  function renderKpis(orderedItems, orderedWeeksCount) {
    const totalEuro = orderedItems.reduce((sum, item) => sum + Number(item.imponibile || 0), 0);
    const machinesSet = new Set(orderedItems.map((item) => String(item.commessaId || "")).filter(Boolean));
    const machinesCount = machinesSet.size;
    const avgKWeek = orderedWeeksCount > 0 ? totalEuro / orderedWeeksCount / 1000 : 0;
    const avgKMachine = machinesCount > 0 ? totalEuro / machinesCount / 1000 : 0;

    if (kpiAvgWeek) kpiAvgWeek.textContent = `${numberFmt.format(avgKWeek)} kEUR`;
    if (kpiMachines) kpiMachines.textContent = String(machinesCount);
    if (kpiAvgMachine) kpiAvgMachine.textContent = `${numberFmt.format(avgKMachine)} kEUR`;
    if (kpiTotalEuro) kpiTotalEuro.textContent = moneyFmt.format(totalEuro);
  }

  function applyYearFilterAndRender() {
    const selectedYear = yearFilter ? yearFilter.value || "2026" : "2026";
    const filteredItems =
      selectedYear === "all"
        ? rawWeeklyItems.slice()
        : rawWeeklyItems.filter((item) => String(item.year) === String(selectedYear));

    const totals = new Map();
    const details = new Map();

    filteredItems.forEach((item) => {
      const agg = totals.get(item.weekKey) || {
        key: item.weekKey,
        label: item.weekLabel,
        sortValue: item.weekSortValue,
        orderedTotal: 0,
        plannedTotal: 0,
      };

      if (item.category === "ordered") agg.orderedTotal += item.imponibile;
      else agg.plannedTotal += item.imponibile;
      totals.set(item.weekKey, agg);

      const bucket = details.get(item.weekKey) || { ordered: [], planned: [] };
      if (item.category === "ordered") bucket.ordered.push({ codice: item.codice, imponibile: item.imponibile });
      else bucket.planned.push({ codice: item.codice, imponibile: item.imponibile });
      details.set(item.weekKey, bucket);
    });

    lastWeekDetails = details;
    const weekRows = Array.from(totals.values()).sort((a, b) => a.sortValue - b.sortValue);
    renderChart(weekRows);
    renderWeekList(weekRows);

    const orderedItems = filteredItems.filter((item) => item.category === "ordered");
    const orderedWeeksCount = new Set(orderedItems.map((item) => item.weekKey)).size;
    renderKpis(orderedItems, orderedWeeksCount);
  }

  async function loadWeeklyChart() {
    if (refreshBtn) refreshBtn.disabled = true;
    try {
      const { data, error } = await supabase
        .from("commessa_imponibili")
        .select("imponibile, commesse!inner(id,codice,telaio_consegnato,data_ordine_telaio,data_consegna_telaio_effettiva)");
      if (error) throw error;

      const items = [];
      (data || []).forEach((row) => {
        const commessa = row.commesse || null;
        const imponibile = Number(row.imponibile);
        if (!Number.isFinite(imponibile)) return;

        const isOrdered = Boolean(commessa?.telaio_consegnato);
        const orderedAt = parseIsoDateOnly(commessa?.data_consegna_telaio_effettiva);
        const plannedAt = parseIsoDateOnly(commessa?.data_ordine_telaio);

        if (isOrdered && orderedAt) {
          const week = getIsoWeekInfo(orderedAt);
          items.push({
            category: "ordered",
            year: week.year,
            commessaId: commessa?.id || "",
            weekKey: week.key,
            weekLabel: week.label,
            weekSortValue: week.sortValue,
            codice: commessa?.codice || "",
            imponibile,
          });
          return;
        }

        if (!isOrdered && plannedAt) {
          const week = getIsoWeekInfo(plannedAt);
          items.push({
            category: "planned",
            year: week.year,
            commessaId: commessa?.id || "",
            weekKey: week.key,
            weekLabel: week.label,
            weekSortValue: week.sortValue,
            codice: commessa?.codice || "",
            imponibile,
          });
        }
      });

      rawWeeklyItems = items;
      populateYearFilter();
      applyYearFilterAndRender();
    } catch (err) {
      setImportStatus(`Errore caricamento grafico: ${err?.message || err}`, "error");
      rawWeeklyItems = [];
      if (yearFilter) {
        yearFilter.innerHTML = `<option value="2026">2026</option><option value="all">Tutti</option>`;
        yearFilter.value = "2026";
      }
      lastWeekDetails = new Map();
      renderChart([]);
      renderWeekList([]);
      renderKpis([], 0);
    } finally {
      if (refreshBtn) refreshBtn.disabled = false;
    }
  }

  if (importBtn) {
    importBtn.addEventListener("click", () => {
      void importImponibiliFromPaste(pasteInput?.value || "");
    });
  }

  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      void loadWeeklyChart();
    });
  }

  if (yearFilter) {
    yearFilter.addEventListener("change", () => {
      applyYearFilterAndRender();
    });
  }

  if (targetInput) {
    targetInput.addEventListener("change", () => {
      const parsed = parseMoney(targetInput.value);
      if (parsed == null || parsed < 0) {
        targetInput.value = String(Math.round(getWeeklyTarget()));
        setImportStatus("Target settimanale non valido.", "error");
        return;
      }
      setWeeklyTarget(parsed);
      targetInput.value = String(Math.round(parsed));
      applyYearFilterAndRender();
    });
  }

  if (weekList) {
    weekList.addEventListener("click", (event) => {
      const item = event.target.closest(".reports-week-item[data-week-key]");
      if (!item) return;
      const weekKey = item.dataset.weekKey || "";
      weekList.querySelectorAll(".reports-week-item.active").forEach((el) => el.classList.remove("active"));
      item.classList.add("active");
      renderWeekDetail(weekKey);
    });
  }

  void (async () => {
    const ok = await ensureAdminAccess();
    if (!ok) return;
    if (yearFilter) {
      yearFilter.innerHTML = `<option value="2026">2026</option><option value="all">Tutti</option>`;
      yearFilter.value = "2026";
    }
    renderTargetInputValue();
    await loadWeeklyChart();
  })();
})();
