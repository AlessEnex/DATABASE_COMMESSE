(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

  const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: { persistSession: true, autoRefreshToken: true },
  });

  const pasteInput = document.getElementById("reportsPasteInput");
  const copyHeaderBtn = document.getElementById("reportsCopyHeaderBtn");
  const headerGuideCode = document.getElementById("reportsHeaderGuideCode");
  const importBtn = document.getElementById("reportsImportBtn");
  const refreshBtn = document.getElementById("reportsRefreshBtn");
  const exportPdfBtn = document.getElementById("reportsExportPdfBtn");
  const yearFilter = document.getElementById("reportsYearFilter");
  const importStatus = document.getElementById("reportsImportStatus");
  const weekList = document.getElementById("reportsWeekList");
  const weekDetail = document.getElementById("reportsWeekDetail");
  const chartCanvas = document.getElementById("reportsChart");
  const ingressChartCanvas = document.getElementById("reportsIngressChart");
  const futureTrendChartCanvas = document.getElementById("reportsFutureTrendChart");
  const backlogChartCanvas = document.getElementById("reportsBacklogChart");
  const targetInput = document.getElementById("reportsTargetInput");
  const chartTitle = document.getElementById("reportsChartTitle");
  const granularityControl = document.getElementById("reportsGranularity");
  const granularityWeekBtn = document.getElementById("reportsGranularityWeek");
  const granularityDayBtn = document.getElementById("reportsGranularityDay");
  const clockDateEl = document.getElementById("reportsClockDate");
  const clockWeekEl = document.getElementById("reportsClockWeek");
  const clockTimeEl = document.getElementById("reportsClockTime");

  const kpiAvgWeek = document.getElementById("kpiAvgWeek");
  const kpiAvgPeriodLabel = document.getElementById("kpiAvgPeriodLabel");
  const kpiMachines = document.getElementById("kpiMachines");
  const kpiAvgMachine = document.getElementById("kpiAvgMachine");
  const kpiTotalEuro = document.getElementById("kpiTotalEuro");

  const gate = document.getElementById("reportsGate");
  const gateTitle = document.getElementById("reportsGateTitle");
  const gateText = document.getElementById("reportsGateText");

  const TARGET_STORAGE_KEY = "reports_weekly_target_eur";
  const GRANULARITY_STORAGE_KEY = "reports_chart_granularity";
  const THEME_STORAGE_KEY = "commesse_theme";
  const DEFAULT_WEEKLY_TARGET_EUR = 1_000_000;

  let chart = null;
  let ingressChart = null;
  let futureTrendChart = null;
  let backlogChart = null;
  let currentProfile = null;
  let rawWeeklyItems = [];
  let rawIngressItems = [];
  let lastWeekDetails = new Map();
  let currentGranularity = localStorage.getItem(GRANULARITY_STORAGE_KEY) === "day" ? "day" : "week";
  let currentWeeklyAverage = 0;
  let currentIngressWeeklyAverage = 0;

  function applyStoredTheme() {
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);
    const prefersDark =
      window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    const theme = savedTheme || (prefersDark ? "dark" : "light");
    document.body.classList.toggle("theme-dark", theme === "dark");
  }

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
  const moneyAxisFmt = new Intl.NumberFormat("it-IT", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
  const numberAxisFmt = new Intl.NumberFormat("it-IT", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
  const clockDateFmt = new Intl.DateTimeFormat("it-IT", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric",
  });
  const clockTimeFmt = new Intl.DateTimeFormat("it-IT", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });

  function setImportStatus(message, tone = "") {
    if (!importStatus) return;
    importStatus.textContent = message;
    importStatus.classList.remove("ok", "error");
    if (tone) importStatus.classList.add(tone);
  }

  async function copyReportsImportHeader() {
    const text = String(headerGuideCode?.textContent || "codice macchina\timponibile").trim();
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      setImportStatus("Header copiato negli appunti.", "ok");
    } catch (_err) {
      setImportStatus("Copia non disponibile. Seleziona e copia l'header manualmente.", "error");
    }
  }

  async function handleExportChartsPdf() {
    if (!exportPdfBtn) return;
    const jsPdfApi = window.jspdf && window.jspdf.jsPDF ? window.jspdf.jsPDF : null;
    if (!jsPdfApi) {
      setImportStatus("Export PDF non disponibile (libreria mancante).", "error");
      return;
    }

    const defaultMainTitle = "Somma imponibili per settimana (solo telaio ordinato)";
    const chartEntries = [
      {
        title: chartTitle ? String(chartTitle.textContent || "").trim() || defaultMainTitle : defaultMainTitle,
        canvas: chartCanvas,
      },
      {
        title: "Imponibile entrato per settimana (data ingresso ordine)",
        canvas: ingressChartCanvas,
      },
      {
        title: "Trend ordinato + pianificato (settimana a ∞)",
        canvas: futureTrendChartCanvas,
      },
      {
        title: "Backlog andamento nel tempo",
        canvas: backlogChartCanvas,
      },
    ].filter((entry) => entry.canvas && entry.canvas.width > 0 && entry.canvas.height > 0);

    if (!chartEntries.length) {
      setImportStatus("Nessun grafico disponibile da esportare.", "error");
      return;
    }

    const originalLabel = exportPdfBtn.textContent;
    exportPdfBtn.disabled = true;
    exportPdfBtn.textContent = "Esporto PDF...";

    try {
      const doc = new jsPdfApi({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
        compress: true,
      });
      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const margin = 12;
      const maxImageW = pageW - margin * 2;
      const topBlockH = 22;
      const maxImageH = pageH - topBlockH - margin;
      const stamp = new Date().toLocaleString("it-IT", { hour12: false });
      const yearLabel = yearFilter ? String(yearFilter.value || "all").toUpperCase() : "ALL";
      const granularityLabel = currentGranularity === "day" ? "Giornaliera" : "Settimanale";
      const metaLine = `Anno: ${yearLabel} | Vista: ${granularityLabel} | Export: ${stamp}`;

      chartEntries.forEach((entry, index) => {
        if (index > 0) doc.addPage("a4", "landscape");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(14);
        doc.text(entry.title, margin, 12);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.text(metaLine, margin, 17);

        const ratio = entry.canvas.width / Math.max(1, entry.canvas.height);
        let drawW = maxImageW;
        let drawH = drawW / ratio;
        if (drawH > maxImageH) {
          drawH = maxImageH;
          drawW = drawH * ratio;
        }
        const x = (pageW - drawW) / 2;
        const y = topBlockH + (maxImageH - drawH) / 2;
        const imageData = entry.canvas.toDataURL("image/png", 1);
        doc.addImage(imageData, "PNG", x, y, drawW, drawH, undefined, "FAST");
      });

      const filename = `reports-grafici-${new Date().toISOString().slice(0, 10)}.pdf`;
      doc.save(filename);
      setImportStatus(`PDF esportato (${chartEntries.length} grafici).`, "ok");
    } catch (err) {
      setImportStatus(`Errore export PDF: ${err?.message || err}`, "error");
    } finally {
      exportPdfBtn.disabled = false;
      exportPdfBtn.textContent = originalLabel;
    }
  }

  function formatMillionEuroShort(value) {
    const numeric = Number(value || 0);
    if (!Number.isFinite(numeric)) return "0.00M€";
    return `${(numeric / 1_000_000).toFixed(2)}M€`;
  }

  function makeBarValueLabelsPlugin(id) {
    return {
      id,
      afterDatasetsDraw(chartInstance) {
        const { ctx } = chartInstance;
        const area = chartInstance.chartArea;
        const isDayView = currentGranularity === "day";
        ctx.save();
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.font = "600 10px 'IBM Plex Sans', sans-serif";
        chartInstance.data.datasets.forEach((dataset, datasetIndex) => {
          const meta = chartInstance.getDatasetMeta(datasetIndex);
          if (!meta || meta.type !== "bar" || meta.hidden) return;
          meta.data.forEach((barEl, idx) => {
            const raw = Number(dataset.data?.[idx] || 0);
            if (!Number.isFinite(raw) || raw <= 0) return;
            const pos = barEl.tooltipPosition();
            const text = formatMillionEuroShort(raw);
            const textWidth = ctx.measureText(text).width;
            const padX = 5;
            const boxH = 14;
            const boxW = textWidth + padX * 2;
            if (isDayView) {
              const props = barEl.getProps(["x", "y", "base"], true);
              const barHeight = Math.abs(Number(props.base || 0) - Number(props.y || 0));
              if (barHeight < 24) return;
              const topY = Number(props.y || pos.y);
              const anchorY = Math.max(area.top + boxW / 2, topY - 28);
              if (anchorY < area.top || anchorY > area.bottom) return;
              ctx.save();
              ctx.translate(Number(props.x || pos.x), anchorY);
              ctx.rotate(-Math.PI / 2);
              ctx.fillStyle = "rgba(255, 255, 255, 0.72)";
              ctx.beginPath();
              ctx.roundRect(-boxW / 2, -boxH / 2, boxW, boxH, 7);
              ctx.fill();
              ctx.fillStyle = "#475569";
              ctx.fillText(text, 0, 0);
              ctx.restore();
              return;
            }
            const x = pos.x - boxW / 2;
            const y = pos.y - boxH - 2;
            ctx.fillStyle = "rgba(255, 255, 255, 0.72)";
            ctx.beginPath();
            ctx.roundRect(x, y, boxW, boxH, 7);
            ctx.fill();
            ctx.fillStyle = "#475569";
            ctx.fillText(text, pos.x, y + boxH / 2);
          });
        });
        ctx.restore();
      },
    };
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
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!match) return null;
    const year = Number.parseInt(match[1], 10);
    const month = Number.parseInt(match[2], 10);
    const day = Number.parseInt(match[3], 10);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
    const d = new Date(year, month - 1, day, 0, 0, 0, 0);
    if (
      Number.isNaN(d.getTime()) ||
      d.getFullYear() !== year ||
      d.getMonth() !== month - 1 ||
      d.getDate() !== day
    ) {
      return null;
    }
    return d;
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

  function refreshClock() {
    if (!clockDateEl || !clockWeekEl || !clockTimeEl) return;
    const now = new Date();
    const week = getIsoWeekInfo(now);
    clockDateEl.textContent = clockDateFmt.format(now);
    clockWeekEl.textContent = `W${String(week?.key || "").match(/W(\d{1,2})$/i)?.[1] || "--"}`;
    clockTimeEl.textContent = clockTimeFmt.format(now);
  }

  function getDayInfo(dateInput) {
    const d = new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
    if (Number.isNaN(d.getTime())) return null;
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return {
      key: `${year}-${month}-${day}`,
      label: `${day}/${month}`,
      year,
      sortValue: year * 10000 + Number(month) * 100 + Number(day),
    };
  }

  function updateChartTitle() {
    if (!chartTitle) return;
    chartTitle.textContent =
      currentGranularity === "day"
        ? "Somma imponibili per giorno (solo telaio ordinato)"
        : "Somma imponibili per settimana (solo telaio ordinato)";
    if (kpiAvgPeriodLabel) {
      kpiAvgPeriodLabel.textContent =
        currentGranularity === "day" ? "Media kEUR / giorno" : "Media kEUR / settimana";
    }
  }

  function updateGranularityUi() {
    if (!granularityControl) return;
    const isDay = currentGranularity === "day";
    if (granularityWeekBtn) granularityWeekBtn.classList.toggle("active", !isDay);
    if (granularityDayBtn) granularityDayBtn.classList.toggle("active", isDay);
    granularityControl.style.setProperty("--granularity-active-index", isDay ? "1" : "0");
    updateChartTitle();
  }

  function setGranularity(mode, { persist = true, render = true } = {}) {
    currentGranularity = mode === "day" ? "day" : "week";
    if (persist) localStorage.setItem(GRANULARITY_STORAGE_KEY, currentGranularity);
    updateGranularityUi();
    if (render) applyYearFilterAndRender();
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
    rawIngressItems.forEach((item) => {
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
        weekDetail.innerHTML = `<div class="reports-week-detail-title">Nessun periodo disponibile.</div>`;
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

    const todayKey =
      currentGranularity === "day" ? (getDayInfo(new Date())?.key || "") : getIsoWeekInfo(new Date()).key;
    if (todayKey && selectWeekInList(todayKey, { scrollIntoView: true })) return;

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
      weekDetail.innerHTML = `<div class="reports-week-detail-title">Nessuna commessa trovata per il periodo selezionato.</div>`;
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

  function selectWeekInList(weekKey, { scrollIntoView = true } = {}) {
    if (!weekList || !weekKey) return false;
    const item = weekList.querySelector(`.reports-week-item[data-week-key="${String(weekKey)}"]`);
    if (!item) return false;
    weekList.querySelectorAll(".reports-week-item.active").forEach((el) => el.classList.remove("active"));
    item.classList.add("active");
    renderWeekDetail(weekKey);
    if (scrollIntoView) {
      const periodCard = weekList.closest(".reports-card");
      if (periodCard && typeof periodCard.scrollIntoView === "function") {
        periodCard.scrollIntoView({ behavior: "smooth", block: "start" });
      }
      window.requestAnimationFrame(() => {
        if (typeof item.scrollIntoView === "function") {
          item.scrollIntoView({ behavior: "smooth", block: "nearest" });
        } else {
          weekList.scrollTop = Math.max(0, item.offsetTop - 4);
        }
      });
    }
    return true;
  }

  function renderChart(weekRows) {
    if (!chartCanvas) return;
    const today = new Date();
    const currentWeekKey = getIsoWeekInfo(today).key;
    const currentDayKey = getDayInfo(today)?.key || "";
    const currentIndex = weekRows.findIndex((row) =>
      currentGranularity === "day" ? row.key === currentDayKey : row.key === currentWeekKey
    );
    const isCurrentTick = (index) => {
      const row = weekRows[index];
      if (!row) return false;
      return currentGranularity === "day" ? row.key === currentDayKey : row.key === currentWeekKey;
    };
    const labels = weekRows.map((row) =>
      currentGranularity === "day"
        ? String(row.label || "")
        : String(row.label || "").match(/W\d{1,2}$/i)?.[0]?.toUpperCase() || String(row.label || "")
    );
    const orderedValues = weekRows.map((row) => Number(row.orderedTotal || 0));
    const plannedValues = weekRows.map((row) => Number(row.plannedTotal || 0));
    const targetValue = currentGranularity === "day" ? getWeeklyTarget() / 5 : getWeeklyTarget();
    const targetSeries = labels.map(() => targetValue);
    const avgWeeklyValue = currentGranularity === "day" ? currentWeeklyAverage / 5 : currentWeeklyAverage;
    const avgWeeklySeries = labels.map(() => avgWeeklyValue);
    const intersectionDotsPlugin = {
      id: "intersectionDots",
      afterDraw(chartInstance) {
        const xScale = chartInstance.scales?.x;
        const yScale = chartInstance.scales?.y;
        if (!xScale || !yScale) return;
        const xTicks = xScale.ticks || [];
        const yTicks = yScale.ticks || [];
        const area = chartInstance.chartArea;
        if (!area) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.fillStyle = "rgba(148, 163, 184, 0.45)";
        const radius = 1.35;
        xTicks.forEach((_, xIndex) => {
          const x = xScale.getPixelForTick(xIndex);
          if (x < area.left || x > area.right) return;
          yTicks.forEach((_, yIndex) => {
            const y = yScale.getPixelForTick(yIndex);
            if (y < area.top || y > area.bottom) return;
            ctx.beginPath();
            ctx.arc(x, y, radius, 0, Math.PI * 2);
            ctx.fill();
          });
        });
        ctx.restore();
      },
    };
    const currentVerticalLinePlugin = {
      id: "currentVerticalLineMain",
      afterDraw(chartInstance) {
        if (currentIndex < 0) return;
        const xScale = chartInstance.scales?.x;
        const area = chartInstance.chartArea;
        if (!xScale || !area) return;
        const x = xScale.getPixelForTick(currentIndex);
        if (!Number.isFinite(x) || x < area.left || x > area.right) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.strokeStyle = "rgba(31, 92, 53, 0.45)";
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.beginPath();
        ctx.moveTo(x, area.top);
        ctx.lineTo(x, area.bottom);
        ctx.stroke();
        ctx.restore();
      },
    };
    const barValueLabelsPlugin = makeBarValueLabelsPlugin("barValueLabelsMain");
    const avgLineValuePlugin = {
      id: "avgLineValueMain",
      afterDatasetsDraw(chartInstance) {
        const yScale = chartInstance.scales?.y;
        const area = chartInstance.chartArea;
        if (!yScale || !area) return;
        const y = yScale.getPixelForValue(avgWeeklyValue);
        if (!Number.isFinite(y) || y < area.top || y > area.bottom) return;
        const text = `Media ${formatMillionEuroShort(avgWeeklyValue)}`;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.font = "700 10px 'IBM Plex Sans', sans-serif";
        const textWidth = ctx.measureText(text).width;
        const padX = 6;
        const boxH = 16;
        const boxW = textWidth + padX * 2;
        const x = Math.max(area.left + 4, area.right - boxW - 2);
        const yTop = Math.max(area.top + 2, y - boxH - 2);
        ctx.fillStyle = "rgba(255, 255, 255, 0.72)";
        ctx.beginPath();
        ctx.roundRect(x, yTop, boxW, boxH, 8);
        ctx.fill();
        ctx.fillStyle = "#1f5c35";
        ctx.textBaseline = "middle";
        ctx.fillText(text, x + padX, yTop + boxH / 2);
        ctx.restore();
      },
    };

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
            label: currentGranularity === "day" ? "Target giornaliero" : "Target settimanale",
            data: targetSeries,
            borderColor: "rgba(108, 95, 199, 0.95)",
            borderWidth: 1.2,
            borderDash: [8, 6],
            stack: "guide_target",
            pointRadius: 0,
            pointHoverRadius: 0,
            tension: 0,
          },
          {
            type: "line",
            label: currentGranularity === "day" ? "Media settimanale (÷5)" : "Media settimanale",
            data: avgWeeklySeries,
            borderColor: "rgba(47, 111, 61, 0.9)",
            borderWidth: 1.2,
            borderDash: [4, 4],
            stack: "guide_avg",
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
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              maxRotation: 0,
              autoSkipPadding: 10,
              color: (ctx) => (isCurrentTick(ctx.index) ? "#1f5c35" : "#6b7280"),
              font: (ctx) => ({
                size: 11,
                weight: isCurrentTick(ctx.index) ? "700" : "500",
              }),
            },
          },
          y: {
            beginAtZero: true,
            stacked: true,
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              callback: (value) => moneyAxisFmt.format(Number(value || 0)),
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 12,
              boxHeight: 12,
              padding: 14,
            },
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${moneyFmt.format(Number(ctx.raw || 0))}`,
            },
          },
        },
        onClick: (_event, elements) => {
          if (!elements || !elements.length) return;
          const row = weekRows[elements[0].index];
          if (!row) return;
          selectWeekInList(row.key, { scrollIntoView: true });
        },
      },
      plugins: [intersectionDotsPlugin, currentVerticalLinePlugin, avgLineValuePlugin, barValueLabelsPlugin],
    });
  }

  function renderIngressChart(ingressRows) {
    if (!ingressChartCanvas) return;
    const currentWeekKey = getIsoWeekInfo(new Date()).key;
    const currentIndex = ingressRows.findIndex((row) => row.key === currentWeekKey);
    const isCurrentWeekTick = (index) => {
      const row = ingressRows[index];
      return Boolean(row) && row.key === currentWeekKey;
    };
    const labels = ingressRows.map((row) => String(row.label || "").match(/W\d{1,2}$/i)?.[0]?.toUpperCase() || String(row.label || ""));
    const values = ingressRows.map((row) => (row.isAfterToday ? null : Number(row.ingressTotal || 0)));
    const avgSeries = ingressRows.map((row) => (row.isAfterToday ? null : currentIngressWeeklyAverage));

    const intersectionDotsPlugin = {
      id: "intersectionDotsIngress",
      afterDraw(chartInstance) {
        const xScale = chartInstance.scales?.x;
        const yScale = chartInstance.scales?.y;
        if (!xScale || !yScale) return;
        const xTicks = xScale.ticks || [];
        const yTicks = yScale.ticks || [];
        const area = chartInstance.chartArea;
        if (!area) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.fillStyle = "rgba(148, 163, 184, 0.42)";
        const radius = 1.3;
        xTicks.forEach((_, xIndex) => {
          const x = xScale.getPixelForTick(xIndex);
          if (x < area.left || x > area.right) return;
          yTicks.forEach((_, yIndex) => {
            const y = yScale.getPixelForTick(yIndex);
            if (y < area.top || y > area.bottom) return;
            ctx.beginPath();
            ctx.arc(x, y, radius, 0, Math.PI * 2);
            ctx.fill();
          });
        });
        ctx.restore();
      },
    };
    const currentVerticalLinePlugin = {
      id: "currentVerticalLineIngress",
      afterDraw(chartInstance) {
        if (currentIndex < 0) return;
        const xScale = chartInstance.scales?.x;
        const area = chartInstance.chartArea;
        if (!xScale || !area) return;
        const x = xScale.getPixelForTick(currentIndex);
        if (!Number.isFinite(x) || x < area.left || x > area.right) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.strokeStyle = "rgba(31, 92, 53, 0.45)";
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.beginPath();
        ctx.moveTo(x, area.top);
        ctx.lineTo(x, area.bottom);
        ctx.stroke();
        ctx.restore();
      },
    };
    if (ingressChart) ingressChart.destroy();
    ingressChart = new window.Chart(ingressChartCanvas, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Imponibile entrato",
            data: values,
            backgroundColor: "rgba(99, 102, 241, 0.35)",
            borderColor: "rgba(79, 70, 229, 0.95)",
            borderWidth: 1.3,
            borderRadius: 8,
            categoryPercentage: 0.82,
            barPercentage: 0.9,
            maxBarThickness: 44,
          },
          {
            type: "line",
            label: "Media settimanale entrato",
            data: avgSeries,
            borderColor: "rgba(47, 111, 61, 0.9)",
            borderWidth: 1.2,
            borderDash: [4, 4],
            stack: "guide_ingress_avg",
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
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              maxRotation: 0,
              autoSkipPadding: 10,
              color: (ctx) => (isCurrentWeekTick(ctx.index) ? "#1f5c35" : "#6b7280"),
              font: (ctx) => ({
                size: 11,
                weight: isCurrentWeekTick(ctx.index) ? "700" : "500",
              }),
            },
          },
          y: {
            beginAtZero: true,
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              callback: (value) => moneyAxisFmt.format(Number(value || 0)),
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 12,
              boxHeight: 12,
              padding: 14,
            },
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${moneyFmt.format(Number(ctx.raw || 0))}`,
            },
          },
        },
        onClick: (_event, elements) => {
          if (!elements || !elements.length) return;
          const row = ingressRows[elements[0].index];
          if (!row) return;
          selectWeekInList(row.key, { scrollIntoView: true });
        },
      },
      plugins: [intersectionDotsPlugin, currentVerticalLinePlugin],
    });
  }

  function renderFutureTrendChart(ingressRows) {
    if (!futureTrendChartCanvas) return;
    const currentWeekKey = getIsoWeekInfo(new Date()).key;
    const currentIndex = ingressRows.findIndex((row) => row.key === currentWeekKey);
    const isCurrentWeekTick = (index) => {
      const row = ingressRows[index];
      return Boolean(row) && row.key === currentWeekKey;
    };
    const labels = ingressRows.map(
      (row) => String(row.label || "").match(/W\d{1,2}$/i)?.[0]?.toUpperCase() || String(row.label || "")
    );
    const values = ingressRows.map((row) => Number(row.cumulativeFutureTotal || 0));
    const intersectionDotsPlugin = {
      id: "intersectionDotsFutureTrend",
      afterDraw(chartInstance) {
        const xScale = chartInstance.scales?.x;
        const yScale = chartInstance.scales?.y;
        if (!xScale || !yScale) return;
        const xTicks = xScale.ticks || [];
        const yTicks = yScale.ticks || [];
        const area = chartInstance.chartArea;
        if (!area) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.fillStyle = "rgba(148, 163, 184, 0.42)";
        const radius = 1.3;
        xTicks.forEach((_, xIndex) => {
          const x = xScale.getPixelForTick(xIndex);
          if (x < area.left || x > area.right) return;
          yTicks.forEach((_, yIndex) => {
            const y = yScale.getPixelForTick(yIndex);
            if (y < area.top || y > area.bottom) return;
            ctx.beginPath();
            ctx.arc(x, y, radius, 0, Math.PI * 2);
            ctx.fill();
          });
        });
        ctx.restore();
      },
    };
    const currentVerticalLinePlugin = {
      id: "currentVerticalLineFutureTrend",
      afterDraw(chartInstance) {
        if (currentIndex < 0) return;
        const xScale = chartInstance.scales?.x;
        const area = chartInstance.chartArea;
        if (!xScale || !area) return;
        const x = xScale.getPixelForTick(currentIndex);
        if (!Number.isFinite(x) || x < area.left || x > area.right) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.strokeStyle = "rgba(31, 92, 53, 0.45)";
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.beginPath();
        ctx.moveTo(x, area.top);
        ctx.lineTo(x, area.bottom);
        ctx.stroke();
        ctx.restore();
      },
    };

    if (futureTrendChart) futureTrendChart.destroy();
    futureTrendChart = new window.Chart(futureTrendChartCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: "Ordinato + pianificato (settimana a ∞)",
            data: values,
            borderColor: "rgba(245, 158, 11, 0.95)",
            backgroundColor: "rgba(245, 158, 11, 0.12)",
            borderWidth: 2,
            tension: 0.2,
            pointRadius: 2.5,
            pointHoverRadius: 3.5,
            fill: false,
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
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              maxRotation: 0,
              autoSkipPadding: 10,
              color: (ctx) => (isCurrentWeekTick(ctx.index) ? "#1f5c35" : "#6b7280"),
              font: (ctx) => ({
                size: 11,
                weight: isCurrentWeekTick(ctx.index) ? "700" : "500",
              }),
            },
          },
          y: {
            beginAtZero: true,
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              callback: (value) => moneyAxisFmt.format(Number(value || 0)),
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 12,
              boxHeight: 12,
              padding: 14,
            },
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${moneyFmt.format(Number(ctx.raw || 0))}`,
            },
          },
        },
        onClick: (_event, elements) => {
          if (!elements || !elements.length) return;
          const row = ingressRows[elements[0].index];
          if (!row) return;
          selectWeekInList(row.key, { scrollIntoView: true });
        },
      },
      plugins: [intersectionDotsPlugin, currentVerticalLinePlugin],
    });
  }

  function renderBacklogChart(backlogRows) {
    if (!backlogChartCanvas) return;
    const currentWeekKey = getIsoWeekInfo(new Date()).key;
    const currentIndex = backlogRows.findIndex((row) => row.key === currentWeekKey);
    const labels = backlogRows.map(
      (row) => String(row.label || "").match(/W\d{1,2}$/i)?.[0]?.toUpperCase() || String(row.label || "")
    );
    const backlogValues = backlogRows.map((row) => Number(row.backlogTotal || 0));
    const backlogWeeksValues = backlogRows.map((row) => Number(row.backlogWeeks || 0));

    const intersectionDotsPlugin = {
      id: "intersectionDotsBacklog",
      afterDraw(chartInstance) {
        const xScale = chartInstance.scales?.x;
        const yScale = chartInstance.scales?.y;
        if (!xScale || !yScale) return;
        const xTicks = xScale.ticks || [];
        const yTicks = yScale.ticks || [];
        const area = chartInstance.chartArea;
        if (!area) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.fillStyle = "rgba(148, 163, 184, 0.42)";
        const radius = 1.3;
        xTicks.forEach((_, xIndex) => {
          const x = xScale.getPixelForTick(xIndex);
          if (x < area.left || x > area.right) return;
          yTicks.forEach((_, yIndex) => {
            const y = yScale.getPixelForTick(yIndex);
            if (y < area.top || y > area.bottom) return;
            ctx.beginPath();
            ctx.arc(x, y, radius, 0, Math.PI * 2);
            ctx.fill();
          });
        });
        ctx.restore();
      },
    };
    const currentVerticalLinePlugin = {
      id: "currentVerticalLineBacklog",
      afterDraw(chartInstance) {
        if (currentIndex < 0) return;
        const xScale = chartInstance.scales?.x;
        const area = chartInstance.chartArea;
        if (!xScale || !area) return;
        const x = xScale.getPixelForTick(currentIndex);
        if (!Number.isFinite(x) || x < area.left || x > area.right) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.strokeStyle = "rgba(31, 92, 53, 0.45)";
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.beginPath();
        ctx.moveTo(x, area.top);
        ctx.lineTo(x, area.bottom);
        ctx.stroke();
        ctx.restore();
      },
    };
    const zeroBacklogLinePlugin = {
      id: "zeroBacklogLine",
      afterDraw(chartInstance) {
        const yScale = chartInstance.scales?.y;
        const area = chartInstance.chartArea;
        if (!yScale || !area) return;
        const y = yScale.getPixelForValue(0);
        if (!Number.isFinite(y) || y < area.top || y > area.bottom) return;
        const ctx = chartInstance.ctx;
        ctx.save();
        ctx.strokeStyle = "rgba(100, 116, 139, 0.7)";
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.beginPath();
        ctx.moveTo(area.left, y);
        ctx.lineTo(area.right, y);
        ctx.stroke();
        ctx.restore();
      },
    };
    if (backlogChart) backlogChart.destroy();
    backlogChart = new window.Chart(backlogChartCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: "Backlog cumulato",
            data: backlogValues,
            backgroundColor: "rgba(107, 114, 128, 0.08)",
            borderColor: "rgba(75, 85, 99, 0.95)",
            borderWidth: 1.6,
            tension: 0.2,
            pointRadius: 2.5,
            pointHoverRadius: 3.5,
            fill: false,
            yAxisID: "y",
            segment: {
              borderColor: (ctx) => {
                const p0 = Number(ctx?.p0?.parsed?.y || 0);
                const p1 = Number(ctx?.p1?.parsed?.y || 0);
                if (p1 > p0) return "rgba(220, 38, 38, 0.95)";
                if (p1 < p0) return "rgba(22, 163, 74, 0.95)";
                return "rgba(75, 85, 99, 0.95)";
              },
            },
          },
          {
            label: "Settimane backlog",
            data: backlogWeeksValues,
            borderColor: "rgba(59, 130, 246, 0.9)",
            backgroundColor: "rgba(59, 130, 246, 0.12)",
            borderWidth: 1.2,
            borderDash: [6, 4],
            tension: 0.2,
            pointRadius: 0,
            pointHoverRadius: 2,
            fill: false,
            yAxisID: "yBacklogWeeks",
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
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              maxRotation: 0,
              autoSkipPadding: 10,
              color: (ctx) => (ctx.index === currentIndex ? "#1f5c35" : "#6b7280"),
              font: (ctx) => ({
                size: 11,
                weight: ctx.index === currentIndex ? "700" : "500",
              }),
            },
          },
          y: {
            beginAtZero: false,
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              callback: (value) => moneyAxisFmt.format(Number(value || 0)),
            },
          },
          yBacklogWeeks: {
            position: "right",
            beginAtZero: false,
            grid: {
              drawOnChartArea: false,
              drawTicks: false,
            },
            ticks: {
              callback: (value) => `${numberAxisFmt.format(Number(value || 0))} w`,
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 12,
              boxHeight: 12,
              padding: 14,
            },
          },
          tooltip: {
            callbacks: {
              title: (items) => {
                const idx = items?.[0]?.dataIndex ?? -1;
                const row = backlogRows[idx];
                if (!row) return "";
                const shortWeek = String(row.label || "").match(/W\d{1,2}$/i)?.[0]?.toUpperCase() || row.label;
                return `Settimana ${shortWeek}`;
              },
              beforeBody: (items) => {
                const idx = items?.[0]?.dataIndex ?? -1;
                const row = backlogRows[idx];
                if (!row) return [];
                return [
                  `Ordini in ingresso: ${moneyFmt.format(Number(row.ingressTotal || 0))}`,
                  `Telaio pianificato: ${moneyFmt.format(Number(row.plannedTotal || 0))}`,
                  `Telaio ordinato: ${moneyFmt.format(Number(row.orderedTotal || 0))}`,
                  `Delta settimanale: ${moneyFmt.format(Number(row.deltaTotal || 0))}`,
                ];
              },
              label: (ctx) => {
                const row = backlogRows[ctx.dataIndex];
                if (!row) return "";
                if (ctx.datasetIndex === 0) return `Backlog cumulato: ${moneyFmt.format(Number(row.backlogTotal || 0))}`;
                if (ctx.datasetIndex === 1) return `Settimane backlog: ${numberFmt.format(Number(row.backlogWeeks || 0))}`;
                return "";
              },
            },
          },
        },
        onClick: (_event, elements) => {
          if (!elements || !elements.length) return;
          const row = backlogRows[elements[0].index];
          if (!row) return;
          selectWeekInList(row.key, { scrollIntoView: true });
        },
      },
      plugins: [intersectionDotsPlugin, currentVerticalLinePlugin, zeroBacklogLinePlugin],
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
    const filteredIngressItems =
      selectedYear === "all"
        ? rawIngressItems.slice()
        : rawIngressItems.filter((item) => String(item.year) === String(selectedYear));

    const totals = new Map();
    const details = new Map();

    filteredItems.forEach((item) => {
      const key = currentGranularity === "day" ? item.dayKey : item.weekKey;
      const label = currentGranularity === "day" ? item.dayLabel : item.weekLabel;
      const sortValue = currentGranularity === "day" ? item.daySortValue : item.weekSortValue;
      const agg = totals.get(key) || {
        key,
        label,
        sortValue,
        orderedTotal: 0,
        plannedTotal: 0,
      };

      if (item.category === "ordered") agg.orderedTotal += item.imponibile;
      else agg.plannedTotal += item.imponibile;
      totals.set(key, agg);

      const bucket = details.get(key) || { ordered: [], planned: [] };
      if (item.category === "ordered") bucket.ordered.push({ codice: item.codice, imponibile: item.imponibile });
      else bucket.planned.push({ codice: item.codice, imponibile: item.imponibile });
      details.set(key, bucket);
    });

    const today = new Date();
    const orderedItemsToToday = filteredItems.filter((item) => {
      if (item.category !== "ordered") return false;
      const d = parseIsoDateOnly(item.eventDateKey);
      return d && d <= today;
    });

    const weeklyTotals = new Map();
    orderedItemsToToday
      .forEach((item) => {
        const current = weeklyTotals.get(item.weekKey) || 0;
        weeklyTotals.set(item.weekKey, current + Number(item.imponibile || 0));
      });
    const weeklyValues = Array.from(weeklyTotals.values());
    currentWeeklyAverage = weeklyValues.length
      ? weeklyValues.reduce((sum, value) => sum + Number(value || 0), 0) / weeklyValues.length
      : 0;

    const ingressTotals = new Map();
    filteredIngressItems.forEach((item) => {
      const agg = ingressTotals.get(item.weekKey) || {
        key: item.weekKey,
        label: item.weekLabel,
        sortValue: item.weekSortValue,
        total: 0,
      };
      agg.total += Number(item.imponibile || 0);
      ingressTotals.set(item.weekKey, agg);
    });
    const orderedPlannedTotals = new Map();
    filteredItems.forEach((item) => {
      const agg = orderedPlannedTotals.get(item.weekKey) || {
        key: item.weekKey,
        label: item.weekLabel,
        sortValue: item.weekSortValue,
        total: 0,
      };
      agg.total += Number(item.imponibile || 0);
      orderedPlannedTotals.set(item.weekKey, agg);
    });
    const allWeekKeys = new Set([...ingressTotals.keys(), ...orderedPlannedTotals.keys()]);
    const allWeeks = Array.from(allWeekKeys)
      .map((key) => ingressTotals.get(key) || orderedPlannedTotals.get(key))
      .filter(Boolean)
      .sort((a, b) => a.sortValue - b.sortValue);
    const cumulativeMap = new Map();
    let running = 0;
    for (let i = allWeeks.length - 1; i >= 0; i -= 1) {
      const row = allWeeks[i];
      const thisWeekTotal = Number(orderedPlannedTotals.get(row.key)?.total || 0);
      running += thisWeekTotal;
      cumulativeMap.set(row.key, running);
    }
    const todayWeekSort = getIsoWeekInfo(today).sortValue;
    const ingressRowsAll = allWeeks.map((row) => ({
      key: row.key,
      label: row.label,
      sortValue: row.sortValue,
      ingressTotal: Number(ingressTotals.get(row.key)?.total || 0),
      cumulativeFutureTotal: Number(cumulativeMap.get(row.key) || 0),
      isAfterToday: Number(row.sortValue || 0) > todayWeekSort,
    }));
    const ingressRows = ingressRowsAll.slice();
    currentIngressWeeklyAverage = ingressRows.length
      ? ingressRows
          .filter((row) => !row.isAfterToday)
          .reduce((sum, row) => sum + Number(row.ingressTotal || 0), 0) /
        Math.max(1, ingressRows.filter((row) => !row.isAfterToday).length)
      : 0;

    lastWeekDetails = details;
    const plannedByWeek = new Map();
    const orderedByWeek = new Map();
    filteredItems.forEach((item) => {
      if (item.category === "planned") {
        const existing = plannedByWeek.get(item.weekKey) || 0;
        plannedByWeek.set(item.weekKey, existing + Number(item.imponibile || 0));
      } else if (item.category === "ordered") {
        const existing = orderedByWeek.get(item.weekKey) || 0;
        orderedByWeek.set(item.weekKey, existing + Number(item.imponibile || 0));
      }
    });

    const ingressByWeek = new Map();
    filteredIngressItems.forEach((item) => {
      const agg = ingressByWeek.get(item.weekKey) || {
        key: item.weekKey,
        label: item.weekLabel,
        sortValue: item.weekSortValue,
        total: 0,
      };
      agg.total += Number(item.imponibile || 0);
      ingressByWeek.set(item.weekKey, agg);
    });

    const backlogBaseRows = allWeeks.map((week) => ({
      key: week.key,
      label: week.label,
      sortValue: week.sortValue,
      ingressTotal: Number(ingressByWeek.get(week.key)?.total || 0),
      plannedTotal: Number(plannedByWeek.get(week.key) || 0),
      orderedTotal: Number(orderedByWeek.get(week.key) || 0),
    }));

    let backlogRunning = 0;
    const backlogRows = backlogBaseRows.map((row) => {
      const deltaTotal = Number(row.ingressTotal || 0) - Number(row.plannedTotal || 0);
      backlogRunning += deltaTotal;
      return {
        ...row,
        deltaTotal,
        backlogTotal: backlogRunning,
        backlogWeeks: backlogRunning / 1_000_000,
      };
    });

    const weekRowsForList = Array.from(totals.values()).sort((a, b) => a.sortValue - b.sortValue);
    const weekRows =
      currentGranularity === "day"
        ? weekRowsForList
        : allWeeks.map((week) => {
            const fromTotals = totals.get(week.key);
            return {
              key: week.key,
              label: week.label,
              sortValue: week.sortValue,
              orderedTotal: Number(fromTotals?.orderedTotal || 0),
              plannedTotal: Number(fromTotals?.plannedTotal || 0),
            };
          });
    renderChart(weekRows);
    renderIngressChart(ingressRows);
    renderFutureTrendChart(ingressRowsAll);
    renderBacklogChart(backlogRows);
    renderWeekList(weekRowsForList);

    const orderedPeriodsCount = new Set(
      orderedItemsToToday.map((item) => (currentGranularity === "day" ? item.dayKey : item.weekKey))
    ).size;
    renderKpis(orderedItemsToToday, orderedPeriodsCount);
  }

  async function loadWeeklyChart() {
    if (refreshBtn) refreshBtn.disabled = true;
    try {
      const { data, error } = await supabase
        .from("commessa_imponibili")
        .select("imponibile, commesse!inner(id,codice,telaio_consegnato,data_ordine_telaio,data_consegna_telaio_effettiva,data_ingresso)");
      if (error) throw error;

      const items = [];
      const ingressItems = [];
      const telaioInconsistencies = [];
      (data || []).forEach((row) => {
        const commessa = row.commesse || null;
        const imponibile = Number(row.imponibile);
        if (!Number.isFinite(imponibile)) return;

        const isOrdered = Boolean(commessa?.telaio_consegnato);
        const orderedAt = parseIsoDateOnly(commessa?.data_consegna_telaio_effettiva);
        const plannedAt = parseIsoDateOnly(commessa?.data_consegna_telaio_effettiva);
        const targetAt = parseIsoDateOnly(commessa?.data_ordine_telaio);
        const ingressoAt = parseIsoDateOnly(commessa?.data_ingresso);
        const codice = String(commessa?.codice || "-");

        if (ingressoAt) {
          const weekIngresso = getIsoWeekInfo(ingressoAt);
          ingressItems.push({
            year: weekIngresso.year,
            weekKey: weekIngresso.key,
            weekLabel: weekIngresso.label,
            weekSortValue: weekIngresso.sortValue,
            imponibile,
          });
        }

        if (isOrdered && orderedAt) {
          const week = getIsoWeekInfo(orderedAt);
          const day = getDayInfo(orderedAt);
          if (!day) return;
          items.push({
            category: "ordered",
            year: week.year,
            commessaId: commessa?.id || "",
            weekKey: week.key,
            weekLabel: week.label,
            weekSortValue: week.sortValue,
            dayKey: day.key,
            dayLabel: day.label,
            daySortValue: day.sortValue,
            eventDateKey: day.key,
            codice: commessa?.codice || "",
            imponibile,
          });
          return;
        }
        if (isOrdered && !orderedAt) {
          telaioInconsistencies.push(`${codice}: ordinato=true senza data telaio ordinato`);
          return;
        }

        if (!isOrdered && plannedAt) {
          const week = getIsoWeekInfo(plannedAt);
          const day = getDayInfo(plannedAt);
          if (!day) return;
          items.push({
            category: "planned",
            year: week.year,
            commessaId: commessa?.id || "",
            weekKey: week.key,
            weekLabel: week.label,
            weekSortValue: week.sortValue,
            dayKey: day.key,
            dayLabel: day.label,
            daySortValue: day.sortValue,
            eventDateKey: day.key,
            codice: commessa?.codice || "",
            imponibile,
          });
          return;
        }

        if (!isOrdered && targetAt && !plannedAt) {
          telaioInconsistencies.push(`${codice}: target presente ma data telaio pianificato assente`);
        }
      });

      if (telaioInconsistencies.length) {
        const sample = telaioInconsistencies.slice(0, 3).join(" | ");
        const more = telaioInconsistencies.length > 3 ? " ..." : "";
        setImportStatus(
          `Incoerenze telai rilevate: ${telaioInconsistencies.length}. ${sample}${more}`,
          "error"
        );
      }

      rawWeeklyItems = items;
      rawIngressItems = ingressItems;
      populateYearFilter();
      applyYearFilterAndRender();
    } catch (err) {
      setImportStatus(`Errore caricamento grafico: ${err?.message || err}`, "error");
      rawWeeklyItems = [];
      rawIngressItems = [];
      if (yearFilter) {
        yearFilter.innerHTML = `<option value="2026">2026</option><option value="all">Tutti</option>`;
        yearFilter.value = "2026";
      }
      lastWeekDetails = new Map();
      renderChart([]);
      renderIngressChart([]);
      renderFutureTrendChart([]);
      renderBacklogChart([]);
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

  if (copyHeaderBtn) {
    copyHeaderBtn.addEventListener("click", () => {
      void copyReportsImportHeader();
    });
  }

  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      void loadWeeklyChart();
    });
  }

  if (exportPdfBtn) {
    exportPdfBtn.addEventListener("click", () => {
      void handleExportChartsPdf();
    });
  }

  if (yearFilter) {
    yearFilter.addEventListener("change", () => {
      applyYearFilterAndRender();
    });
  }

  if (granularityWeekBtn) {
    granularityWeekBtn.addEventListener("click", () => {
      setGranularity("week");
    });
  }

  if (granularityDayBtn) {
    granularityDayBtn.addEventListener("click", () => {
      setGranularity("day");
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
      selectWeekInList(weekKey, { scrollIntoView: false });
    });
  }

  applyStoredTheme();

  void (async () => {
    refreshClock();
    window.setInterval(refreshClock, 1000);
    const ok = await ensureAdminAccess();
    if (!ok) return;
    setGranularity(currentGranularity, { persist: false, render: false });
    if (yearFilter) {
      yearFilter.innerHTML = `<option value="2026">2026</option><option value="all">Tutti</option>`;
      yearFilter.value = "2026";
    }
    renderTargetInputValue();
    await loadWeeklyChart();
  })();
})();
