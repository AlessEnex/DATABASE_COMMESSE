(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";
  const AUTH_STORAGE_KEY = "sb-bsceqirconhqmxwipbyl-auth-token";

  const statusEl = document.getElementById("sheetStatus");
  const subtitleEl = document.getElementById("commessaSubtitle");
  const schedaForm = document.getElementById("schedaForm");
  const saveSchedaBtn = document.getElementById("saveSchedaBtn");
  const closeSchedaBtn = document.getElementById("closeSchedaBtn");
  const tipoMacchina = document.getElementById("tipoMacchina");
  const numCompressoriMt = document.getElementById("numCompressoriMt");
  const numCompressoriLt = document.getElementById("numCompressoriLt");
  const numCompressoriAux = document.getElementById("numCompressoriAux");
  const numCompressoriMt2 = document.getElementById("numCompressoriMt2");
  const numCompressoriLt2 = document.getElementById("numCompressoriLt2");

  const setStatus = (message, tone = "") => {
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.classList.remove("ok", "error", "hidden");
    if (tone) statusEl.classList.add(tone);
  };

  const getToken = () => {
    const raw = localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return null;
    try {
      return JSON.parse(raw)?.access_token || null;
    } catch {
      return null;
    }
  };

  const restFetch = async (path, options = {}) => {
    const token = getToken();
    if (!token) throw new Error("Token non disponibile.");
    const res = await fetch(`${SUPABASE_URL}${path}`, {
      ...options,
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });
    return res;
  };

  const parseIntOrNull = (value) => {
    const n = Number.parseInt(value, 10);
    return Number.isFinite(n) ? n : null;
  };

  const loadCommessa = async (commessaId) => {
    const res = await restFetch(`/rest/v1/commesse?id=eq.${commessaId}&select=*`);
    if (!res.ok) throw new Error("Errore caricamento commessa.");
    const data = await res.json();
    return data[0] || null;
  };

  const loadScheda = async (commessaId) => {
    const res = await restFetch(`/rest/v1/commessa_schede?commessa_id=eq.${commessaId}&select=*`);
    if (!res.ok) throw new Error("Errore caricamento scheda.");
    const data = await res.json();
    return data[0] || null;
  };

  const saveScheda = async (commessaId) => {
    const payload = {
      commessa_id: commessaId,
      tipo_macchina: tipoMacchina.value || null,
      num_compressori_mt: parseIntOrNull(numCompressoriMt.value),
      num_compressori_lt: parseIntOrNull(numCompressoriLt.value),
      num_compressori_aux: parseIntOrNull(numCompressoriAux.value),
      num_compressori_mt2_pdc_ac: parseIntOrNull(numCompressoriMt2.value),
      num_compressori_lt2: parseIntOrNull(numCompressoriLt2.value),
    };

    const res = await restFetch("/rest/v1/commessa_schede", {
      method: "POST",
      headers: {
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(text || "Errore salvataggio.");
    }
    return res.json();
  };

  const init = async () => {
    const params = new URLSearchParams(window.location.search);
    const commessaId = params.get("id");
    if (!commessaId) {
      setStatus("ID commessa mancante.", "error");
      return;
    }

    try {
      const commessa = await loadCommessa(commessaId);
      if (!commessa) {
        setStatus("Commessa non trovata.", "error");
        return;
      }
      const title = commessa.titolo ? ` - ${commessa.titolo}` : "";
      if (subtitleEl) subtitleEl.textContent = `${commessa.codice}${title}`;

      const scheda = await loadScheda(commessaId);
      if (scheda) {
        tipoMacchina.value = scheda.tipo_macchina || "";
        numCompressoriMt.value = scheda.num_compressori_mt ?? "";
        numCompressoriLt.value = scheda.num_compressori_lt ?? "";
        numCompressoriAux.value = scheda.num_compressori_aux ?? "";
        numCompressoriMt2.value = scheda.num_compressori_mt2_pdc_ac ?? "";
        numCompressoriLt2.value = scheda.num_compressori_lt2 ?? "";
      }
    } catch (err) {
      setStatus(err?.message || "Errore caricamento.", "error");
    }
  };

  if (schedaForm) {
    schedaForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const params = new URLSearchParams(window.location.search);
      const commessaId = params.get("id");
      if (!commessaId) {
        setStatus("ID commessa mancante.", "error");
        return;
      }
      if (saveSchedaBtn) saveSchedaBtn.disabled = true;
      try {
        await saveScheda(commessaId);
        setStatus("Scheda salvata.", "ok");
      } catch (err) {
        setStatus(err?.message || "Errore salvataggio.", "error");
      } finally {
        if (saveSchedaBtn) saveSchedaBtn.disabled = false;
      }
    });
  }

  if (closeSchedaBtn) {
    closeSchedaBtn.addEventListener("click", () => {
      window.close();
      if (!window.closed) {
        window.location.href = "./";
      }
    });
  }

  init();
})();
