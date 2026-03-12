(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

  const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
    },
  });

  const permessiList = document.getElementById("permessiList");
  const permessiStatus = document.getElementById("permessiStatus");
  const permessiSummary = document.getElementById("permessiSummary");
  const permessiSqlRules = document.getElementById("permessiSqlRules");
  const permessiFrontendRules = document.getElementById("permessiFrontendRules");
  const gate = document.getElementById("permessiGate");
  const gateTitle = document.getElementById("permessiGateTitle");
  const gateText = document.getElementById("permessiGateText");

  const setStatus = (message, tone = "") => {
    if (!permessiStatus) return;
    permessiStatus.textContent = message;
    permessiStatus.classList.remove("ok", "error", "hidden");
    if (tone) permessiStatus.classList.add(tone);
  };

  const showGate = (title, text) => {
    if (gateTitle) gateTitle.textContent = title;
    if (gateText) gateText.textContent = text;
    if (gate) gate.classList.remove("hidden");
  };

  const hideGate = () => {
    if (gate) gate.classList.add("hidden");
  };

  const roles = ["admin", "responsabile", "planner", "operatore", "viewer"];
  const fields = [
    { key: "can_move_matrix", label: "Spostare attivita (Matrice)" },
    { key: "can_delete_matrix", label: "Eliminare attivita (Matrice)" },
    { key: "can_move_gantt", label: "Spostare attivita (Gantt)" },
    { key: "can_delete_gantt", label: "Eliminare attivita (Gantt)" },
  ];

  const ensureAuth = async () => {
    const { data: sessionData } = await supabase.auth.getSession();
    const session = sessionData?.session || null;
    if (!session) {
      showGate("Accesso richiesto", "Effettua il login nella app per continuare.");
      return null;
    }
    const { data: userData } = await supabase.auth.getUser();
    const user = userData?.user || session.user;
    if (!user) {
      showGate("Accesso richiesto", "Effettua il login nella app per continuare.");
      return null;
    }
    const { data: profile } = await supabase.from("utenti").select("ruolo").eq("id", user.id).single();
    const role = String(profile?.ruolo || "").trim().toLowerCase() || "viewer";
    hideGate();
    return role;
  };

  const renderSqlRules = (data, isAdmin) => {
    if (!permessiSqlRules) return;
    permessiSqlRules.innerHTML = "";
    if (!data || !data.length) return;

    const intro = document.createElement("div");
    intro.className = "permessi-rules-title";
    intro.textContent = "Regole lato SQL";
    permessiSqlRules.appendChild(intro);

    const list = document.createElement("div");
    list.className = "permessi-rules-list";

    data.forEach((row) => {
      const enabled = fields.filter((f) => row[f.key]).map((f) => f.label);
      const card = document.createElement("div");
      card.className = "permessi-rules-card";
      const title = document.createElement("div");
      title.className = "permessi-role";
      title.textContent = row.ruolo;
      const body = document.createElement("div");
      body.className = "permessi-rules-body";
      body.textContent = enabled.length
        ? `Puo: ${enabled.join(", ")}.`
        : "Nessuna operazione abilitata.";
      card.appendChild(title);
      card.appendChild(body);
      list.appendChild(card);
    });

    const extra = document.createElement("div");
    extra.className = "permessi-rules-note";
    extra.textContent =
      "Regola extra: l'operatore puo agire solo sulle attivita della propria risorsa. Admin e responsabile possono gestire tutte le risorse. Il ruolo planner ha permessi limitati al Gantt (ordine/consegna).";

    const editNote = document.createElement("div");
    editNote.className = "permessi-rules-note";
    editNote.textContent = isAdmin
      ? "Puoi modificare i permessi usando i toggle sotto."
      : "Solo gli admin possono modificare i permessi: qui li vedi in sola lettura.";

    permessiSqlRules.appendChild(list);
    permessiSqlRules.appendChild(extra);
    permessiSqlRules.appendChild(editNote);
  };

  const renderFrontendRules = () => {
    if (!permessiFrontendRules) return;
    permessiFrontendRules.innerHTML = "";

    const intro = document.createElement("div");
    intro.className = "permessi-rules-title";
    intro.textContent = "Regole lato frontend";
    permessiFrontendRules.appendChild(intro);

    const list = document.createElement("ul");
    list.className = "permessi-front-list";
    list.innerHTML = `
      <li>In Matrice, l'operatore puo spostare/eliminare solo le attivita della propria risorsa.</li>
      <li>In Gantt, admin e planner possono modificare T lavanda (produzione) e consegna macchina.</li>
      <li>In Gantt, admin, planner e responsabili possono modificare l'arrivo kit cavi.</li>
      <li>In Gantt, admin e responsabili possono modificare i T pianificati (azzurri/verdi).</li>
      <li>Quando il telaio e segnato come ordinato, il T pianificato resta bloccato.</li>
      <li>To Dos (Preliminare CAD): admin, responsabile CAD e planner possono tracciare invio cliente, conferma e silenzio assenso.</li>
    `;

    permessiFrontendRules.appendChild(list);
  };

  const render = (data, isAdmin) => {
    if (!permessiList) return;
    permessiList.innerHTML = "";
    if (permessiSummary) permessiSummary.textContent = "";
    if (!data || !data.length) {
      permessiList.innerHTML = `<div class="matrix-empty">Nessun dato.</div>`;
      return;
    }
    if (permessiSummary) {
      const totals = data.map((row) => {
        const enabled = fields.reduce((acc, f) => acc + (row[f.key] ? 1 : 0), 0);
        return `${row.ruolo}: ${enabled}/${fields.length} attive`;
      });
      permessiSummary.textContent = `Totale regole per ruolo: ${totals.join(" | ")}`;
    }
    data.forEach((row) => {
      const card = document.createElement("div");
      card.className = "permessi-card";
      const title = document.createElement("div");
      title.className = "permessi-role";
      title.textContent = row.ruolo;
      card.appendChild(title);

      fields.forEach((f) => {
        const line = document.createElement("label");
        line.className = "permessi-line";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = Boolean(row[f.key]);
        checkbox.disabled = !isAdmin;
        checkbox.addEventListener("change", async () => {
          if (!isAdmin) return;
          const { error } = await supabase
            .from("permessi_ruolo")
            .update({ [f.key]: checkbox.checked })
            .eq("ruolo", row.ruolo);
          if (error) {
            setStatus(`Errore salvataggio: ${error.message}`, "error");
            checkbox.checked = !checkbox.checked;
            return;
          }
          setStatus("Permessi aggiornati.", "ok");
        });
        const text = document.createElement("span");
        text.textContent = f.label;
        line.appendChild(checkbox);
        line.appendChild(text);
        card.appendChild(line);
      });

      permessiList.appendChild(card);
    });
  };

  const load = async () => {
    const role = await ensureAuth();
    if (!role) return;
    const isAdmin = role === "admin";
    const { data, error } = await supabase.from("permessi_ruolo").select("*").order("ruolo");
    if (error) {
      setStatus(`Errore caricamento: ${error.message}`, "error");
      return;
    }
    renderSqlRules(data || [], isAdmin);
    renderFrontendRules();
    render(data || [], isAdmin);
  };

  load();
})();
