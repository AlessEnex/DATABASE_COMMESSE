# Migrazioni SQL – descrizioni human friendly

Quando aggiungi una nuova migrazione, aggiungi qui una riga con la descrizione in linguaggio semplice.

| File | Descrizione |
| --- | --- |
| `20260227130500_init_commesse.sql` | Crea lo schema base (tabelle, enum, funzioni, trigger, RLS) e inserisce dati iniziali per reparti, risorse e una commessa di esempio. |
| `20260303143000_add_ore_assenza.sql` | Aggiunge il campo `ore_assenza` nelle attività e rende `commessa_id` opzionale per gestire assenze non legate a una commessa. |
| `20260304130000_add_reparto_to_risorse.sql` | Aggiunge il collegamento tra risorse e reparti e crea l’indice relativo. |
| `20260310120000_add_utente_to_risorse.sql` | Collega ogni risorsa a un utente (uno a uno) tramite `utente_id` con vincolo e indice univoco. |
| `20260310130000_add_permessi_ruolo.sql` | Introduce i permessi per ruolo (matrice e gantt), con RLS e valori di default. |
| `20260310133000_rls_attivita_per_risorsa.sql` | Limita le modifiche alle attività: operatore può intervenire solo sulla propria risorsa, admin/responsabile su tutte. |
| `20260310134500_rls_restrict_commesse.sql` | Restringe le operazioni di scrittura su commesse/risorse/utenti/reparti/assegnazioni a admin e responsabili. |
| `20260311101500_add_planner_role.sql` | Aggiunge il ruolo `planner`, i suoi permessi base e una funzione che consente di aggiornare solo le date ordine/consegna commessa. |
| `20260311113000_add_data_arrivo_kit_cavi.sql` | Aggiunge la data arrivo kit cavi alle commesse ed estende la funzione planner per aggiornarla. |
| `20260311121000_update_v_commesse_add_kit_cavi.sql` | Aggiorna la vista `v_commesse` per includere `data_arrivo_kit_cavi`. |
| `20260311133000_update_planner_rpc_add_prelievo.sql` | Estende la funzione planner per aggiornare anche la data prelievo materiali. |
