# Database commesse (Supabase/Postgres)

Questo schema e pensato per un MVP con realtime: commesse, reparti, whitelist email, assegnazioni e log.

## Setup rapido (Supabase)
1. Apri il progetto Supabase e vai su SQL Editor.
2. Esegui lo script `schema.sql`.
3. Inserisci la prima email amministratore in whitelist:

```sql
insert into whitelist_email (email, ruolo_predefinito, reparto_id_predefinito)
values (
  'admin@azienda.it',
  'admin',
  (select id from reparti where nome = 'tutti')
);
```

4. Registra un utente con quella email. Il trigger creera il profilo in `utenti`.

Nota: lo SQL Editor usa il service role e bypassa l'RLS, quindi puoi inserire la whitelist iniziale senza problemi.

## Regole principali
- Solo email in `whitelist_email` possono leggere/scrivere i dati.
- Solo utenti `admin` possono modificare la whitelist.
- Gli utenti `viewer` hanno accesso in sola lettura.
- `codice` commessa: formato `YYYY_N` (es: `2026_001`).

## Tabelle principali
- `commesse`: anagrafica commessa.
- `reparti`: elenco reparti.
- `risorse`: elenco persone per la matrice.
- `commesse_reparti`: stato/note per reparto.
- `utenti`: profili interni collegati a `auth.users`.
- `assegnazioni`: chi fa cosa.
- `log_commessa`: audit minimale.
- `attivita`: eventi per calendario e matrice.

## Vista comoda
`v_commesse` aggrega i reparti in JSON per ogni commessa.

## Interfaccia web (MVP)
Cartella: `web/`

1. Apri `web/app.js` e imposta `SUPABASE_URL` e `SUPABASE_ANON_KEY`.
2. Apri `web/index.html` in browser.
3. Login con email in whitelist (verra inviato un link OTP).

Nota: gli utenti con ruolo `viewer` hanno sola lettura.

## Matrice risorse
- Tabella `risorse` con elenco persone.
- Sezione "Matrice risorse" per assegnare attivita alle commesse per giorno (lun-ven).
