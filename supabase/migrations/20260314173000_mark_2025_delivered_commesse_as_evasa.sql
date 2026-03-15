-- Storico: tutte le commesse con consegna macchina nel 2025 vanno in stato EVASA (db: chiusa).

update commesse
set stato = 'chiusa',
    updated_at = now()
where data_consegna_macchina >= date '2025-01-01'
  and data_consegna_macchina < date '2026-01-01'
  and stato is distinct from 'chiusa';
