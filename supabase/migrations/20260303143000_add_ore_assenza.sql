alter table if exists public.attivita
  add column if not exists ore_assenza numeric;

alter table if exists public.attivita
  alter column commessa_id drop not null;
