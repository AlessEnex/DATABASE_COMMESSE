-- Add reparto to risorse
alter table if exists risorse
  add column if not exists reparto_id smallint references reparti(id) on delete set null;

create index if not exists idx_risorse_reparto_id on risorse (reparto_id);
