-- Link risorse to utenti (one-to-one)
alter table if exists risorse
  add column if not exists utente_id uuid;

do $$
begin
  if not exists (
    select 1 from pg_constraint where conname = 'risorse_utente_id_fkey'
  ) then
    alter table risorse
      add constraint risorse_utente_id_fkey
      foreign key (utente_id) references utenti(id) on delete set null;
  end if;
end $$;

create unique index if not exists idx_risorse_utente_id_unique on risorse (utente_id);
