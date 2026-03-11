do $$
begin
  if not exists (select 1 from pg_type where typname = 'attivita_stato') then
    create type attivita_stato as enum ('pianificata', 'in_corso', 'completata', 'annullata', 'non_necessaria');
  else
    if not exists (
      select 1 from pg_enum
      where enumlabel = 'non_necessaria'
        and enumtypid = 'attivita_stato'::regtype
    ) then
      alter type attivita_stato add value 'non_necessaria';
    end if;
  end if;
end $$;

create table if not exists commessa_attivita_override (
  commessa_id uuid not null references commesse(id) on delete cascade,
  titolo text not null,
  stato attivita_stato not null,
  updated_at timestamptz not null default now(),
  primary key (commessa_id, titolo),
  constraint commessa_attivita_override_only_non_necessaria check (stato::text = 'non_necessaria')
);

create index if not exists idx_commessa_attivita_override_commessa on commessa_attivita_override (commessa_id);

alter table commessa_attivita_override enable row level security;

drop policy if exists commessa_attivita_override_select on commessa_attivita_override;
create policy commessa_attivita_override_select on commessa_attivita_override
  for select using (is_whitelisted_email());

drop policy if exists commessa_attivita_override_write on commessa_attivita_override;
create policy commessa_attivita_override_write on commessa_attivita_override
  for insert with check (can_write());

create policy commessa_attivita_override_update on commessa_attivita_override
  for update using (can_write()) with check (can_write());

create policy commessa_attivita_override_delete on commessa_attivita_override
  for delete using (can_write());
