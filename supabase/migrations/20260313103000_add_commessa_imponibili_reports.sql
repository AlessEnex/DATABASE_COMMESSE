create table if not exists commessa_imponibili (
  commessa_id uuid primary key references commesse(id) on delete cascade,
  imponibile numeric(14,2) not null default 0,
  updated_by uuid references utenti(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_commessa_imponibili_updated_at on commessa_imponibili (updated_at desc);

drop trigger if exists trg_commessa_imponibili_updated_at on commessa_imponibili;
create trigger trg_commessa_imponibili_updated_at
before update on commessa_imponibili
for each row execute function set_updated_at();

alter table commessa_imponibili enable row level security;

drop policy if exists commessa_imponibili_select on commessa_imponibili;
create policy commessa_imponibili_select on commessa_imponibili
  for select using (
    exists (
      select 1
      from public.utenti u
      where u.id = auth.uid()
        and u.ruolo = 'admin'
    ) and is_whitelisted_email()
  );

drop policy if exists commessa_imponibili_insert on commessa_imponibili;
create policy commessa_imponibili_insert on commessa_imponibili
  for insert with check (
    exists (
      select 1
      from public.utenti u
      where u.id = auth.uid()
        and u.ruolo = 'admin'
    ) and is_whitelisted_email()
  );

drop policy if exists commessa_imponibili_update on commessa_imponibili;
create policy commessa_imponibili_update on commessa_imponibili
  for update using (
    exists (
      select 1
      from public.utenti u
      where u.id = auth.uid()
        and u.ruolo = 'admin'
    ) and is_whitelisted_email()
  ) with check (
    exists (
      select 1
      from public.utenti u
      where u.id = auth.uid()
        and u.ruolo = 'admin'
    ) and is_whitelisted_email()
  );

drop policy if exists commessa_imponibili_delete on commessa_imponibili;
create policy commessa_imponibili_delete on commessa_imponibili
  for delete using (
    exists (
      select 1
      from public.utenti u
      where u.id = auth.uid()
        and u.ruolo = 'admin'
    ) and is_whitelisted_email()
  );
