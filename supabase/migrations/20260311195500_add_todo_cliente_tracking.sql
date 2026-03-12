do $$
begin
  if not exists (select 1 from pg_type where typname = 'attivita_cliente_esito') then
    create type attivita_cliente_esito as enum ('in_attesa', 'confermato', 'silenzio_assenso', 'respinto');
  end if;
end $$;

create or replace function can_manage_commessa_attivita_cliente()
returns boolean
language sql
stable
security definer
set search_path = public, auth
as $$
  select exists (
    select 1
    from public.utenti u
    where u.id = auth.uid()
      and u.ruolo in ('admin', 'responsabile', 'planner')
  ) and is_whitelisted_email();
$$;

create table if not exists commessa_attivita_cliente (
  commessa_id uuid not null references commesse(id) on delete cascade,
  titolo text not null,
  reparto text not null default '',
  inviato_il date,
  scadenza_il date,
  esito attivita_cliente_esito not null default 'in_attesa',
  confermato_il date,
  updated_by uuid references utenti(id) on delete set null,
  updated_at timestamptz not null default now(),
  primary key (commessa_id, titolo, reparto),
  constraint commessa_attivita_cliente_due_after_sent check (
    scadenza_il is null or inviato_il is null or scadenza_il >= inviato_il
  ),
  constraint commessa_attivita_cliente_confirmed_date check (
    esito::text <> 'confermato' or confermato_il is not null
  )
);

create index if not exists idx_commessa_attivita_cliente_commessa on commessa_attivita_cliente (commessa_id);
create index if not exists idx_commessa_attivita_cliente_esito on commessa_attivita_cliente (esito);

alter table commessa_attivita_cliente enable row level security;

drop policy if exists commessa_attivita_cliente_select on commessa_attivita_cliente;
create policy commessa_attivita_cliente_select on commessa_attivita_cliente
  for select using (is_whitelisted_email());

drop policy if exists commessa_attivita_cliente_insert on commessa_attivita_cliente;
create policy commessa_attivita_cliente_insert on commessa_attivita_cliente
  for insert with check (can_manage_commessa_attivita_cliente());

drop policy if exists commessa_attivita_cliente_update on commessa_attivita_cliente;
create policy commessa_attivita_cliente_update on commessa_attivita_cliente
  for update using (can_manage_commessa_attivita_cliente()) with check (can_manage_commessa_attivita_cliente());

drop policy if exists commessa_attivita_cliente_delete on commessa_attivita_cliente;
create policy commessa_attivita_cliente_delete on commessa_attivita_cliente
  for delete using (can_manage_commessa_attivita_cliente());
