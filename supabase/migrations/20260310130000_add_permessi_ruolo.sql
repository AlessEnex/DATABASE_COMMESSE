create table if not exists permessi_ruolo (
  ruolo ruolo_utente primary key,
  can_move_matrix boolean not null default true,
  can_delete_matrix boolean not null default true,
  can_move_gantt boolean not null default true,
  can_delete_gantt boolean not null default true,
  updated_at timestamptz not null default now()
);

alter table permessi_ruolo enable row level security;

drop policy if exists permessi_ruolo_select on permessi_ruolo;
create policy permessi_ruolo_select on permessi_ruolo
  for select using (is_whitelisted_email());
drop policy if exists permessi_ruolo_admin_write on permessi_ruolo;
create policy permessi_ruolo_admin_write on permessi_ruolo
  for insert with check (is_admin());
create policy permessi_ruolo_admin_update on permessi_ruolo
  for update using (is_admin()) with check (is_admin());
create policy permessi_ruolo_admin_delete on permessi_ruolo
  for delete using (is_admin());

insert into permessi_ruolo (ruolo, can_move_matrix, can_delete_matrix, can_move_gantt, can_delete_gantt)
values
  ('admin', true, true, true, true),
  ('responsabile', true, true, true, true),
  ('operatore', false, false, false, false),
  ('viewer', false, false, false, false)
on conflict (ruolo) do nothing;
