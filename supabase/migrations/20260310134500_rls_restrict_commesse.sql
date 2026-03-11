create or replace function can_manage_commesse()
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
      and u.ruolo in ('admin', 'responsabile')
  ) and is_whitelisted_email();
$$;

create or replace function can_manage_risorse()
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
      and u.ruolo in ('admin', 'responsabile')
  ) and is_whitelisted_email();
$$;

drop policy if exists commesse_write on commesse;
create policy commesse_write on commesse
  for insert with check (can_manage_commesse());
drop policy if exists commesse_update on commesse;
create policy commesse_update on commesse
  for update using (can_manage_commesse()) with check (can_manage_commesse());
drop policy if exists commesse_delete on commesse;
create policy commesse_delete on commesse
  for delete using (can_manage_commesse());

drop policy if exists commessa_schede_write on commessa_schede;
create policy commessa_schede_write on commessa_schede
  for insert with check (can_manage_commesse());
drop policy if exists commessa_schede_update on commessa_schede;
create policy commessa_schede_update on commessa_schede
  for update using (can_manage_commesse()) with check (can_manage_commesse());
drop policy if exists commessa_schede_delete on commessa_schede;
create policy commessa_schede_delete on commessa_schede
  for delete using (can_manage_commesse());

drop policy if exists commesse_reparti_write on commesse_reparti;
create policy commesse_reparti_write on commesse_reparti
  for insert with check (can_manage_commesse());
drop policy if exists commesse_reparti_update on commesse_reparti;
create policy commesse_reparti_update on commesse_reparti
  for update using (can_manage_commesse()) with check (can_manage_commesse());
drop policy if exists commesse_reparti_delete on commesse_reparti;
create policy commesse_reparti_delete on commesse_reparti
  for delete using (can_manage_commesse());

drop policy if exists reparti_write on reparti;
create policy reparti_write on reparti
  for insert with check (can_manage_risorse());
drop policy if exists reparti_update on reparti;
create policy reparti_update on reparti
  for update using (can_manage_risorse()) with check (can_manage_risorse());
drop policy if exists reparti_delete on reparti;
create policy reparti_delete on reparti
  for delete using (can_manage_risorse());

drop policy if exists risorse_write on risorse;
create policy risorse_write on risorse
  for insert with check (can_manage_risorse());
drop policy if exists risorse_update on risorse;
create policy risorse_update on risorse
  for update using (can_manage_risorse()) with check (can_manage_risorse());
drop policy if exists risorse_delete on risorse;
create policy risorse_delete on risorse
  for delete using (can_manage_risorse());

drop policy if exists utenti_write on utenti;
create policy utenti_write on utenti
  for insert with check (can_manage_risorse());
drop policy if exists utenti_update on utenti;
create policy utenti_update on utenti
  for update using (can_manage_risorse()) with check (can_manage_risorse());
drop policy if exists utenti_delete on utenti;
create policy utenti_delete on utenti
  for delete using (can_manage_risorse());

drop policy if exists assegnazioni_write on assegnazioni;
create policy assegnazioni_write on assegnazioni
  for insert with check (can_manage_commesse());
drop policy if exists assegnazioni_update on assegnazioni;
create policy assegnazioni_update on assegnazioni
  for update using (can_manage_commesse()) with check (can_manage_commesse());
drop policy if exists assegnazioni_delete on assegnazioni;
create policy assegnazioni_delete on assegnazioni
  for delete using (can_manage_commesse());
