create or replace function current_risorsa_id()
returns smallint
language sql
stable
security definer
set search_path = public, auth
as $$
  select r.id
  from public.risorse r
  where r.utente_id = auth.uid()
  limit 1;
$$;

create or replace function can_edit_attivita(target_risorsa_id smallint)
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
      and (
        u.ruolo in ('admin', 'responsabile')
        or (u.ruolo = 'operatore' and current_risorsa_id() = target_risorsa_id)
      )
  ) and is_whitelisted_email();
$$;

drop policy if exists attivita_write on attivita;
create policy attivita_write on attivita
  for insert with check (can_edit_attivita(risorsa_id));
drop policy if exists attivita_update on attivita;
create policy attivita_update on attivita
  for update using (can_edit_attivita(risorsa_id)) with check (can_edit_attivita(risorsa_id));
drop policy if exists attivita_delete on attivita;
create policy attivita_delete on attivita
  for delete using (can_edit_attivita(risorsa_id));
