create or replace function can_edit_attivita_row(
  target_attivita_id uuid,
  target_risorsa_id smallint,
  target_assegnato_a uuid
)
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
        or (
          u.ruolo = 'operatore'
          and (
            current_risorsa_id() = target_risorsa_id
            or target_assegnato_a = auth.uid()
          )
        )
      )
  ) and is_whitelisted_email();
$$;

drop policy if exists attivita_update on attivita;
create policy attivita_update on attivita
  for update using (can_edit_attivita_row(id, risorsa_id, assegnato_a))
  with check (can_edit_attivita_row(id, risorsa_id, assegnato_a));

drop policy if exists attivita_delete on attivita;
create policy attivita_delete on attivita
  for delete using (can_edit_attivita_row(id, risorsa_id, assegnato_a));
