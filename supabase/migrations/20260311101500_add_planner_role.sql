-- Add planner role and limited commessa update function

do $$
begin
  if not exists (
    select 1
    from pg_type t
    join pg_enum e on e.enumtypid = t.oid
    where t.typname = 'ruolo_utente' and e.enumlabel = 'planner'
  ) then
    alter type ruolo_utente add value 'planner';
  end if;
end $$;

insert into permessi_ruolo (ruolo, can_move_matrix, can_delete_matrix, can_move_gantt, can_delete_gantt)
values ('planner', false, false, true, false)
on conflict (ruolo) do nothing;

create or replace function update_commessa_planner_dates(
  p_commessa_id uuid,
  p_data_ordine_telaio date,
  p_data_consegna_macchina date
)
returns void
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  v_role text;
begin
  select u.ruolo::text into v_role
  from public.utenti u
  where u.id = auth.uid();

  if v_role is null or v_role not in ('planner', 'admin', 'responsabile') then
    raise exception 'unauthorized';
  end if;

  update public.commesse
  set data_ordine_telaio = p_data_ordine_telaio,
      data_consegna_macchina = p_data_consegna_macchina
  where id = p_commessa_id;
end;
$$;

grant execute on function update_commessa_planner_dates(uuid, date, date) to authenticated;
