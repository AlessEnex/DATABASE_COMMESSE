-- Extend planner RPC to handle data_prelievo

create or replace function update_commessa_planner_dates(
  p_commessa_id uuid,
  p_data_ordine_telaio date,
  p_data_consegna_macchina date,
  p_data_arrivo_kit_cavi date,
  p_data_prelievo date
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
      data_consegna_macchina = p_data_consegna_macchina,
      data_arrivo_kit_cavi = p_data_arrivo_kit_cavi,
      data_prelievo = p_data_prelievo
  where id = p_commessa_id;
end;
$$;

grant execute on function update_commessa_planner_dates(uuid, date, date, date, date) to authenticated;
