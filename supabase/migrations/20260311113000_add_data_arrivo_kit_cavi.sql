-- Add data_arrivo_kit_cavi and extend planner RPC

alter table if exists public.commesse
  add column if not exists data_arrivo_kit_cavi date;

create or replace function update_commessa_planner_dates(
  p_commessa_id uuid,
  p_data_ordine_telaio date,
  p_data_consegna_macchina date,
  p_data_arrivo_kit_cavi date
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
      data_arrivo_kit_cavi = p_data_arrivo_kit_cavi
  where id = p_commessa_id;
end;
$$;

grant execute on function update_commessa_planner_dates(uuid, date, date, date) to authenticated;
