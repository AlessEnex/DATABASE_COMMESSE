-- Structured machine type on commesse + readable motore tipologie for whitelisted users

alter table commesse
  add column if not exists tipo_macchina text,
  add column if not exists variante_macchina text;

update commesse
set tipo_macchina = 'Altro tipo'
where coalesce(btrim(tipo_macchina), '') = '';

update commesse
set variante_macchina = 'standard'
where coalesce(btrim(variante_macchina), '') = '';

alter table commesse
  alter column tipo_macchina set default 'Altro tipo',
  alter column variante_macchina set default 'standard',
  alter column tipo_macchina set not null,
  alter column variante_macchina set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'commesse_tipo_macchina_not_empty'
      and conrelid = 'commesse'::regclass
  ) then
    alter table commesse
      add constraint commesse_tipo_macchina_not_empty
      check (btrim(tipo_macchina) <> '');
  end if;

  if not exists (
    select 1
    from pg_constraint
    where conname = 'commesse_variante_macchina_not_empty'
      and conrelid = 'commesse'::regclass
  ) then
    alter table commesse
      add constraint commesse_variante_macchina_not_empty
      check (btrim(variante_macchina) <> '');
  end if;
end $$;

create index if not exists idx_commesse_tipo_macchina on commesse (tipo_macchina);
create index if not exists idx_commesse_tipo_variante on commesse (tipo_macchina, variante_macchina);

drop view if exists v_commesse;

create view v_commesse as
select
  c.*,
  coalesce(
    jsonb_agg(
      jsonb_build_object(
        'reparto_id', r.id,
        'reparto_nome', r.nome,
        'stato', cr.stato,
        'note_reparto', cr.note_reparto,
        'updated_at', cr.updated_at
      )
      order by r.nome
    ) filter (where r.id is not null),
    '[]'::jsonb
  ) as reparti
from commesse c
left join commesse_reparti cr on cr.commessa_id = c.id
left join reparti r on r.id = cr.reparto_id
group by c.id;

grant select on v_commesse to authenticated;

-- Keep write restricted to admin, but allow read for all whitelisted users.
drop policy if exists motore_tipologie_select on motore_tipologie;
create policy motore_tipologie_select on motore_tipologie
  for select using (is_whitelisted_email());
