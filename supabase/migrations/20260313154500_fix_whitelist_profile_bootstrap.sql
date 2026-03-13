-- Fix whitelist/profile bootstrap with normalized email matching
-- 1) Compare emails case-insensitively and trimming spaces
-- 2) Backfill missing rows in public.utenti for already-existing auth.users

create or replace function is_whitelisted_email()
returns boolean
language sql
stable
security definer
set search_path = public, auth
as $$
  select exists (
    select 1
    from public.whitelist_email w
    where lower(trim(w.email)) = lower(trim(auth.jwt() ->> 'email'))
      and w.attiva = true
  );
$$;

create or replace function handle_new_user()
returns trigger as $$
declare
  wl record;
begin
  select * into wl
  from public.whitelist_email w
  where lower(trim(w.email)) = lower(trim(new.email))
    and w.attiva = true
  limit 1;

  if found then
    insert into public.utenti (id, email, ruolo, reparto_id, nome)
    values (
      new.id,
      new.email,
      wl.ruolo_predefinito,
      wl.reparto_id_predefinito,
      nullif(new.raw_user_meta_data ->> 'full_name', '')
    )
    on conflict (id) do update
      set email = excluded.email,
          ruolo = coalesce(public.utenti.ruolo, excluded.ruolo),
          reparto_id = coalesce(public.utenti.reparto_id, excluded.reparto_id),
          nome = coalesce(public.utenti.nome, excluded.nome);
  end if;

  return new;
end;
$$ language plpgsql security definer set search_path = public, auth;

insert into public.utenti (id, email, ruolo, reparto_id, nome)
select
  au.id,
  au.email,
  coalesce(w.ruolo_predefinito, 'viewer'::ruolo_utente),
  w.reparto_id_predefinito,
  nullif(au.raw_user_meta_data ->> 'full_name', '')
from auth.users au
join public.whitelist_email w
  on lower(trim(w.email)) = lower(trim(au.email))
 and w.attiva = true
left join public.utenti u
  on u.id = au.id
where u.id is null;
