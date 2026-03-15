-- Motore planning data model:
-- - Ensure core motore tables exist (tipologie/varianti/regole)
-- - Add canvas tables for planner inputs (priorita, flussi, skill/capacita)
-- - Restrict access to admin users via RLS

create table if not exists motore_tipologie (
  id uuid primary key default gen_random_uuid(),
  nome text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists motore_varianti (
  id uuid primary key default gen_random_uuid(),
  tipologia_id uuid not null references motore_tipologie(id) on delete cascade,
  nome text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists motore_regole (
  variante_id uuid not null references motore_varianti(id) on delete cascade,
  attivita_titolo text not null,
  ore numeric(8,2),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  primary key (variante_id, attivita_titolo)
);

alter table motore_tipologie add column if not exists created_at timestamptz not null default now();
alter table motore_tipologie add column if not exists updated_at timestamptz not null default now();

alter table motore_varianti add column if not exists created_at timestamptz not null default now();
alter table motore_varianti add column if not exists updated_at timestamptz not null default now();

alter table motore_regole add column if not exists created_at timestamptz not null default now();
alter table motore_regole add column if not exists updated_at timestamptz not null default now();

create index if not exists idx_motore_varianti_tipologia on motore_varianti (tipologia_id);
create index if not exists idx_motore_regole_variante on motore_regole (variante_id);

drop trigger if exists trg_motore_tipologie_updated_at on motore_tipologie;
create trigger trg_motore_tipologie_updated_at
before update on motore_tipologie
for each row execute function set_updated_at();

drop trigger if exists trg_motore_varianti_updated_at on motore_varianti;
create trigger trg_motore_varianti_updated_at
before update on motore_varianti
for each row execute function set_updated_at();

drop trigger if exists trg_motore_regole_updated_at on motore_regole;
create trigger trg_motore_regole_updated_at
before update on motore_regole
for each row execute function set_updated_at();

create table if not exists motore_canvas_config (
  variante_id uuid primary key references motore_varianti(id) on delete cascade,
  priority_due integer not null default 40 check (priority_due >= 0 and priority_due <= 100),
  priority_delay integer not null default 30 check (priority_delay >= 0 and priority_delay <= 100),
  priority_strategic integer not null default 20 check (priority_strategic >= 0 and priority_strategic <= 100),
  priority_complexity integer not null default 10 check (priority_complexity >= 0 and priority_complexity <= 100),
  rule_hard_skills boolean not null default true,
  rule_prefer_continuity boolean not null default true,
  rule_balance_load boolean not null default true,
  planner_notes text,
  updated_by uuid references utenti(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists motore_canvas_flow (
  variante_id uuid not null references motore_varianti(id) on delete cascade,
  attivita_titolo text not null,
  depends_on text,
  constraint_level text not null default 'none' check (constraint_level in ('hard', 'soft', 'none')),
  allow_parallel boolean not null default false,
  notes text,
  updated_by uuid references utenti(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  primary key (variante_id, attivita_titolo)
);

create table if not exists motore_canvas_skill (
  variante_id uuid not null references motore_varianti(id) on delete cascade,
  risorsa_id smallint not null references risorse(id) on delete cascade,
  enabled boolean not null default false,
  skill_level text not null default 'medio' check (skill_level in ('base', 'medio', 'esperto')),
  allowed_activities text not null default '',
  max_hours_per_day numeric(5,2) not null default 8 check (max_hours_per_day >= 0 and max_hours_per_day <= 24),
  updated_by uuid references utenti(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  primary key (variante_id, risorsa_id)
);

create index if not exists idx_motore_canvas_flow_variante on motore_canvas_flow (variante_id);
create index if not exists idx_motore_canvas_skill_variante on motore_canvas_skill (variante_id);
create index if not exists idx_motore_canvas_skill_risorsa on motore_canvas_skill (risorsa_id);

drop trigger if exists trg_motore_canvas_config_updated_at on motore_canvas_config;
create trigger trg_motore_canvas_config_updated_at
before update on motore_canvas_config
for each row execute function set_updated_at();

drop trigger if exists trg_motore_canvas_flow_updated_at on motore_canvas_flow;
create trigger trg_motore_canvas_flow_updated_at
before update on motore_canvas_flow
for each row execute function set_updated_at();

drop trigger if exists trg_motore_canvas_skill_updated_at on motore_canvas_skill;
create trigger trg_motore_canvas_skill_updated_at
before update on motore_canvas_skill
for each row execute function set_updated_at();

create or replace function can_manage_motore()
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
      and u.ruolo = 'admin'
  ) and is_whitelisted_email();
$$;

alter table motore_tipologie enable row level security;
alter table motore_varianti enable row level security;
alter table motore_regole enable row level security;
alter table motore_canvas_config enable row level security;
alter table motore_canvas_flow enable row level security;
alter table motore_canvas_skill enable row level security;

drop policy if exists motore_tipologie_select on motore_tipologie;
create policy motore_tipologie_select on motore_tipologie
  for select using (can_manage_motore());
drop policy if exists motore_tipologie_insert on motore_tipologie;
create policy motore_tipologie_insert on motore_tipologie
  for insert with check (can_manage_motore());
drop policy if exists motore_tipologie_update on motore_tipologie;
create policy motore_tipologie_update on motore_tipologie
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_tipologie_delete on motore_tipologie;
create policy motore_tipologie_delete on motore_tipologie
  for delete using (can_manage_motore());

drop policy if exists motore_varianti_select on motore_varianti;
create policy motore_varianti_select on motore_varianti
  for select using (can_manage_motore());
drop policy if exists motore_varianti_insert on motore_varianti;
create policy motore_varianti_insert on motore_varianti
  for insert with check (can_manage_motore());
drop policy if exists motore_varianti_update on motore_varianti;
create policy motore_varianti_update on motore_varianti
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_varianti_delete on motore_varianti;
create policy motore_varianti_delete on motore_varianti
  for delete using (can_manage_motore());

drop policy if exists motore_regole_select on motore_regole;
create policy motore_regole_select on motore_regole
  for select using (can_manage_motore());
drop policy if exists motore_regole_insert on motore_regole;
create policy motore_regole_insert on motore_regole
  for insert with check (can_manage_motore());
drop policy if exists motore_regole_update on motore_regole;
create policy motore_regole_update on motore_regole
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_regole_delete on motore_regole;
create policy motore_regole_delete on motore_regole
  for delete using (can_manage_motore());

drop policy if exists motore_canvas_config_select on motore_canvas_config;
create policy motore_canvas_config_select on motore_canvas_config
  for select using (can_manage_motore());
drop policy if exists motore_canvas_config_insert on motore_canvas_config;
create policy motore_canvas_config_insert on motore_canvas_config
  for insert with check (can_manage_motore());
drop policy if exists motore_canvas_config_update on motore_canvas_config;
create policy motore_canvas_config_update on motore_canvas_config
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_canvas_config_delete on motore_canvas_config;
create policy motore_canvas_config_delete on motore_canvas_config
  for delete using (can_manage_motore());

drop policy if exists motore_canvas_flow_select on motore_canvas_flow;
create policy motore_canvas_flow_select on motore_canvas_flow
  for select using (can_manage_motore());
drop policy if exists motore_canvas_flow_insert on motore_canvas_flow;
create policy motore_canvas_flow_insert on motore_canvas_flow
  for insert with check (can_manage_motore());
drop policy if exists motore_canvas_flow_update on motore_canvas_flow;
create policy motore_canvas_flow_update on motore_canvas_flow
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_canvas_flow_delete on motore_canvas_flow;
create policy motore_canvas_flow_delete on motore_canvas_flow
  for delete using (can_manage_motore());

drop policy if exists motore_canvas_skill_select on motore_canvas_skill;
create policy motore_canvas_skill_select on motore_canvas_skill
  for select using (can_manage_motore());
drop policy if exists motore_canvas_skill_insert on motore_canvas_skill;
create policy motore_canvas_skill_insert on motore_canvas_skill
  for insert with check (can_manage_motore());
drop policy if exists motore_canvas_skill_update on motore_canvas_skill;
create policy motore_canvas_skill_update on motore_canvas_skill
  for update using (can_manage_motore()) with check (can_manage_motore());
drop policy if exists motore_canvas_skill_delete on motore_canvas_skill;
create policy motore_canvas_skill_delete on motore_canvas_skill
  for delete using (can_manage_motore());

with tipologie_seed(nome) as (
  values
    ('TAGO'),
    ('MINIBOOSTER SWISS'),
    ('DRAVA'),
    ('SENNA'),
    ('SENNA-XS'),
    ('NEVA'),
    ('ELBA'),
    ('LT-UNIT'),
    ('AH'),
    ('GH'),
    ('YUKON')
)
insert into motore_tipologie (nome)
select s.nome
from tipologie_seed s
where not exists (
  select 1 from motore_tipologie t where t.nome = s.nome
);

with varianti_seed(tipologia_nome, nome) as (
  values
    ('TAGO', 'standard'),
    ('MINIBOOSTER SWISS', 'standard'),
    ('DRAVA', 'standard'),
    ('SENNA', 'standard'),
    ('SENNA-XS', 'standard'),
    ('NEVA', 'standard'),
    ('NEVA', 'custom'),
    ('ELBA', 'standard'),
    ('ELBA', 'custom'),
    ('LT-UNIT', 'standard'),
    ('LT-UNIT', 'custom'),
    ('AH', 'standard'),
    ('GH', 'standard'),
    ('YUKON', 'standard'),
    ('YUKON', 'custom')
)
insert into motore_varianti (tipologia_id, nome)
select t.id, s.nome
from varianti_seed s
join motore_tipologie t on t.nome = s.tipologia_nome
where not exists (
  select 1
  from motore_varianti v
  where v.tipologia_id = t.id
    and v.nome = s.nome
);
