-- ============================================================
-- WERKBEGROTING APP — Supabase schema
-- Uitvoeren in Supabase SQL Editor
-- ============================================================

-- Projecten tabel: één rij per project
create table if not exists wb_projecten (
  id            uuid primary key default gen_random_uuid(),
  naam          text not null,
  projectnummer text,
  opdrachtgever text,
  aangemaakt_op timestamptz default now(),
  bijgewerkt_op timestamptz default now(),

  -- Volledig projectdossier als JSON blobs
  posts         jsonb default '[]',      -- verrijkte besteksposten
  staartkosten  jsonb default '[]',      -- staartkosten met pct + interneKosten
  taken         jsonb default '[]',      -- planning werkpakketten
  koppelingen   jsonb default '{}'       -- {taskId: [{hoofdstuk, percentage}]}
);

-- Automatisch bijgewerkt_op updaten
create or replace function wb_update_bijgewerkt_op()
returns trigger language plpgsql as $$
begin
  new.bijgewerkt_op = now();
  return new;
end;
$$;

create trigger wb_projecten_bijgewerkt
  before update on wb_projecten
  for each row execute function wb_update_bijgewerkt_op();

-- Row Level Security: iedereen mag lezen/schrijven (anon key)
alter table wb_projecten enable row level security;

create policy "Publiek lezen" on wb_projecten
  for select using (true);

create policy "Publiek aanmaken" on wb_projecten
  for insert with check (true);

create policy "Publiek bijwerken" on wb_projecten
  for update using (true);

create policy "Publiek verwijderen" on wb_projecten
  for delete using (true);
