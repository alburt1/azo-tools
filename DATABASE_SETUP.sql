-- ============================================================
-- AZO Tools — Database schema (Fase 2, 3 en 4)
-- ============================================================
-- Voer dit volledige script EENMALIG uit in Supabase:
-- Dashboard -> SQL Editor -> New query -> plak dit -> Run
-- ============================================================


-- ------------------------------------------------------------
-- FASE 2 — Koepellijst centraal beheren
-- ------------------------------------------------------------
create table if not exists koepellijst_entries (
  id bigint generated always as identity primary key,
  naam text not null,
  gemeente text,
  bron text,
  created_at timestamptz not null default now()
);

alter table koepellijst_entries enable row level security;

-- Elke ingelogde gebruiker mag de lijst lezen (nodig voor de Ledenlijst Samenvoeger-tool).
create policy "Ingelogde gebruikers kunnen koepellijst lezen"
  on koepellijst_entries for select
  to authenticated
  using (true);

-- Elke ingelogde gebruiker mag de lijst vernieuwen (uploaden/vervangen).
-- Wil je dit later beperken tot één beheerder, dan passen we deze regel aan.
create policy "Ingelogde gebruikers kunnen koepellijst bijwerken"
  on koepellijst_entries for insert
  to authenticated
  with check (true);

create policy "Ingelogde gebruikers kunnen koepellijst verwijderen"
  on koepellijst_entries for delete
  to authenticated
  using (true);

create index if not exists idx_koepellijst_naam on koepellijst_entries (naam);


-- ------------------------------------------------------------
-- FASE 3 — Documentenhistoriek
-- ------------------------------------------------------------
create table if not exists document_history (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) default auth.uid(),
  user_email text not null,
  doc_type text not null,           -- bv. "selectieleidraad", "bestek-minicompetitie"
  doc_label text,                   -- leesbare naam, bv. "Selectieleidraad"
  project_ref text,                 -- de ingevulde projectreferentie
  opdracht_titel text,              -- de ingevulde opdrachttitel, indien aanwezig
  created_at timestamptz not null default now()
);

alter table document_history enable row level security;

-- Iedereen die ingelogd is ziet de volledige historiek (gedeeld team-overzicht).
create policy "Ingelogde gebruikers kunnen documentenhistoriek lezen"
  on document_history for select
  to authenticated
  using (true);

-- Elke gebruiker mag enkel een rij toevoegen met zijn eigen user_id (kan niet
-- doen alsof iemand anders een document gegenereerd heeft).
create policy "Gebruikers kunnen eigen documenten loggen"
  on document_history for insert
  to authenticated
  with check (auth.uid() = user_id);

create index if not exists idx_dochist_created on document_history (created_at desc);


-- ------------------------------------------------------------
-- FASE 4 — Bewaarde samenvoegingen (Ledenlijst Samenvoeger)
-- ------------------------------------------------------------
create table if not exists merged_lists (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) default auth.uid(),
  user_email text not null,
  azo_filename text,                -- naam van het geuploade AZO-bestand
  total_count int,
  unique_added_count int,
  auto_merged_count int,
  reviewed_count int,
  threshold int,                    -- ingestelde gevoeligheid (%)
  result_data jsonb not null,       -- de volledige samengevoegde lijst, om later te herbekijken/downloaden
  created_at timestamptz not null default now()
);

alter table merged_lists enable row level security;

create policy "Ingelogde gebruikers kunnen samenvoegingen lezen"
  on merged_lists for select
  to authenticated
  using (true);

create policy "Gebruikers kunnen eigen samenvoegingen bewaren"
  on merged_lists for insert
  to authenticated
  with check (auth.uid() = user_id);

create index if not exists idx_mergedlists_created on merged_lists (created_at desc);

-- ============================================================
-- Klaar. Dit script mag je maar één keer uitvoeren; nogmaals
-- uitvoeren is onschadelijk dankzij "if not exists", maar de
-- policies zouden dan een foutmelding geven dat ze al bestaan
-- (dat is normaal en onschuldig).
-- ============================================================
