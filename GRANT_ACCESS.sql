-- ============================================================
-- AZO Tools — API-toegang geven aan de nieuwe tabellen
-- ============================================================
-- Nodig omdat "Automatically expose new tables" bewust uitstond bij het
-- aanmaken van het project (veiligere standaardinstelling) — daardoor
-- moet elke tabel apart toestemming krijgen om via de API bereikbaar
-- te zijn. Voer dit EENMALIG uit, na DATABASE_SETUP.sql.
-- ============================================================

grant usage on schema public to authenticated;

grant select, insert, delete on public.koepellijst_entries to authenticated;
grant select, insert on public.document_history to authenticated;
grant select, insert on public.merged_lists to authenticated;

-- Nodig zodat de "id"-kolommen (auto-increment) correct kunnen aangevuld worden bij nieuwe rijen.
grant usage, select on all sequences in schema public to authenticated;

-- ============================================================
-- Klaar. Herlaad daarna de Beheer-pagina — de foutmelding zou moeten
-- verdwijnen.
-- ============================================================
