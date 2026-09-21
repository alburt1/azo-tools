-- ============================================================
-- AZO Tools — Documentenhistoriek: volledige data bewaren
-- ============================================================
-- Voer dit EENMALIG uit, na de eerdere setup-scripts.
-- Voegt een kolom toe waarin de volledige ingevulde data van elk
-- gegenereerd document bewaard wordt, zodat je het later kan
-- heropenen (niet enkel titel/datum zoals voorheen).
-- ============================================================

alter table document_history
  add column if not exists form_data jsonb;

alter table document_history
  add column if not exists file_name text;

-- ============================================================
-- Klaar. Bestaande, oudere rijen krijgen gewoon NULL voor deze
-- nieuwe kolom — die blijven gewoon zichtbaar in de historiek,
-- enkel zonder "Openen"-knop (want geen data om te herladen).
-- ============================================================
