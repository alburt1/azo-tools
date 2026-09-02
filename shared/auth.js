/* ============================================================
   AZO Tools — Supabase-configuratie en toegangsbeveiliging
   ============================================================
   Dit bestand:
   1. Maakt de verbinding met de Supabase-backend.
   2. Controleert bij elke pagina of er een geldige login-sessie is.
      Geen sessie? -> stuurt door naar login.html.
   3. Voorziet een uitlog-knop-helper.

   Let op: de "publishable key" hieronder is bewust bedoeld om
   publiek/in browsercode te staan (vergelijkbaar met een API-sleutel
   voor een mobiele app) — de eigenlijke beveiliging gebeurt via
   Row Level Security-regels aan de databasekant, niet door deze
   sleutel geheim te houden.
   ============================================================ */

const SUPABASE_URL = "https://gjsoepafdujpyclicini.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_T-ZrRD4W3vxjofab7RkpUw_gm-ltZYl";

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);

// Bepaalt het relatieve pad naar login.html, afhankelijk van hoe diep de
// huidige pagina in de mappenstructuur zit (root, of één niveau dieper
// zoals Documentgenerator/ of Ledenlijst-samenvoeger/).
function _loginPath() {
  const depth = window.__AZO_PAGE_DEPTH__ || 0;
  return depth === 0 ? "login.html" : "../login.html";
}

// Controleert of er een ingelogde gebruiker is. Zo niet: stuur door naar
// de inlogpagina en onthoud waar de gebruiker vandaan kwam, zodat we na
// het inloggen daar weer naartoe kunnen sturen.
async function requireAuth() {
  const { data: { session } } = await supabaseClient.auth.getSession();
  if (!session) {
    const redirectTo = encodeURIComponent(window.location.pathname);
    window.location.replace(_loginPath() + "?redirect=" + redirectTo);
    return null;
  }
  return session;
}

async function signOutAndRedirect() {
  await supabaseClient.auth.signOut();
  window.location.href = _loginPath();
}

// Vult elk element met data-user-email="" in met het e-mailadres van de
// ingelogde gebruiker. Gebruikt een MutationObserver omdat de titelbalk op
// sommige pagina's pas ná dit script wordt ingeladen (bv. Documentgenerator,
// waar de titelbalk via JS wordt geïnjecteerd) — zo werkt het altijd,
// ongeacht de volgorde waarin dingen op de pagina verschijnen.
let _azoSession = null;

function _fillUserEmails() {
  if (!_azoSession) return;
  document.querySelectorAll("[data-user-email]:not([data-filled])").forEach(el => {
    el.textContent = _azoSession.user.email;
    el.setAttribute("data-filled", "true");
  });
}

function wireAuthUI(session) {
  _azoSession = session;
  _fillUserEmails();
  const observer = new MutationObserver(_fillUserEmails);
  observer.observe(document.documentElement, { childList: true, subtree: true });
}

// Uitlog-knop: event delegation op document zelf in plaats van de knop
// rechtstreeks te zoeken bij het laden — zo werkt de knop ook als hij pas
// later aan de pagina wordt toegevoegd (zelfde reden als hierboven).
document.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-signout-btn]");
  if (btn) {
    e.preventDefault();
    signOutAndRedirect();
  }
});

// Fase 3 — Documentenhistoriek: logt één rij per gegenereerd document naar
// de gedeelde database. Werkt alleen als er een ingelogde sessie is (die
// hebben we altijd, want elke pagina zit al achter requireAuth()); faalt
// stil bij problemen zodat het downloaden van het document zelf nooit
// geblokkeerd wordt door een logging-fout.
async function logDocumentToHistory({docType, docLabel, projectRef, opdrachtTitel}){
  try{
    const { data: { session } } = await supabaseClient.auth.getSession();
    if(!session) return;
    await supabaseClient.from('document_history').insert({
      user_id: session.user.id,
      user_email: session.user.email,
      doc_type: docType,
      doc_label: docLabel || docType,
      project_ref: projectRef || null,
      opdracht_titel: opdrachtTitel || null,
    });
  } catch(e){
    // stil falen — loggen mag de eigenlijke download nooit blokkeren
  }
}

// Fase 4 — Bewaarde samenvoegingen: slaat het resultaat van een Ledenlijst-
// samenvoeging op in de database, zodat die later terug te vinden is.
async function saveMergedListToHistory({azoFilename, totalCount, uniqueAddedCount, autoMergedCount, reviewedCount, threshold, resultData}){
  try{
    const { data: { session } } = await supabaseClient.auth.getSession();
    if(!session) return null;
    const { data, error } = await supabaseClient.from('merged_lists').insert({
      user_id: session.user.id,
      user_email: session.user.email,
      azo_filename: azoFilename || null,
      total_count: totalCount ?? null,
      unique_added_count: uniqueAddedCount ?? null,
      auto_merged_count: autoMergedCount ?? null,
      reviewed_count: reviewedCount ?? null,
      threshold: threshold ?? null,
      result_data: resultData,
    }).select('id').single();
    if(error) return null;
    return data;
  } catch(e){
    return null;
  }
}
