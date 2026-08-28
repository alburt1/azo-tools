# AZO Documentgenerator

Webtool om de standaardbrieven en -verslagen van AZO in te vullen en als correct
opgemaakt Word-document te downloaden. Alles werkt lokaal, in de browser — er is
geen internetverbinding of server nodig om de tool te gebruiken.

## Snelste manier om te openen (aangeraden)

Geen VS Code, geen Live Server nodig:

1. Pak de zip **volledig uit** naar een map op je computer.
2. Dubbelklik op **`index.html`**.
3. Klaar — dit opent in je standaardbrowser en alles werkt.

Dit werkt omdat de tool geen `fetch()`, geen API-calls en geen ES-modules gebruikt.
Alles zit in gewone `<script src="...">`-tags, en die werken perfect via het
`file://`-protocol, zonder server.

## Openen via VS Code + Live Server

Als je toch liever met Live Server werkt (bv. voor verdere ontwikkeling), volg
dan deze volgorde exact — de meest voorkomende fout is dat Live Server een **los
bestand** probeert te serveren in plaats van de **map**:

1. **File → Open Folder…** en kies de uitgepakte map (dus niet "Open File").
   Je moet de mappenstructuur (`index.html`, `shared/`, `assets/`, …) links in
   de Explorer zien staan.
2. Rechtsklik op `index.html` in de Explorer-sidebar → **"Open with Live Server"**.
   *(Of: klik op de "Go Live"-knop rechtsonder in de statusbalk.)*
3. Je browser opent automatisch op `http://127.0.0.1:5501/index.html`.

Dit project bevat een `.vscode/settings.json` die de poort en de rootmap al
voor je instelt, zodat dit meteen zou moeten werken.

### Live Server lukt nog steeds niet? Checklist

| Symptoom | Oorzaak & oplossing |
|---|---|
| Geen "Open with Live Server" in het rechtsklikmenu | De extensie is niet geïnstalleerd. Ga naar Extensions (Ctrl+Shift+X), zoek **"Live Server"** van *Ritwick Dey* en installeer. |
| Geen "Go Live" knop rechtsonder | Je hebt een los bestand geopend i.p.v. een map. Sluit het bestand, gebruik **File → Open Folder…**. |
| "Port 5501 is already in use" | Een andere Live Server-sessie draait nog. Klik rechtsonder op **"Port: 5501"** → *Stop Server*, en start opnieuw. Of wijzig de poort in `.vscode/settings.json`. |
| Pagina laadt maar toont een bestandslijst i.p.v. de tool | Je bent op de verkeerde map gestart. Zorg dat `index.html` in de **root** van de geopende map staat (dat is hier het geval). |
| Extensie start wel, maar niets gebeurt | Herlaad VS Code volledig (Ctrl+Shift+P → "Reload Window") en probeer opnieuw. |

## Alternatief: eender welke andere lokale server

Werkt ook prima, voor wie geen VS Code gebruikt:

```bash
# Python (meestal al geïnstalleerd)
python3 -m http.server 8000
# → open http://localhost:8000

# Node (als je npx hebt)
npx serve .
```

## Projectstructuur

```
index.html                              → overzichtspagina met zoekfunctie
brief-selectie-positief.html            → 1 pagina per documenttype
brief-selectie-negatief.html
brief-selectie-gemengd.html
brief-uitnodiging-onderhandeling.html
brief-gunning.html
brief-niet-gunning.html
brief-sluiting.html
verslag-selectie.html
verslag-gunning.html
selectieleidraad.html
shared/style.css                        → gedeeld design system
shared/engine.js                        → documentdata, formulier- en exportlogica
assets/azo_full.png                     → volledig logo (voor documenten & hero)
assets/azo_icon.png                     → bergicoon (voor titelbalk & voettekst)
```

Elke documentpagina toont links een gegroepeerd invulformulier en rechts een
live A4-voorbeeld dat **exact** overeenkomt met het Word-document dat je
downloadt via "Genereer Word-document" — je ziet dus tijdens het invullen al
hoe de brief er definitief zal uitzien.

## Een document toevoegen of aanpassen

Alle documentdefinities (velden, secties, brief-tekst) staan centraal in
`shared/engine.js`, in het object `buildDocs()`. Een nieuwe pagina toevoegen:

1. Voeg een nieuw item toe aan `buildDocs()` met `sections` (velden) en een
   `render(d)`-functie die de brief-HTML teruggeeft.
2. Voeg een regel toe aan de `PAGES`-lijst bovenaan `engine.js`.
3. Kopieer één van de bestaande `.html`-paginabestanden, pas de titel aan en
   verander `initPage('...')` naar de nieuwe document-id.

De titelbalk, de zoekfunctie op de indexpagina en de Word-export werken dan
automatisch mee, zonder verdere aanpassingen.
