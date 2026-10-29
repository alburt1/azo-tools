/* ======================================================================
   AZO Doorgeefluik — leest een Word-document (.docx) en zet de variabele
   gegevens door naar het volgende document in de keten:

     selectieleidraad  ->  bestek (4 soorten)  ->  gunningsverslag (2 soorten)

   Werkwijze: het .docx wordt in de browser opengemaakt (JSZip), de tekst
   wordt als blokken (alinea's en tabellen) ingelezen, en de variabele
   gegevens worden herkend aan de vaste zinnen van de AZO-sjablonen.
   Niets verlaat de browser. De uitkomst is altijd controleerbaar en
   aanpasbaar voordat ze naar de Documentgenerator gaat.
   ====================================================================== */
(function (root) {
  "use strict";

  /* ---------------------------------------------------------------
     1. .docx -> blokken
     --------------------------------------------------------------- */
  function parseXml(xml) {
    return new root.DOMParser().parseFromString(xml, "application/xml");
  }

  function paraText(p) {
    let out = "";
    (function walk(n) {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const nm = c.nodeName;
        if (nm === "w:t") out += c.textContent;
        else if (nm === "w:tab") out += " ";
        else if (nm === "w:br" || nm === "w:cr") out += "\n";
        else if (nm === "w:noBreakHyphen") out += "-";
        else if (nm === "w:delText" || nm === "w:instrText" || nm === "w:del") continue;
        else walk(c);
      }
    })(p);
    return out.replace(/ /g, " ");
  }

  function paraStyle(p) {
    for (let c = p.firstChild; c; c = c.nextSibling) {
      if (c.nodeName === "w:pPr") {
        for (let d = c.firstChild; d; d = d.nextSibling) {
          if (d.nodeName === "w:pStyle") return d.getAttribute("w:val") || "";
        }
      }
    }
    return "";
  }

  function cellParas(tc) {
    const out = [];
    (function walk(n) {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.nodeName === "w:p") out.push(paraText(c));
        else if (c.nodeName === "w:tbl" || c.nodeName === "w:tr" || c.nodeName === "w:tc" || c.nodeName === "w:sdt" || c.nodeName === "w:sdtContent") walk(c);
      }
    })(tc);
    return out;
  }

  function readBlocks(docXml) {
    const doc = parseXml(docXml);
    const body = doc.getElementsByTagName("w:body")[0];
    const blocks = [];
    (function walk(parent) {
      for (let c = parent.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.nodeName === "w:p") {
          blocks.push({ type: "p", text: paraText(c), style: paraStyle(c) });
        } else if (c.nodeName === "w:tbl") {
          const rows = [];
          for (let tr = c.firstChild; tr; tr = tr.nextSibling) {
            if (tr.nodeName !== "w:tr") continue;
            const cells = [];
            for (let tc = tr.firstChild; tc; tc = tc.nextSibling) {
              if (tc.nodeName === "w:tc") cells.push(cellParas(tc));
            }
            rows.push(cells);
          }
          blocks.push({ type: "tbl", rows });
        } else if (c.nodeName === "w:sdt") {
          for (let d = c.firstChild; d; d = d.nextSibling) if (d.nodeName === "w:sdtContent") walk(d);
        }
      }
    })(body);
    return blocks;
  }

  async function readDocx(arrayBuffer) {
    const zip = await root.JSZip.loadAsync(arrayBuffer);
    const f = zip.file("word/document.xml");
    if (!f) throw new Error("Dit lijkt geen Word-document (.docx) te zijn.");
    return readBlocks(await f.async("string"));
  }

  /* ---------------------------------------------------------------
     2. Kleine tekst-hulpjes
     --------------------------------------------------------------- */
  const clean = (s) => (s || "").replace(/[ \t]+/g, " ").replace(/\s*\n\s*/g, "\n").trim();
  const flat = (s) => (s || "").replace(/\s+/g, " ").trim();
  const unq = (s) => flat(s).replace(/^[‘’'"“”«]+|[‘’'"“”»]+$/g, "").trim();
  const pText = (b) => (b.type === "p" ? flat(b.text) : "");

  function findP(blocks, re, from) {
    for (let i = from || 0; i < blocks.length; i++) {
      if (blocks[i].type === "p" && re.test(flat(blocks[i].text))) return i;
    }
    return -1;
  }
  // Alle platte tekstregels van de eerste N blokken (voorblad zit vaak in tabellen)
  function topLines(blocks, max) {
    const lines = [];
    for (let i = 0; i < Math.min(blocks.length, max); i++) {
      const b = blocks[i];
      if (b.type === "p") { const t = flat(b.text); if (t) lines.push(t); }
      else b.rows.forEach(r => r.forEach(c => c.forEach(t => { t = flat(t); if (t) lines.push(t); })));
    }
    return lines;
  }
  function allText(blocks) {
    return blocks.map(b => b.type === "p" ? b.text : b.rows.map(r => r.map(c => c.join(" ")).join(" ")).join(" ")).join("\n");
  }

  const NL_GETAL = { een: 1, "één": 1, twee: 2, drie: 3, vier: 4, vijf: 5, zes: 6, zeven: 7, acht: 8, negen: 9, tien: 10 };

  // "1 en 2", "1, 2 en 4", "1 tot en met 3", "beide" -> [1,2]; "alle" -> null
  function parsePerceelNummers(s) {
    const t = (s || "").toLowerCase();
    if (/\b(beide|alle)\b/.test(t) && !/\d/.test(t)) return null;
    const nums = [];
    const re = /(\d+)\s*(?:tot en met|t\.e\.m\.|-|–)\s*(\d+)|(\d+)/g;
    let m;
    while ((m = re.exec(t))) {
      if (m[1]) for (let i = +m[1]; i <= +m[2]; i++) nums.push(i);
      else nums.push(+m[3]);
    }
    return [...new Set(nums)].sort((a, b) => a - b);
  }

  // "1.000.000,00 euro excl. btw voor perceel 1;" -> "1.000.000,00 euro"
  function cleanBedrag(s) {
    let t = flat(s).replace(/[;.]\s*$/, "");
    t = t.replace(/\s*[(\[]?\s*excl(usief|\.)?\s*\.?\s*btw\.?\s*[)\]]?/i, " ");
    t = t.replace(/\s+voor\s+(de\s+|het\s+)?perce(el|len).*$/i, "");
    return flat(t);
  }

  /* ---------------------------------------------------------------
     3. Selectieleidraad -> dossiergegevens
     --------------------------------------------------------------- */
  function parseCriteriaTables(blocks, startIdx, stopRe) {
    // groepen: [{label, percelen:[..]|null, rows:[{nr,naam,gewicht,details}]}]
    const groups = [];
    let label = "";
    for (let i = startIdx; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === "p") {
        const t = flat(b.text);
        if (stopRe && stopRe.test(t)) break;
        if (/^(voor|perceel)\b/i.test(t) && t.length < 140) label = t.replace(/[:：]\s*$/, "");
        continue;
      }
      // tabel: herken Nr./Beschrijving/Gewicht
      const head = (b.rows[0] || []).map(c => flat(c.join(" ")).toLowerCase());
      if (!(head[0] || "").startsWith("nr") || head.length < 2) continue;
      const rows = [];
      let totaal = null;
      b.rows.slice(1).forEach(r => {
        const nr = flat((r[0] || []).join(" "));
        const desc = (r[1] || []).map(clean).filter(Boolean);
        const gew = flat((r[2] || []).join(" "));
        if (/totaal/i.test(nr)) { totaal = gew || flat((r[1] || []).join(" ")); return; }
        if (!nr && !desc.length) return;
        if (!nr && desc.length) {
          // detailrij onder een criterium
          const last = rows[rows.length - 1];
          const det = desc.join("\n");
          if (last && !/^details worden nader gepreciseerd/i.test(det)) last.details = (last.details ? last.details + "\n" : "") + det;
          return;
        }
        rows.push({
          nr,
          naam: desc[0] || "",
          gewicht: gew,
          details: desc.slice(1).join("\n").replace(/^details worden nader gepreciseerd in het bijzonder bestek\.?$/i, "").trim(),
        });
      });
      groups.push({ label, percelen: label ? parsePerceelNummers(label) : null, rows, totaal });
      label = "";
    }
    return groups;
  }

  function parseSelectieleidraad(blocks) {
    const out = { fields: {}, percelen: [], criteria: [], meta: {} };
    const F = out.fields;
    const lines = topLines(blocks, 12);

    // Voorblad
    const iV = lines.findIndex(l => /^voor de raamovereenkomst voor/i.test(l) || /^voor de (opdracht|overeenkomst) voor/i.test(l));
    if (iV >= 0) F.type = lines[iV + 1] || "";
    const iM = lines.findIndex(l => /^met als voorwerp/i.test(l));
    if (iM >= 0) {
      F.titel = unq(lines[iM + 1] || "");
      const ref = lines.slice(iM + 2, iM + 4).find(l => /^ref/i.test(l));
      if (ref) F.referte = ref.replace(/^ref\.?\s*:?\s*/i, "").trim();
    }
    const proc = lines.find(l => /PROCEDURE/.test(l) && l === l.toUpperCase());
    if (proc) F.procedure = proc;

    // Contactpersoon
    const iC = findP(blocks, /^contactpersoon\s*:/i);
    if (iC >= 0) F.contact_naam = flat(blocks[iC].text).replace(/^contactpersoon\s*:\s*/i, "");
    const iE = findP(blocks, /^e-?mail\s*:/i);
    if (iE >= 0) F.contact_email = flat(blocks[iE].text).replace(/^e-?mail\s*:\s*/i, "").replace(/\s+@/g, "@");

    // Beschrijving van de opdracht
    const iB = findP(blocks, /^raamovereenkomst voor\b/i);
    if (iB >= 0) {
      const t = flat(blocks[iB].text);
      const m = t.match(/^raamovereenkomst voor\s+[‘'“"]?(.+?)[’'”"]?\s+met als referentie\s*:?\s*(.+)$/i);
      if (m) { F.titel_klein = unq(m[1]); if (!F.referte) F.referte = m[2].trim(); }
      else F.titel_klein = unq(t.replace(/^raamovereenkomst voor\s*/i, ""));
    }
    const iI = findP(blocks, /opgedeeld in volgende percelen|verdeeld in (verschillende )?percelen/i);
    if (iI >= 0) {
      F.percelen_inleiding = flat(blocks[iI].text);
      for (let i = iI + 1; i < blocks.length; i++) {
        const t = pText(blocks[i]);
        const m = t.match(/^perceel\s*(\d+)\s*[:\-–]\s*(.*)$/i);
        if (!m) { if (t) break; else continue; }
        out.percelen.push({ naam: unq(m[2]), max_bedrag: "" });
      }
    }

    const iL = findP(blocks, /^de looptijd van deze raamovereenkomst/i);
    if (iL >= 0) {
      const t = flat(blocks[iL].text);
      let m = t.match(/periode van\s+(.+?)(?:,\s*per perceel)?\.\s*aansluitend/i);
      if (m) F.looptijd = m[1].trim();
      m = t.match(/kan de opdracht(?:,\s*per perceel,)?\s+(.+?)\s+verlengd worden met een periode van telkens\s+(.+?)\s*,/i);
      if (m) { F.verlengingen = m[1].trim(); F.verlengperiode = m[2].trim(); }
    }
    const iD = findP(blocks, /^het doel van deze opdracht/i);
    if (iD >= 0) F.doel = flat(blocks[iD].text);

    const iMax = findP(blocks, /maximale bestelhoeveelheid onder de raamovereenkomst bedraagt/i);
    if (iMax >= 0) {
      const t = flat(blocks[iMax].text);
      const m = t.match(/bedraagt\s*:?\s*(.+?)(?:\.\s|$)/i);
      if (m && m[1] && /\d/.test(m[1])) F.bestelhoeveelheid = cleanBedrag(m[1]);
      for (let i = iMax + 1; i < blocks.length; i++) {
        const l = pText(blocks[i]);
        if (!l) continue;
        let mm = l.match(/^(.+?)\s+voor\s+(?:het\s+)?perce(?:el|len)\s+(\d+)/i);
        if (mm) { const idx = +mm[2] - 1; if (out.percelen[idx]) out.percelen[idx].max_bedrag = cleanBedrag(mm[1]); else out.percelen[idx] = { naam: "", max_bedrag: cleanBedrag(mm[1]) }; continue; }
        mm = l.match(/^perceel\s*(\d+)\s*[:\-–]\s*(.+)$/i);
        if (mm) { const idx = +mm[1] - 1; if (out.percelen[idx]) out.percelen[idx].max_bedrag = cleanBedrag(mm[2]); continue; }
        break;
      }
    }
    const iP = findP(blocks, /^(leveringsplaats|plaats van de leveringen|plaats van uitvoering)\s*:/i);
    if (iP >= 0) F.leveringsplaats = flat(blocks[iP].text).replace(/^[^:]+:\s*/, "");

    // Gunningscriteria
    const iG = findP(blocks, /volgende criteria zijn van toepassing/i);
    if (iG >= 0) out.criteria = parseCriteriaTables(blocks, iG + 1, /^aan elk criterium wordt een gewicht|^voorwaarden van het bestek|^aan elk criterium werd/i);

    // Raming en bijdrage
    const iR = findP(blocks, /^de maximale raming van deze opdracht bedraagt/i);
    if (iR >= 0) {
      const m = flat(blocks[iR].text).match(/bedraagt\s+(.+?)\.\s+de raming/i);
      if (m) F.raming = cleanBedrag(m[1]);
    }
    const txt = allText(blocks);
    const mb = txt.match(/administratieve bijdrage leveren ten belope van\s*([\d.,]+)\s*%/i);
    if (mb) F.bijdrage = mb[1];

    return out;
  }

  /* ---------------------------------------------------------------
     4. Bestek -> dossiergegevens (voor de stap naar het gunningsverslag)
     --------------------------------------------------------------- */
  function parseBestek(blocks) {
    const out = { fields: {}, percelen: [], criteria: [] };
    const F = out.fields;
    // Voorblad: "BESTEK / VOOR / TYPE / MET ALS VOORWERP / “TITEL” / PROCEDURE / Ref.: ..."
    const lines = [];
    blocks.forEach(b => { if (b.type === "p") { const t = flat(b.text); if (t) lines.push(t); } });
    const iB = lines.findIndex((l, i) => /^bestek$/i.test(l) && /^voor$/i.test(lines[i + 1] || ""));
    if (iB >= 0) {
      F.type = lines[iB + 2] || "";
      const iM = lines.findIndex((l, i) => i > iB && /^met als voorwerp/i.test(l));
      if (iM >= 0) {
        F.titel = unq(lines[iM + 1] || "");
        F.procedure = lines[iM + 2] || "";
        const ref = lines.slice(iM + 2, iM + 5).find(l => /^ref/i.test(l));
        if (ref) F.referte = ref.replace(/^ref\.?\s*:?\s*/i, "").trim();
      }
    }
    const iC = findP(blocks, /^naam\s*:\s*samenaankoop/i);
    const txt = allText(blocks);
    let m = txt.match(/Contactpersoon\s*:\s*([^\n]+?)\s*(?:E-?mail|$)/i);
    if (m) F.contact_naam = flat(m[1]);
    m = txt.match(/E-?mail\s*:?\s*([\w.\-+]+@[\w.\-]+\.[a-z]{2,})/i);
    if (m) F.contact_email = m[1];

    const iV = findP(blocks, /^voorwerp van deze (leveringen|diensten|werken|opdracht)/i);
    if (iV >= 0) { const t = pText(blocks[iV + 1] || {}); if (t) F.voorwerp_zin = t; }

    const iI = findP(blocks, /verdeeld in (verschillende )?percelen|opgedeeld in volgende percelen/i);
    if (iI >= 0) {
      for (let i = iI + 1; i < blocks.length; i++) {
        const t = pText(blocks[i]);
        const mm = t.match(/^perceel\s*(\d+)\s*[:\-–]\s*(.*)$/i);
        if (!mm) { if (t) break; else continue; }
        out.percelen.push({ naam: unq(mm[2]), max_bedrag: "" });
      }
    }
    const iMax = findP(blocks, /maximale bestelhoeveelheid onder de raamovereenkomst bedraagt/i);
    if (iMax >= 0) {
      const t = flat(blocks[iMax].text);
      const mm = t.match(/bedraagt\s*:?\s*(.+?)(?:\.\s|$)/i);
      if (mm && /\d/.test(mm[1])) F.bestelhoeveelheid = cleanBedrag(mm[1]);
      for (let i = iMax + 1; i < blocks.length; i++) {
        const l = pText(blocks[i]);
        if (!l) continue;
        const m2 = l.match(/^perceel\s*(\d+)\s*[:\-–]\s*(.+)$/i);
        if (m2) { const idx = +m2[1] - 1; if (out.percelen[idx]) out.percelen[idx].max_bedrag = cleanBedrag(m2[2]); continue; }
        break;
      }
    }
    // Gunningscriteria: kop "Gunningscriteria" (Heading) gevolgd door tabel(len)
    let iG = -1;
    for (let i = 0; i < blocks.length; i++) {
      if (blocks[i].type === "p" && /^volgende criteria zijn van toepassing/i.test(flat(blocks[i].text))) { iG = i; break; }
    }
    if (iG >= 0) out.criteria = parseCriteriaTables(blocks, iG + 1, /^aan elk criterium werd|^aan elk criterium wordt/i);
    const mj = txt.match(/Voorziene begindatum van de leveringen\s*:\s*([^\n.]+)/i);
    if (mj) F.begindatum = flat(mj[1]);
    return out;
  }

  /* ---------------------------------------------------------------
     5. Doorzetten naar de velden van het gekozen doeldocument
        Elke mapper geeft { data, items } terug:
          data  = waarden per formulierveld van de Documentgenerator
          items = [{label, value, ok}] voor het controle-overzicht
     --------------------------------------------------------------- */
  const sentenceCase = (s) => { s = (s || "").toLowerCase(); return s.charAt(0).toUpperCase() + s.slice(1); };
  const num = (s) => { const v = parseFloat((s || "").toString().replace(/\./g, "").replace(",", ".")); return isNaN(v) ? null : v; };
  const keerNaarGetal = (s) => {
    const t = (s || "").toLowerCase(); const m = t.match(/\d+/);
    if (m) return +m[0];
    const w = Object.keys(NL_GETAL).sort((a, b) => b.length - a.length).find(k => t.includes(k));
    return w ? NL_GETAL[w] : null;
  };
  const naarMaanden = (s) => {
    const m = (s || "").toLowerCase().match(/(\d+)\s*(jaar|jaren|maand|maanden)/);
    return m ? (/jaar|jaren/.test(m[2]) ? +m[1] * 12 : +m[1]) : null;
  };
  function totaalBedrag(percelen) {
    const vals = percelen.map(p => num(p.max_bedrag));
    if (!vals.length || vals.some(v => v === null)) return "";
    const sum = vals.reduce((a, b) => a + b, 0);
    return sum.toLocaleString("nl-BE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " euro";
  }

  // criteria -> pgroups-structuur van de Documentgenerator
  function criteriaToGroups(criteria, aantalPercelen) {
    const rowOf = (r) => ({ nr: r.nr || "", naam: r.naam || "", gewicht: r.gewicht || "", details: r.details || "" });
    const filled = criteria.filter(g => g.rows.length);
    if (!filled.length) return null;
    const alle = filled.length === 1 && (!filled[0].percelen || !aantalPercelen || filled[0].percelen.length >= aantalPercelen || filled[0].percelen === null);
    if (alle || aantalPercelen < 2) {
      return { mode: "alle", groups: [{ percelen: [], v: { rows: filled[0].rows.map(rowOf) } }] };
    }
    return {
      mode: "groepen", aantal: aantalPercelen,
      groups: filled.map(g => ({ percelen: g.percelen || [], v: { rows: g.rows.map(rowOf) } })),
    };
  }

  function mapCommon(sl, d, items) {
    const F = sl.fields;
    const put = (key, val, label) => { d[key] = val || ""; items.push({ label, value: val || "", ok: !!val }); };
    put("opdracht_type", F.type, "Type opdracht");
    put("opdracht_titel", F.titel, "Voorwerp (voorblad)");
    put("procedure", F.procedure, "Plaatsingsprocedure");
    put("project_contact_naam", F.contact_naam, "Contactpersoon");
    put("project_contact_email", F.contact_email, "E-mail contactpersoon");
    if (F.referte) { d.project_ref = F.referte; }
  }

  function slToBestekCascadePercelen(sl) {
    const items = [], d = {}, F = sl.fields;
    mapCommon(sl, d, items);
    d.referte = F.referte || ""; items.push({ label: "Referte", value: d.referte, ok: !!d.referte });
    d.voorwerp_zin = F.titel_klein ? `Raamovereenkomst voor ‘${F.titel_klein}’.` : "";
    items.push({ label: "Voorwerp (zin onder 'Voorwerp van deze leveringen')", value: d.voorwerp_zin, ok: !!d.voorwerp_zin });
    d.percelen_inleiding = F.percelen_inleiding || "";
    items.push({ label: "Inleidende zin boven de percelen", value: d.percelen_inleiding, ok: !!d.percelen_inleiding });
    d.percelen = sl.percelen.map(p => ({ naam: p.naam, max_bedrag: p.max_bedrag }));
    items.push({ label: "Percelen en maximale bestelhoeveelheid", value: sl.percelen.map((p, i) => `Perceel ${i + 1}: ${p.naam}${p.max_bedrag ? " — " + p.max_bedrag : ""}`).join("\n"), ok: sl.percelen.length > 0 });
    d.looptijd = F.looptijd || ""; d.verlengingen = F.verlengingen || ""; d.verlengperiode = F.verlengperiode || "";
    items.push({ label: "Looptijd en verlengingen", value: [F.looptijd, F.verlengingen && `${F.verlengingen} verlengbaar, telkens ${F.verlengperiode || F.looptijd || ""}`].filter(Boolean).join(" — "), ok: !!F.looptijd });
    let doel = F.doel || "";
    if (doel && sl.percelen.length > 1) doel = doel.replace(/bovenvermelde opdracht,?\s+waarbij/i, "bovenvermelde percelen, waarbij").replace(/bovenvermelde opdracht/i, "bovenvermelde percelen");
    d.doel_tekst = doel; items.push({ label: "Doel van de opdracht", value: doel, ok: !!doel });
    d.leveringsplaats = F.leveringsplaats || ""; items.push({ label: "Plaats van de leveringen", value: d.leveringsplaats, ok: !!d.leveringsplaats });
    d.bijdrage_percentage = F.bijdrage || ""; items.push({ label: "Administratieve bijdrage (%)", value: d.bijdrage_percentage, ok: !!d.bijdrage_percentage });
    const g = criteriaToGroups(sl.criteria, sl.percelen.length);
    if (g) d.gunningscriteria_g = g;
    items.push({ label: "Gunningscriteria (namen)", value: sl.criteria.map(c => (c.label ? c.label + ": " : "") + c.rows.map(r => r.naam).join(", ")).join("\n"), ok: !!g });
    return { data: d, items, todo: ["Gewicht en uitwerking van elk gunningscriterium", "Eventueel verschillende criteria per perceel nauwkeuriger invullen"] };
  }

  function slToBestekEenvoud(sl, opts) {
    const items = [], d = {}, F = sl.fields;
    mapCommon(sl, d, items);
    d.looptijd = F.looptijd || ""; d.verlengingen = F.verlengingen || ""; d.verlengperiode = F.verlengperiode || "";
    items.push({ label: "Looptijd en verlengingen", value: [F.looptijd, F.verlengingen && `${F.verlengingen} verlengbaar, telkens ${F.verlengperiode || F.looptijd || ""}`].filter(Boolean).join(" — "), ok: !!F.looptijd });
    let best = F.bestelhoeveelheid || "";
    if (!best && sl.percelen.length) best = totaalBedrag(sl.percelen);
    d.bestelhoeveelheid_max = best;
    items.push({ label: "Maximale bestelhoeveelheid", value: best + (sl.percelen.length > 1 && best ? " (totaal van de percelen)" : ""), ok: !!best });
    d.bijdrage_percentage = F.bijdrage || ""; items.push({ label: "Administratieve bijdrage (%)", value: d.bijdrage_percentage, ok: !!d.bijdrage_percentage });
    if (opts.leveringsplaats) { d.leveringsplaats = F.leveringsplaats || ""; items.push({ label: "Leveringsplaats", value: d.leveringsplaats, ok: !!d.leveringsplaats }); }
    return { data: d, items, todo: ["De gunningscriteria staan in dit bestektype per onderdeel uitgeschreven en worden niet automatisch overgenomen"] };
  }

  function slToBestekOpenbaar(sl) {
    const items = [], d = {}, F = sl.fields;
    d.opdracht_type = F.type || ""; items.push({ label: "Soort opdracht", value: d.opdracht_type, ok: !!d.opdracht_type });
    d.voorwerp = F.titel_klein || F.titel || ""; items.push({ label: "Voorwerp van de opdracht", value: d.voorwerp, ok: !!d.voorwerp });
    d.project_contact_naam = F.contact_naam || ""; items.push({ label: "Dossierbeheerder", value: d.project_contact_naam, ok: !!d.project_contact_naam });
    d.project_contact_email = F.contact_email || ""; items.push({ label: "E-mail dossierbeheerder", value: d.project_contact_email, ok: !!d.project_contact_email });
    if (F.referte) d.project_ref = F.referte;
    d.percelen = sl.percelen.map((p, i) => ({
      titel: `Perceel ${i + 1}: ${p.naam}`,
      omschrijving: p.max_bedrag ? `Geraamde maximale waarde over de volledige looptijd: ${p.max_bedrag} excl. btw` : "",
    }));
    items.push({ label: "Percelen", value: d.percelen.map(p => p.titel + (p.omschrijving ? " — " + p.omschrijving : "")).join("\n"), ok: d.percelen.length > 0 });
    const lt = naarMaanden(F.looptijd), v = keerNaarGetal(F.verlengingen), vp = naarMaanden(F.verlengperiode) || lt;
    if (lt && v != null && vp) d.looptijd_maanden = String(lt + v * vp);
    else if (lt) d.looptijd_maanden = String(lt);
    items.push({ label: "Maximale looptijd (maanden)", value: d.looptijd_maanden || "", ok: !!d.looptijd_maanden });
    // gunningscriteria per perceel
    const rows = [];
    const aantal = sl.percelen.length;
    sl.criteria.forEach(g => {
      const perc = (g.percelen && g.percelen.length) ? g.percelen : (aantal > 1 ? Array.from({ length: aantal }, (_, i) => i + 1) : [null]);
      perc.forEach(n => g.rows.filter(r => !/\./.test(r.nr || "")).forEach(r => rows.push({
        perceel: n && aantal > 1 ? `Perceel ${n}: ${(sl.percelen[n - 1] || {}).naam || ""}`.trim() : "",
        naam: r.naam, gewicht: r.gewicht || "", toelichting: r.details || "",
      })));
    });
    if (rows.length) d.gunningscriteria = rows;
    items.push({ label: "Gunningscriteria (namen)", value: rows.map(r => (r.perceel ? r.perceel + " — " : "") + r.naam).join("\n"), ok: rows.length > 0 });
    return { data: d, items, todo: ["Gewichten en evaluatiemethode van de gunningscriteria", "Begin- en einddatum van de leveringen", "Selectiecriteria (die staan al in de selectieleidraad)"] };
  }

  /* ---------------------------------------------------------------
     6. Bestek -> gunningsverslag
     --------------------------------------------------------------- */
  function bestekToCriteriaRows(bs) {
    const aantal = bs.percelen.length;
    const rows = [];
    bs.criteria.forEach(g => {
      const perc = (g.percelen && g.percelen.length) ? g.percelen : (aantal > 1 && !g.label ? [null] : [g.percelen && g.percelen[0] || null]);
      const lab = g.label ? g.label.replace(/[:：]\s*$/, "") : "";
      g.rows.forEach(r => rows.push({ perceelLabel: lab, nr: r.nr, naam: r.naam, gewicht: r.gewicht, details: r.details }));
    });
    return rows;
  }
  function evaluatieTekst(details) {
    const lines = (details || "").split("\n").map(flat).filter(Boolean);
    const i = lines.findIndex(l => /^(evaluatie|beoordeling|score offerte)/i.test(l));
    return (i >= 0 ? lines.slice(i, i + 3) : []).join(" ").slice(0, 400);
  }

  function bestekToVerslagGunning(bs) {
    const items = [], d = {}, F = bs.fields;
    d.opdracht_titel = F.titel || ""; items.push({ label: "Titel opdracht", value: d.opdracht_titel, ok: !!d.opdracht_titel });
    d.besteknummer = F.referte || ""; items.push({ label: "Referte opdracht", value: d.besteknummer, ok: !!d.besteknummer });
    d.voorwerp = F.titel || F.voorwerp_zin || ""; items.push({ label: "Voorwerp van de opdracht", value: d.voorwerp, ok: !!d.voorwerp });
    if (F.referte) d.project_ref = F.referte;
    d.gunningswijze = F.procedure ? sentenceCase(F.procedure) : ""; items.push({ label: "Wijze van gunning", value: d.gunningswijze, ok: !!d.gunningswijze });
    const rows = bestekToCriteriaRows(bs).filter(r => r.naam);
    d.gunningscriteria = rows.map(r => ({
      naam: (r.perceelLabel ? r.perceelLabel + " — " : "") + (r.nr ? r.nr + " " : "") + r.naam,
      punten: r.gewicht, methode: evaluatieTekst(r.details), scores: "",
    }));
    items.push({ label: "Gunningscriteria met maximumpunten", value: d.gunningscriteria.map(r => `${r.naam} (${r.punten || "?"})`).join("\n"), ok: rows.length > 0 });
    if (bs.percelen.length) {
      d.besluit_gunning = bs.percelen.map((p, i) => ({ perceel: `Perceel ${i + 1}: ${p.naam}`, gekozen: "", motivering: "" }));
      items.push({ label: "Gunning per perceel (rijen)", value: d.besluit_gunning.map(r => r.perceel).join("\n"), ok: true });
    }
    return { data: d, items, todo: ["Ontvangen offertes (inschrijvers)", "Nazicht, scores per inschrijver en het gemotiveerd besluit"] };
  }

  function bestekToVerslagOpenbaar(bs) {
    const items = [], d = {}, F = bs.fields;
    d.voorwerp = F.titel || F.voorwerp_zin || ""; items.push({ label: "Voorwerp van de opdracht", value: d.voorwerp, ok: !!d.voorwerp });
    if (F.referte) d.project_ref = F.referte;
    if (F.procedure && !/openbare/i.test(F.procedure)) { d.wijze_gunning = sentenceCase(F.procedure); }
    items.push({ label: "Wijze van gunning", value: d.wijze_gunning || "(standaardtekst openbare procedure blijft staan)", ok: !!d.wijze_gunning });
    d.percelen = bs.percelen.map((p, i) => ({ titel: `Perceel ${i + 1}: ${p.naam}` }));
    items.push({ label: "Percelen", value: d.percelen.map(p => p.titel).join("\n"), ok: d.percelen.length > 0 });
    const rows = bestekToCriteriaRows(bs).filter(r => r.naam);
    d.gunningscriteria = rows.map(r => ({
      perceel: r.perceelLabel, naam: (r.nr ? r.nr + " " : "") + r.naam, gewicht: r.gewicht,
      toelichting: evaluatieTekst(r.details),
    }));
    items.push({ label: "Gunningscriteria met maximumpunten", value: d.gunningscriteria.map(r => `${r.perceel ? r.perceel + " — " : ""}${r.naam} (${r.gewicht || "?"})`).join("\n"), ok: rows.length > 0 });
    return { data: d, items, todo: ["Inschrijvers, regelmatigheid, scores en rangschikking", "Controle uitsluitingsgronden en selectiecriteria per inschrijver", "Gunningsvoorstel"] };
  }

  /* ---------------------------------------------------------------
     7. Overzicht van de ketens en doeldocumenten
     --------------------------------------------------------------- */
  const CHAINS = {
    sl2bestek: {
      title: "Selectieleidraad → Bestek",
      intro: "Upload de selectieleidraad (.docx). De gele delen — voorwerp, percelen, looptijd, bestelhoeveelheid, gunningscriteria, bijdrage — gaan door naar het bestek van je keuze.",
      parse: parseSelectieleidraad,
      sourceName: "selectieleidraad",
      targets: [
        { id: "bestek_cascade_percelen", page: "bestek-cascade-percelen.html", label: "Bestek — cascade met percelen", desc: "Raamopdracht met meerdere deelnemers per perceel, toewijzing via cascade.", map: slToBestekCascadePercelen },
        { id: "bestek_cascade", page: "bestek-cascade.html", label: "Bestek — cascade (zonder percelen)", desc: "Raamopdracht met meerdere deelnemers via een cascadesysteem, één geheel.", map: (sl) => slToBestekEenvoud(sl, { leveringsplaats: true }) },
        { id: "bestek_minicompetitie", page: "bestek-minicompetitie.html", label: "Bestek — minicompetitie", desc: "Raamopdracht waarbij de afroepen via een minicompetitie worden toegewezen.", map: (sl) => slToBestekEenvoud(sl, {}) },
        { id: "bestek_openbare_procedure", page: "bestek-openbare-procedure.html", label: "Bestek — openbare procedure", desc: "Bestek voor een openbare procedure, met selectie- en gunningscriteria in één document.", map: slToBestekOpenbaar },
      ],
    },
    bestek2verslag: {
      title: "Bestek → Gunningsverslag",
      intro: "Upload het bestek (.docx). Voorwerp, referte, percelen en gunningscriteria met hun maximumpunten gaan door naar het gunningsverslag van je keuze.",
      parse: parseBestek,
      sourceName: "bestek",
      targets: [
        { id: "verslag_gunning", page: "verslag-gunning.html", label: "Gunningsverslag (mededingingsprocedure / cascade)", desc: "Verslag na een mededingingsprocedure met onderhandeling of een cascade.", map: bestekToVerslagGunning },
        { id: "verslag_gunning_openbaar", page: "verslag-gunning-openbaar.html", label: "Gunningsverslag — openbare procedure", desc: "Verslag met regelmatigheidscontrole, scores en rangschikking per perceel.", map: bestekToVerslagOpenbaar },
      ],
    },
  };

  root.DG = { readDocx, readBlocks, parseSelectieleidraad, parseBestek, CHAINS, parsePerceelNummers, cleanBedrag, criteriaToGroups };
})(typeof window !== "undefined" ? window : globalThis);
