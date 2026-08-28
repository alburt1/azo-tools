/* =========================================================================
   HTML → docx.js converter
   Zet de HTML die render() produceert (zelfde functies die de live preview
   voeden) om naar echte docx.js-objecten (Paragraph/Table/TextRun/...) voor
   een écht .docx-bestand — geen HTML-die-doet-alsof-.doc meer.
   Dekt exact de tags/klassen die de render()-functies in engine.js gebruiken
   (zie shared/engine.js voor de brontekst).
   ========================================================================= */
(function (global) {
  const docx = global.docx;

  const PT = (pt) => Math.round(pt * 20); // punten -> twips/half-points basis (1pt = 20 twips)
  const HALFPT = (pt) => Math.round(pt * 2); // punten -> half-points (voor TextRun size)

  const COLOR = {
    ink: "171A15",
    headingDark: "21261E",
    green: "7DB73F",
    greenDark: "4C7527",
    tableHeaderText: "35521C",
    tableHeaderBg: "EEF5E6",
    tableBorder: "B9C2AE",
    muted: "8A927F",
    mutedFn: "5B6152",
    lhRef: "5B6455",
    lhRefB: "21261E",
    fillOk: "1A3D0D",
    fillEmptyText: "97662A",
    fillEmptyBg: "FBF1DE",
  };

  function dataUriToUint8(dataUri) {
    const b64 = dataUri.split(",")[1] || "";
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }

  function parseStyleAttr(styleStr) {
    const out = {};
    (styleStr || "").split(";").forEach((decl) => {
      const idx = decl.indexOf(":");
      if (idx === -1) return;
      const prop = decl.slice(0, idx).trim().toLowerCase();
      const val = decl.slice(idx + 1).trim();
      if (!prop || !val) return;
      out[prop] = val;
    });
    return out;
  }

  function ptFromCss(val) {
    if (!val) return null;
    const m = /([\d.]+)\s*pt/.exec(val);
    return m ? parseFloat(m[1]) : null;
  }

  /* -------------------------------------------------------------------
     Inline (run-level) opbouw
     ------------------------------------------------------------------- */
  function runPropsFromFmt(fmt) {
    const props = {
      bold: !!fmt.bold,
      italics: !!fmt.italics,
      size: fmt.size || 20,
      font: fmt.font || "Arial",
      color: fmt.color || undefined,
      superScript: !!fmt.superScript,
      allCaps: !!fmt.allCaps,
    };
    if (fmt.underline) props.underline = {};
    if (fmt.highlight) props.shading = { fill: fmt.highlight };
    return props;
  }

  // CSS .paper{ line-height:1.15 } — geldt voor alle lopende tekst. In docx is
  // dat spacing.line uitgedrukt in 240sten van een regel: 240*1.15 = 276.
  const LINE_115 = { line: 276, lineRule: docx.LineRuleType.AUTO };

  function makeImageRun(imgEl) {
    const src = imgEl.getAttribute("src") || "";
    if (!src.startsWith("data:image")) return null;
    const w = parseInt(imgEl.getAttribute("width") || "40", 10);
    const h = parseInt(imgEl.getAttribute("height") || "20", 10);
    const type = src.includes("image/png") ? "png" : "jpg";
    try {
      return new docx.ImageRun({
        type,
        data: dataUriToUint8(src),
        transformation: { width: w, height: h },
      });
    } catch (e) {
      return null;
    }
  }

  /* Loopt over de inline-kinderen van een blok-element (p, li, td, ...) en
     bouwt een platte lijst van TextRun/ImageRun, met opgestapelde opmaak
     (bold/underline/italic/superscript/highlight) via tag- en klasse-naam. */
  function buildRuns(node, fmt, runs, topLevel) {
    fmt = fmt || {};
    if (topLevel === undefined) topLevel = true; // enkel de allereerste (paragraaf-niveau) aanroep telt als topLevel
    const children = Array.from(node.childNodes);
    children.forEach((child, i) => {
      if (child.nodeType === 3) {
        let text = child.textContent.replace(/[\t\n\r ]+/g, " ");
        if (text === "") return;
        // Randspaties zijn HTML-opmaak-inspringing, geen betekenisvolle tekst —
        // maar enkel relevant op het niveau van de alinea zelf: een spatie die
        // toevallig de volledige inhoud van een eigen <strong>/<u>/... vormt
        // (bv. "<strong>A</strong><strong> </strong><strong>B</strong>") is wél
        // een betekenisvolle scheiding tussen woorden en mag niet wegvallen.
        if (topLevel) {
          const prev = children[i - 1];
          const isFirst = i === 0;
          const afterBreak = prev && prev.nodeType === 1 && prev.tagName.toLowerCase() === "br";
          if ((isFirst || afterBreak) && text.charAt(0) === " ") text = text.slice(1);
          const isLast = i === children.length - 1;
          if (isLast && text.charAt(text.length - 1) === " ") text = text.slice(0, -1);
        }
        if (text === "") return;
        runs.push(new docx.TextRun(Object.assign({ text }, runPropsFromFmt(fmt))));
        return;
      }
      if (child.nodeType !== 1) return;
      const tag = child.tagName.toLowerCase();
      const cls = (child.getAttribute && child.getAttribute("class")) || "";

      if (tag === "br") {
        runs.push(new docx.TextRun(Object.assign({ text: "", break: 1 }, runPropsFromFmt(fmt))));
        return;
      }
      if (tag === "ul" || tag === "ol") {
        // geneste lijsten worden elders als eigen paragrafen verwerkt, nooit inline
        return;
      }
      if (tag === "img") {
        const run = makeImageRun(child);
        if (run) runs.push(run);
        return;
      }
      if (tag === "sup") {
        buildRuns(child, Object.assign({}, fmt, { superScript: true, size: Math.round((fmt.size || 20) * 0.72) }), runs, false);
        return;
      }
      if (tag === "strong" || tag === "b") {
        buildRuns(child, Object.assign({}, fmt, { bold: true }), runs, false);
        return;
      }
      if (tag === "u") {
        buildRuns(child, Object.assign({}, fmt, { underline: true }), runs, false);
        return;
      }
      if (tag === "em" || tag === "i") {
        buildRuns(child, Object.assign({}, fmt, { italics: true }), runs, false);
        return;
      }
      if (tag === "mark") {
        // Op vraag: geen highlight-markering meer, gewoon platte tekst
        // (de omliggende opmaak — bv. <strong>/<u> — blijft wel behouden).
        buildRuns(child, fmt, runs, false);
        return;
      }
      if (cls.indexOf("field-fill") !== -1) {
        const isEmpty = cls.indexOf("empty") !== -1;
        const extra = isEmpty
          ? { italics: true, bold: true, color: COLOR.fillEmptyText, highlight: "FBF1DE" }
          : { bold: true, color: COLOR.fillOk };
        buildRuns(child, Object.assign({}, fmt, extra), runs, false);
        return;
      }
      // span, a, en alle andere inline-achtige tags: opmaak doorgeven, gewoon induiken
      buildRuns(child, fmt, runs, false);
    });
  }

  function inlineParagraphOptions(el, extraProps) {
    const runs = [];
    buildRuns(el, {}, runs);
    if (runs.length === 0) runs.push(new docx.TextRun({ text: "" }));
    return Object.assign({ children: runs }, extraProps || {});
  }

  /* -------------------------------------------------------------------
     Blok-niveau opbouw
     ------------------------------------------------------------------- */
  const ACCENT_BORDER = { style: "single", size: 18, color: COLOR.green, space: 4 };
  const ACCENT_BORDER_THIN = { style: "single", size: 8, color: COLOR.green, space: 2 };
  const DOCX_NO_BORDER = { style: "none", size: 0, color: "FFFFFF" };
  const DOCX_TABLE_NO_BORDERS = {
    top: DOCX_NO_BORDER, bottom: DOCX_NO_BORDER, left: DOCX_NO_BORDER, right: DOCX_NO_BORDER,
    insideHorizontal: DOCX_NO_BORDER, insideVertical: DOCX_NO_BORDER,
  };

  // Strip een handmatig cijfer-voorvoegsel ("1. " of "3.1. ") van de
  // allereerste tekst — die kan een rechtstreeks kind zijn van de alinea, of
  // genest zitten in een opmaak-wrapper (bv. <p><em><strong>4.1. Perceel...).
  // Dat voorvoegsel stond al in de brontekst; Word's eigen Heading-nummering
  // vervangt dat nu.
  function stripLeadingNumberFromClone(el) {
    const clone = el.cloneNode(true);
    let node = clone;
    while (node) {
      const first = node.childNodes[0];
      if (!first) break;
      if (first.nodeType === 3) {
        const stripped = first.textContent.replace(/^\s*\d+(\.\d+)*\.\s*/, "");
        if (stripped !== first.textContent) first.textContent = stripped;
        break;
      }
      if (first.nodeType === 1 && ["strong", "b", "u", "em", "i"].indexOf(first.tagName.toLowerCase()) !== -1) {
        node = first; // induiken in de opmaak-wrapper op zoek naar de tekst
        continue;
      }
      break; // onbekend eerste kind: niets te strippen
    }
    return clone;
  }

  // Detecteert een alinea die VOLLEDIG bestaat uit vetgedrukte tekst,
  // gecombineerd met ofwel onderstreping (criteria, bv. "2. DROPKOST") ofwel
  // cursief (perceel-niveau, bv. "4.1. Perceel 1: ..."). Enkel gebruikt in
  // combinatie met een cijfer-voorvoegsel-check hieronder, zodat algemene
  // labels als "Beoordeling:" of bedrijfsnamen (die nooit met een cijfer
  // beginnen) er nooit door gevangen worden.
  function detectEmphasisStyle(el) {
    let hasStrong = false, hasU = false, hasEm = false;
    function walk(node) {
      for (const c of node.childNodes) {
        if (c.nodeType === 3) continue;
        if (c.nodeType === 1 && c.tagName.toLowerCase() === "br") continue;
        if (c.nodeType !== 1) return false;
        const t = c.tagName.toLowerCase();
        if (t === "strong" || t === "b") { hasStrong = true; if (!walk(c)) return false; }
        else if (t === "u") { hasU = true; if (!walk(c)) return false; }
        else if (t === "em" || t === "i") { hasEm = true; if (!walk(c)) return false; }
        else return false;
      }
      return true;
    }
    if (!walk(el) || !hasStrong) return null;
    if (hasU) return "underline";
    if (hasEm) return "italic";
    return null;
  }

  /* Bouwt een paragraaf met een echte Word-koptekststijl (Heading 1/2/3/4)
     i.p.v. handmatige bold/kleur-opmaak — enkel wanneer ctx.nativeHeadings
     actief is (de 3 verslagen). Zo kan Word zelf een inhoudsopgave +
     automatische nummering (1., 1.1, ...) genereren. Geeft null terug als
     deze paragraaf geen koptekst is, zodat de aanroeper op de normale
     afhandeling terugvalt. */
  function buildHeadingParagraph(el, cls, ctx) {
    if (!ctx || !ctx.nativeHeadings) return null;
    let headingLevel = null;
    let stripNumber = false;
    let disableNumbering = false;

    if (cls.indexOf("verslag-h") !== -1) {
      headingLevel = docx.HeadingLevel.HEADING_1;
      stripNumber = true;
    } else if (cls.indexOf("verslag-sub") !== -1) {
      headingLevel = docx.HeadingLevel.HEADING_2;
    } else if (cls.indexOf("bijlage-h") !== -1) {
      headingLevel = docx.HeadingLevel.HEADING_1;
      disableNumbering = true; // "BIJLAGE A: ..." behoudt haar eigen tekst, geen extra Word-cijfer ervoor
    } else {
      // Herkent expliciet genummerde tussentitels in de lopende tekst die zelf
      // nooit een echte koptekst-klasse kregen, bv. "4.1. Perceel 1: ..."
      // (cursief+vet = perceel-niveau) of "2. DROPKOST (10 punten)" /
      // "3.1. 'Must have'..." (onderstreept+vet = gunningscriterium-niveau).
      // Het cijfer-voorvoegsel is de betrouwbare marker: labels als
      // "Beoordeling:" of een bedrijfsnaam beginnen nooit met een cijfer.
      const text = el.textContent.trim();
      const numMatch = /^(\d+(?:\.\d+)*)\.\s/.exec(text);
      const style = detectEmphasisStyle(el);
      if (numMatch && style && parseStyleAttr(el.getAttribute("style")||"")["text-align"] !== "center") {
        if (style === "italic") {
          headingLevel = docx.HeadingLevel.HEADING_2;
        } else {
          const depth = numMatch[1].split(".").length; // "2" -> 1, "3.1" -> 2
          headingLevel = depth === 1 ? docx.HeadingLevel.HEADING_3 : docx.HeadingLevel.HEADING_4;
        }
        stripNumber = true;
      }
      if (!headingLevel) return null;
    }

    const source = stripNumber ? stripLeadingNumberFromClone(el) : el;
    const runs = [];
    // Platte tekst gebruiken (geen <strong>/<u>/<em> uit de brontekst
    // meenemen) — anders botst de originele inline-opmaak met de opmaak die
    // de Heading-stijl zelf al oplegt (bv. dubbel onderstreept).
    const plainText = source.textContent.replace(/[\t\n\r ]+/g, " ").trim();
    if (plainText) runs.push(new docx.TextRun({ text: plainText }));
    if (runs.length === 0) runs.push(new docx.TextRun({ text: "" }));

    const props = { heading: headingLevel, spacing: LINE_115 };
    if (cls.indexOf("bijlage-h") !== -1) {
      props.pageBreakBefore = true;
      props.border = { top: ACCENT_BORDER_THIN };
      props.spacing = Object.assign({}, LINE_115, { before: PT(14) }); // ruimte tussen lijn en titel
    }
    if (disableNumbering) props.numbering = false;
    return new docx.Paragraph(Object.assign({ children: runs }, props));
  }

  function blockParagraphProps(el) {
    const cls = (el.getAttribute && el.getAttribute("class")) || "";
    const style = parseStyleAttr(el.getAttribute ? el.getAttribute("style") : "");
    // CSS: p{ margin:0 0 8pt; } — geen ruimte boven, 8pt onder, en overal de
    // 1.15 regelafstand van .paper.
    const props = { spacing: Object.assign({ before: 0, after: PT(8) }, LINE_115) };

    if (cls.indexOf("verslag-h") !== -1) {
      Object.assign(props, { spacing: Object.assign({ before: PT(16), after: PT(6) }, LINE_115) });
      return { props, fmt: { bold: true, size: HALFPT(10.5), color: COLOR.headingDark } };
    }
    if (cls.indexOf("verslag-sub") !== -1) {
      Object.assign(props, { spacing: Object.assign({ before: PT(12), after: PT(4) }, LINE_115) });
      return { props, fmt: { bold: true } };
    }
    if (cls.indexOf("bijlage-h") !== -1) {
      // CSS: margin:26pt 0 10pt; padding-top:14pt; border-top — de 14pt
      // padding zit tussen de lijn en de titel (border "space"), de 26pt
      // margin komt vóór de lijn zelf.
      Object.assign(props, {
        pageBreakBefore: true,
        spacing: Object.assign({ before: PT(26), after: PT(10) }, LINE_115),
        border: { top: { style: "single", size: 8, color: COLOR.green, space: 14 } },
      });
      return { props, fmt: { bold: true, size: HALFPT(12), color: COLOR.headingDark } };
    }
    if (cls.indexOf("doc-title-h3") !== -1) {
      Object.assign(props, { alignment: docx.AlignmentType.CENTER, spacing: Object.assign({ before: 0, after: PT(14) }, LINE_115) });
      return { props, fmt: { bold: true, size: HALFPT(11) } };
    }
    if (cls.indexOf("scope-note") !== -1) {
      Object.assign(props, { spacing: Object.assign({ before: PT(6), after: PT(6) }, LINE_115) });
      return { props, fmt: { italics: true, color: COLOR.lhRef, size: HALFPT(9) } };
    }
    if (cls.indexOf("aangetekend") !== -1) {
      Object.assign(props, { spacing: Object.assign({ before: 0, after: PT(14) }, LINE_115) });
      return { props, fmt: { bold: true } };
    }
    if (cls.indexOf("betreft") !== -1) {
      Object.assign(props, { spacing: Object.assign({ before: 0, after: PT(16) }, LINE_115) });
      return { props, fmt: {} };
    }
    if (style["text-align"] === "center") {
      Object.assign(props, { alignment: docx.AlignmentType.CENTER });
    }
    let fmt = {};
    if (style["font-weight"] === "700" || style["font-weight"] === "bold") fmt.bold = true;
    const sz = ptFromCss(style["font-size"]);
    if (sz) fmt.size = HALFPT(sz);
    return { props, fmt };
  }

  function convertList(listEl, ordered, level, ctx) {
    level = level || 0;
    ctx = ctx || {};
    const baseFmt = ctx.baseFmt || {};
    const paras = [];
    let n = parseInt(listEl.getAttribute("start") || "1", 10);
    const liEls = Array.from(listEl.children).filter((c) => c.tagName.toLowerCase() === "li");
    liEls.forEach((li, liIndex) => {
      const isLastLi = liIndex === liEls.length - 1;
      const directChildren = Array.from(li.children);
      const nestedList = directChildren.find((c) => c.tagName === "UL" || c.tagName === "OL");
      const nonListChildren = directChildren.filter((c) => c.tagName !== "UL" && c.tagName !== "OL");

      // Pandoc wikkelt li-inhoud vaak in exact één <p>: gebruik die p rechtstreeks
      // als bron voor de bullet-tekst. Anders: li's eigen inline content (tekst-
      // nodes + inline tags), met geneste UL/OL/P bewust overgeslagen door buildRuns.
      const inlineSource = (nonListChildren.length === 1 && nonListChildren[0].tagName === "P")
        ? nonListChildren[0]
        : li;

      const runs = [];
      buildRuns(inlineSource, baseFmt, runs);
      if (runs.length === 0) runs.push(new docx.TextRun({ text: "" }));

      // CSS: .doc-body ul/ol{ margin:0 0 10pt; } — geen ruimte tussen de
      // items onderling, enkel 10pt ná de volledige lijst (laatste item),
      // maar alleen als er geen geneste sub-lijst of extra blok-inhoud volgt
      // (die zorgt dan zelf voor de afsluitende ruimte).
      const isTrailingSpacer = isLastLi && !nestedList && inlineSource === li;
      const listItemSpacing = Object.assign({ after: isTrailingSpacer ? PT(10) : 0 }, LINE_115);

      if (ordered) {
        runs.unshift(new docx.TextRun(Object.assign({ text: `${n}. ` }, runPropsFromFmt(baseFmt))));
        paras.push(new docx.Paragraph({ children: runs, indent: { left: PT(18 + level * 14) }, spacing: listItemSpacing }));
      } else {
        paras.push(new docx.Paragraph({ children: runs, bullet: { level }, spacing: listItemSpacing }));
      }
      n++;

      // overige blok-content binnen deze li (bv. een tweede <p> naast de eerste) apart toevoegen
      if (inlineSource !== li) {
        nonListChildren.filter((c) => c !== inlineSource).forEach((extra) => {
          paras.push(...htmlNodeToBlocks(wrapSingle(extra), ctx));
        });
      }
      if (nestedList) {
        paras.push(...convertList(nestedList, nestedList.tagName === "OL", level + 1, ctx));
      }
    });
    return paras;
  }

  // Hulpfunctie: stopt één element in een tijdelijke wrapper-div zodat
  // htmlNodeToBlocks() (die over .children van een container itereert) het kan verwerken.
  function wrapSingle(el) {
    const wrapper = global.document.createElement("div");
    wrapper.appendChild(el.cloneNode(true));
    return wrapper;
  }

  function cellWidthFromStyle(td) {
    const style = parseStyleAttr(td.getAttribute("style"));
    const pt = ptFromCss(style["width"]);
    if (pt) return { size: PT(pt), type: docx.WidthType.DXA };
    // HTML width-attribuut (bv. width="54%" of width="120"): niet elke tabel
    // gebruikt inline style, sommige (bv. de rechterkolom-truc voor het
    // adresblok) zetten de breedte via het klassieke HTML-attribuut.
    const attr = td.getAttribute("width");
    if (attr) {
      const pct = /^(\d+(\.\d+)?)%$/.exec(attr.trim());
      if (pct) return { size: `${pct[1]}%`, type: docx.WidthType.PERCENTAGE };
      const px = parseFloat(attr);
      if (!isNaN(px)) return { size: Math.round(px * 15), type: docx.WidthType.DXA }; // ruwe px->twips
    }
    return undefined;
  }

function cellToBlocks(td, ctx){
    ctx = ctx || {};
    if (hasBlockChild(td)) return htmlNodeToBlocks(td, ctx);
    const runs = [];
    buildRuns(td, ctx.baseFmt || {}, runs);
    if (runs.length === 0) return [];
    return [ new docx.Paragraph({ spacing: LINE_115, children: runs }) ];
  }

  function convertTable(tableEl, opts) {
    opts = opts || {};
    const cls = tableEl.getAttribute("class") || "";
    const isDataTable = cls.indexOf("doc-table") !== -1; // .doc-table / .doc-table-checklist
    const rows = [];
    const theadRows = tableEl.querySelectorAll("thead tr");
    const bodyRows = tableEl.querySelectorAll("tbody tr");
    const allRows = theadRows.length || bodyRows.length
      ? [...theadRows, ...bodyRows]
      : Array.from(tableEl.querySelectorAll("tr"));

    // Op vraag: kolommen die geen enkele expliciete breedte meekregen,
    // worden gelijk over de tabel verdeeld i.p.v. Word ze te laten
    // auto-schatten op basis van de inhoud (wat vaak erg ongelijke, lelijke
    // kolombreedtes oplevert). Enkel van toepassing als GEEN enkele cel in
    // de hele tabel al een eigen breedte heeft (anders respecteren we die
    // bewust ingestelde breedtes).
    const anyExplicitWidth = Array.from(tableEl.querySelectorAll("td, th")).some((td) => cellWidthFromStyle(td) !== undefined);
    let equalColPct = null;
    if (isDataTable && !anyExplicitWidth && allRows.length) {
      const firstRowCols = Array.from(allRows[0].children)
        .reduce((sum, td) => sum + parseInt(td.getAttribute("colspan") || "1", 10), 0);
      if (firstRowCols > 0) equalColPct = (100 / firstRowCols).toFixed(2) + "%";
    }

    // Komt overeen met table.doc-table th CSS: hoofdletters, groene kleur,
    // kleiner Calibri-lettertype, vet — en dat geldt ook voor eventuele
    // opsommingen/alinea's die binnen zo'n koptekstcel genest zitten.
    const headerCtx = { baseFmt: { bold: true, allCaps: true, color: COLOR.tableHeaderText, size: HALFPT(8), font: "Calibri" } };
    // Voor tabellen zonder echte <th>: de bovenste rij oogt visueel toch als
    // titelrij (bv. "Nr. | Naam | Adres..."), dus die krijgt dezelfde groene
    // vet-opmaak — enkel zonder de hoofdletters/kleiner lettertype, dat past
    // niet bij elke titelrij (bv. een samengevoegde titelcel).
    const topRowCtx = { baseFmt: { bold: true, color: COLOR.tableHeaderText } };

    allRows.forEach((tr, rowIndex) => {
      const cells = [];
      const rowHasRealTh = isDataTable && Array.from(tr.children).some((td) => td.tagName.toLowerCase() === "th");
      Array.from(tr.children).forEach((td) => {
        const isRealHeader = isDataTable && td.tagName.toLowerCase() === "th";
        const isTopRowHeader = isDataTable && rowIndex === 0 && !rowHasRealTh;
        const isHeader = isRealHeader || isTopRowHeader;
        const colspan = parseInt(td.getAttribute("colspan") || "1", 10);
        const cellCtx = isRealHeader ? headerCtx : (isTopRowHeader ? topRowCtx : {});
        const blocks = cellToBlocks(td, cellCtx);
        const cellChildren = blocks.length ? blocks : [new docx.Paragraph({ children: [new docx.TextRun("")] })];
        cells.push(
          new docx.TableCell({
            children: cellChildren,
            columnSpan: colspan > 1 ? colspan : undefined,
            width: cellWidthFromStyle(td) || (equalColPct ? { size: equalColPct, type: docx.WidthType.PERCENTAGE } : undefined),
            shading: isHeader ? { fill: COLOR.tableHeaderBg } : undefined,
            borders: isDataTable ? undefined : DOCX_TABLE_NO_BORDERS,
            verticalAlign: docx.VerticalAlign.TOP,
            // CSS: table.doc-table th/td{ padding:5pt 7pt; } (1pt = 20 twips) —
            // iets extra verticale ademruimte (6,5pt i.p.v. 5pt) voor leesbaarheid.
            margins: isDataTable
              ? { top: PT(6.5), bottom: PT(6.5), left: PT(7), right: PT(7) }
              : { top: 0, bottom: 0, left: 0, right: 0 },
          })
        );
      });
      if (cells.length) rows.push(new docx.TableRow({ children: cells }));
    });

    if (!rows.length) return null;

    // CSS: table.doc-table th/td{ border:0.75pt solid #b9c2ac; } — op vraag
    // iets dikker dan de preview: 1pt (size in achtsten van een punt: 1*8=8).
    const dataBorder = { style: "single", size: 8, color: COLOR.tableBorder };
    return new docx.Table({
      rows,
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: equalColPct ? docx.TableLayoutType.FIXED : undefined,
      borders: isDataTable
        ? {
            top: dataBorder,
            bottom: dataBorder,
            left: dataBorder,
            right: dataBorder,
            insideHorizontal: dataBorder,
            insideVertical: dataBorder,
          }
        : DOCX_TABLE_NO_BORDERS,
    });
  }

  function convertFootnotes(asideEl) {
    const blocks = [];
    // CSS: .doc-footnotes{ margin-top:22pt; padding-top:8pt; border-top:0.75pt solid #b9c2ac; }
    blocks.push(
      new docx.Paragraph({
        border: { top: { style: "single", size: 6, color: COLOR.tableBorder, space: 8 } },
        spacing: Object.assign({ before: PT(22), after: 0 }, LINE_115),
        children: [new docx.TextRun("")],
      })
    );
    const items = asideEl.querySelectorAll("li");
    items.forEach((li, i) => {
      const runs = [];
      buildRuns(li, { size: HALFPT(8.2), color: COLOR.mutedFn }, runs);
      blocks.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: `${i + 1}. `, size: HALFPT(8.2), color: COLOR.mutedFn }), ...runs],
          spacing: Object.assign({ after: PT(4) }, LINE_115),
        })
      );
    });
    return blocks;
  }

  const BLOCK_TAGS = ["P", "TABLE", "UL", "OL", "DIV", "HR", "ASIDE", "BLOCKQUOTE"];
  function hasBlockChild(el) {
    return Array.from(el.children).some((c) => BLOCK_TAGS.indexOf(c.tagName) !== -1);
  }

  /* Zet één DOM-node (blok-element) om naar 0..n docx blocks (Paragraph/Table). */
  function htmlNodeToBlocks(node, ctx) {
    ctx = ctx || {};
    const baseFmt = ctx.baseFmt || {};
    const out = [];
    Array.from(node.children).forEach((el) => {
      const tag = el.tagName.toLowerCase();
      const cls = el.getAttribute("class") || "";

      if (tag === "table") {
        const t = convertTable(el);
        if (t) {
          out.push(t);
          // CSS: table.doc-table{ margin:10pt 0 14pt; } — docx.Table zelf
          // ondersteunt geen alinea-marge, dus een kleine lege regel erna
          // voor ademruimte t.o.v. de volgende tekst.
          const isDataTable = (el.getAttribute("class") || "").indexOf("doc-table") !== -1;
          if (isDataTable) {
            out.push(new docx.Paragraph({ spacing: Object.assign({ before: 0, after: PT(6) }, LINE_115), children: [new docx.TextRun("")] }));
          }
        }
        return;
      }
      if (tag === "ul") {
        out.push(...convertList(el, false, 0, ctx));
        return;
      }
      if (tag === "ol") {
        out.push(...convertList(el, true, 0, ctx));
        return;
      }
      if (tag === "hr") {
        const isBottom = cls.indexOf("doc-accent-line-bottom") !== -1;
        out.push(
          new docx.Paragraph({
            border: { bottom: isBottom ? ACCENT_BORDER_THIN : ACCENT_BORDER },
            spacing: { before: 0, after: isBottom ? PT(6) : PT(10) },
            children: [new docx.TextRun("")],
          })
        );
        return;
      }
      if (tag === "aside" && cls.indexOf("doc-footnotes") !== -1) {
        out.push(...convertFootnotes(el));
        return;
      }
      if (tag === "blockquote") {
        const pChildren = Array.from(el.children).filter((c) => c.tagName === "P");
        if (pChildren.length > 0) {
          pChildren.forEach((p) => {
            const runs = [];
            buildRuns(p, Object.assign({ italics: true }, baseFmt), runs);
            if (runs.length) out.push(new docx.Paragraph({ children: runs, indent: { left: PT(14) }, spacing: { after: PT(6) } }));
          });
        } else {
          const runs = [];
          buildRuns(el, Object.assign({ italics: true }, baseFmt), runs);
          if (runs.length) out.push(new docx.Paragraph({ children: runs, indent: { left: PT(14) }, spacing: { after: PT(6) } }));
        }
        return;
      }
      if (tag === "div") {
        const clsDiv = el.getAttribute("class") || "";
        if (clsDiv.indexOf("red-box") !== -1) {
          // Rode omkaderde box (bv. de stuknummer/bestandsnaam-instructies) —
          // een tabel met 1 cel en rode rand geeft een betrouwbaar echt kader.
          const innerBlocks = htmlNodeToBlocks(el, ctx);
          const cellChildren = innerBlocks.length ? innerBlocks : [new docx.Paragraph({ children: [new docx.TextRun("")] })];
          const redBorder = { style: "single", size: 10, color: "B02318" };
          out.push(new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            borders: { top: redBorder, bottom: redBorder, left: redBorder, right: redBorder, insideHorizontal: redBorder, insideVertical: redBorder },
            rows: [ new docx.TableRow({ children: [ new docx.TableCell({
              margins: { top: 140, bottom: 140, left: 160, right: 160 },
              children: cellChildren,
            }) ] }) ],
          }));
          // ademruimte na de box
          out.push(new docx.Paragraph({ spacing: { before: 0, after: PT(6) }, children: [new docx.TextRun("")] }));
          return;
        }
        if (hasBlockChild(el)) {
          // structurele wrapper (doc-body, addr-block, sign, ...): induiken
          out.push(...htmlNodeToBlocks(el, ctx));
        } else {
          // div zonder blok-kinderen = een enkele tekstregel (bv. adresregel)
          const runs = [];
          buildRuns(el, baseFmt, runs);
          if (runs.length) out.push(new docx.Paragraph({ children: runs, spacing: { after: PT(2) } }));
        }
        return;
      }
      if (tag === "p") {
        const cls2 = el.getAttribute("class") || "";
        const headingPara = buildHeadingParagraph(el, cls2, ctx);
        if (headingPara) { out.push(headingPara); return; }
        const { props, fmt } = blockParagraphProps(el);
        const mergedFmt = Object.assign({}, baseFmt, fmt);
        const runs = [];
        buildRuns(el, mergedFmt, runs);
        if (runs.length === 0) return; // lege paragrafen (bv. spacer <p>) overslaan
        out.push(new docx.Paragraph(Object.assign({ children: runs }, props)));
        return;
      }
      // onbekend blok-element (bv. span/a direct als "blok" gebruikt): geen verdere
      // blok-kinderen? behandel als platte tekstregel i.p.v. stilzwijgend te verdwijnen.
      if (hasBlockChild(el)) {
        out.push(...htmlNodeToBlocks(el, ctx));
      } else {
        const runs = [];
        buildRuns(el, baseFmt, runs);
        if (runs.length) out.push(new docx.Paragraph({ children: runs }));
      }
    });
    return out;
  }

  /* Publieke API: HTML-string -> array van docx.Paragraph/docx.Table */
  function htmlStringToDocxBlocks(html, opts) {
    const wrapper = global.document.createElement("div");
    wrapper.innerHTML = html;
    return htmlNodeToBlocks(wrapper, opts || {});
  }

  global.AZO_HTML_TO_DOCX = { htmlStringToDocxBlocks, dataUriToUint8, PT, HALFPT, COLOR };
})(typeof window !== "undefined" ? window : globalThis);
