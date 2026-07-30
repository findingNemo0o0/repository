const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "SKP Marktanalyse";
pres.title = "Schulbegleitung im Umbruch";

// ---- Palette -------------------------------------------------------------
const INK = "0E2E38";
const INK_SOFT = "1B4A57";
const INK_CARD = "13424F";
const TXT = "182E36";
const MUT = "5F7278";
const CARD = "F1F6F8";
const CARD2 = "E7EFF3";
const BLUE = "2A78D6";     // validierte Serienfarbe 1
const ORANGE = "EB6834";   // validierte Serienfarbe 2
const WHITE = "FFFFFF";
const PALE = "CFE3EA";

const HFONT = "Cambria";
const BFONT = "Calibri";
const M = 0.62;
const W = 13.33;

// ---- Helfer ---------------------------------------------------------------
const shadow = () => ({ type: "outer", color: "0E2E38", blur: 10, offset: 2, angle: 90, opacity: 0.10 });

function titleBar(slide, kicker, title) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: M, y: 0.34, w: 11, h: 0.26, fontFace: BFONT, fontSize: 11, bold: true,
      color: ORANGE, charSpacing: 1.6, margin: 0,
    });
  }
  slide.addText(title, {
    x: M, y: 0.58, w: W - 2 * M, h: 0.86, fontFace: HFONT, fontSize: 26, bold: true,
    color: TXT, margin: 0, valign: "top",
  });
}
function darkTitle(slide, kicker, title) {
  slide.addText(kicker.toUpperCase(), {
    x: M, y: 0.66, w: 11, h: 0.28, fontFace: BFONT, fontSize: 11, bold: true,
    color: ORANGE, charSpacing: 1.6, margin: 0,
  });
  slide.addText(title, {
    x: M, y: 0.98, w: W - 2 * M, h: 0.8, fontFace: HFONT, fontSize: 28, bold: true, color: WHITE, margin: 0,
  });
}
function keyline(slide, text, y) {
  slide.addText(text, {
    x: M, y: y, w: W - 2 * M, h: 0.4, fontFace: BFONT, fontSize: 13.5, italic: true, color: INK_SOFT, margin: 0,
  });
}
function srcNote(slide, text, y) {
  slide.addText(text, {
    x: M, y: y || 6.94, w: W - 2 * M, h: 0.3, fontFace: BFONT, fontSize: 9, color: MUT, margin: 0,
  });
}
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, fill: { color: fill || CARD }, rectRadius: 0.08, line: { color: fill || CARD }, shadow: shadow(),
  });
}
function chip(slide, x, y, w, label, color) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h: 0.34, fill: { color }, rectRadius: 0.06, line: { color },
  });
  slide.addText(label, {
    x, y, w, h: 0.34, fontFace: BFONT, fontSize: 11, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0,
  });
}
// Zweispaltige Vergleichstabelle
function compareTable(slide, y0, rows, colX, colW, headColors, rowH, step, fs, tblW) {
  let y = y0;
  const tw = tblW || (W - 2 * M);
  rows.forEach((r, i) => {
    if (i > 0) card(slide, M, y, tw, rowH, i % 2 === 0 ? CARD : "FFFFFF");
    r.forEach((cell, j) => {
      slide.addText(cell, {
        x: colX[j] + 0.24, y: i === 0 ? y : y + 0.08, w: colW[j], h: i === 0 ? 0.34 : rowH - 0.16,
        fontFace: BFONT, fontSize: i === 0 ? 11.5 : fs,
        bold: i === 0 || j === 0,
        color: i === 0 ? (headColors[j] || MUT) : (j === 2 ? INK_SOFT : TXT),
        charSpacing: i === 0 ? 1 : 0, margin: 0, valign: "middle",
      });
    });
    y += i === 0 ? 0.40 : step;
  });
  return y;
}

/* =======================================================================
   1 — Titel
   ===================================================================== */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("Schulbegleitung im Umbruch", {
    x: M, y: 2.00, w: 11.6, h: 0.95, fontFace: HFONT, fontSize: 44, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Marktanalyse SKP · Kreis Steinburg · Hamburg · Segeberg · Ostholstein · Neumünster", {
    x: M, y: 3.00, w: 11.6, h: 0.5, fontFace: BFONT, fontSize: 17, color: "BFD6DE", margin: 0,
  });
  card(s, M, 4.00, 9.2, 1.25, INK_SOFT);
  s.addText([
    { text: "Die Kernthese:  ", options: { bold: true, color: WHITE } },
    { text: "Der Markt wächst — aber der Zugang zu ihm wird gerade neu vergeben.", options: { color: "DCEAF0" } },
  ], { x: M + 0.3, y: 4.20, w: 8.6, h: 0.85, fontFace: BFONT, fontSize: 16, margin: 0, valign: "middle" });
  s.addText("Stand: Juli 2026  ·  Folien 15–19 sind Anhang zum Nachschlagen", {
    x: M, y: 6.42, w: 8, h: 0.3, fontFace: BFONT, fontSize: 11, color: "7FA3B0", margin: 0,
  });
  s.addNotes("Einstieg in 30 Sekunden: SKP verkauft Schulbegleitung an Kreise. Der Bedarf steigt seit zehn Jahren ununterbrochen. Gleichzeitig stellen die Kostenträger gerade das Vergabesystem um — von Einzelfallbewilligung auf Standortvertrag. Wer die Umstellung übersteht, hat ein größeres und planbareres Geschäft. Wer sie verpasst, verliert Standorte vollständig. Hinweis geben: Detailfragen sind im Anhang beantwortet.");
}

/* =======================================================================
   2 — Geschäftsmodell: Rollen & Geltungsbereich
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Das Geschäft in 60 Sekunden", "Drei Akteure — und nur einer ist Vertragspartner");

  const roles = [
    ["Kreis / kreisfreie Stadt", "DER VERTRAGSPARTNER", BLUE,
      "Rechtsträger, Zahler und Unterzeichner. In Schleswig-Holstein erstattet das Land rund 84 % der SGB-IX-Kosten. In Hamburg tritt an seine Stelle die Schulbehörde BSFB."],
    ["Jugendamt / Amt für Teilhabe", "DIE FACHBEHÖRDE", BLUE,
      "Kein eigenes Rechtssubjekt, sondern ein Amt innerhalb des Kreises: Bedarfsfeststellung, Bewilligung, Trägerauswahl, Fachaufsicht."],
    ["Schule", "NIE VERTRAGSPARTNER", ORANGE,
      "Aber Einsatzort mit faktischem Vetorecht: Ein Pool kommt nur „im Einvernehmen mit der Schulleitung“ zustande."],
  ];
  roles.forEach((r, i) => {
    const x = M + i * 4.06;
    card(s, x, 1.54, 3.90, 2.42);
    chip(s, x + 0.26, 1.74, 2.0, r[1], r[2]);
    s.addText(r[0], {
      x: x + 0.26, y: 2.18, w: 3.38, h: 0.38, fontFace: BFONT, fontSize: 14, bold: true, color: TXT, margin: 0,
    });
    s.addText(r[3], {
      x: x + 0.26, y: 2.58, w: 3.38, h: 1.24, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
    });
  });

  card(s, M, 4.12, W - 2 * M, 1.16, CARD2);
  s.addText("Geltungsbereich: alle Schularten", {
    x: M + 0.3, y: 4.26, w: 11.5, h: 0.3, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Schulbegleitung gilt für Klasse 1 bis 13 — Regelschule wie Förderschule, berufsbildende Schulen und den offenen Ganztag (§ 112 SGB IX). Die Poolreform beginnt meist an Grundschulen, weil dort die Schulische Assistenz des Landes ansetzt; sie ist aber nicht darauf beschränkt.",
    { x: M + 0.3, y: 4.58, w: 11.5, h: 0.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  card(s, M, 5.42, W - 2 * M, 1.16);
  s.addText("So entsteht heute Umsatz", {
    x: M + 0.3, y: 5.56, w: 11.5, h: 0.3, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Vergütet wird pro Kind und Stunde auf Basis eines Bewilligungsbescheids. Klassenfahrten und Ausflüge sind als Zusatzstunden abrechenbar — in Pinneberg 8 Stunden je Tag, ohne vorherigen Antrag; das monatliche Stundensoll erhöht sich entsprechend.",
    { x: M + 0.3, y: 5.88, w: 11.5, h: 0.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  srcNote(s, "Quellen: SGB VIII, SGB IX; Landtag SH Drs. 20/2643; Trägerschreiben Kreis Pinneberg 20.03.2024; Empfehlungen Bayerischer Bezirketag/StMUK 2025.");
  s.addNotes("Drei Akteure, ein Vertragspartner. Der wichtigste Punkt ist die dritte Karte: Die Schule unterschreibt nie, entscheidet aber mit. Deshalb ist die Beziehung zu den Schulleitungen der eigentliche Vermögenswert. Zweiter Punkt für Rückfragen: Schulbegleitung ist nicht auf die Grundschule beschränkt — nur die Poolreform startet dort.");
}

/* =======================================================================
   3 — Nachfrage (Chart A)
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Nachfrage", "Mehr Kinder mit Förderbedarf — und mehr davon an Regelschulen");

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [1.1, 3.5, 3.4] },
    { name: "Exklusionsquote (Förderschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [4.8, 4.3, 4.2] },
  ], {
    x: M, y: 1.54, w: 7.5, h: 4.55,
    barDir: "col", barGrouping: "stacked", chartColors: [BLUE, ORANGE],
    showTitle: true, title: "Förderquote Deutschland, in % aller Schüler:innen",
    titleFontSize: 12, titleColor: MUT, titleFontFace: BFONT,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 11,
    dataLabelColor: WHITE, dataLabelFontFace: BFONT, dataLabelFormatCode: "0.0",
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 12, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 10, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMaxVal: 9, valAxisMajorUnit: 3,
  });

  const facts = [
    ["+1,7 pp", "Förderquote insgesamt:\n5,9 % → 7,6 % aller Schüler:innen", BLUE],
    ["+2,3 pp", "Inklusionsquote — Kinder mit\nFörderbedarf an Regelschulen", BLUE],
    ["−0,6 pp", "Exklusionsquote: nur ein Viertel\ndes Zuwachses kommt von dort", ORANGE],
  ];
  let y = 1.66;
  facts.forEach((f) => {
    card(s, 8.45, y, 4.26, 1.34);
    s.addText(f[0], {
      x: 8.72, y: y + 0.13, w: 3.8, h: 0.5, fontFace: HFONT, fontSize: 25, bold: true, color: f[2], margin: 0,
    });
    s.addText(f[1], {
      x: 8.72, y: y + 0.64, w: 3.8, h: 0.58, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
    });
    y += 1.49;
  });
  keyline(s, "Drei Viertel des Inklusionszuwachses sind neu diagnostizierte Förderbedarfe — nicht Kinder, die von der Förderschule wechseln.", 6.20);
  srcNote(s, "Quelle: Bertelsmann Stiftung auf Grundlage KMK (2024) / Klemm; Schuljahre 2008/09 und 2022/23 (ohne Saarland). pp = Prozentpunkte.");
  s.addNotes("Kernbotschaft: Der Nachfragesockel wächst strukturell. Wichtig gegen die naheliegende Rückfrage: Die Förderschulen leeren sich kaum — die Exklusionsquote fiel nur um 0,6 Punkte, während die Inklusionsquote um 2,3 Punkte stieg. Drei Viertel des Zuwachses sind also zusätzliche Diagnosen. Beide Bewegungen erzeugen Schulbegleitungsbedarf.");
}

/* =======================================================================
   4 — Standort (Chart B)
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Standort", "SKPs Regionen sind deutlich inklusiver als der Bundesschnitt");

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [3.4, 4.4, 5.2] },
    { name: "Exklusionsquote (Förderschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [4.2, 2.3, 2.7] },
  ], {
    x: M, y: 1.54, w: 7.5, h: 4.55,
    barDir: "col", barGrouping: "stacked", chartColors: [BLUE, ORANGE],
    showTitle: true, title: "Schuljahr 2022/23, in % aller Schüler:innen",
    titleFontSize: 12, titleColor: MUT, titleFontFace: BFONT,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 11,
    dataLabelColor: WHITE, dataLabelFontFace: BFONT, dataLabelFormatCode: "0.0",
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 11, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 10, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMaxVal: 9, valAxisMajorUnit: 3,
  });

  card(s, 8.45, 1.66, 4.26, 2.30);
  s.addText("Warum das den Marktwert trägt", {
    x: 8.72, y: 1.82, w: 3.8, h: 0.34, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Schulbegleitung entsteht ganz überwiegend an Regelschulen. Dort werden in Hamburg 5,2 % und in Schleswig-Holstein 4,4 % aller Schüler:innen mit Förderbedarf unterrichtet — gegenüber 3,4 % im Bund.",
    { x: 8.72, y: 2.20, w: 3.8, h: 1.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  card(s, 8.45, 4.10, 4.26, 1.99, CARD2);
  s.addText("Die Kehrseite", {
    x: 8.72, y: 4.26, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Beide Regionen haben die niedrigsten Förderschulanteile. Die Reserve an Kindern, die noch von der Förderschule wechseln könnten, ist damit weitgehend ausgeschöpft — künftiges Wachstum kommt fast nur noch aus neuen Diagnosen.",
    { x: 8.72, y: 4.58, w: 3.8, h: 1.4, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });

  srcNote(s, "Quelle: Bertelsmann Stiftung auf Grundlage KMK (2024), Schuljahr 2022/23. Förderquote = Inklusions- + Exklusionsquote.");
  s.addNotes("Antwort auf die Frage „warum ausgerechnet hier?“: Weil in SH und Hamburg überdurchschnittlich viele Kinder mit Förderbedarf an Regelschulen sind — und Schulbegleitung genau dort gebraucht wird. Der Standort ist ein Asset. Die Kehrseite ehrlich mitnennen: Der Umschichtungseffekt ist hier fast ausgereizt.");
}

/* =======================================================================
   5 — Marktvolumen (Chart C)
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Marktvolumen", "Die Ausgaben wachsen fast doppelt so schnell wie die Fallzahlen");

  s.addChart(pres.ChartType.line, [
    { name: "Ausgaben", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 122, 148, 170, 199, 234, 286, 384] },
    { name: "Leistungsempfänger", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 111, 131, 144, 162, 189, 204, 245] },
  ], {
    x: M, y: 1.54, w: 7.5, h: 4.35,
    chartColors: [ORANGE, BLUE], lineDataSymbol: "circle", lineDataSymbolSize: 7, lineSize: 3,
    showTitle: true, title: "Schulbegleitung Schleswig-Holstein, Index 2014 = 100",
    titleFontSize: 12, titleColor: MUT, titleFontFace: BFONT,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 11, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 10, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 80, valAxisMaxVal: 400, valAxisMajorUnit: 80,
  });

  card(s, 8.45, 1.66, 4.26, 2.16, CARD2);
  s.addText("Schleswig-Holstein 2022", {
    x: 8.72, y: 1.82, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0,
  });
  s.addText("119,7 Mio €", {
    x: 8.72, y: 2.12, w: 3.8, h: 0.6, fontFace: HFONT, fontSize: 31, bold: true, color: ORANGE, margin: 0,
  });
  s.addText("Gesamtausgaben (SGB VIII + SGB IX); 2014 waren es 27,1 Mio €.\n6.545 Leistungsempfänger (2021).",
    { x: 8.72, y: 2.76, w: 3.8, h: 0.9, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });

  card(s, 8.45, 3.96, 4.26, 2.13);
  s.addText("Hamburg 2014/15 → 2025/26", {
    x: 8.72, y: 4.12, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0,
  });
  s.addText([
    { text: "Bewilligungen   1.574 → 4.011", options: { breakLine: true, bold: true, color: TXT } },
    { text: "Kosten   6,75 → 42,15 Mio €", options: { breakLine: true, bold: true, color: TXT } },
    { text: "Kosten je Fall   ~4.300 → ~10.500 €", options: { color: ORANGE, bold: true } },
  ], { x: 8.72, y: 4.46, w: 3.8, h: 1.2, fontFace: BFONT, fontSize: 12.5, margin: 0, lineSpacingMultiple: 1.3 });

  keyline(s, "Nicht die Menge treibt die Kosten, sondern der Preis je Fall — und genau daran setzt die Reform an.", 6.14);
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643 (Landkreistag/Städteverband, 2014–2022); BSFB Hamburg. Kosten je Fall: eigene Berechnung. Für 2022 fehlt die Empfängerzahl einer Stadt.");
  s.addNotes("Der wichtigste Chart. Orange: Ausgaben, Blau: Fallzahlen. Beide steigen — die Ausgaben fast doppelt so schnell. Kosten je Fall haben sich in Hamburg auf das Zweieinhalbfache erhöht. Das ist der Auslöser der Reform. Hier eine Pause machen.");
}

/* =======================================================================
   6 — Die Reform
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Die Reform", "Von der 1:1-Begleitung zum Pool — auf drei Ebenen gleichzeitig");

  card(s, M, 1.52, 5.50, 1.94, CARD2);
  s.addText("HEUTE  ·  Einzelfall", {
    x: M + 0.28, y: 1.70, w: 4.9, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: MUT, charSpacing: 1, margin: 0,
  });
  s.addText("Ein Kind — ein Antrag — ein Bescheid — eine Begleitperson. Bezahlt wird pro Kind und Stunde.",
    { x: M + 0.28, y: 2.04, w: 4.9, h: 1.2, fontFace: BFONT, fontSize: 14, color: TXT, margin: 0 });

  card(s, 6.90, 1.52, 5.81, 1.94);
  s.addText("KÜNFTIG  ·  Pool / Infrastruktur", {
    x: 7.18, y: 1.70, w: 5.2, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: ORANGE, charSpacing: 1, margin: 0,
  });
  s.addText("Die Schule erhält ein Budget und ein Team. Kein Antrag, kein Bescheid. Bezahlt wird pro Klasse oder Standort.",
    { x: 7.18, y: 2.04, w: 5.2, h: 1.2, fontFace: BFONT, fontSize: 14, color: TXT, margin: 0 });

  s.addShape(pres.ShapeType.rightArrow, {
    x: 6.32, y: 2.22, w: 0.46, h: 0.54, fill: { color: ORANGE }, line: { color: ORANGE },
  });

  const lvl = [
    ["Bund", BLUE, "Referentenentwurf 1. KJHSRG (März 2026): „infrastrukturelle Bildungsassistenz“ als Regelfall; Individualanspruch nur noch, wenn ausschließlich eine 1:1-Begleitung hilft. Geplant ab 01.01.2028 — aber ohne Personalschlüssel."],
    ["Land SH", BLUE, "Erklärtes Ziel der Landesregierung: Schulassistenz und Schulbegleitung in Poollösungen zusammenführen. Ein Endtermin wird ausdrücklich nicht genannt."],
    ["Kreise", ORANGE, "Handeln bereits eigenständig aus Kostendruck — nahezu alle Kreise in SH erproben oder betreiben Poolmodelle. Hamburg stellt zum Schuljahr 2026/27 um."],
  ];
  let y = 3.62;
  lvl.forEach((l) => {
    card(s, M, y, W - 2 * M, 0.94);
    chip(s, M + 0.26, y + 0.30, 1.32, l[0], l[1]);
    s.addText(l[2], {
      x: M + 1.78, y: y + 0.10, w: 10.0, h: 0.74, fontFace: BFONT, fontSize: 11.5, color: TXT, margin: 0, valign: "middle",
    });
    y += 1.04;
  });
  srcNote(s, "Quellen: Referentenentwurf 1. KJHSRG (§§ 35d Abs. 4, 80a SGB VIII-E); Landtag SH Drs. 20/2643 und 20/3271; BSFB Hamburg.");
  s.addNotes("Keine einzelne Gesetzesänderung, auf die man warten könnte: Drei Ebenen laufen parallel, und die unterste — die Kreise — ist bereits in der Umsetzung. Wenn das Bundesgesetz scheitert, ändert das operativ wenig. Wichtig für Rückfragen: § 80a schreibt keinen Personalschlüssel vor.");
}

/* =======================================================================
   7 — Konsequenz für den Anbieter
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Konsequenz für den Anbieter", "Aus vielen kleinen Bewilligungen werden wenige große Verträge");

  const rows = [
    ["", "Einzelfall (heute)", "Pool (künftig)"],
    ["Vertragspartner", "Kreis, je Kind", "Kreis, je Standort oder Region"],
    ["Auftragserteilung", "Bewilligungsbescheid", "Ausschreibung oder Interessenbekundung"],
    ["Laufzeit", "12 Monate, rollierend", "bis 5 Jahre, plus Verlängerungsoption"],
    ["Vergütung", "pro Kind und Stunde", "Festbudget je Klasse oder Standort"],
    ["Klassenfahrten", "8 h/Tag zusätzlich abrechenbar", "im Festbudget enthalten"],
    ["Mengenrisiko", "beim Kreis", "beim Träger"],
    ["Standortverlust", "einzelne Fälle fallen weg", "gesamtes Volumen fällt weg"],
  ];
  compareTable(s, 1.58, rows, [M, M + 3.1, M + 7.6], [3.0, 4.4, 4.5], [MUT, MUT, ORANGE], 0.50, 0.56, 12);

  card(s, M, 5.96, W - 2 * M, 0.78, CARD2);
  s.addText([
    { text: "Mengenrisiko heißt:  ", options: { bold: true, color: ORANGE } },
    { text: "Kommt ein Kind dazu, gab es früher einen neuen Bescheid und zusätzliches Geld. Im Pool steht das Budget fest — Krankheit, Vertretung, Klassenfahrten und Zuzüge trägt der Träger. ", options: { color: TXT } },
    { text: "An jedem Standort entscheidet es sich damit binär: koordinierender Träger — oder dort raus.", options: { bold: true, color: TXT } },
  ], { x: M + 0.3, y: 5.96, w: 11.5, h: 0.78, fontFace: BFONT, fontSize: 12, margin: 0, valign: "middle" });

  srcNote(s, "Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg; TED 246925-2024; Trägerschreiben Kreis Pinneberg 20.03.2024; DISW-Evaluation Ostholstein.", 6.94);
  s.addNotes("Hier liegt Risiko und Chance zugleich. Vorher war der Umsatz granular. Künftig ist er binär: Standort gewonnen heißt Jahre, Standort verloren heißt alles weg. In Ostholstein ist dokumentiert, dass auch die Mitarbeitenden zum neuen Träger wechseln.");
}

/* =======================================================================
   8 — Rechenbeispiel
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Rechenbeispiel", "Was das Klassenmodell konkret bedeutet — und warum es die Ausnahme ist");

  const rows = [
    ["", "Einzelfall (bisher)", "Klassenmodell Pinneberg"],
    ["Finanzierte Stellen", "223 Schulbegleitungen", "431 Klassenassistenzen + 38 Teamleitungen"],
    ["Bezugsgröße", "Kinder mit Bewilligung", "alle Klassen (rund 470)"],
    ["Quote", "ca. 1 Fall je 2 Klassen", "1 Stelle je Klasse"],
    ["Umfang je Stelle", "individuell bewilligt", "24 h (Kl. 1–2) / 30 h (Kl. 3–4)"],
    ["Relation je Kind mit Bedarf", "1 : 1", "1 : ganze Klasse"],
  ];
  compareTable(s, 1.58, rows, [M, M + 2.9, M + 5.9], [2.75, 2.85, 3.0], [MUT, MUT, ORANGE], 0.54, 0.60, 11, 9.13);

  card(s, 9.95, 1.58, 2.76, 2.06);
  s.addText("Warum das gut wäre", {
    x: 10.18, y: 1.74, w: 2.35, h: 0.32, fontFace: BFONT, fontSize: 12.5, bold: true, color: BLUE, margin: 0,
  });
  s.addText("Rund die Hälfte der Klassen hatte bisher keinen Fall — und damit keinen Umsatz. Im Klassenmodell wird jede Klasse zur Abrechnungseinheit.",
    { x: 10.18, y: 2.10, w: 2.35, h: 1.4, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0 });

  card(s, 9.95, 3.76, 2.76, 2.06, CARD2);
  s.addText("Warum es nicht die Basis ist", {
    x: 10.18, y: 3.92, w: 2.35, h: 0.32, fontFace: BFONT, fontSize: 12.5, bold: true, color: ORANGE, margin: 0,
  });
  s.addText("Nur Pinneberg rechnet klassenbasiert — und hat zweimal verschoben. Ostholstein bemisst nach der Summe der bisherigen Einzelbedarfe, Hamburg nach Kombinationsmaßnahmen.",
    { x: 10.18, y: 4.28, w: 2.35, h: 1.45, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0 });

  card(s, M, 5.94, W - 2 * M, 0.84, CARD);
  s.addText([
    { text: "Basisannahme für die Bewertung:  ", options: { bold: true, color: TXT } },
    { text: "gleiches Volumen, mehr Kinder, weniger Stunden je Kind — so wie in Ostholstein. Das Klassenmodell ist Upside, nicht Grundlage.", options: { color: TXT } },
  ], { x: M + 0.3, y: 5.94, w: 11.5, h: 0.84, fontFace: BFONT, fontSize: 12.5, margin: 0, valign: "middle" });

  srcNote(s, "Die 223 sind ein Durchschnitt 2017–2020 und nur SGB VIII; aktuell eher 350–500. Die Klassenzahl (~470) ist aus den 431 Stellen zzgl. abgezogener Landeskräfte abgeleitet und steht in keiner Quelle.", 6.98);
  s.addNotes("Diese Folie beantwortet die naheliegendste Rückfrage: Wäre eine Kraft je Klasse nicht sogar besser? Ja — in Pinneberg verdoppelt sich die Stellenzahl fast. Aber der Satz ist gedeckelt, vorhandene Landeskräfte werden abgezogen, und nur ein einziger Kreis rechnet so. Deshalb Upside, nicht Basisannahme.");
}

/* =======================================================================
   9 — Regionale Betroffenheit
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Regionale Betroffenheit", "Fünf Regionen, drei Geschwindigkeiten");

  const regs = [
    ["Hamburg", "AKUT", ORANGE, "Umstellung auf Kombinationsmaßnahmen läuft seit Schuljahr 2026/27, ohne Beschränkung auf eine Schulart. Höher qualifizierte Kräfte nur noch im Ausnahmefall."],
    ["Kreis Ostholstein", "HOCH", ORANGE, "Poolmodell seit 2020/21, seit 2024 Ausweitung auf alle Schulstandorte beschlossen. „Konzentration der Leistungserbringer“ ist erklärtes Projektziel."],
    ["Kreis Segeberg", "MITTEL", "C98500", "Pool an den Förderzentren seit 2021/22, Ausweitung auf rund fünf Grundschulen ab 2025/26 — Auswahl über Interessenbekundung, nicht über Ausschreibung."],
    ["Neumünster", "NIEDRIG", BLUE, "Kein Poolmodell dokumentiert. Als kreisfreie Stadt aber schnell umsteuerbar — Flensburg brauchte vom Pilot bis flächendeckend zwei Jahre."],
    ["Kreis Steinburg", "NIEDRIG", BLUE, "Steuerungs-AG ruht. Seit Februar 2025 sind Jugend- und Sozialamtsaufgaben im Amt für Teilhabe gebündelt — die strukturelle Vorstufe."],
  ];
  let y = 1.56;
  regs.forEach((r) => {
    card(s, M, y, W - 2 * M, 0.92);
    s.addShape(pres.ShapeType.ellipse, { x: M + 0.3, y: y + 0.32, w: 0.28, h: 0.28, fill: { color: r[2] }, line: { color: r[2] } });
    s.addText(r[0], { x: M + 0.72, y: y + 0.13, w: 2.7, h: 0.34, fontFace: BFONT, fontSize: 14, bold: true, color: TXT, margin: 0 });
    s.addText(r[1], { x: M + 0.72, y: y + 0.48, w: 2.7, h: 0.28, fontFace: BFONT, fontSize: 10.5, bold: true, color: r[2], charSpacing: 1.2, margin: 0 });
    s.addText(r[3], { x: M + 3.55, y: y + 0.12, w: 8.2, h: 0.68, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0, valign: "middle" });
    y += 1.02;
  });
  keyline(s, "Am wenigsten exponiert ist das Geschäft an weiterführenden Schulen und Förderzentren — dort kommen die Poolmodelle später und lückenhafter.", 6.20);
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643 (Stand 11/2024); DISW-Evaluation Ostholstein; BSFB Hamburg; Kreis Steinburg. Der Stand kann sich kurzfristig ändern.");
  s.addNotes("Keine einheitliche Frist: Zwei Regionen laufen bereits, zwei noch gar nicht. „Niedrig“ heißt nicht „kein Handlungsbedarf“. Und der Hinweis unten ist bewertungsrelevant: Der Umsatzanteil an weiterführenden Schulen und Förderzentren bleibt am längsten stabil.");
}

/* =======================================================================
   10 — Chancen
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Chancen", "Warum der Markt für den richtigen Anbieter größer wird");

  const items = [
    ["Der Gesamtmarkt wächst weiter", "Fallzahlen und Ausgaben steigen in allen belastbaren Datenreihen. Kein Kreis erwartet einen Rückgang."],
    ["Größere, planbarere Aufträge", "Statt hunderter Einzelbewilligungen wenige Standortverträge über Jahre — mit geringerem Verwaltungsaufwand je Umsatzeuro."],
    ["Regionale Verankerung zählt formal", "Kreis Pinneberg macht Kenntnis der regionalen Strukturen und bestehende Kooperationen zu nachzuweisenden Zuschlagskriterien."],
    ["Qualität schlägt Preis", "Kreis Düren gewichtet Qualität mit 60 von 100 Punkten. Ein dokumentiertes Qualifizierungskonzept ist der billigste Differenzierer."],
    ["Zeitfenster in zwei Regionen", "In Steinburg und Neumünster existiert noch kein Modell — dort lässt sich das Verfahren mitgestalten statt sich darauf zu bewerben."],
    ["Der Markt ist offen für Unternehmen", "Zuschläge gingen zuletzt auch an eine GmbH & Co. KG und eine gGmbH — nicht nur an Wohlfahrtsverbände."],
  ];
  items.forEach((it, i) => {
    const x = M + (i % 2) * 6.15;
    const yy = 1.56 + Math.floor(i / 2) * 1.70;
    card(s, x, yy, 5.94, 1.50);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.26, y: yy + 0.24, w: 0.3, h: 0.3, fill: { color: BLUE }, rectRadius: 0.05, line: { color: BLUE } });
    s.addText(it[0], { x: x + 0.68, y: yy + 0.20, w: 5.0, h: 0.36, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0 });
    s.addText(it[1], { x: x + 0.68, y: yy + 0.58, w: 5.0, h: 0.80, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });
  });
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643; TED-Bekanntmachungen Kreis Pinneberg, Düren, Euskirchen, Rhein-Erft, Essen.");
  s.addNotes("Nicht euphorisch werden — alle Chancen hängen an einer Bedingung: Das Unternehmen muss vergabefähig werden. Punkt 3 und 4 sind entscheidend, weil ein etablierter lokaler Anbieter genau das günstiger liefern kann als ein Konzern.");
}

/* =======================================================================
   11 — Risiken
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Risiken", "Was den Wert des Unternehmens gefährden kann");

  const items = [
    ["Standortverlust ist total", "Wer den Zuschlag nicht bekommt, verliert das gesamte Volumen an diesem Standort — zum Schuljahresbeginn, ohne Übergang."],
    ["Personal wandert mit", "In Ostholstein dokumentiert: Mitarbeitende werden an den koordinierenden Träger „übergeben“. Rechtlich freiwillig, faktisch die Regel."],
    ["Kein Rechtsschutz", "Das OVG Schleswig wies die Klage von vier Trägern gegen das Pinneberger Vergabeverfahren zurück. Unanfechtbar. Bestandsschutz gibt es nicht."],
    ["Vergabefähigkeit fehlt oft", "10 Mio € Betriebshaftpflicht, Referenzliste über vier Jahre, getrennte Umsatzausweisung, Präqualifizierung — harte Ausschlusskriterien (Anhang, Folie 18)."],
  ];
  let y = 1.56;
  items.forEach((it) => {
    card(s, M, y, W - 2 * M, 0.96);
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.28, y: y + 0.33, w: 0.3, h: 0.3, fill: { color: ORANGE }, rectRadius: 0.05, line: { color: ORANGE } });
    s.addText(it[0], { x: M + 0.72, y: y + 0.15, w: 3.6, h: 0.68, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0, valign: "middle" });
    s.addText(it[1], { x: M + 4.5, y: y + 0.15, w: 7.25, h: 0.68, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0, valign: "middle" });
    y += 1.06;
  });

  card(s, M, 5.90, W - 2 * M, 0.78, CARD2);
  s.addText([
    { text: "Das größte Einzelrisiko:  ", options: { bold: true, color: ORANGE } },
    { text: "Preisbindung ohne Tarifausgleich — nächste Folie.", options: { color: TXT } },
  ], { x: M + 0.3, y: 5.90, w: 11.5, h: 0.78, fontFace: BFONT, fontSize: 13, margin: 0, valign: "middle" });

  srcNote(s, "Quellen: OVG Schleswig, Beschluss vom 03.09.2024 – 5 MB 7/24 (unanfechtbar); TED 246925-2024; DISW-Evaluation Ostholstein.");
  s.addNotes("Ehrlich vortragen — das ist der Grund, warum das Unternehmen jetzt Investition braucht. Besonders Punkt 3: Es gibt keinen Rechtsweg gegen die Umstellung. Die einzige wirksame Verteidigung ist, das Verfahren zu gewinnen.");
}

/* =======================================================================
   12 — Preisbindung (dunkel)
   ===================================================================== */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  darkTitle(s, "Das eine, das man verstehen muss", "Ein Festpreis über bis zu neun Jahre — ohne Tarifausgleich");

  s.addText("„Anpassungen durch Tarifsteigerungen sind durch die Mischkalkulation bereits abgegolten und erfolgen im Rahmen der Vertragslaufzeit nicht.“", {
    x: M, y: 1.94, w: 7.4, h: 0.9, fontFace: HFONT, fontSize: 15, italic: true, color: PALE, margin: 0,
  });
  s.addText("Konzept „Klassenassistenz“, Kreis Pinneberg", {
    x: M, y: 2.84, w: 7.4, h: 0.3, fontFace: BFONT, fontSize: 11, color: "7FA3B0", margin: 0,
  });

  const tiles = [
    ["01.11.2024 – 31.07.2029", "Grundlaufzeit des ausgeschriebenen Vertrags"],
    ["+ 4 Jahre", "einseitige Verlängerungsoption des Kreises — der Träger kann sie nicht ablehnen"],
    ["≈ 36 %", "höhere Personalkosten nach neun Jahren bei 3,5 % Tarifsteigerung p. a., bei unverändertem Erlös"],
  ];
  let y = 3.34;
  tiles.forEach((t) => {
    card(s, M, y, 7.4, 0.90, INK_SOFT);
    s.addText(t[0], { x: M + 0.3, y: y + 0.11, w: 3.0, h: 0.32, fontFace: HFONT, fontSize: 16, bold: true, color: ORANGE, margin: 0 });
    s.addText(t[1], { x: M + 0.3, y: y + 0.44, w: 6.8, h: 0.4, fontFace: BFONT, fontSize: 11.5, color: PALE, margin: 0 });
    y += 1.02;
  });

  card(s, 8.35, 1.94, 4.36, 4.42, INK_CARD);
  s.addText("Aber: Es verschwindet nicht überall", {
    x: 8.65, y: 2.14, w: 3.8, h: 0.34, fontFace: BFONT, fontSize: 13.5, bold: true, color: WHITE, margin: 0,
  });
  s.addText([
    { text: "Verhandelte Vergütungen mit jährlicher Anpassung bleiben dort bestehen, wo der Kreis den Vereinbarungsweg geht — etwa Segeberg und Schleswig-Flensburg. Dort schützt § 124 SGB IX die Tarifbindung, und bei Nichteinigung entscheidet die Schiedsstelle.", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "Nur im Vergabeweg gilt beides nicht. SKP wird künftig beide Systeme parallel führen.", options: { bold: true } },
  ], { x: 8.65, y: 2.54, w: 3.8, h: 3.6, fontFace: BFONT, fontSize: 12, color: PALE, margin: 0 });

  s.addText("Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg, Kap. 5b; TED 246925-2024; §§ 124, 126 SGB IX; § 78g SGB VIII. Der Prozentwert ist eine eigene Modellrechnung.", {
    x: M, y: 6.60, w: 12.1, h: 0.4, fontFace: BFONT, fontSize: 9, color: "7FA3B0", margin: 0,
  });
  s.addNotes("Langsam vortragen. Die Ausschreibungslogik belohnt kurzfristig den, der zu knapp kalkuliert. Konsequenz: Vor jedem Angebot Laufzeit, Verlängerungsoption und mögliche Indexierungsklausel klären — und eine Preisuntergrenze definieren. Rechte Karte ist die Antwort auf die Frage, ob Tarifverhandlungen ganz verschwinden: nein, nur im Vergabeweg.");
}

/* =======================================================================
   13 — Was zu tun ist
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Was zu tun ist", "Investitionsbedarf und Zeitachse");

  const inv = [
    "Teamleitung mit pädagogischer Fachqualifikation aufbauen — ohne sie ist SKP nicht zuschlagsfähig",
    "Schriftliches Fortbildungs- und Qualifizierungskonzept — hartes Zuschlagskriterium",
    "Betriebshaftpflicht auf 10 Mio € prüfen; Referenzliste über vier Jahre anlegen",
    "Buchhaltung: Schulbegleitungsumsatz getrennt ausweisen; Präqualifizierung beantragen",
    "Kalkulationsmodell für Festbudgets ohne Tarifgleitklausel",
  ];
  card(s, M, 1.56, 6.35, 3.50);
  s.addText("Investitionsbedarf", { x: M + 0.3, y: 1.74, w: 5.7, h: 0.34, fontFace: BFONT, fontSize: 15, bold: true, color: TXT, margin: 0 });
  s.addText(inv.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i !== inv.length - 1 } })), {
    x: M + 0.3, y: 2.16, w: 5.75, h: 2.72, fontFace: BFONT, fontSize: 11.5, color: TXT, margin: 0, paraSpaceAfter: 8,
  });

  const mile = [
    ["läuft", "Hamburg: Kombinationsmaßnahmen ab Schuljahr 2026/27"],
    ["Sommer 2026", "1. KJHSRG: Kabinett, danach Bundestag und Bundesrat"],
    ["01.01.2027", "Frist: bis dahin muss das Bundesgesetz verkündet sein"],
    ["ab 2027/28", "Kreis Pinneberg: Klassenassistenz Stufe 1 (abgeleitet)"],
    ["01.01.2028", "Gesamtzuständigkeit der Jugendämter, Bildungsassistenz"],
  ];
  card(s, 7.28, 1.56, 5.43, 3.50, CARD2);
  s.addText("Meilensteine", { x: 7.58, y: 1.74, w: 4.8, h: 0.34, fontFace: BFONT, fontSize: 15, bold: true, color: TXT, margin: 0 });
  let my = 2.20;
  mile.forEach((m, i) => {
    s.addText(m[0], { x: 7.58, y: my, w: 1.55, h: 0.5, fontFace: BFONT, fontSize: 11, bold: true, color: i < 2 ? ORANGE : INK_SOFT, margin: 0, valign: "top" });
    s.addText(m[1], { x: 9.18, y: my, w: 3.3, h: 0.5, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0, valign: "top" });
    my += 0.56;
  });

  card(s, M, 5.22, W - 2 * M, 1.32, CARD);
  s.addText("Der günstigste Frühwarnindikator liegt bereits im Haus", {
    x: M + 0.3, y: 5.36, w: 11.5, h: 0.32, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Wenn ein Kreis beginnt, alle Bewilligungsbescheide einheitlich auf einen Schuljahresbeginn zu befristen, ist die Systemumstellung beschlossen — Monate bevor sie öffentlich wird. In Pinneberg war genau das das erste erkennbare Signal. Die laufenden Bescheide darauf zu prüfen, kostet nichts.",
    { x: M + 0.3, y: 5.70, w: 11.5, h: 0.76, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });

  srcNote(s, "Quellen: TED 246925-2024 (Eignungskriterien); Kabinettzeitplanung der Bundesregierung 25.06.2026; Art. 10 Abs. 3 KJSG; Trägerschreiben Kreis Pinneberg 12/2023 und 03/2024.");
  s.addNotes("Der Investitionsbedarf ist einmalig und überschaubar: im Wesentlichen eine Fachkraftstelle plus Konzeptarbeit. Das ist die Kaufthese — ein etabliertes regionales Unternehmen, dem die Vergabefähigkeit fehlt. Die lässt sich in 12 bis 18 Monaten herstellen.");
}

/* =======================================================================
   14 — Fazit (dunkel)
   ===================================================================== */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  darkTitle(s, "Fazit", "Ein wachsender Markt, dessen Zugang gerade neu vergeben wird");

  const cols = [
    ["Was für den Kauf spricht", [
      "Struktureller Nachfragezuwachs seit über zehn Jahren",
      "Überdurchschnittlich inklusive Regionen — größerer adressierbarer Markt",
      "Künftig größere, mehrjährige und planbarere Verträge",
      "Regionale Verankerung ist formales Zuschlagskriterium",
      "Weiterführende Schulen und Förderzentren bleiben länger stabil",
    ], BLUE],
    ["Was zu klären ist", [
      "Vergabefähigkeit: Haftpflicht, Referenzen, Qualifizierungskonzept",
      "Standortliste mit Umsatz je Schule, Schulart und Poolstatus",
      "Kalkulation für Festbudgets ohne Tarifausgleich",
      "Arbeitsrechtliche Folgen eines Standortverlusts (§ 613a BGB)",
      "Kreisscharfe Fallzahlen — nicht öffentlich, direkt zu erfragen",
    ], ORANGE],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 6.15;
    card(s, x, 2.00, 5.94, 3.72, INK_SOFT);
    s.addText(c[0], { x: x + 0.32, y: 2.22, w: 5.3, h: 0.36, fontFace: BFONT, fontSize: 15, bold: true, color: c[2], margin: 0 });
    s.addText(c[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== c[1].length - 1 } })), {
      x: x + 0.32, y: 2.66, w: 5.3, h: 2.86, fontFace: BFONT, fontSize: 12, color: "DCEAF0", margin: 0, paraSpaceAfter: 9,
    });
  });

  s.addText([
    { text: "Datenbasis:  ", options: { bold: true, color: WHITE } },
    { text: "Landtagsdrucksachen SH, EU-Vergabebekanntmachungen, Konzeptunterlagen Kreis Pinneberg, wissenschaftliche Evaluationen (DISW, Universität Regensburg), Bertelsmann Stiftung. Als Schätzung gekennzeichnete Werte sind vor einer Kaufentscheidung zu verifizieren.", options: { color: "9FBECA" } },
  ], { x: M, y: 5.96, w: 12.1, h: 0.9, fontFace: BFONT, fontSize: 10.5, margin: 0 });

  s.addNotes("Kaufthese in einem Satz: ein etabliertes regionales Unternehmen in einem strukturell wachsenden Markt, dem eine überschaubare, aber zeitkritische Investition in Vergabefähigkeit fehlt. Offene Punkte aktiv benennen.");
}

/* =======================================================================
   15 — Anhang-Trenner
   ===================================================================== */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("ANHANG", {
    x: M, y: 2.4, w: 11, h: 0.4, fontFace: BFONT, fontSize: 12, bold: true, color: ORANGE, charSpacing: 2, margin: 0,
  });
  s.addText("Zum Nachschlagen", {
    x: M, y: 2.82, w: 11.6, h: 0.8, fontFace: HFONT, fontSize: 36, bold: true, color: WHITE, margin: 0,
  });
  const idx = [
    ["16", "Vergütungsmechanik — Zuschläge, Abschläge, Klassenfahrten"],
    ["17", "Zwei Wege zum Vertrag — Vergabe oder Vereinbarung"],
    ["18", "Vergabefähigkeit im Klartext"],
    ["19", "Glossar"],
  ];
  let y = 3.90;
  idx.forEach((r) => {
    s.addText(r[0], { x: M, y: y, w: 0.6, h: 0.36, fontFace: HFONT, fontSize: 16, bold: true, color: ORANGE, margin: 0 });
    s.addText(r[1], { x: M + 0.7, y: y, w: 10.5, h: 0.36, fontFace: BFONT, fontSize: 15, color: PALE, margin: 0 });
    y += 0.56;
  });
  s.addNotes("Ab hier nur auf Nachfrage zeigen.");
}

/* =======================================================================
   16 — Anhang A1: Vergütungsmechanik
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 1", "Vergütungsmechanik: Zuschläge, Abschläge und Klassenfahrten");

  card(s, M, 1.56, 6.05, 2.30);
  s.addText("Zuschläge auf die Personalkosten", { x: M + 0.28, y: 1.72, w: 5.5, h: 0.32, fontFace: BFONT, fontSize: 13.5, bold: true, color: BLUE, margin: 0 });
  s.addText([
    { text: "Kreis Pinneberg: 10 % Sachkosten + 5 % Verwaltungskostenzuschlag", options: { bullet: true, breakLine: true } },
    { text: "Rahmenleistungsbeschreibung NRW: 10 % Leitung und Verwaltung + 5 % Sachkosten (Plausibilitätswerte)", options: { bullet: true, breakLine: true } },
    { text: "Einen bundesweiten Standard gibt es nicht — die Sätze sind vertragsindividuell", options: { bullet: true } },
  ], { x: M + 0.28, y: 2.10, w: 5.5, h: 1.62, fontFace: BFONT, fontSize: 11.5, color: TXT, margin: 0, paraSpaceAfter: 6 });

  card(s, 6.98, 1.56, 5.73, 2.30, CARD2);
  s.addText("Abschläge — die relevantere Seite", { x: 7.26, y: 1.72, w: 5.2, h: 0.32, fontFace: BFONT, fontSize: 13.5, bold: true, color: ORANGE, margin: 0 });
  s.addText([
    { text: "Abzug vorhandener Landeskräfte (Schulassistenz, FSJ, Bufdi) mindert die finanzierte Stellenzahl", options: { bullet: true, breakLine: true } },
    { text: "Rückforderungsvorbehalt bei nicht erbrachten Leistungen", options: { bullet: true, breakLine: true } },
    { text: "Kein Tarifausgleich über die Laufzeit — faktisch 3–4 % Realabschlag pro Jahr", options: { bullet: true } },
  ], { x: 7.26, y: 2.10, w: 5.2, h: 1.62, fontFace: BFONT, fontSize: 11.5, color: TXT, margin: 0, paraSpaceAfter: 6 });

  card(s, M, 4.00, W - 2 * M, 2.40);
  s.addText("Klassenfahrten und Ausflüge — der konkreteste Einzelposten", {
    x: M + 0.3, y: 4.18, w: 11.5, h: 0.32, fontFace: BFONT, fontSize: 14, bold: true, color: TXT, margin: 0,
  });
  const kf = [
    ["Heute (Einzelfall)", "8 Stunden je Tag zusätzlich abrechenbar, ohne vorherigen Antrag; das monatliche Stundensoll erhöht sich entsprechend. Eine Bestätigung der Schule genügt. Kein Sondersatz — mehr Stunden zum normalen Satz. Zum Vergleich: 5 Tage × 8 h = 40 Stunden gegenüber 24–30 regulären Wochenstunden."],
    ["Im Pool", "Ausdrücklich im Festbudget enthalten („Hierzu zählen auch Komponenten wie Klassenfahrten oder Ausflüge“). Der Aufwand bleibt, die Zusatzvergütung entfällt. Ostholstein berichtet, schulinterne Lösungen seien „nicht immer zufriedenstellend“ — Extrabewilligungen werden angeregt."],
    ["Offen", "Ob Fahrt-, Übernachtungs- und Verpflegungskosten der Begleitperson erstattet werden, ist in keiner öffentlichen Quelle geregelt. Das bestimmt die jeweilige Vereinbarung — Due-Diligence-Frage."],
  ];
  let y = 4.56;
  kf.forEach((r, i) => {
    s.addText(r[0], { x: M + 0.3, y: y, w: 1.9, h: 0.56, fontFace: BFONT, fontSize: 11.5, bold: true, color: i === 2 ? ORANGE : TXT, margin: 0, valign: "top" });
    s.addText(r[1], { x: M + 2.3, y: y, w: 9.5, h: 0.56, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0, valign: "top" });
    y += 0.60;
  });

  srcNote(s, "Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg, Kap. 5b; Trägerschreiben 20.03.2024; Rahmenleistungsbeschreibung Schulbegleitung NRW (Anlage A.2.6); DISW-Evaluation Ostholstein.");
  s.addNotes("Anhangfolie. Antwort auf die Frage, ob Klassenfahrten bisher vergütet wurden: ja, aber als zusätzliche Stunden, nicht als Zuschlag.");
}

/* =======================================================================
   17 — Anhang A2: Zwei Wege zum Vertrag
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 2", "Zwei Wege zum Vertrag — und nur einer friert den Preis ein");

  const rows = [
    ["", "Vergaberecht", "Vereinbarungsrecht"],
    ["Auswahl", "EU-Ausschreibung, Verhandlungsverfahren", "Interessenbekundung, dann Verhandlung"],
    ["Preisfindung", "Angebotspreis, einmalig", "Vergütungsvereinbarung, prospektiv je Periode"],
    ["Laufzeit", "bis 5 Jahre + einseitige Option", "in der Regel jährlich"],
    ["Tarifsteigerung", "trägt der Auftragnehmer", "§ 124 SGB IX schützt die Tarifbindung"],
    ["Bei Nichteinigung", "kein Verfahren — es gibt nichts zu verhandeln", "Schiedsstelle (§ 126 SGB IX, § 78g SGB VIII)"],
    ["Vertragsende", "Neuausschreibung, kein Bestandsschutz", "Vereinbarung läuft weiter"],
    ["Beispiele", "Kreis Pinneberg, Essen, Rhein-Erft", "Segeberg, Schleswig-Flensburg"],
  ];
  compareTable(s, 1.58, rows, [M, M + 3.1, M + 7.6], [3.0, 4.4, 4.5], [MUT, ORANGE, BLUE], 0.50, 0.56, 11.5);

  card(s, M, 5.96, W - 2 * M, 0.78, CARD2);
  s.addText([
    { text: "Wichtig:  ", options: { bold: true, color: TXT } },
    { text: "Verhandelte Vergütungen verschwinden weder bundesweit noch regionsweit — sondern vertragsweise. Welchen Weg ein Kreis wählt, entscheidet über Preisrisiko und Verhandlungsmacht. SKP wird beide Welten parallel führen.", options: { color: TXT } },
  ], { x: M + 0.3, y: 5.96, w: 11.5, h: 0.78, fontFace: BFONT, fontSize: 12, margin: 0, valign: "middle" });

  srcNote(s, "Quellen: §§ 123–126 SGB IX; §§ 78a–78g SGB VIII; § 20 Abs. 2 Landesrahmenvertrag SH; TED 246925-2024; Kreispräsentation Schleswig-Flensburg.", 7.00);
  s.addNotes("Anhangfolie. Antwort auf die Frage, ob jährliche Tarifverhandlungen verschwinden: nur im Vergabeweg. Ergänzend: Auch im Vereinbarungsrecht sind nachträgliche Ausgleiche unzulässig — verhandelt wird immer prospektiv.");
}

/* =======================================================================
   18 — Anhang A3: Vergabefähigkeit
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 3", "Vergabefähigkeit im Klartext — die Ausschlusskriterien");

  const items = [
    ["Betriebshaftpflicht 10 Mio €", "Mindestdeckungssumme im Pinneberger Verfahren. Bei Bietergemeinschaften muss jedes Mitglied sie nachweisen. Nicht kurzfristig heilbar."],
    ["Getrennte Umsatzausweisung", "Gefordert ist der Gesamtumsatz UND separat der Umsatz der Leistungsart Schulbegleitung, je drei Geschäftsjahre. Die Buchhaltung muss das trennen können — nachträglich ist es kaum rekonstruierbar."],
    ["Präqualifizierung", "Eintrag im amtlichen Verzeichnis präqualifizierter Unternehmen (DIHK, pq-vol.de). Eignungsnachweise werden einmal zentral hinterlegt; im Verfahren genügt dann die PQ-Nummer statt eines Stapels Bescheinigungen. Ohne PQ: Formblatt je Verfahren, Nachreichfrist sechs Kalendertage."],
    ["Referenzliste über vier Jahre", "Je Schulbegleitungsverhältnis mit Leistungsumfang, öffentlichem Auftraggeber und Ansprechpartner. Im laufenden Verfahren nicht mehr aufzubauen."],
    ["Keine Honorarkräfte", "„Es werden keine Honorarkräfte eingesetzt“ — Festanstellung ist Qualitätsmerkmal und Vertragsbedingung."],
  ];
  let y = 1.56;
  items.forEach((it) => {
    card(s, M, y, W - 2 * M, 0.98);
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.28, y: y + 0.34, w: 0.3, h: 0.3, fill: { color: ORANGE }, rectRadius: 0.05, line: { color: ORANGE } });
    s.addText(it[0], { x: M + 0.72, y: y + 0.14, w: 3.0, h: 0.70, fontFace: BFONT, fontSize: 12.5, bold: true, color: TXT, margin: 0, valign: "middle" });
    s.addText(it[1], { x: M + 3.9, y: y + 0.14, w: 7.85, h: 0.70, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0, valign: "middle" });
    y += 1.06;
  });
  srcNote(s, "Quelle: EU-Bekanntmachung TED 246925-2024, Eignungskriterien; Konzept „Klassenassistenz“ Kreis Pinneberg, Kap. 5a.", 6.96);
  s.addNotes("Anhangfolie. Klartext zu den Begriffen aus der Risikofolie. Die ersten beiden Punkte sind diejenigen, die man nicht mehr heilen kann, wenn das Verfahren einmal läuft.");
}

/* =======================================================================
   19 — Anhang A4: Glossar
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 4", "Glossar");

  const terms = [
    ["Eingliederungshilfe", "Sozialleistung für Menschen mit Behinderung; Schulbegleitung ist eine Form davon."],
    ["Schulbegleitung", "Individuelle Unterstützung eines Kindes im Unterricht. Keine geschützte Berufsbezeichnung, kein Fachkräftegebot."],
    ["Schulische Assistenz", "Systemische Unterstützung aller Kinder, nur an Grundschulen, vom Land finanziert. Nicht dasselbe wie Schulbegleitung."],
    ["Poolmodell", "Mehrere Kinder werden gemeinsam begleitet; die Schule erhält ein Budget statt Einzelbewilligungen."],
    ["Infrastrukturangebot / Bildungsassistenz", "Pool außerhalb des individuellen Sozialrechts — kein Antrag, kein Bescheid. Im Bundesentwurf § 80a SGB VIII-E."],
    ["Leistungs- und Vergütungs­vereinbarung", "Vertragsform des Sozialrechts (§§ 123 ff. SGB IX): Leistung, Qualität und Preis werden verhandelt, nicht ausgeschrieben."],
    ["Teilnahmewettbewerb", "Erste Stufe einer EU-Ausschreibung: Eignungsprüfung. Erst danach werden Angebote eingeholt."],
    ["Los", "Teilpaket eines Auftrags (z. B. eine Stadt). Kein Losentscheid — jedes Los wird nach Qualität und Preis vergeben."],
    ["Förder-, Inklusions-, Exklusionsquote", "Anteil aller Schüler:innen mit Förderbedarf / davon an Regelschulen / davon an Förderschulen."],
  ];
  let y = 1.52;
  terms.forEach((t, i) => {
    if (i % 2 === 0) card(s, M, y, W - 2 * M, 0.56, CARD);
    s.addText(t[0], { x: M + 0.28, y: y + 0.06, w: 3.5, h: 0.44, fontFace: BFONT, fontSize: 11.5, bold: true, color: TXT, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 3.95, y: y + 0.06, w: 7.8, h: 0.44, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0, valign: "middle" });
    y += 0.58;
  });
  srcNote(s, "Definitionen nach SGB VIII / SGB IX, Landtag SH Drs. 20/2643, BAGüS-Orientierungshilfe 2019 und § 97 Abs. 4 GWB.");
  s.addNotes("Anhangfolie für Begriffsfragen.");
}

pres.writeFile({ fileName: "/home/user/repository/praesentation/SKP-Marktanalyse-Schulbegleitung.pptx" })
  .then((f) => console.log("geschrieben:", f));
