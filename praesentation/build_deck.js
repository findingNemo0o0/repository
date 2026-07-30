const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "SKP Marktanalyse";
pres.title = "Schulbegleitung im Umbruch";

// ---- Palette -------------------------------------------------------------
const INK = "0E2E38";
const INK_SOFT = "1B4A57";
const TXT = "182E36";
const MUT = "5F7278";
const CARD = "F1F6F8";
const CARD2 = "E7EFF3";
const BLUE = "2A78D6";     // validierte Serienfarbe 1
const ORANGE = "EB6834";   // validierte Serienfarbe 2
const WHITE = "FFFFFF";

const HFONT = "Cambria";
const BFONT = "Calibri";
const M = 0.62;
const W = 13.33;

// ---- Helfer ---------------------------------------------------------------
const shadow = () => ({ type: "outer", color: "0E2E38", blur: 10, offset: 2, angle: 90, opacity: 0.10 });

function titleBar(slide, kicker, title) {
  slide.addText(kicker.toUpperCase(), {
    x: M, y: 0.32, w: 11, h: 0.26, fontFace: BFONT, fontSize: 11, bold: true,
    color: ORANGE, charSpacing: 1.6, margin: 0,
  });
  slide.addText(title, {
    x: M, y: 0.56, w: W - 2 * M, h: 0.80, fontFace: HFONT, fontSize: 26, bold: true,
    color: TXT, margin: 0, valign: "top",
  });
}
function keyline(slide, text, y) {
  slide.addText(text, {
    x: M, y: y, w: W - 2 * M, h: 0.40, fontFace: BFONT, fontSize: 13, italic: true, color: INK_SOFT, margin: 0,
  });
}
function srcNote(slide, text, y) {
  slide.addText(text, {
    x: M, y: y || 6.92, w: W - 2 * M, h: 0.34, fontFace: BFONT, fontSize: 9, color: MUT, margin: 0,
  });
}
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, fill: { color: fill || CARD }, rectRadius: 0.08, line: { color: fill || CARD }, shadow: shadow(),
  });
}
function chip(slide, x, y, w, label, color) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.32, fill: { color }, rectRadius: 0.06, line: { color } });
  slide.addText(label, {
    x, y, w, h: 0.32, fontFace: BFONT, fontSize: 10, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0, charSpacing: 0.8,
  });
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
  s.addText("Stand: Juli 2026  ·  Folien 7 und 8 sind Anhang zum Nachschlagen", {
    x: M, y: 6.42, w: 8, h: 0.3, fontFace: BFONT, fontSize: 11, color: "7FA3B0", margin: 0,
  });
  s.addNotes("Einstieg in 30 Sekunden: SKP verkauft Schulbegleitung an Kreise. Der Bedarf steigt seit zehn Jahren ununterbrochen. Gleichzeitig stellen die Kostenträger gerade das Vergabesystem um — von Einzelfallbewilligung auf Standortvertrag. Wer die Umstellung übersteht, hat ein größeres und planbareres Geschäft. Wer sie verpasst, verliert Standorte vollständig.");
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
    card(s, x, 1.50, 3.90, 2.42);
    chip(s, x + 0.26, 1.70, 2.0, r[1], r[2]);
    s.addText(r[0], { x: x + 0.26, y: 2.14, w: 3.38, h: 0.38, fontFace: BFONT, fontSize: 14, bold: true, color: TXT, margin: 0 });
    s.addText(r[3], { x: x + 0.26, y: 2.54, w: 3.38, h: 1.24, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });
  });

  card(s, M, 4.08, W - 2 * M, 1.16, CARD2);
  s.addText("Geltungsbereich: alle Schularten", {
    x: M + 0.3, y: 4.22, w: 11.5, h: 0.3, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Schulbegleitung gilt für Klasse 1 bis 13 — Regelschule wie Förderschule, berufsbildende Schulen und den offenen Ganztag (§ 112 SGB IX). Die Poolreform beginnt meist an Grundschulen, weil dort die Schulische Assistenz des Landes ansetzt; sie ist aber nicht darauf beschränkt.",
    { x: M + 0.3, y: 4.54, w: 11.5, h: 0.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  card(s, M, 5.38, W - 2 * M, 1.16);
  s.addText("So entsteht heute Umsatz", {
    x: M + 0.3, y: 5.52, w: 11.5, h: 0.3, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Vergütet wird pro Kind und Stunde auf Basis eines Bewilligungsbescheids. Klassenfahrten und Ausflüge sind als Zusatzstunden abrechenbar — in Pinneberg 8 Stunden je Tag, ohne vorherigen Antrag; das monatliche Stundensoll erhöht sich entsprechend. Im Pool entfällt diese Zusatzvergütung.",
    { x: M + 0.3, y: 5.84, w: 11.5, h: 0.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  srcNote(s, "Quellen: SGB VIII, SGB IX; Landtag SH Drs. 20/2643; Trägerschreiben Kreis Pinneberg 20.03.2024; Empfehlungen Bayerischer Bezirketag/StMUK 2025.");
  s.addNotes("Drei Akteure, ein Vertragspartner. Wichtigster Punkt ist die dritte Karte: Die Schule unterschreibt nie, entscheidet aber mit. Deshalb ist die Beziehung zu den Schulleitungen der eigentliche Vermögenswert. Zweiter Punkt: Schulbegleitung ist nicht auf die Grundschule beschränkt — nur die Poolreform startet dort.");
}

/* =======================================================================
   3 — Nachfrage und Standort (zusammengeführt)
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Nachfrage und Standort", "Der Bedarf wächst — und SKPs Regionen sind besonders inklusiv");

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [1.1, 3.5, 3.4] },
    { name: "Exklusionsquote (Förderschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [4.8, 4.3, 4.2] },
  ], {
    x: M, y: 1.44, w: 5.88, h: 3.30,
    barDir: "col", barGrouping: "stacked", chartColors: [BLUE, ORANGE],
    showTitle: true, title: "Förderquote Deutschland, in % aller Schüler:innen",
    titleFontSize: 11.5, titleColor: MUT, titleFontFace: BFONT,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 10,
    dataLabelColor: WHITE, dataLabelFontFace: BFONT, dataLabelFormatCode: "0.0",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 10.5, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 9.5, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMaxVal: 9, valAxisMajorUnit: 3,
  });

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [3.4, 4.4, 5.2] },
    { name: "Exklusionsquote (Förderschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [4.2, 2.3, 2.7] },
  ], {
    x: 6.83, y: 1.44, w: 5.88, h: 3.30,
    barDir: "col", barGrouping: "stacked", chartColors: [BLUE, ORANGE],
    showTitle: true, title: "Schuljahr 2022/23, im Ländervergleich",
    titleFontSize: 11.5, titleColor: MUT, titleFontFace: BFONT,
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 10,
    dataLabelColor: WHITE, dataLabelFontFace: BFONT, dataLabelFormatCode: "0.0",
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 10, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 9.5, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMaxVal: 9, valAxisMajorUnit: 3,
  });

  const facts = [
    ["5,9 → 7,6 %", "Förderquote Deutschland seit 2008/09. Drei Viertel des Zuwachses sind neue Diagnosen — nicht Wechsler von der Förderschule.", BLUE],
    ["4,4 / 5,2 %", "Inklusionsquote Schleswig-Holstein und Hamburg gegenüber 3,4 % im Bund. Schulbegleitung entsteht fast nur an Regelschulen.", BLUE],
    ["≈ 2 Jahre", "Durchschnittliche Hilfedauer je Kind (§ 35a SGB VIII) — leicht steigend: 22 Monate (2013) auf 24,6 Monate (2023).", ORANGE],
  ];
  facts.forEach((f, i) => {
    const x = M + i * 4.06;
    card(s, x, 4.86, 3.90, 1.24);
    s.addText(f[0], { x: x + 0.26, y: 4.98, w: 3.38, h: 0.42, fontFace: HFONT, fontSize: 21, bold: true, color: f[2], margin: 0 });
    s.addText(f[1], { x: x + 0.26, y: 5.42, w: 3.38, h: 0.60, fontFace: BFONT, fontSize: 10.5, color: MUT, margin: 0 });
  });

  keyline(s, "Kehrseite: Beide Regionen haben die niedrigsten Förderschulanteile — die Reserve an Wechslern ist weitgehend ausgeschöpft.", 6.22);
  srcNote(s, "Quellen: Bertelsmann Stiftung auf Grundlage KMK (2024) / Klemm, Schuljahre 2008/09–2022/23 (ohne Saarland); Hilfedauer: AKJStat/HzE-Monitor TU Dortmund, beendete Hilfen nach § 35a SGB VIII — alle Hilfearten, nicht nur Schulbegleitung.", 6.70);
  s.addNotes("Zwei Botschaften auf einer Folie. Erstens: Der Nachfragesockel wächst strukturell — und der Zuwachs kommt zu drei Vierteln aus neuen Diagnosen, nicht aus dem Abbau der Förderschulen. Zweitens: SH und Hamburg liegen bei der Inklusionsquote deutlich über dem Bund, der adressierbare Markt ist hier also größer. Die Hilfedauer von rund zwei Jahren ist die Rechengröße für die nächste Folie.");
}

/* =======================================================================
   4 — Marktvolumen, Löhne, Kosten je Kind
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Marktvolumen und Kosten", "Ausgaben wachsen schneller als Fälle — Löhne schneller als beide");

  s.addChart(pres.ChartType.line, [
    { name: "Ausgaben", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 122, 148, 170, 199, 234, 286, 384] },
    { name: "Leistungsempfänger", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 111, 131, 144, 162, 189, 204, 245] },
  ], {
    x: M, y: 1.42, w: 5.88, h: 3.02,
    chartColors: [ORANGE, BLUE], lineDataSymbol: "circle", lineDataSymbolSize: 6, lineSize: 3,
    showTitle: true, title: "Schulbegleitung Schleswig-Holstein, Index 2014 = 100",
    titleFontSize: 11.5, titleColor: MUT, titleFontFace: BFONT,
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 10, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 9.5, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 80, valAxisMaxVal: 400, valAxisMajorUnit: 80,
  });

  s.addChart(pres.ChartType.bar, [
    { name: "Mindestlohn", labels: ["2015", "2017", "2019", "2021", "2023", "2024", "2025", "2026", "2027"], values: [8.50, 8.84, 9.19, 9.60, 12.00, 12.41, 12.82, 13.90, 14.60] },
  ], {
    x: 6.83, y: 1.42, w: 5.88, h: 3.02,
    barDir: "col", chartColors: [ORANGE],
    showTitle: true, title: "Gesetzlicher Mindestlohn, Euro brutto je Stunde",
    titleFontSize: 11.5, titleColor: MUT, titleFontFace: BFONT,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9,
    dataLabelColor: TXT, dataLabelFontFace: BFONT, dataLabelFormatCode: "0.00",
    showLegend: false,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 10, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 9.5, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: 18, valAxisMajorUnit: 6,
  });

  card(s, M, 4.58, 5.88, 2.02, CARD2);
  s.addText("Was der Kreis je Kind zahlt", {
    x: M + 0.28, y: 4.70, w: 5.3, h: 0.30, fontFace: BFONT, fontSize: 13, bold: true, color: TXT, margin: 0,
  });
  s.addText("gerechnet auf die durchschnittliche Hilfedauer von 2,05 Jahren", {
    x: M + 0.28, y: 4.98, w: 5.3, h: 0.24, fontFace: BFONT, fontSize: 10, italic: true, color: MUT, margin: 0,
  });
  s.addText([
    { text: "Hamburg, alle Schularten (2025/26)", options: { breakLine: true, color: MUT } },
    { text: "10.500 € je Jahr   →   21.500 € je Kind", options: { breakLine: true, bold: true, color: TXT } },
    { text: "Kreis Pinneberg, Grundschule (Ø 2017–2020)", options: { breakLine: true, color: MUT } },
    { text: "18.700 € je Jahr   →   38.300 € je Kind", options: { bold: true, color: TXT } },
  ], { x: M + 0.28, y: 5.26, w: 5.3, h: 1.20, fontFace: BFONT, fontSize: 11.5, margin: 0, lineSpacingMultiple: 1.12 });

  card(s, 6.83, 4.58, 5.88, 2.02);
  s.addText("Was ein Kind an Personalkosten kostet", {
    x: 7.11, y: 4.70, w: 5.3, h: 0.30, fontFace: BFONT, fontSize: 13, bold: true, color: TXT, margin: 0,
  });
  s.addText("Mindestlohn 13,90 € + 22 % Arbeitgeberanteil = 17,00 € je Stunde", {
    x: 7.11, y: 4.98, w: 5.3, h: 0.24, fontFace: BFONT, fontSize: 10, italic: true, color: MUT, margin: 0,
  });
  s.addText([
    { text: "20 Wochenstunden × 39 Schulwochen = 780 h", options: { breakLine: true, color: MUT } },
    { text: "13.200 € je Jahr   →   27.100 € je Kind", options: { breakLine: true, bold: true, color: ORANGE } },
    { text: "Gegenprobe: Hamburgs Durchschnitt trägt rund 16 Wochenstunden, Pinnebergs Satz rund 28 — genau der Korridor der Klassenassistenz (24–30 h).", options: { color: MUT } },
  ], { x: 7.11, y: 5.26, w: 5.3, h: 1.20, fontFace: BFONT, fontSize: 11.5, margin: 0, lineSpacingMultiple: 1.12 });

  srcNote(s, "Quellen: Landtag SH Drs. 20/2643; BSFB Hamburg; Konzept Klassenassistenz Kreis Pinneberg; Mindestlohnkommission/BMAS. Die Kosten je Kind sind eigene Berechnungen auf Basis der ausgewiesenen Annahmen. Der Mindestlohn ist die Untergrenze — das durchschnittliche Bruttogehalt in der Schulbegleitung liegt bei 2.700–3.100 € im Monat.", 6.70);
  s.addNotes("Drei Aussagen. Erstens: Die Ausgaben wachsen fast doppelt so schnell wie die Fallzahlen — nicht die Menge treibt die Kosten, sondern der Preis je Fall. Zweitens: Der Mindestlohn ist seit 2015 um 72 Prozent gestiegen, allein 2026 um 8,4 Prozent. Drittens die Gegenüberstellung unten. Die Gegenprobe rechts ist wichtig: Pinnebergs historischer Satz entspricht 28 Wochenstunden und landet damit exakt im Korridor der geplanten Klassenassistenz — das bestätigt die Rechnung.");
}

/* =======================================================================
   5 — Die vier Modelle
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Die Modelle", "Vier Wege, die derzeit zur Debatte stehen");

  const models = [
    ["Klassenmodell", "Kreis Pinneberg", "UPSIDE", BLUE,
      "Budget = Anzahl der Klassen, abzüglich vorhandener Landeskräfte. Eine Assistenz je Klasse, 24 h (Kl. 1–2) bzw. 30 h (Kl. 3–4), Vergütung nach S2 Stufe 5.",
      "Stellenzahl steigt deutlich (223 → 431), aber Satz und Stunden sind gedeckelt. Nur Grundschulen. Bereits zweimal verschoben."],
    ["Poolmodell", "Kreis Ostholstein", "BASISANNAHME", ORANGE,
      "Budget = Summe der bisherigen Einzelbedarfe, schulscharf. Ein koordinierender Träger je Schule; Mitarbeitende wechseln zu ihm.",
      "Gleiches Volumen, mehr Kinder, weniger Stunden je Kind. „Konzentration der Leistungserbringer“ ist erklärtes Projektziel."],
    ["Kombinationsmaßnahmen", "Hamburg", "LÄUFT SEIT 2026/27", ORANGE,
      "Eine Kraft für mehrere Kinder einer Klasse oder Schule. Freiwilligendienstleistende werden vollständig darüber organisiert.",
      "Stunden und Qualifikation je Kind sinken. Höher qualifizierte Kräfte nur noch in begründeten Ausnahmefällen."],
    ["Bildungsassistenz", "Bundesentwurf 1. KJHSRG", "OFFEN — AB 2028", BLUE,
      "§ 80a SGB VIII-E: reine Planungspflicht ohne Personalschlüssel. Individualanspruch nur noch, wenn ausschließlich eine 1:1-Begleitung hilft.",
      "Kein gesetzlicher Mindestumfang — ein Kreis kann ein deutlich dünneres Angebot aufstellen als Pinneberg."],
  ];
  models.forEach((m, i) => {
    const x = M + (i % 2) * 6.15;
    const y = 1.44 + Math.floor(i / 2) * 2.58;
    card(s, x, y, 5.94, 2.42);
    s.addText(m[0], { x: x + 0.28, y: y + 0.14, w: 3.4, h: 0.32, fontFace: BFONT, fontSize: 14.5, bold: true, color: TXT, margin: 0 });
    s.addText(m[1], { x: x + 0.28, y: y + 0.46, w: 3.4, h: 0.28, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0 });
    chip(s, x + 3.76, y + 0.16, 1.90, m[2], m[3]);
    s.addText(m[4], { x: x + 0.28, y: y + 0.82, w: 5.38, h: 0.74, fontFace: BFONT, fontSize: 11, color: TXT, margin: 0 });
    s.addText(m[5], { x: x + 0.28, y: y + 1.60, w: 5.38, h: 0.66, fontFace: BFONT, fontSize: 11, italic: true, color: MUT, margin: 0 });
  });

  srcNote(s, "Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg; DISW-Gesamtevaluation Ostholstein 2023; BSFB Hamburg; Referentenentwurf 1. KJHSRG (§§ 35d Abs. 4, 80a SGB VIII-E), Kabinettbefassung Sommer 2026.", 6.68);
  s.addNotes("Vier Modelle, ein Muster: Alle ersetzen die Einzelfallbewilligung durch ein Budget. Der Unterschied liegt in der Bemessung — und die entscheidet über SKPs Volumen. Pinneberg rechnet nach Klassen und ist das großzügigste Modell; es ist Upside, nicht Basis. Ostholstein ist die realistische Annahme. Hamburg läuft bereits. Und der Bundesentwurf schreibt gar keinen Schlüssel vor — das ist die eigentliche Unsicherheit.");
}

/* =======================================================================
   6 — Chancen und Risiken
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Bewertung", "Chancen und Risiken auf einen Blick");

  const chancen = [
    ["Der Gesamtmarkt wächst weiter", "Fallzahlen und Ausgaben steigen in allen belastbaren Datenreihen. Kein Kreis erwartet einen Rückgang."],
    ["Größere, planbarere Aufträge", "Statt hunderter Einzelbewilligungen wenige Standortverträge über Jahre — mit geringerem Verwaltungsaufwand je Umsatzeuro."],
    ["Regionale Verankerung zählt formal", "Kenntnis der regionalen Strukturen und bestehende Kooperationen sind nachzuweisende Zuschlagskriterien."],
    ["Qualität schlägt Preis", "Kreis Düren gewichtet Qualität mit 60 von 100 Punkten. Ein dokumentiertes Qualifizierungskonzept ist der billigste Differenzierer."],
    ["Zeitfenster und stabile Segmente", "Steinburg und Neumünster haben noch kein Modell. Weiterführende Schulen und Förderzentren bleiben am längsten stabil."],
  ];
  const risiken = [
    ["Standortverlust ist total", "Wer den Zuschlag nicht bekommt, verliert das gesamte Volumen dort — zum Schuljahresbeginn, ohne Übergang."],
    ["Personal wandert mit", "In Ostholstein dokumentiert: Mitarbeitende werden an den koordinierenden Träger „übergeben“. Rechtlich freiwillig, faktisch die Regel."],
    ["Kein Rechtsschutz", "Das OVG Schleswig wies die Klage von vier Trägern gegen das Pinneberger Vergabeverfahren zurück — unanfechtbar."],
    ["Preisbindung ohne Tarifausgleich", "Pinneberg: Festpreis bis 31.07.2029, einseitig um vier Jahre verlängerbar. Tarifsteigerungen trägt allein der Träger."],
    ["Vergabefähigkeit fehlt oft", "10 Mio € Betriebshaftpflicht, Referenzliste über vier Jahre, getrennte Umsatzausweisung, Präqualifizierung — Details auf Folie 7."],
  ];

  [["Chancen", chancen, BLUE, CARD], ["Risiken", risiken, ORANGE, CARD2]].forEach((col, ci) => {
    const x = M + ci * 6.15;
    card(s, x, 1.44, 5.94, 5.02, col[3]);
    s.addText(col[0], { x: x + 0.30, y: 1.58, w: 5.3, h: 0.36, fontFace: BFONT, fontSize: 16, bold: true, color: col[2], margin: 0 });
    let y = 2.04;
    col[1].forEach((it) => {
      s.addShape(pres.ShapeType.roundRect, {
        x: x + 0.30, y: y + 0.05, w: 0.22, h: 0.22, fill: { color: col[2] }, rectRadius: 0.04, line: { color: col[2] },
      });
      s.addText(it[0], { x: x + 0.64, y: y, w: 4.95, h: 0.28, fontFace: BFONT, fontSize: 12.5, bold: true, color: TXT, margin: 0 });
      s.addText(it[1], { x: x + 0.64, y: y + 0.30, w: 4.95, h: 0.52, fontFace: BFONT, fontSize: 10.5, color: MUT, margin: 0 });
      y += 0.86;
    });
  });

  srcNote(s, "Quellen: Landtag SH Drs. 20/2643; TED-Bekanntmachungen Pinneberg, Düren, Euskirchen, Rhein-Erft, Essen; OVG Schleswig 5 MB 7/24; DISW-Evaluation Ostholstein.", 6.62);
  s.addNotes("Beide Spalten ehrlich vortragen. Alle Chancen hängen an einer Bedingung: Das Unternehmen muss vergabefähig werden. Bei den Risiken sind Punkt 3 und 4 die entscheidenden — es gibt keinen Rechtsweg gegen die Umstellung, und der Festpreis kann bis zu neun Jahre laufen. Wer diese beiden Punkte selbst vorträgt, gewinnt Vertrauen.");
}

/* =======================================================================
   7 — Anhang: Vergabefähigkeit
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 1", "Vergabefähigkeit im Klartext — die Ausschlusskriterien");

  const items = [
    ["Betriebshaftpflicht 10 Mio €", "Mindestdeckungssumme im Pinneberger Verfahren. Bei Bietergemeinschaften muss jedes Mitglied sie nachweisen. Nicht kurzfristig heilbar."],
    ["Getrennte Umsatzausweisung", "Gefordert ist der Gesamtumsatz UND separat der Umsatz der Leistungsart Schulbegleitung, je drei Geschäftsjahre. Die Buchhaltung muss das trennen können — nachträglich kaum rekonstruierbar."],
    ["Präqualifizierung", "Eintrag im amtlichen Verzeichnis präqualifizierter Unternehmen (DIHK, pq-vol.de). Nachweise werden einmal zentral hinterlegt; im Verfahren genügt dann die PQ-Nummer. Ohne PQ: Formblatt je Verfahren, Nachreichfrist sechs Kalendertage."],
    ["Referenzliste über vier Jahre", "Je Schulbegleitungsverhältnis mit Leistungsumfang, öffentlichem Auftraggeber und Ansprechpartner. Im laufenden Verfahren nicht mehr aufzubauen."],
    ["Keine Honorarkräfte", "„Es werden keine Honorarkräfte eingesetzt“ — Festanstellung ist Qualitätsmerkmal und Vertragsbedingung."],
    ["Teamleitung als Fachkraft", "Ausschließlich Sozialpädagog:innen, Pädagog:innen, Erzieher:innen oder gleichwertig. Schlüssel 1:15. Ohne diese Funktion kein Zuschlag."],
  ];
  let y = 1.44;
  items.forEach((it) => {
    card(s, M, y, W - 2 * M, 0.86);
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.28, y: y + 0.30, w: 0.28, h: 0.28, fill: { color: ORANGE }, rectRadius: 0.05, line: { color: ORANGE } });
    s.addText(it[0], { x: M + 0.70, y: y + 0.10, w: 3.0, h: 0.66, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0, valign: "middle" });
    s.addText(it[1], { x: M + 3.86, y: y + 0.10, w: 7.90, h: 0.66, fontFace: BFONT, fontSize: 10.5, color: MUT, margin: 0, valign: "middle" });
    y += 0.94;
  });
  srcNote(s, "Quelle: EU-Bekanntmachung TED 246925-2024 (Eignungskriterien); Konzept „Klassenassistenz“ Kreis Pinneberg, Kap. 5a.", 7.02);
  s.addNotes("Anhangfolie. Die ersten beiden Punkte sind die, die man nicht mehr heilen kann, sobald ein Verfahren läuft.");
}

/* =======================================================================
   8 — Anhang: Glossar
   ===================================================================== */
{
  const s = pres.addSlide();
  titleBar(s, "Anhang 2", "Glossar");

  const terms = [
    ["Eingliederungshilfe", "Sozialleistung für Menschen mit Behinderung; Schulbegleitung ist eine Form davon."],
    ["Schulbegleitung", "Individuelle Unterstützung eines Kindes im Unterricht. Keine geschützte Berufsbezeichnung, kein Fachkräftegebot."],
    ["Schulische Assistenz", "Systemische Unterstützung aller Kinder, nur an Grundschulen, vom Land finanziert. Nicht dasselbe wie Schulbegleitung."],
    ["Poolmodell", "Mehrere Kinder werden gemeinsam begleitet; die Schule erhält ein Budget statt Einzelbewilligungen."],
    ["Infrastrukturangebot / Bildungsassistenz", "Pool außerhalb des individuellen Sozialrechts — kein Antrag, kein Bescheid. Im Bundesentwurf § 80a SGB VIII-E."],
    ["Leistungs- und Vergütungsvereinbarung", "Vertragsform des Sozialrechts (§§ 123 ff. SGB IX): Leistung, Qualität und Preis werden verhandelt, nicht ausgeschrieben. Bei Nichteinigung entscheidet die Schiedsstelle."],
    ["Teilnahmewettbewerb", "Erste Stufe einer EU-Ausschreibung: Eignungsprüfung. Erst danach werden Angebote eingeholt."],
    ["Los", "Teilpaket eines Auftrags, etwa eine Stadt. Kein Losentscheid — jedes Los wird nach Qualität und Preis vergeben."],
    ["Förder-, Inklusions-, Exklusionsquote", "Anteil aller Schüler:innen mit Förderbedarf / davon an Regelschulen / davon an Förderschulen."],
    ["Mengenrisiko", "Wer zahlt, wenn mehr Bedarf entsteht als kalkuliert. Im Einzelfall der Kreis, im Pool der Träger."],
  ];
  let y = 1.42;
  terms.forEach((t, i) => {
    if (i % 2 === 0) card(s, M, y, W - 2 * M, 0.54, CARD);
    s.addText(t[0], { x: M + 0.26, y: y + 0.05, w: 3.5, h: 0.44, fontFace: BFONT, fontSize: 11, bold: true, color: TXT, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 3.92, y: y + 0.05, w: 7.85, h: 0.44, fontFace: BFONT, fontSize: 10.5, color: MUT, margin: 0, valign: "middle" });
    y += 0.56;
  });
  srcNote(s, "Definitionen nach SGB VIII / SGB IX, Landtag SH Drs. 20/2643, BAGüS-Orientierungshilfe 2019 und § 97 Abs. 4 GWB.", 7.06);
  s.addNotes("Anhangfolie für Begriffsfragen.");
}

pres.writeFile({ fileName: "/home/user/repository/praesentation/SKP-Marktanalyse-Schulbegleitung.pptx" })
  .then((f) => console.log("geschrieben:", f));
