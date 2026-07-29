const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "SKP Marktanalyse";
pres.title = "Schulbegleitung im Umbruch";

// ---- Palette -------------------------------------------------------------
const INK = "0E2E38";      // deep petrol, dark slides
const INK_SOFT = "1B4A57";
const TXT = "182E36";
const MUT = "5F7278";
const CARD = "F1F6F8";
const CARD2 = "E7EFF3";
const BLUE = "2A78D6";     // validated series 1
const ORANGE = "EB6834";   // validated series 2
const WHITE = "FFFFFF";

const HFONT = "Cambria";
const BFONT = "Calibri";

const M = 0.62;            // page margin
const W = 13.33;

// ---- helpers -------------------------------------------------------------
function shadow() {
  return { type: "outer", color: "0E2E38", blur: 10, offset: 2, angle: 90, opacity: 0.10 };
}

function titleBar(slide, kicker, title) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: M, y: 0.36, w: 11, h: 0.26, fontFace: BFONT, fontSize: 11, bold: true,
      color: ORANGE, charSpacing: 1.6, margin: 0,
    });
  }
  slide.addText(title, {
    x: M, y: 0.60, w: W - 2 * M, h: 0.92, fontFace: HFONT, fontSize: 27, bold: true,
    color: TXT, margin: 0, valign: "top",
  });
}

function keyline(slide, text, y) {
  slide.addText(text, {
    x: M, y: y, w: W - 2 * M, h: 0.42, fontFace: BFONT, fontSize: 14, italic: true,
    color: INK_SOFT, margin: 0,
  });
}

function srcNote(slide, text, y) {
  slide.addText(text, {
    x: M, y: y || 6.92, w: W - 2 * M, h: 0.3, fontFace: BFONT, fontSize: 9,
    color: MUT, margin: 0,
  });
}

function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, fill: { color: fill || CARD }, rectRadius: 0.08, line: { color: fill || CARD },
    shadow: shadow(),
  });
}

function numDot(slide, x, y, n, color) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: 0.38, h: 0.38, fill: { color: color || BLUE }, line: { color: color || BLUE },
  });
  slide.addText(String(n), {
    x, y, w: 0.38, h: 0.38, fontFace: BFONT, fontSize: 14, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0,
  });
}

// =========================================================================
// 1 — Titel
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("Schulbegleitung im Umbruch", {
    x: M, y: 2.05, w: 11.6, h: 0.95, fontFace: HFONT, fontSize: 44, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Marktanalyse SKP · Kreis Steinburg · Hamburg · Segeberg · Ostholstein · Neumünster", {
    x: M, y: 3.05, w: 11.6, h: 0.5, fontFace: BFONT, fontSize: 17, color: "BFD6DE", margin: 0,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.05, w: 8.9, h: 1.25, fill: { color: INK_SOFT }, rectRadius: 0.08, line: { color: INK_SOFT },
  });
  s.addText([
    { text: "Die Kernthese:  ", options: { bold: true, color: WHITE } },
    { text: "Der Markt wächst — aber der Zugang zu ihm wird gerade neu vergeben.", options: { color: "DCEAF0" } },
  ], { x: M + 0.3, y: 4.28, w: 8.3, h: 0.8, fontFace: BFONT, fontSize: 16, margin: 0, valign: "middle" });
  s.addText("Stand: Juli 2026", {
    x: M, y: 6.5, w: 6, h: 0.3, fontFace: BFONT, fontSize: 11, color: "7FA3B0", margin: 0,
  });
  s.addNotes("Einstieg in 30 Sekunden: SKP verkauft Schulbegleitung an Kreise. Der Bedarf steigt seit zehn Jahren ununterbrochen. Gleichzeitig stellen die Kostenträger gerade das Vergabesystem um — von Einzelfallbewilligung auf Standortvertrag. Wer diese Umstellung übersteht, hat ein größeres und planbareres Geschäft als heute. Wer sie verpasst, verliert Standorte vollständig.");
}

// =========================================================================
// 2 — Geschäftsmodell
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Das Geschäft in 60 Sekunden", "Wer bezahlt wofür — und wo SKP darin steht");

  const rows = [
    ["Das Kind hat einen Rechtsanspruch", "Kinder mit (drohender) Behinderung haben Anspruch auf Unterstützung im Unterricht — § 35a SGB VIII (seelisch) oder § 112 SGB IX (körperlich/geistig)."],
    ["Der Kreis bewilligt und bezahlt", "Zuständig sind Kreise und kreisfreie Städte. In Schleswig-Holstein erstattet das Land rund 84 % der SGB-IX-Kosten."],
    ["SKP erbringt die Leistung", "Freie Träger und Unternehmen stellen das Personal. Vergütet wird bisher pro Kind und Stunde."],
    ["Die Schule ist der eigentliche Kunde", "Ohne Zustimmung der Schulleitung läuft nichts — sie hat faktisches Vetorecht bei der Trägerauswahl."],
  ];
  let y = 1.62;
  rows.forEach((r, i) => {
    card(s, M, y, W - 2 * M, 1.14);
    numDot(s, M + 0.32, y + 0.38, i + 1, i === 3 ? ORANGE : BLUE);
    s.addText(r[0], {
      x: M + 0.92, y: y + 0.17, w: 10.9, h: 0.34, fontFace: BFONT, fontSize: 15, bold: true, color: TXT, margin: 0,
    });
    s.addText(r[1], {
      x: M + 0.92, y: y + 0.53, w: 10.9, h: 0.5, fontFace: BFONT, fontSize: 12.5, color: MUT, margin: 0,
    });
    y += 1.28;
  });
  srcNote(s, "Quellen: SGB VIII, SGB IX; Bericht der Landesregierung SH „Schulbegleitung 2024“ (Drs. 20/2643).");
  s.addNotes("Für den Käufer wichtig: Das ist kein Endkundengeschäft. Es gibt drei Parteien — Kind/Eltern, Kreis als Zahler, Schule als Ort. Und der vierte Punkt ist der, den man beim Kauf am leichtesten übersieht: Die Beziehung zu den Schulleitungen ist der eigentliche Vermögenswert, nicht der Vertrag mit dem Kreis.");
}

// =========================================================================
// 3 — Nachfragetreiber (Chart A)
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Nachfrage", "Mehr Kinder mit Förderbedarf — und mehr davon an Regelschulen");

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [1.1, 3.5, 3.4] },
    { name: "Exklusionsquote (Förderschule)", labels: ["2008/09", "2021/22", "2022/23"], values: [4.8, 4.3, 4.2] },
  ], {
    x: M, y: 1.62, w: 7.5, h: 4.55,
    barDir: "col", barGrouping: "stacked",
    chartColors: [BLUE, ORANGE],
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
    ["+29 %", "mehr Kinder mit Förderbedarf\nseit 2008/09 (482.155 → 581.265)"],
    ["3,1×", "so hohe Inklusionsquote:\n1,1 % → 3,4 % an Regelschulen"],
    ["−0,6 pp", "Exklusionsquote nur leicht gesunken\n(4,8 % → 4,2 %)"],
  ];
  let y = 1.75;
  facts.forEach((f) => {
    card(s, 8.45, y, 4.26, 1.35);
    s.addText(f[0], {
      x: 8.72, y: y + 0.14, w: 3.8, h: 0.5, fontFace: HFONT, fontSize: 26, bold: true, color: BLUE, margin: 0,
    });
    s.addText(f[1], {
      x: 8.72, y: y + 0.66, w: 3.8, h: 0.58, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
    });
    y += 1.5;
  });
  keyline(s, "Der Zuwachs an Inklusion entsteht fast vollständig aus zusätzlichen Förderbedarfen — nicht aus dem Abbau der Förderschulen.", 6.24);
  srcNote(s, "Quelle: Bertelsmann Stiftung auf Grundlage KMK (2024) / Klemm; Schuljahre 2008/09, 2021/22, 2022/23 (ohne Saarland).");
  s.addNotes("Die eine Botschaft dieser Folie: Der Nachfragesockel wächst strukturell, nicht konjunkturell. Beide Balken zusammen sind die Förderquote — sie steigt von 5,9 auf 7,6 Prozent. Und der hellblaue Teil, die Kinder an Regelschulen, hat sich verdreifacht. Genau dort — und nur dort — entsteht Bedarf an Schulbegleitung. Wichtig für den Käufer: Das ist kein Markt, der wegbricht.");
}

// =========================================================================
// 4 — Standortvorteil (Chart B)
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Standort", "SKPs Regionen sind deutlich inklusiver als der Bundesschnitt");

  s.addChart(pres.ChartType.bar, [
    { name: "Inklusionsquote (Regelschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [3.4, 4.4, 5.2] },
    { name: "Exklusionsquote (Förderschule)", labels: ["Deutschland", "Schleswig-Holstein", "Hamburg"], values: [4.2, 2.3, 2.7] },
  ], {
    x: M, y: 1.62, w: 7.5, h: 4.55,
    barDir: "col", barGrouping: "stacked",
    chartColors: [BLUE, ORANGE],
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

  card(s, 8.45, 1.75, 4.26, 2.35);
  s.addText("Warum das der Kern des Marktwerts ist", {
    x: 8.72, y: 1.92, w: 3.8, h: 0.34, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText([
    { text: "Schulbegleitung entsteht fast ausschließlich an ", options: {} },
    { text: "Regelschulen", options: { bold: true } },
    { text: ". In Hamburg sitzen 5,2 % aller Schüler:innen mit Förderbedarf dort, in Schleswig-Holstein 4,4 % — gegenüber 3,4 % im Bund.", options: {} },
  ], { x: 8.72, y: 2.3, w: 3.8, h: 1.6, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0 });

  card(s, 8.45, 4.28, 4.26, 1.9, CARD2);
  s.addText("Umgekehrt gilt:", {
    x: 8.72, y: 4.45, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0,
  });
  s.addText("Beide Regionen haben die niedrigsten Förderschulanteile. Der adressierbare Markt je Einwohner ist damit strukturell größer als im Bundesdurchschnitt — und weniger anfällig für Rückverlagerung.", {
    x: 8.72, y: 4.75, w: 3.8, h: 1.3, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
  });

  srcNote(s, "Quelle: Bertelsmann Stiftung auf Grundlage KMK (2024); Schuljahr 2022/23. Förderquote = Inklusions- + Exklusionsquote.");
  s.addNotes("Diese Folie beantwortet die Frage: Warum ausgerechnet hier? Antwort: Weil in Schleswig-Holstein und Hamburg überdurchschnittlich viele Kinder mit Förderbedarf an Regelschulen unterrichtet werden — und Schulbegleitung genau dort gebraucht wird. Der Standort ist ein Wettbewerbsvorteil, kein Zufall.");
}

// =========================================================================
// 5 — Marktvolumen (Chart C + Hamburg-Tiles)
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Marktvolumen", "Die Ausgaben wachsen fast doppelt so schnell wie die Fallzahlen");

  s.addChart(pres.ChartType.line, [
    { name: "Ausgaben", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 122, 148, 170, 199, 234, 286, 384] },
    { name: "Leistungsempfänger", labels: ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"], values: [100, 111, 131, 144, 162, 189, 204, 245] },
  ], {
    x: M, y: 1.62, w: 7.5, h: 4.35,
    chartColors: [ORANGE, BLUE], lineDataSymbol: "circle", lineDataSymbolSize: 7, lineSize: 3,
    showTitle: true, title: "Schulbegleitung Schleswig-Holstein, Index 2014 = 100",
    titleFontSize: 12, titleColor: MUT, titleFontFace: BFONT,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: TXT, legendFontFace: BFONT,
    catAxisLabelColor: TXT, catAxisLabelFontSize: 11, catAxisLabelFontFace: BFONT,
    valAxisLabelColor: MUT, valAxisLabelFontSize: 10, valAxisLabelFontFace: BFONT,
    valGridLine: { color: "E3EAED", size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 80, valAxisMaxVal: 400, valAxisMajorUnit: 80,
  });

  card(s, 8.45, 1.72, 4.26, 2.2, CARD2);
  s.addText("Schleswig-Holstein 2022", {
    x: 8.72, y: 1.88, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0,
  });
  s.addText("119,7 Mio €", {
    x: 8.72, y: 2.2, w: 3.8, h: 0.62, fontFace: HFONT, fontSize: 32, bold: true, color: ORANGE, margin: 0,
  });
  s.addText("Gesamtausgaben für Schulbegleitung (SGB VIII + SGB IX). 2014 waren es 27,1 Mio €.\n6.545 Leistungsempfänger (2021).", {
    x: 8.72, y: 2.85, w: 3.8, h: 0.95, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
  });

  card(s, 8.45, 4.08, 4.26, 2.1);
  s.addText("Hamburg 2014/15 → 2025/26", {
    x: 8.72, y: 4.24, w: 3.8, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: TXT, margin: 0,
  });
  s.addText([
    { text: "Bewilligungen   1.574 → 4.011", options: { breakLine: true, bold: true, color: TXT } },
    { text: "Kosten   6,75 → 42,15 Mio €", options: { breakLine: true, bold: true, color: TXT } },
    { text: "Kosten je Fall   ~4.300 → ~10.500 €", options: { color: ORANGE, bold: true } },
  ], { x: 8.72, y: 4.58, w: 3.8, h: 1.1, fontFace: BFONT, fontSize: 12.5, margin: 0, lineSpacingMultiple: 1.25 });

  keyline(s, "Genau diese Schere — Kosten je Fall × 2,4 — ist der Auslöser der Reform.", 6.2);
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643 (Landkreistag/Städteverband, 2014–2022); Behörde für Schule, Familie und Berufsbildung Hamburg. Kosten je Fall: eigene Berechnung.");
  s.addNotes("Der wichtigste Chart des Decks. Die orange Linie sind die Ausgaben, die blaue die Fallzahlen. Beide steigen — aber die Ausgaben fast doppelt so schnell. Übersetzt: Nicht die Menge treibt die Kosten, sondern der Preis je Fall. Und genau da setzt die Reform an, die auf den nächsten Folien kommt. Wer das verstanden hat, versteht alles Weitere.");
}

// =========================================================================
// 6 — Die Reform
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Die Reform", "Von der 1:1-Begleitung zum Pool — auf drei Ebenen gleichzeitig");

  // Vorher / Nachher
  card(s, M, 1.6, 5.5, 2.0, CARD2);
  s.addText("HEUTE  ·  Einzelfall", {
    x: M + 0.28, y: 1.78, w: 4.9, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: MUT, charSpacing: 1, margin: 0,
  });
  s.addText("Ein Kind — ein Antrag — ein Bescheid — eine Begleitperson. Bezahlt wird pro Kind und Stunde.",
    { x: M + 0.28, y: 2.12, w: 4.9, h: 1.2, fontFace: BFONT, fontSize: 14, color: TXT, margin: 0 });

  card(s, 6.90, 1.6, 5.81, 2.0);
  s.addText("KÜNFTIG  ·  Pool / Infrastruktur", {
    x: 7.18, y: 1.78, w: 5.2, h: 0.3, fontFace: BFONT, fontSize: 12, bold: true, color: ORANGE, charSpacing: 1, margin: 0,
  });
  s.addText("Die Schule erhält ein Budget und ein Team. Kein Antrag, kein Bescheid. Bezahlt wird pro Klasse oder pro Standort.",
    { x: 7.18, y: 2.12, w: 5.2, h: 1.2, fontFace: BFONT, fontSize: 14, color: TXT, margin: 0 });

  s.addShape(pres.ShapeType.rightArrow, {
    x: 6.32, y: 2.32, w: 0.46, h: 0.56, fill: { color: ORANGE }, line: { color: ORANGE },
  });

  const lvl = [
    ["Bund", "Referentenentwurf 1. KJHSRG (März 2026): „infrastrukturelle Bildungsassistenz“ als Regelfall, Individualanspruch nur noch, wenn ausschließlich 1:1 hilft. Geplant ab 01.01.2028."],
    ["Land SH", "Erklärtes Ziel der Landesregierung: Schulassistenz und Schulbegleitung in Poollösungen zusammenführen. Kein Endtermin genannt."],
    ["Kreise", "Handeln bereits eigenständig aus Kostendruck — nahezu alle Kreise in SH erproben oder betreiben Poolmodelle. Hamburg stellt zum Schuljahr 2026/27 um."],
  ];
  let y = 3.82;
  lvl.forEach((l, i) => {
    card(s, M, y, W - 2 * M, 0.86);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.26, y: y + 0.24, w: 1.32, h: 0.38, fill: { color: i === 2 ? ORANGE : BLUE }, rectRadius: 0.06,
      line: { color: i === 2 ? ORANGE : BLUE },
    });
    s.addText(l[0], {
      x: M + 0.26, y: y + 0.24, w: 1.32, h: 0.38, fontFace: BFONT, fontSize: 12, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(l[1], {
      x: M + 1.78, y: y + 0.14, w: 10.0, h: 0.6, fontFace: BFONT, fontSize: 12, color: TXT, margin: 0, valign: "middle",
    });
    y += 0.98;
  });
  srcNote(s, "Quellen: Referentenentwurf 1. KJHSRG (§§ 35d Abs. 4, 80a SGB VIII-E); Landtag SH Drs. 20/2643 und 20/3271; BSFB Hamburg.");
  s.addNotes("Wichtig zu betonen: Das ist keine einzelne Gesetzesänderung, auf die man warten könnte. Drei Ebenen laufen parallel, und die unterste — die Kreise — ist bereits in der Umsetzung. Das Bundesgesetz beschleunigt, aber es verhindert nichts, wenn es scheitert.");
}

// =========================================================================
// 7 — Was sich für den Anbieter ändert
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Konsequenz für den Anbieter", "Aus vielen kleinen Bewilligungen werden wenige große Verträge");

  const cols = [
    ["", "Einzelfall (heute)", "Pool (künftig)"],
    ["Vertragspartner", "Kreis, je Kind", "Kreis, je Standort oder Region"],
    ["Auftragserteilung", "Bewilligungsbescheid", "Ausschreibung oder Interessenbekundung"],
    ["Laufzeit", "12 Monate, rollierend", "bis zu 5 Jahre, plus Verlängerungsoption"],
    ["Vergütung", "pro Kind und Stunde", "Festbudget pro Klasse / Standort"],
    ["Mengenrisiko", "beim Kreis", "beim Träger"],
    ["Verlust eines Standorts", "einzelne Fälle fallen weg", "gesamtes Volumen fällt weg"],
  ];
  const colX = [M, M + 3.1, M + 7.6];
  const colW = [3.0, 4.4, 4.5];
  let y = 1.66;
  cols.forEach((r, i) => {
    if (i > 0) {
      card(s, M, y, W - 2 * M, 0.60, i % 2 === 0 ? CARD : "FFFFFF");
    }
    r.forEach((cell, j) => {
      s.addText(cell, {
        x: colX[j] + (i > 0 ? 0.24 : 0.24), y: y + (i === 0 ? 0 : 0.12), w: colW[j], h: i === 0 ? 0.36 : 0.4,
        fontFace: BFONT, fontSize: i === 0 ? 12 : 12.5,
        bold: i === 0 || j === 0,
        color: i === 0 ? (j === 2 ? ORANGE : MUT) : (j === 2 ? INK_SOFT : TXT),
        charSpacing: i === 0 ? 1 : 0, margin: 0, valign: "middle",
      });
    });
    y += i === 0 ? 0.42 : 0.68;
  });

  card(s, M, 6.24, W - 2 * M, 0.66, CARD2);
  s.addText([
    { text: "Der eine Satz, der zählt:  ", options: { bold: true, color: TXT } },
    { text: "An jedem Standort entscheidet sich künftig binär — SKP wird koordinierender Träger oder ist dort raus.", options: { color: TXT } },
  ], { x: M + 0.3, y: 6.24, w: 11.5, h: 0.66, fontFace: BFONT, fontSize: 13, margin: 0, valign: "middle" });

  srcNote(s, "Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg; EU-Bekanntmachung TED 246925-2024; DISW-Evaluation Ostholstein.", 7.02);
  s.addNotes("Hier liegt das eigentliche unternehmerische Risiko und zugleich die Chance. Vorher war der Umsatz granular — verliert man ein Kind, verliert man ein Kind. Künftig ist es binär: gewinnt man den Standort, hat man ihn für Jahre; verliert man ihn, ist das gesamte Volumen dort weg. Und in Ostholstein ist dokumentiert, dass auch die Mitarbeitenden zum neuen Träger wechseln.");
}

// =========================================================================
// 8 — Regionale Betroffenheit
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Regionale Betroffenheit", "Fünf Regionen, drei Geschwindigkeiten");

  const regs = [
    ["Hamburg", "AKUT", ORANGE, "Umstellung auf Kombinationsmaßnahmen läuft seit Schuljahr 2026/27. Höher qualifizierte Kräfte nur noch im Ausnahmefall."],
    ["Kreis Ostholstein", "HOCH", ORANGE, "Poolmodell seit 2020/21, seit 2024 phasenweise Ausweitung auf alle Schulstandorte. „Konzentration der Leistungserbringer“ ist erklärtes Projektziel."],
    ["Kreis Segeberg", "MITTEL", "C98500", "Pool an den Förderzentren seit 2021/22; Ausweitung auf rund fünf Grundschulen ab 2025/26 über Interessenbekundungsverfahren."],
    ["Neumünster", "NIEDRIG", BLUE, "Kein Poolmodell dokumentiert. Als kreisfreie Stadt aber schnell umsteuerbar — Flensburg brauchte dafür zwei Jahre."],
    ["Kreis Steinburg", "NIEDRIG", BLUE, "Steuerungs-AG ruht. Seit Februar 2025 sind Jugend- und Sozialamtsaufgaben im neuen Amt für Teilhabe gebündelt — die strukturelle Vorstufe."],
  ];
  let y = 1.62;
  regs.forEach((r) => {
    card(s, M, y, W - 2 * M, 0.92);
    s.addShape(pres.ShapeType.ellipse, { x: M + 0.3, y: y + 0.32, w: 0.28, h: 0.28, fill: { color: r[2] }, line: { color: r[2] } });
    s.addText(r[0], {
      x: M + 0.72, y: y + 0.15, w: 2.7, h: 0.34, fontFace: BFONT, fontSize: 14.5, bold: true, color: TXT, margin: 0,
    });
    s.addText(r[1], {
      x: M + 0.72, y: y + 0.5, w: 2.7, h: 0.3, fontFace: BFONT, fontSize: 10.5, bold: true, color: r[2], charSpacing: 1.2, margin: 0,
    });
    s.addText(r[3], {
      x: M + 3.55, y: y + 0.14, w: 8.2, h: 0.7, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0, valign: "middle",
    });
    y += 1.02;
  });
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643 (Stand 11/2024); DISW-Evaluation Ostholstein; BSFB Hamburg; Kreis Steinburg. Stand kann sich kurzfristig ändern.");
  s.addNotes("Die Botschaft: Es gibt keine einheitliche Frist. Zwei der fünf Regionen sind bereits in der Umstellung, zwei noch gar nicht. Aber „niedrig“ heißt nicht „kein Handlungsbedarf“ — Flensburg ging vom Modellversuch an einer Schule bis zur flächendeckenden Einführung in zwei Jahren.");
}

// =========================================================================
// 9 — Chancen
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Chancen", "Warum der Markt für den richtigen Anbieter größer wird");

  const items = [
    ["Der Gesamtmarkt wächst weiter", "Fallzahlen und Ausgaben steigen in allen belastbaren Datenreihen. Kein Kreis erwartet einen Rückgang."],
    ["Größere, planbarere Aufträge", "Statt hunderter Einzelbewilligungen wenige Standortverträge über Jahre — mit deutlich geringerem Verwaltungsaufwand je Umsatzeuro."],
    ["Regionale Verankerung zählt formal", "Kreis Pinneberg macht Kenntnis der regionalen Strukturen und bestehende Kooperationen zu nachzuweisenden Zuschlagskriterien."],
    ["Qualität schlägt Preis", "Kreis Düren gewichtet Qualität mit 60 von 100 Punkten. Ein dokumentiertes Qualifizierungskonzept ist der billigste Differenzierer."],
    ["Zeitfenster in zwei Regionen", "In Steinburg und Neumünster existiert noch kein Modell — dort lässt sich das Verfahren mitgestalten, statt sich später darauf zu bewerben."],
    ["Der Markt ist offen für Unternehmen", "Zuschläge gingen zuletzt auch an eine GmbH & Co. KG und eine gGmbH — nicht nur an Wohlfahrtsverbände."],
  ];
  let y = 1.62;
  items.forEach((it, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = M + col * 6.15;
    const yy = 1.62 + row * 1.72;
    card(s, x, yy, 5.94, 1.52);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.26, y: yy + 0.24, w: 0.3, h: 0.3, fill: { color: BLUE }, rectRadius: 0.05, line: { color: BLUE },
    });
    s.addText(it[0], {
      x: x + 0.68, y: yy + 0.2, w: 5.0, h: 0.38, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
    });
    s.addText(it[1], {
      x: x + 0.68, y: yy + 0.6, w: 5.0, h: 0.8, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0,
    });
  });
  srcNote(s, "Quellen: Landtag SH Drs. 20/2643; TED-Bekanntmachungen Kreis Pinneberg, Düren, Euskirchen, Rhein-Erft, Essen.");
  s.addNotes("Hier nicht euphorisch werden. Die Chancen sind real, hängen aber alle an einer Bedingung: Das Unternehmen muss vergabefähig werden. Punkt 3 und 4 sind die entscheidenden — regionale Verankerung und Qualifizierungskonzept sind genau die Dinge, die ein etablierter lokaler Anbieter günstiger liefern kann als ein Konzern.");
}

// =========================================================================
// 10 — Risiken
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Risiken", "Was den Wert des Unternehmens gefährden kann");

  const items = [
    ["Standortverlust ist total, nicht anteilig", "Wer den Zuschlag nicht bekommt, verliert das gesamte Volumen an diesem Standort — zum Schuljahresbeginn, ohne Übergang."],
    ["Personal wandert mit", "In Ostholstein dokumentiert: Mitarbeitende werden an den koordinierenden Träger „übergeben“. Rechtlich freiwillig, faktisch die Regel."],
    ["Kein Rechtsschutz", "Das OVG Schleswig hat die Klage von vier Trägern gegen das Pinneberger Vergabeverfahren zurückgewiesen. Unanfechtbar. Bestandsschutz gibt es nicht."],
    ["Vergabefähigkeit fehlt oft", "10 Mio € Betriebshaftpflicht, Referenzliste über 4 Jahre, getrennte Umsatzausweisung, Präqualifizierung — harte Ausschlusskriterien."],
  ];
  let y = 1.62;
  items.forEach((it, i) => {
    card(s, M, y, W - 2 * M, 0.96);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.28, y: y + 0.33, w: 0.3, h: 0.3, fill: { color: ORANGE }, rectRadius: 0.05, line: { color: ORANGE },
    });
    s.addText(it[0], {
      x: M + 0.72, y: y + 0.15, w: 4.1, h: 0.68, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: M + 5.0, y: y + 0.15, w: 6.75, h: 0.68, fontFace: BFONT, fontSize: 12, color: MUT, margin: 0, valign: "middle",
    });
    y += 1.06;
  });

  card(s, M, 5.95, W - 2 * M, 0.72, CARD2);
  s.addText([
    { text: "Das größte Einzelrisiko:  ", options: { bold: true, color: ORANGE } },
    { text: "Preisbindung ohne Tarifausgleich — siehe nächste Folie.", options: { color: TXT } },
  ], { x: M + 0.3, y: 5.95, w: 11.5, h: 0.72, fontFace: BFONT, fontSize: 13, margin: 0, valign: "middle" });

  srcNote(s, "Quellen: OVG Schleswig, Beschluss vom 03.09.2024 – 5 MB 7/24; TED 246925-2024; DISW-Evaluation Ostholstein.");
  s.addNotes("Diese Folie ehrlich vortragen — sie ist der Grund, warum das Unternehmen jetzt Investition braucht. Besonders Punkt 3: Es gibt keinen Rechtsweg gegen die Umstellung. Vier Träger haben geklagt und verloren, unanfechtbar. Die einzige wirksame Verteidigung ist, das Verfahren zu gewinnen.");
}

// =========================================================================
// 11 — Preisbindung
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("DAS EINE, DAS MAN VERSTEHEN MUSS", {
    x: M, y: 0.66, w: 11, h: 0.3, fontFace: BFONT, fontSize: 11, bold: true, color: ORANGE, charSpacing: 1.6, margin: 0,
  });
  s.addText("Ein Festpreis über bis zu neun Jahre — ohne Tarifausgleich", {
    x: M, y: 1.0, w: W - 2 * M, h: 0.7, fontFace: HFONT, fontSize: 30, bold: true, color: WHITE, margin: 0,
  });

  s.addText("„Anpassungen durch Tarifsteigerungen sind durch die Mischkalkulation bereits abgegolten und erfolgen im Rahmen der Vertragslaufzeit nicht.“", {
    x: M, y: 1.92, w: 7.4, h: 1.0, fontFace: HFONT, fontSize: 15, italic: true, color: "CFE3EA", margin: 0,
  });
  s.addText("Konzept „Klassenassistenz“, Kreis Pinneberg", {
    x: M, y: 2.9, w: 7.4, h: 0.3, fontFace: BFONT, fontSize: 11, color: "7FA3B0", margin: 0,
  });

  const tiles = [
    ["01.11.2024 – 31.07.2029", "Grundlaufzeit des ausgeschriebenen Vertrags"],
    ["+ 4 Jahre", "einseitige Verlängerungsoption des Kreises — der Träger kann sie nicht ablehnen"],
    ["≈ 36 %", "höhere Personalkosten nach neun Jahren bei 3,5 % Tarifsteigerung p. a., bei unverändertem Erlös"],
  ];
  let y = 3.42;
  tiles.forEach((t) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: y, w: 7.4, h: 0.92, fill: { color: INK_SOFT }, rectRadius: 0.07, line: { color: INK_SOFT },
    });
    s.addText(t[0], {
      x: M + 0.3, y: y + 0.12, w: 3.0, h: 0.34, fontFace: HFONT, fontSize: 16, bold: true, color: ORANGE, margin: 0,
    });
    s.addText(t[1], {
      x: M + 0.3, y: y + 0.46, w: 6.8, h: 0.4, fontFace: BFONT, fontSize: 11.5, color: "CFE3EA", margin: 0,
    });
    y += 1.04;
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 8.35, y: 1.92, w: 4.36, h: 4.54, fill: { color: "13424F" }, rectRadius: 0.08, line: { color: "13424F" },
  });
  s.addText("Warum das im Kaufpreis stehen muss", {
    x: 8.65, y: 2.15, w: 3.8, h: 0.34, fontFace: BFONT, fontSize: 13.5, bold: true, color: WHITE, margin: 0,
  });
  s.addText([
    { text: "In der Ausschreibung gewinnt der günstigste tragfähige Preis. Wer den Tarifanstieg einpreist, ist teurer und verliert. Wer ihn nicht einpreist, gewinnt — und zahlt später drauf.", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "Im heutigen Vereinbarungssystem schützt § 124 SGB IX die Tarifbindung und bei Streit entscheidet die Schiedsstelle. Im Vergabesystem gilt beides nicht.", options: {} },
  ], { x: 8.65, y: 2.55, w: 3.8, h: 3.3, fontFace: BFONT, fontSize: 12, color: "CFE3EA", margin: 0 });

  s.addText("Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg, Kap. 5b; TED 246925-2024. Der Prozentwert ist eine eigene Modellrechnung.", {
    x: M, y: 6.92, w: 12.1, h: 0.3, fontFace: BFONT, fontSize: 9, color: "7FA3B0", margin: 0,
  });
  s.addNotes("Diesen Punkt langsam vortragen. Er ist der Grund, warum ein unvorbereiteter Käufer hier Geld verlieren kann. Die Ausschreibungslogik belohnt kurzfristig den, der zu knapp kalkuliert. Konsequenz für die Kaufentscheidung: Vor jedem Angebot müssen Laufzeit, Verlängerungsoption und eine mögliche Indexierungsklausel geklärt sein — und es braucht eine definierte Preisuntergrenze.");
}

// =========================================================================
// 12 — Investitionsbedarf & Fahrplan
// =========================================================================
{
  const s = pres.addSlide();
  titleBar(s, "Was zu tun ist", "Investitionsbedarf und Zeitachse");

  const inv = [
    "Teamleitung / Koordination mit pädagogischer Fachqualifikation aufbauen — ohne sie ist SKP nicht zuschlagsfähig",
    "Schriftliches Fortbildungs- und Qualifizierungskonzept — hartes Zuschlagskriterium",
    "Betriebshaftpflicht auf 10 Mio € prüfen und Referenzliste über 4 Jahre anlegen",
    "Registrierung auf den Vergabeplattformen mit Daueralarm auf CPV 85310000",
    "Kalkulationsmodell für Festbudgets ohne Tarifgleitklausel",
  ];
  card(s, M, 1.62, 6.35, 3.55);
  s.addText("Investitionsbedarf", {
    x: M + 0.3, y: 1.82, w: 5.7, h: 0.34, fontFace: BFONT, fontSize: 15, bold: true, color: TXT, margin: 0,
  });
  s.addText(inv.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i !== inv.length - 1 } })), {
    x: M + 0.3, y: 2.24, w: 5.75, h: 2.75, fontFace: BFONT, fontSize: 11.5, color: TXT, margin: 0, paraSpaceAfter: 8,
  });

  const mile = [
    ["läuft", "Hamburg: Kombinationsmaßnahmen ab Schuljahr 2026/27"],
    ["Sommer 2026", "1. KJHSRG: Kabinett, danach Bundestag und Bundesrat"],
    ["01.01.2027", "Frist: bis dahin muss das Bundesgesetz verkündet sein"],
    ["ab 2027/28", "Kreis Pinneberg: Klassenassistenz Stufe 1 (abgeleitet)"],
    ["01.01.2028", "Gesamtzuständigkeit der Jugendämter, Bildungsassistenz"],
  ];
  card(s, 7.28, 1.62, 5.43, 3.55, CARD2);
  s.addText("Meilensteine", {
    x: 7.58, y: 1.82, w: 4.8, h: 0.34, fontFace: BFONT, fontSize: 15, bold: true, color: TXT, margin: 0,
  });
  let my = 2.28;
  mile.forEach((m, i) => {
    s.addText(m[0], {
      x: 7.58, y: my, w: 1.55, h: 0.5, fontFace: BFONT, fontSize: 11, bold: true,
      color: i < 2 ? ORANGE : INK_SOFT, margin: 0, valign: "top",
    });
    s.addText(m[1], {
      x: 9.18, y: my, w: 3.3, h: 0.5, fontFace: BFONT, fontSize: 11, color: MUT, margin: 0, valign: "top",
    });
    my += 0.56;
  });

  card(s, M, 5.30, W - 2 * M, 1.32, CARD);
  s.addText("Der günstigste Frühwarnindikator liegt bereits im Haus", {
    x: M + 0.3, y: 5.44, w: 11.5, h: 0.32, fontFace: BFONT, fontSize: 13.5, bold: true, color: TXT, margin: 0,
  });
  s.addText("Wenn ein Kreis beginnt, alle Bewilligungsbescheide einheitlich auf einen Schuljahresbeginn zu befristen, ist die Systemumstellung beschlossen — Monate bevor sie öffentlich wird. In Pinneberg war genau das das erste erkennbare Signal. Die laufenden Bescheide von SKP darauf zu prüfen, kostet nichts.",
    { x: M + 0.3, y: 5.78, w: 11.5, h: 0.78, fontFace: BFONT, fontSize: 11.5, color: MUT, margin: 0 });

  srcNote(s, "Quellen: TED 246925-2024 (Eignungskriterien); Kabinettzeitplanung Bundesregierung 25.06.2026; Art. 10 Abs. 3 KJSG; Trägerschreiben Kreis Pinneberg 12/2023 und 03/2024.");
  s.addNotes("Der Investitionsbedarf ist überschaubar und einmalig — im Wesentlichen eine Fachkraftstelle plus Konzeptarbeit. Das ist die eigentliche Kaufthese: ein etabliertes regionales Unternehmen mit Schulbeziehungen, dem die Vergabefähigkeit fehlt. Die lässt sich in zwölf bis achtzehn Monaten herstellen.");
}

// =========================================================================
// 13 — Fazit
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("FAZIT", {
    x: M, y: 0.72, w: 11, h: 0.3, fontFace: BFONT, fontSize: 11, bold: true, color: ORANGE, charSpacing: 1.6, margin: 0,
  });
  s.addText("Ein wachsender Markt, dessen Zugang gerade neu vergeben wird", {
    x: M, y: 1.06, w: W - 2 * M, h: 0.75, fontFace: HFONT, fontSize: 30, bold: true, color: WHITE, margin: 0,
  });

  const cols = [
    ["Was für den Kauf spricht", ["Struktureller Nachfragezuwachs seit über zehn Jahren", "Überdurchschnittlich inklusive Regionen — größerer adressierbarer Markt", "Künftig größere, mehrjährige und planbarere Verträge", "Regionale Verankerung ist formales Zuschlagskriterium"], BLUE],
    ["Was zu klären ist", ["Vergabefähigkeit: Haftpflicht, Referenzen, Qualifizierungskonzept", "Standortliste mit Umsatz je Schule und Poolstatus des Kreises", "Kalkulation für Festbudgets ohne Tarifausgleich", "Arbeitsrechtliche Folgen eines Standortverlusts (§ 613a BGB)"], ORANGE],
  ];
  cols.forEach((c, i) => {
    const x = M + i * 6.15;
    s.addShape(pres.ShapeType.roundRect, {
      x: x, y: 2.05, w: 5.94, h: 3.5, fill: { color: INK_SOFT }, rectRadius: 0.08, line: { color: INK_SOFT },
    });
    s.addText(c[0], {
      x: x + 0.32, y: 2.28, w: 5.3, h: 0.36, fontFace: BFONT, fontSize: 15, bold: true, color: c[2], margin: 0,
    });
    s.addText(c[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== c[1].length - 1 } })), {
      x: x + 0.32, y: 2.75, w: 5.3, h: 2.6, fontFace: BFONT, fontSize: 12.5, color: "DCEAF0", margin: 0, paraSpaceAfter: 10,
    });
  });

  s.addText([
    { text: "Datenbasis:  ", options: { bold: true, color: WHITE } },
    { text: "Landtagsdrucksachen SH, EU-Vergabebekanntmachungen, Konzeptunterlagen Kreis Pinneberg, wissenschaftliche Evaluationen (DISW, Universität Regensburg), Bertelsmann Stiftung. Fallzahlen je Kreis werden in Schleswig-Holstein nicht veröffentlicht — regionale Volumina sind Schätzungen und vor einer Kaufentscheidung bei den Kreisen zu verifizieren.", options: { color: "9FBECA" } },
  ], { x: M, y: 5.8, w: 12.1, h: 1.0, fontFace: BFONT, fontSize: 10.5, margin: 0 });

  s.addNotes("Abschluss: Die Kaufthese in einem Satz — ein etabliertes regionales Unternehmen in einem strukturell wachsenden Markt, dem eine überschaubare, aber zeitkritische Investition in Vergabefähigkeit fehlt. Und ganz klar sagen, was noch offen ist: die kreisscharfen Fallzahlen und die arbeitsrechtliche Frage beim Standortverlust.");
}

pres.writeFile({ fileName: "/home/user/repository/praesentation/SKP-Marktanalyse-Schulbegleitung.pptx" })
  .then((f) => console.log("geschrieben:", f));
