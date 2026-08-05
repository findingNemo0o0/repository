const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
const W = 13.33, H = 7.5;

// ---- Palette ----
const INK   = "082A31"; // darkest (dark slide bg)
const PRIM  = "0E5A66"; // primary petrol/teal
const TEAL2 = "2A9D8F"; // seafoam / positive
const AMBER = "E8A13A"; // accent (gates, highlights)
const RED   = "C0442E"; // risk / scam
const LIGHT = "F5F8F8"; // light content bg
const CARD  = "FFFFFF";
const MUTE  = "5C7075"; // muted text
const LINE  = "D5E1E1"; // hairline
const WHITE = "FFFFFF";
const INKSOFT = "0C3A43";

const HFONT = "Cambria";   // serif headers
const BFONT = "Calibri";   // sans body

const shadow = () => ({ type: "outer", color: "0A2226", blur: 8, offset: 3, angle: 90, opacity: 0.22 });

function bg(slide, color){ slide.background = { color }; }

function footer(slide, n){
  slide.addText("Interner Prozess-Leitfaden  ·  Kauf insolventer (Intensiv-)Pflegedienste", {
    x:0.5, y:7.08, w:9.5, h:0.3, fontFace:BFONT, fontSize:9, color:MUTE, align:"left", margin:0 });
  slide.addText(String(n), { x:12.4, y:7.08, w:0.45, h:0.3, fontFace:BFONT, fontSize:9, color:MUTE, align:"right", margin:0 });
}

// section title (light slides)
function titleBlock(slide, kicker, title){
  slide.addText(kicker.toUpperCase(), { x:0.5, y:0.42, w:12.3, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:TEAL2, charSpacing:2, align:"left", margin:0 });
  slide.addText(title, { x:0.5, y:0.72, w:12.3, h:0.72, fontFace:HFONT, fontSize:30, bold:true, color:PRIM, align:"left", margin:0 });
}

function iconCircle(slide, x, y, d, fill, label, lsize){
  slide.addShape(p.ShapeType.ellipse, { x, y, w:d, h:d, fill:{color:fill}, line:{type:"none"}, shadow:shadow() });
  slide.addText(label, { x, y, w:d, h:d, fontFace:BFONT, fontSize:lsize||14, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
}

// =====================================================================
// SLIDE 1 — TITLE
// =====================================================================
let s = p.addSlide(); bg(s, INK);
// subtle motif: large translucent rings
s.addShape(p.ShapeType.ellipse, { x:9.6, y:-1.7, w:5.2, h:5.2, fill:{type:"none"}, line:{color:PRIM, width:1.5, transparency:35} });
s.addShape(p.ShapeType.ellipse, { x:10.7, y:-0.6, w:3.4, h:3.4, fill:{type:"none"}, line:{color:TEAL2, width:1.5, transparency:45} });
s.addShape(p.ShapeType.ellipse, { x:-1.4, y:5.2, w:4.2, h:4.2, fill:{type:"none"}, line:{color:AMBER, width:1.5, transparency:55} });

s.addText("M&A-LEITFADEN  ·  DISTRESSED HEALTHCARE", { x:0.7, y:1.75, w:10, h:0.35, fontFace:BFONT, fontSize:13, bold:true, color:AMBER, charSpacing:3, margin:0 });
s.addText("Kauf eines insolventen\nPflegedienstes", { x:0.68, y:2.15, w:11.5, h:1.85, fontFace:HFONT, fontSize:52, bold:true, color:WHITE, lineSpacing:52, margin:0 });
s.addText("Prozess, Zeitplan & Achievement-Gates — mit Fokus auf ambulante\nund betreute außerklinische Intensivpflege (AKI, § 132l SGB V)", {
  x:0.7, y:4.35, w:11, h:0.8, fontFace:BFONT, fontSize:17, color:"C7D8DA", lineSpacing:24, margin:0 });

// bottom chips
const chips = ["Verfahrenskunde", "Deal-Strukturen", "Due Diligence", "Scam-Radar", "Integration"];
let cx = 0.7;
chips.forEach(c => {
  const wch = 0.35 + c.length*0.098;
  s.addShape(p.ShapeType.roundRect, { x:cx, y:5.75, w:wch, h:0.45, rectRadius:0.22, fill:{color:INKSOFT}, line:{color:PRIM, width:1} });
  s.addText(c, { x:cx, y:5.75, w:wch, h:0.45, fontFace:BFONT, fontSize:11.5, color:"CFE1E2", align:"center", valign:"middle", margin:0 });
  cx += wch + 0.2;
});
s.addNotes("Titelfolie. Positionierung: interner Leitfaden für den Erwerb insolventer Intensivpflegedienste. Zwei Rechtsgebiete kollidieren: Insolvenzrecht (InsO) und Sozial-/Pflegerecht (SGB V/XI).");

// =====================================================================
// SLIDE 2 — MANAGEMENT SUMMARY / WARUM
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Management Summary", "Warum insolvente AKI-Targets attraktiv sind");

// left: narrative
const bullets2 = [
  { h:"Konsolidierungswelle läuft", t:"Demografie, Personalmangel und IPReG-Verschärfungen treffen kleine Träger überproportional — 2024–2027 gilt als aktivstes Fenster." },
  { h:"Massive Bewertungsrabatte", t:"Distressed-Preise liegen weit unter regulären EBITDA-Multiples; dokumentiert sind Wertverluste bis ~90 %." },
  { h:"Zwei Hebel, klar definiert", t:"In der Pflege sind Mitarbeiter und Patienten die entscheidenden — oft die einzigen — Werttreiber. Das macht Integration planbar." },
  { h:"Buy-and-Build-Fit", t:"Zukauf statt Neugründung: bestehende Patienten, Personal und (im Idealfall) Versorgungsverträge statt Aufbau bei null." },
];
let by = 1.7;
bullets2.forEach((b,i) => {
  s.addShape(p.ShapeType.roundRect, { x:0.5, y:by, w:6.55, h:1.02, rectRadius:0.06, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  iconCircle(s, 0.72, by+0.24, 0.54, PRIM, String(i+1), 16);
  s.addText(b.h, { x:1.42, y:by+0.14, w:5.5, h:0.34, fontFace:BFONT, fontSize:14.5, bold:true, color:PRIM, margin:0 });
  s.addText(b.t, { x:1.42, y:by+0.46, w:5.5, h:0.5, fontFace:BFONT, fontSize:11.5, color:"33474C", lineSpacing:14, margin:0 });
  by += 1.14;
});

// right: stat panel
s.addShape(p.ShapeType.roundRect, { x:7.35, y:1.7, w:5.45, h:4.72, rectRadius:0.08, fill:{color:INK}, line:{type:"none"}, shadow:shadow() });
s.addText("Kernaussage", { x:7.7, y:1.95, w:4.8, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:AMBER, charSpacing:2, margin:0 });
s.addText("Der Versorgungsvertrag ist das eigentliche Asset — nicht das Inventar.", {
  x:7.7, y:2.28, w:4.75, h:1.0, fontFace:HFONT, fontSize:20, bold:true, color:WHITE, lineSpacing:24, margin:0 });

const stats = [
  { n:"2–3 Mon.", l:"Insolvenzgeld-Fenster: Löhne getragen, Betrieb läuft weiter" },
  { n:"§ 132l", l:"Versorgungsvertrag — Voraussetzung, um überhaupt abrechnen zu dürfen" },
  { n:"< 4 Wo.", l:"Typischer Zeitdruck der Distressed-DD gegenüber Monaten im Normal-M&A" },
];
let sy = 3.5;
stats.forEach(st => {
  s.addText(st.n, { x:7.7, y:sy, w:2.15, h:0.7, fontFace:HFONT, fontSize:30, bold:true, color:TEAL2, align:"left", valign:"middle", margin:0 });
  s.addText(st.l, { x:9.9, y:sy, w:2.6, h:0.72, fontFace:BFONT, fontSize:11, color:"C7D8DA", align:"left", valign:"middle", lineSpacing:13, margin:0 });
  sy += 0.92;
});
footer(s,2);
s.addNotes("Attraktivität begründen, aber sofort die Kernthese setzen: In der (Intensiv-)Pflege ist die Kassenzulassung/der Versorgungsvertrag das wertbildende Asset. Ohne ihn kein Umsatz.");

// =====================================================================
// SLIDE 3 — ZWEI WELTEN: NORMAL vs DISTRESSED
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Grundverständnis", "Zwei Welten: normaler M&A vs. Distressed");

const rows3 = [
  ["Zeitrahmen", "Monate, geordneter Prozess", "Wochen — Masse verbrennt, Verwalter drängt"],
  ["Due Diligence", "Vollständiger Datenraum", "Limitierte Infos, Käufer trägt das Risiko"],
  ["Verkäufer", "Eigentümer / Gesellschafter", "Insolvenzverwalter (+ Gläubigerausschuss)"],
  ["Gewährleistung", "Umfangreiche Garantien / W&I", "Praktisch keine — Kauf „wie besichtigt“"],
  ["Preislogik", "EBITDA-Multiple, Zukunftswert", "Asset-/going-concern-nah, Sanierungskosten"],
  ["Personal", "§ 613a voll, keine Erleichterung", "§ 613a gilt, aber: Kündigungsfrist max. 3 Mon."],
  ["Freigabe", "Gesellschafterbeschluss", "Insolvenzgericht + Gläubiger müssen zustimmen"],
];
const colX = [0.5, 3.15, 7.9], colW = [2.55, 4.65, 4.9];
const ry0 = 1.72, rh = 0.66;
// header row
s.addShape(p.ShapeType.roundRect, { x:colX[1], y:ry0-0.02, w:colW[1], h:0.5, rectRadius:0.05, fill:{color:PRIM}, line:{type:"none"} });
s.addShape(p.ShapeType.roundRect, { x:colX[2], y:ry0-0.02, w:colW[2], h:0.5, rectRadius:0.05, fill:{color:INK}, line:{type:"none"} });
s.addText("Normaler M&A", { x:colX[1], y:ry0-0.02, w:colW[1], h:0.5, fontFace:BFONT, fontSize:14, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
s.addText("Distressed / Insolvenz", { x:colX[2], y:ry0-0.02, w:colW[2], h:0.5, fontFace:BFONT, fontSize:14, bold:true, color:AMBER, align:"center", valign:"middle", margin:0 });
let yy = ry0 + 0.58;
rows3.forEach((r,i) => {
  if(i%2===0){ s.addShape(p.ShapeType.rect, { x:colX[0], y:yy, w:12.3, h:rh, fill:{color:"ECF3F3"}, line:{type:"none"} }); }
  s.addText(r[0], { x:colX[0]+0.1, y:yy, w:colW[0], h:rh, fontFace:BFONT, fontSize:12.5, bold:true, color:PRIM, valign:"middle", margin:0 });
  s.addText(r[1], { x:colX[1]+0.15, y:yy, w:colW[1]-0.2, h:rh, fontFace:BFONT, fontSize:12, color:"33474C", valign:"middle", margin:0 });
  s.addText(r[2], { x:colX[2]+0.15, y:yy, w:colW[2]-0.2, h:rh, fontFace:BFONT, fontSize:12, color:"33474C", valign:"middle", margin:0 });
  yy += rh;
});
s.addText("Kernunterschied: Du kaufst unter Zeitdruck, ohne Garantien, von einem Dritten — und musst die Substanz selbst prüfen.", {
  x:0.5, y:yy+0.12, w:12.3, h:0.4, fontFace:BFONT, fontSize:12.5, italic:true, color:MUTE, margin:0 });
footer(s,3);
s.addNotes("Diese Tabelle rahmt alles Folgende. Wichtig: keine Garantien vom Verwalter — deshalb ist die DD überlebenswichtig und die Preislogik eine andere.");

// =====================================================================
// SLIDE 4 — VERFAHRENSKUNDE: 3 ARTEN
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Verfahrenskunde", "Wer ist mein Gegenüber? Drei Verfahrensarten");

const proc = [
  { t:"Regelinsolvenz", head:PRIM, who:"Insolvenzverwalter",
    pts:["Schuldner verliert die Verfügungs­gewalt vollständig","Verwalter führt & verkauft, maximiert die Masse","Neutral, wenig emotional — aber gibt kaum Gewähr"],
    tag:"Dein Partner: der Verwalter" },
  { t:"Eigenverwaltung", head:INK, who:"Alt-Management + Sachwalter",
    pts:["Geschäftsführung bleibt im Amt (§ 270 InsO)","Sachwalter überwacht nur, ersetzt nicht","Dieselben Leute, die die Krise verursacht haben"],
    tag:"Erhöhte Täuschungsgefahr" },
  { t:"Schutzschirm", head:TEAL2, who:"Management + vorl. Sachwalter",
    pts:["Sonderform der vorläufigen Eigenverwaltung","Nur bei drohender, noch nicht eingetretener Insolvenz","Voraussetzung: nachgewiesene Sanierungsfähigkeit"],
    tag:"Frühe Krise, § 270b InsO" },
];
let px = 0.5;
proc.forEach(c => {
  const cw = 3.97;
  s.addShape(p.ShapeType.roundRect, { x:px, y:1.7, w:cw, h:4.75, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  s.addShape(p.ShapeType.roundRect, { x:px, y:1.7, w:cw, h:0.9, rectRadius:0.08, fill:{color:c.head}, line:{type:"none"} });
  s.addShape(p.ShapeType.rect, { x:px, y:2.2, w:cw, h:0.4, fill:{color:c.head}, line:{type:"none"} });
  s.addText(c.t, { x:px, y:1.78, w:cw, h:0.74, fontFace:HFONT, fontSize:21, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText("VERHANDLUNGSPARTNER", { x:px+0.25, y:2.78, w:cw-0.5, h:0.25, fontFace:BFONT, fontSize:9.5, bold:true, color:MUTE, charSpacing:1.5, margin:0 });
  s.addText(c.who, { x:px+0.25, y:3.02, w:cw-0.5, h:0.55, fontFace:BFONT, fontSize:15, bold:true, color:c.head==INK?PRIM:c.head, lineSpacing:17, margin:0 });
  let ly = 3.72;
  c.pts.forEach(pt => {
    s.addShape(p.ShapeType.ellipse, { x:px+0.28, y:ly+0.06, w:0.1, h:0.1, fill:{color:AMBER}, line:{type:"none"} });
    s.addText(pt, { x:px+0.5, y:ly-0.04, w:cw-0.72, h:0.55, fontFace:BFONT, fontSize:11.5, color:"33474C", lineSpacing:13.5, margin:0 });
    ly += 0.62;
  });
  s.addShape(p.ShapeType.roundRect, { x:px+0.25, y:5.85, w:cw-0.5, h:0.42, rectRadius:0.06, fill:{color:"EEF4F4"}, line:{type:"none"} });
  s.addText(c.tag, { x:px+0.25, y:5.85, w:cw-0.5, h:0.42, fontFace:BFONT, fontSize:11, bold:true, italic:true, color:c.head==INK?PRIM:c.head, align:"center", valign:"middle", margin:0 });
  px += cw + 0.22;
});
footer(s,4);
s.addNotes("Erste Weichenstellung: Wer verkauft mir eigentlich? Bei Eigenverwaltung/Schutzschirm sitzt das alte Management gegenüber — strukturell höhere Täuschungsgefahr.");

// =====================================================================
// SLIDE 5 — ZWEI KAUFSTRUKTUREN + PFLEGE-TWIST
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Deal-Architektur", "Zwei Kaufstrukturen — der Pflege-Twist entscheidet");

// Card A Asset Deal
s.addShape(p.ShapeType.roundRect, { x:0.5, y:1.7, w:6.0, h:3.35, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
iconCircle(s, 0.75, 1.95, 0.6, PRIM, "A", 20);
s.addText("Asset Deal", { x:1.5, y:1.94, w:4.8, h:0.4, fontFace:HFONT, fontSize:19, bold:true, color:PRIM, margin:0 });
s.addText("„Übertragende Sanierung“ — der Normalfall", { x:1.5, y:2.34, w:4.8, h:0.3, fontFace:BFONT, fontSize:11.5, italic:true, color:MUTE, margin:0 });
const aPts = [
  ["+","Keine Haftung für Altverbindlichkeiten"],
  ["+","Keine Alt-Personalkosten vor Eröffnung"],
  ["−","Verträge gehen NICHT über (außer § 613a)"],
  ["−","Jeder Kassen-/Lieferantenvertrag neu"],
];
let ay = 2.82;
aPts.forEach(pt => {
  const col = pt[0]==="+"?TEAL2:RED;
  s.addShape(p.ShapeType.ellipse, { x:0.78, y:ay+0.02, w:0.26, h:0.26, fill:{color:col}, line:{type:"none"} });
  s.addText(pt[0], { x:0.78, y:ay+0.02, w:0.26, h:0.26, fontFace:BFONT, fontSize:13, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(pt[1], { x:1.16, y:ay-0.02, w:5.2, h:0.36, fontFace:BFONT, fontSize:12.5, color:"33474C", valign:"middle", margin:0 });
  ay += 0.5;
});

// Card B Insolvenzplan
s.addShape(p.ShapeType.roundRect, { x:6.83, y:1.7, w:6.0, h:3.35, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
iconCircle(s, 7.08, 1.95, 0.6, INK, "B", 20);
s.addText("Share Deal via Insolvenzplan", { x:7.83, y:1.94, w:4.8, h:0.4, fontFace:HFONT, fontSize:17, bold:true, color:PRIM, margin:0 });
s.addText("Rechtsträger überlebt entschuldet", { x:7.83, y:2.34, w:4.8, h:0.3, fontFace:BFONT, fontSize:11.5, italic:true, color:MUTE, margin:0 });
const bPts = [
  ["+","Versorgungsverträge & Zulassungen bleiben"],
  ["+","Kein Abrechnungsbruch, Verordnungen bestehen"],
  ["−","Mehr latente Haftung wird mitgekauft"],
  ["−","Komplexer: Gläubigermehrheiten nötig"],
];
let byy = 2.82;
bPts.forEach(pt => {
  const col = pt[0]==="+"?TEAL2:RED;
  s.addShape(p.ShapeType.ellipse, { x:7.11, y:byy+0.02, w:0.26, h:0.26, fill:{color:col}, line:{type:"none"} });
  s.addText(pt[0], { x:7.11, y:byy+0.02, w:0.26, h:0.26, fontFace:BFONT, fontSize:13, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(pt[1], { x:7.49, y:byy-0.02, w:5.2, h:0.36, fontFace:BFONT, fontSize:12.5, color:"33474C", valign:"middle", margin:0 });
  byy += 0.5;
});

// Twist banner
s.addShape(p.ShapeType.roundRect, { x:0.5, y:5.28, w:12.33, h:1.5, rectRadius:0.08, fill:{color:INK}, line:{type:"none"}, shadow:shadow() });
s.addShape(p.ShapeType.roundRect, { x:0.5, y:5.28, w:2.1, h:1.5, rectRadius:0.08, fill:{color:AMBER}, line:{type:"none"} });
s.addShape(p.ShapeType.rect, { x:2.0, y:5.28, w:0.6, h:1.5, fill:{color:AMBER}, line:{type:"none"} });
s.addText("PFLEGE-\nTWIST", { x:0.5, y:5.28, w:2.1, h:1.5, fontFace:HFONT, fontSize:19, bold:true, color:INK, align:"center", valign:"middle", lineSpacing:20, margin:0 });
s.addText([
  { text:"Der Versorgungsvertrag (§ 132l / § 132a SGB V, § 72 SGB XI) ist an den Rechtsträger gebunden — nicht an den Betrieb.\n", options:{ fontSize:14, bold:true, color:WHITE } },
  { text:"Im Asset Deal geht er NICHT automatisch über → drohende Abrechnungslücke. Betreiberwechsel/neuen Vertrag vorab mit den Kassen klären. In der Intensivpflege kann der Insolvenzplan deshalb überlegen sein.", options:{ fontSize:12.5, color:"C7D8DA" } },
], { x:2.85, y:5.42, w:9.75, h:1.25, valign:"middle", lineSpacing:15, margin:0, align:"left" });
footer(s,5);
s.addNotes("Wichtigste inhaltliche Folie. Kernbotschaft: In der Pflege ist die Kassenzulassung das Asset. Asset Deal = Zulassung weg; Insolvenzplan = Zulassung bleibt. Früh mit Fachanwalt Insolvenz- + Medizinrecht entscheiden.");

// =====================================================================
// SLIDE 6 — PROZESS-MASTER (ZEITSTRAHL + GATES)  [HERO]
// =====================================================================
s = p.addSlide(); bg(s, INK);
s.addText("DER PROZESS AUF EINEN BLICK", { x:0.5, y:0.42, w:12.3, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:AMBER, charSpacing:2, margin:0 });
s.addText("Insolvenz-Verfahrensachse über M&A-Phasen 0–6", { x:0.5, y:0.72, w:12.3, h:0.6, fontFace:HFONT, fontSize:28, bold:true, color:WHITE, margin:0 });

// --- Top lane: Insolvenzverfahren ---
s.addText("INSOLVENZ-VERFAHREN", { x:0.5, y:1.62, w:5, h:0.28, fontFace:BFONT, fontSize:11, bold:true, color:TEAL2, charSpacing:1.5, margin:0 });
const insSeg = [
  { x:0.5,  w:2.15, t:"Krise / drohende ZU", c:INKSOFT },
  { x:2.72, w:1.55, t:"Antrag", c:PRIM },
  { x:4.34, w:4.1,  t:"Vorläufiges Verfahren  ·  Insolvenzgeld (2–3 Mon.)", c:TEAL2 },
  { x:8.51, w:1.9,  t:"Eröffnung", c:PRIM },
  { x:10.48,w:2.35, t:"Berichts- & Prüfungstermin", c:INKSOFT },
];
insSeg.forEach(g => {
  s.addShape(p.ShapeType.roundRect, { x:g.x, y:1.94, w:g.w, h:0.68, rectRadius:0.05, fill:{color:g.c}, line:{color:"14464F", width:1} });
  s.addText(g.t, { x:g.x+0.06, y:1.94, w:g.w-0.12, h:0.68, fontFace:BFONT, fontSize:10.3, bold:true, color:WHITE, align:"center", valign:"middle", lineSpacing:11.5, margin:0 });
});
// arrow across
s.addShape(p.ShapeType.line, { x:0.5, y:2.95, w:12.33, h:0, line:{color:"2E5A62", width:1.5, dashType:"dash"} });

// --- Bottom lane: M&A phases with gate diamonds ---
s.addText("M&A-PROZESS  ·  PHASEN & ACHIEVEMENT-GATES", { x:0.5, y:3.18, w:8, h:0.28, fontFace:BFONT, fontSize:11, bold:true, color:AMBER, charSpacing:1.5, margin:0 });

const phases = [
  { x:0.5,  w:2.15, n:"0", t:"Sourcing & Ansprache" },
  { x:2.72, w:1.55, n:"1", t:"NDA & Interesse" },
  { x:4.34, w:2.0,  n:"2", t:"Unterlagen & Konzept" },
  { x:6.42, w:2.0,  n:"3", t:"DD & Bewertung" },
  { x:8.51, w:1.9,  n:"4", t:"C-Level & Preis" },
  { x:10.48,w:1.15, n:"5", t:"Verhandlung" },
  { x:11.71,w:1.12, n:"6", t:"Closing" },
];
phases.forEach(ph => {
  s.addShape(p.ShapeType.roundRect, { x:ph.x, y:3.55, w:ph.w, h:0.92, rectRadius:0.06, fill:{color:"0F454F"}, line:{color:PRIM, width:1} });
  s.addText(ph.n, { x:ph.x+0.08, y:3.62, w:0.5, h:0.4, fontFace:HFONT, fontSize:20, bold:true, color:AMBER, margin:0 });
  s.addText(ph.t, { x:ph.x+0.06, y:4.02, w:ph.w-0.12, h:0.42, fontFace:BFONT, fontSize:9.8, bold:true, color:"D8E8E9", align:"left", valign:"top", lineSpacing:10.5, margin:0 });
});

// gate diamonds row
// centered under each phase box, label widths clamped to stay on-canvas
const gates = [
  { c:1.575, g:"G0", t:"Target qualifiziert",              lw:1.75 },
  { c:3.495, g:"G1", t:"NDA + Datenraum",                  lw:1.75 },
  { c:5.34,  g:"G2", t:"Indik. Angebot / LOI",             lw:1.9  },
  { c:7.42,  g:"G3", t:"DD & Bewertung fertig",            lw:1.9  },
  { c:9.46,  g:"G4", t:"Freigabe Verwalter /\nGläubigerausschuss", lw:1.95 },
  { c:11.055,g:"G5", t:"Signing",                          lw:1.0  },
  { c:12.27, g:"G6", t:"Closing +\nIntegration",           lw:1.4  },
];
const gy = 5.15;
gates.forEach(gt => {
  const dx = gt.c - 0.31;
  s.addShape(p.ShapeType.diamond, { x:dx, y:gy, w:0.62, h:0.62, fill:{color:AMBER}, line:{color:INK, width:1.5}, shadow:shadow() });
  s.addText(gt.g, { x:gt.c-0.45, y:gy, w:0.9, h:0.62, fontFace:BFONT, fontSize:12, bold:true, color:INK, align:"center", valign:"middle", margin:0 });
  s.addText(gt.t, { x:gt.c-gt.lw/2, y:gy+0.68, w:gt.lw, h:0.72, fontFace:BFONT, fontSize:9, color:"C7D8DA", align:"center", lineSpacing:10, margin:0 });
});

s.addText("Best-Practice-Kaufzeitpunkt: im vorläufigen Verfahren / rund um die Eröffnung — bevor Patienten und Fachkräfte abwandern und solange das Insolvenzgeld die Löhne trägt.", {
  x:0.5, y:6.72, w:12.33, h:0.4, fontFace:BFONT, fontSize:11.5, italic:true, color:AMBER, align:"center", margin:0 });
s.addNotes("Hero-Folie: verknüpft Insolvenz-Zeitachse (oben) mit den M&A-Phasen 0–6 (unten). Rauten = Achievement-Gates, jede Phase endet mit einem klaren Deliverable. G4 ist der insolvenzspezifische Sonderschritt: Zustimmung von Gericht/Gläubigerausschuss.");

// =====================================================================
// SLIDE 7 — PHASEN 0–3 DETAIL
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Ablauf im Detail  ·  Teil 1", "Phasen 0–3: Sourcing bis Bewertung");

const ph1 = [
  { n:"0", t:"Sourcing & Ansprache", d:"Attraktive Targets identifizieren (Kaltakquise, Netzwerk, Verwalter/M&A-Berater). Insolvenz­bekanntmachungen laufend monitoren.", g:"G0 · Target qualifiziert" },
  { n:"1", t:"NDA & Verkaufsinteresse", d:"Verkaufsbereitschaft klären, NDA zeichnen, Zugang zum (schlanken) Datenraum. Verfahrensart & Ansprechpartner feststellen.", g:"G1 · NDA + Datenraum" },
  { n:"2", t:"Unterlagen & Konzept", d:"§ 132l/132a-Vertrag, Anlage 2 (Vergütung), MD-Prüfberichte, BWA. Integrations- & Skalierungskonzept skizzieren.", g:"G2 · Indik. Angebot / LOI" },
  { n:"3", t:"Due Diligence & Bewertung", d:"Dokumente nachrechnen, Regress-/Abrechnungsrisiken prüfen. WP & ggf. externes M&A-Team. Bewertung + Sanierungsplan.", g:"G3 · DD & Bewertung fertig" },
];
let y7 = 1.72;
ph1.forEach(ph => {
  s.addShape(p.ShapeType.roundRect, { x:0.5, y:y7, w:12.33, h:1.17, rectRadius:0.07, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  iconCircle(s, 0.74, y7+0.31, 0.56, PRIM, ph.n, 20);
  s.addText(ph.t, { x:1.5, y:y7+0.16, w:6.6, h:0.4, fontFace:BFONT, fontSize:15.5, bold:true, color:PRIM, margin:0 });
  s.addText(ph.d, { x:1.5, y:y7+0.54, w:7.5, h:0.55, fontFace:BFONT, fontSize:11.5, color:"33474C", lineSpacing:13.5, margin:0 });
  // gate chip
  s.addShape(p.ShapeType.diamond, { x:9.35, y:y7+0.4, w:0.4, h:0.4, fill:{color:AMBER}, line:{type:"none"} });
  s.addShape(p.ShapeType.roundRect, { x:9.85, y:y7+0.36, w:2.9, h:0.48, rectRadius:0.06, fill:{color:"FCF1DC"}, line:{color:AMBER, width:1} });
  s.addText(ph.g, { x:9.9, y:y7+0.36, w:2.8, h:0.48, fontFace:BFONT, fontSize:10.5, bold:true, color:"8A5A10", align:"center", valign:"middle", lineSpacing:11.5, margin:0 });
  y7 += 1.28;
});
footer(s,7);
s.addNotes("Phasen 0–3 = dein Standard-M&A, angereichert um Insolvenz- und Pflege-Spezifika (Bekanntmachungen monitoren, § 132l/Anlage 2/MD prüfen, Regressrisiken nachrechnen).");

// =====================================================================
// SLIDE 8 — PHASEN 4–6 DETAIL
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Ablauf im Detail  ·  Teil 2", "Phasen 4–6: Freigabe bis Integration");

const ph2 = [
  { n:"4", t:"C-Level & interne Kaufpreislogik", d:"Target intern vorstellen, Kaufpreis­komponenten festlegen (Assets, Patientenwert, Sanierungskosten, Risikoabschläge). Finanzierung sichern.", g:"— intern" },
  { n:"4a", t:"Freigabe Insolvenzorgane", d:"Insolvenzspezifischer Sonderschritt: Kaufvertrag steht unter der aufschiebenden Bedingung der Zustimmung von Insolvenzgericht + Gläubigerausschuss/-versammlung.", g:"G4 · Freigabe" , hl:true},
  { n:"5", t:"Verhandlung & Einigung", d:"Preis, Finanzierung, (minimale) Gewährleistungen, Betreiberwechsel/Zulassung mit den Kassen. Personalübergang (§ 613a) klären.", g:"G5 · Signing" },
  { n:"6", t:"Closing & Integration", d:"Vollzug, Übergang von Patienten & Personal, Anbindung an Konzernstruktur, PDL/Fachkräfte binden, Kassenverträge aktivieren.", g:"G6 · Closing" },
];
let y8 = 1.72;
ph2.forEach(ph => {
  const cardFill = ph.hl ? "0E5A66" : CARD;
  const txtCol = ph.hl ? WHITE : "33474C";
  const titleCol = ph.hl ? WHITE : PRIM;
  s.addShape(p.ShapeType.roundRect, { x:0.5, y:y8, w:12.33, h:1.17, rectRadius:0.07, fill:{color:cardFill}, line:{color:ph.hl?PRIM:LINE, width:1}, shadow:shadow() });
  iconCircle(s, 0.74, y8+0.31, 0.56, ph.hl?AMBER:PRIM, ph.n, ph.n.length>1?14:20);
  if(ph.hl){ s.addText(ph.n, { x:0.74, y:y8+0.31, w:0.56, h:0.56, fontFace:BFONT, fontSize:14, bold:true, color:INK, align:"center", valign:"middle", margin:0 }); }
  s.addText(ph.t, { x:1.5, y:y8+0.16, w:7.6, h:0.4, fontFace:BFONT, fontSize:15.5, bold:true, color:titleCol, margin:0 });
  s.addText(ph.d, { x:1.5, y:y8+0.54, w:7.6, h:0.58, fontFace:BFONT, fontSize:11.5, color:txtCol, lineSpacing:13.5, margin:0 });
  s.addShape(p.ShapeType.diamond, { x:9.55, y:y8+0.4, w:0.4, h:0.4, fill:{color:AMBER}, line:{type:"none"} });
  s.addShape(p.ShapeType.roundRect, { x:10.05, y:y8+0.36, w:2.55, h:0.48, rectRadius:0.06, fill:{color:ph.hl?"0A3D46":"FCF1DC"}, line:{color:AMBER, width:1} });
  s.addText(ph.g, { x:10.1, y:y8+0.36, w:2.45, h:0.48, fontFace:BFONT, fontSize:10.5, bold:true, color:ph.hl?AMBER:"8A5A10", align:"center", valign:"middle", margin:0 });
  y8 += 1.28;
});
footer(s,8);
s.addNotes("Phase 4a ist der Unterschied zum Normal-M&A: Ohne Zustimmung der Insolvenzorgane kein wirksamer Kauf. Deshalb Kaufvertrag unter aufschiebender Bedingung. In Phase 5/6: Betreiberwechsel/Zulassung mit Kassen ist erfolgskritisch.");

// =====================================================================
// SLIDE 9 — DUE DILIGENCE FOKUS PFLEGE
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Prüf-Fokus", "Due Diligence: die pflegespezifischen Pflichtchecks");

const dd = [
  { t:"Versorgungsverträge & Zulassung", d:"§ 132l / § 132a SGB V, § 72 SGB XI vorhanden, gültig, übertragbar? Noch in der Kassenliste gelistet?" },
  { t:"Vergütung (Anlage 2)", d:"Vergütungsvereinbarungen nachrechnen — tragfähig, marktüblich, verhandelbar?" },
  { t:"MD-Prüfung & Ruf", d:"Qualitätsprüfberichte, laufende Prüfverfahren, drohende Vertragskündigung, Mängelbescheide." },
  { t:"Abrechnung & Regress", d:"Leistungsnachweise vs. Abrechnung; latente Rückforderungen der Kassen (reichen Jahre zurück)." },
  { t:"Personalqualifikation", d:"Examinierte Fachkräfte, PDL vorhanden? Behandlungs-/Intensivpflege wirklich fachgerecht erbracht?" },
  { t:"Patientenstruktur", d:"Fallzahlen gegen Verordnungen/Nachweise; Konzentrationsrisiko; Loyalität beim Betreiberwechsel." },
];
const dcolW = 3.98, drowH = 1.62;
const dx0 = 0.5, dy0 = 1.72, dgap = 0.2;
dd.forEach((c,i) => {
  const col = i%3, row = Math.floor(i/3);
  const x = dx0 + col*(dcolW+dgap);
  const y = dy0 + row*(drowH+0.22);
  s.addShape(p.ShapeType.roundRect, { x, y, w:dcolW, h:drowH, rectRadius:0.07, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  iconCircle(s, x+0.24, y+0.24, 0.5, TEAL2, String(i+1), 15);
  s.addText(c.t, { x:x+0.86, y:y+0.22, w:dcolW-1.05, h:0.55, fontFace:BFONT, fontSize:13.5, bold:true, color:PRIM, valign:"middle", lineSpacing:15, margin:0 });
  s.addText(c.d, { x:x+0.24, y:y+0.82, w:dcolW-0.48, h:0.68, fontFace:BFONT, fontSize:11, color:"33474C", lineSpacing:13, margin:0 });
});
footer(s,9);
s.addNotes("Checklistencharakter für das DD-Team. Roter Faden: Deckt sich der abgerechnete Umsatz mit tatsächlich erbrachten, fachgerecht dokumentierten Leistungen — und ist die Zulassung sicher übertragbar?");

// =====================================================================
// SLIDE 10 — SCAM-RADAR
// =====================================================================
s = p.addSlide(); bg(s, INK);
s.addText("ROTE FLAGGEN", { x:0.5, y:0.42, w:12.3, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:AMBER, charSpacing:2, margin:0 });
s.addText("Scam-Radar: typische Täuschungen", { x:0.5, y:0.72, w:12.3, h:0.6, fontFace:HFONT, fontSize:30, bold:true, color:WHITE, margin:0 });

// two columns: general vs pflege
s.addText("ALLGEMEIN (INSOLVENZNÄHE)", { x:0.5, y:1.55, w:6, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:TEAL2, charSpacing:1, margin:0 });
const gen = [
  ["Geschönte Zahlen / verschwiegene Verluste", "Aufklärungspflicht verletzt → arglistige Täuschung, Rückabwicklung möglich"],
  ["Assets nicht frei", "Sicherungsübereignung, Eigentumsvorbehalt, Leasing — weniger Substanz als es scheint"],
  ["Cherry-Picking vor Insolvenz", "„Gute“ Verträge in neue Gesellschaft ausgegliedert; du bekommst die Hülle"],
];
let gy2 = 1.92;
gen.forEach(r => {
  s.addShape(p.ShapeType.roundRect, { x:0.5, y:gy2, w:6.0, h:1.15, rectRadius:0.06, fill:{color:INKSOFT}, line:{color:"1D5560", width:1} });
  s.addShape(p.ShapeType.roundRect, { x:0.5, y:gy2, w:0.14, h:1.15, rectRadius:0.03, fill:{color:AMBER}, line:{type:"none"} });
  s.addText(r[0], { x:0.78, y:gy2+0.14, w:5.55, h:0.35, fontFace:BFONT, fontSize:13, bold:true, color:WHITE, margin:0 });
  s.addText(r[1], { x:0.78, y:gy2+0.5, w:5.55, h:0.55, fontFace:BFONT, fontSize:11, color:"BCD2D4", lineSpacing:13, margin:0 });
  gy2 += 1.3;
});

s.addText("PFLEGE-SPEZIFISCH (DIE GEFÄHRLICHSTEN)", { x:6.83, y:1.55, w:6, h:0.3, fontFace:BFONT, fontSize:12, bold:true, color:RED, charSpacing:1, margin:0 });
const pfl = [
  ["Scheinpatienten / Karteileichen", "Aufgeblähte Fallzahlen → gegen Verordnungen & Kassenabrechnungen abgleichen"],
  ["Nicht erbrachte Leistungen", "Latente Regressforderungen der Kassen + § 263 StGB — beim Share Deal geerbt"],
  ["Qualifikations-Fake / Key-Person", "Unqualifiziertes Personal als Fachleistung; PDL wandert ab → Zulassung wackelt"],
];
let py2 = 1.92;
pfl.forEach(r => {
  s.addShape(p.ShapeType.roundRect, { x:6.83, y:py2, w:6.0, h:1.15, rectRadius:0.06, fill:{color:"3A1A16"}, line:{color:"6E2A22", width:1} });
  s.addShape(p.ShapeType.roundRect, { x:6.83, y:py2, w:0.14, h:1.15, rectRadius:0.03, fill:{color:RED}, line:{type:"none"} });
  s.addText(r[0], { x:7.11, y:py2+0.14, w:5.55, h:0.35, fontFace:BFONT, fontSize:13, bold:true, color:WHITE, margin:0 });
  s.addText(r[1], { x:7.11, y:py2+0.5, w:5.55, h:0.55, fontFace:BFONT, fontSize:11, color:"E7C6C1", lineSpacing:13, margin:0 });
  py2 += 1.3;
});

s.addShape(p.ShapeType.roundRect, { x:0.5, y:6.05, w:12.33, h:0.72, rectRadius:0.06, fill:{color:"0F454F"}, line:{color:PRIM, width:1} });
s.addText([
  { text:"Gegenmittel:  ", options:{ bold:true, color:AMBER, fontSize:13 } },
  { text:"Umsatz immer gegen erbrachte, dokumentierte Leistungen prüfen · Zulassung & § 132l-Status direkt bei den Kassen verifizieren · PDL/Fachkräfte vertraglich binden.", options:{ color:"D8E8E9", fontSize:12.5 } },
], { x:0.75, y:6.05, w:11.85, h:0.72, valign:"middle", lineSpacing:14, margin:0 });
s.addNotes("Der Scam-Radar. Wichtig: pflegespezifische Betrugsmuster sind gefährlicher, weil sie regulatorische Konsequenzen (Regress, Kündigung Versorgungsvertrag, Strafrecht) nach sich ziehen, nicht nur finanzielle.");

// =====================================================================
// SLIDE 11 — BEWERTUNG / KAUFPREISLOGIK
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Bewertung", "Kaufpreislogik im Distressed-Kontext");

// bridge-style building blocks
s.addText("Vom Substanzwert zum Angebotspreis — was einfließt:", { x:0.5, y:1.66, w:12, h:0.35, fontFace:BFONT, fontSize:14, bold:true, color:PRIM, margin:0 });

const blocks = [
  { t:"Substanz / Assets", d:"Inventar, Fahrzeuge, IT — nur unbelastete Gegenstände", c:PRIM, sign:"" },
  { t:"+ Going-Concern-Prämie", d:"Wert von Patientenstamm, Personal & laufender Zulassung", c:TEAL2, sign:"+" },
  { t:"− Sanierungskosten", d:"Personalaufbau, Zulassung/Betreiberwechsel, Nachqualifikation", c:AMBER, sign:"−" },
  { t:"− Risikoabschläge", d:"Latente Regresse, DD-Lücken, fehlende Garantien", c:RED, sign:"−" },
];
let bx = 0.5;
blocks.forEach((b,i) => {
  const bw = 2.95;
  s.addShape(p.ShapeType.roundRect, { x:bx, y:2.2, w:bw, h:2.05, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  s.addShape(p.ShapeType.roundRect, { x:bx, y:2.2, w:bw, h:0.62, rectRadius:0.08, fill:{color:b.c}, line:{type:"none"} });
  s.addShape(p.ShapeType.rect, { x:bx, y:2.5, w:bw, h:0.32, fill:{color:b.c}, line:{type:"none"} });
  s.addText(b.t, { x:bx+0.12, y:2.2, w:bw-0.24, h:0.62, fontFace:BFONT, fontSize:13.5, bold:true, color:b.c==AMBER?INK:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(b.d, { x:bx+0.2, y:2.95, w:bw-0.4, h:1.15, fontFace:BFONT, fontSize:11.5, color:"33474C", align:"left", valign:"top", lineSpacing:14, margin:0 });
  if(i<blocks.length-1){ s.addText("→", { x:bx+bw-0.02, y:2.85, w:0.4, h:0.6, fontFace:BFONT, fontSize:22, bold:true, color:MUTE, align:"center", valign:"middle", margin:0 }); }
  bx += bw + 0.18;
});

// result banner
s.addShape(p.ShapeType.roundRect, { x:0.5, y:4.55, w:12.33, h:1.95, rectRadius:0.08, fill:{color:INK}, line:{type:"none"}, shadow:shadow() });
s.addText("Ergebnis: risikoadjustierter Angebotspreis", { x:0.8, y:4.75, w:7, h:0.4, fontFace:HFONT, fontSize:18, bold:true, color:WHITE, margin:0 });
const notes11 = [
  "Kein EBITDA-Multiple wie im Normal-M&A — der Verwalter verkauft substanz-/masseorientiert.",
  "Fehlende Garantien preislich einkalkulieren: Risiken trägt der Käufer.",
  "„Return in max. 3 Monaten“ nur bei gesunder Substanz + gesicherter Zulassung realistisch.",
];
let ny = 5.2;
notes11.forEach(t => {
  s.addShape(p.ShapeType.ellipse, { x:0.85, y:ny+0.06, w:0.12, h:0.12, fill:{color:AMBER}, line:{type:"none"} });
  s.addText(t, { x:1.1, y:ny-0.05, w:11.5, h:0.35, fontFace:BFONT, fontSize:12.5, color:"D8E8E9", margin:0 });
  ny += 0.42;
});
footer(s,11);
s.addNotes("Bewertung: additive „Bridge“ vom Substanzwert. Wichtig, dem C-Level zu vermitteln, dass Sanierungskosten (v. a. Zulassung/Personal) und fehlende Garantien den Preis drücken.");

// =====================================================================
// SLIDE 12 — INTEGRATION & SKALIERUNG (2 HEBEL)
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Nach dem Closing", "Integration & Skalierung — die zwei Hebel");

const levers = [
  { t:"Mitarbeiter", ic:"MA", c:PRIM, pts:[
    "PDL & examinierte Fachkräfte sofort binden (Retention, Kommunikation)",
    "Nachqualifikation sichert Abrechenbarkeit der Behandlungspflege",
    "Personalschlüssel nach IPReG/AKI-Richtlinie erfüllen",
    "Recruiting-Pipeline des Konzerns nutzen" ] },
  { t:"Patienten", ic:"PT", c:TEAL2, pts:[
    "Versorgungskontinuität beim Betreiberwechsel sichern (kein Abwandern)",
    "Verordnungen/Genehmigungen mit Kassen sauber überleiten",
    "Auslastung der Intensiv-WGs erhöhen (Margenhebel)",
    "Zuweiser-/Kliniknetzwerk ausbauen" ] },
];
let lx = 0.5;
levers.forEach(l => {
  const lw = 6.05;
  s.addShape(p.ShapeType.roundRect, { x:lx, y:1.75, w:lw, h:4.6, rectRadius:0.08, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  iconCircle(s, lx+0.35, 2.05, 0.85, l.c, l.ic, 20);
  s.addText(l.t, { x:lx+1.4, y:2.12, w:lw-1.6, h:0.5, fontFace:HFONT, fontSize:24, bold:true, color:l.c, valign:"middle", margin:0 });
  s.addText("Werttreiber #"+(l.t==="Mitarbeiter"?"1":"2"), { x:lx+1.4, y:2.62, w:lw-1.6, h:0.3, fontFace:BFONT, fontSize:11, italic:true, color:MUTE, margin:0 });
  let ly = 3.3;
  l.pts.forEach(pt => {
    s.addShape(p.ShapeType.ellipse, { x:lx+0.4, y:ly+0.05, w:0.14, h:0.14, fill:{color:AMBER}, line:{type:"none"} });
    s.addText(pt, { x:lx+0.7, y:ly-0.08, w:lw-1.0, h:0.6, fontFace:BFONT, fontSize:12.5, color:"33474C", lineSpacing:14.5, margin:0 });
    ly += 0.72;
  });
  lx += lw + 0.23;
});
footer(s,12);
s.addNotes("Integration entlang der zwei Hebel, die der Nutzer selbst als zentral benennt: Personal und Patienten. Beide sind in der Pflege zugleich Werttreiber UND Voraussetzung der Zulassung.");

// =====================================================================
// SLIDE 13 — RISIKO-AMPEL
// =====================================================================
s = p.addSlide(); bg(s, LIGHT);
titleBlock(s, "Steuerung", "Risiko-Ampel: worauf das Deal-Team schaut");

const risks = [
  { r:"Versorgungsvertrag geht nicht über / Abrechnungslücke", lvl:"HOCH", c:RED },
  { r:"Latente Regressforderungen der Kassen", lvl:"HOCH", c:RED },
  { r:"Abwanderung von PDL / Fachkräften", lvl:"HOCH", c:RED },
  { r:"Personalqualifikation nicht IPReG/AKI-konform", lvl:"MITTEL", c:AMBER },
  { r:"Patientenabwanderung beim Betreiberwechsel", lvl:"MITTEL", c:AMBER },
  { r:"Belastete / geleaste Assets", lvl:"MITTEL", c:AMBER },
  { r:"Zustimmung Gläubigerausschuss verzögert", lvl:"NIEDRIG", c:TEAL2 },
  { r:"Altverbindlichkeiten (bei sauberem Asset Deal)", lvl:"NIEDRIG", c:TEAL2 },
];
const colX2 = [0.5, 6.83];
const rcw = 6.0, rrh = 0.72;
risks.forEach((rk,i) => {
  const col = Math.floor(i/4), row = i%4;
  const x = colX2[col], y = 1.85 + row*(rrh+0.2);
  s.addShape(p.ShapeType.roundRect, { x, y, w:rcw, h:rrh, rectRadius:0.06, fill:{color:CARD}, line:{color:LINE, width:1}, shadow:shadow() });
  s.addShape(p.ShapeType.ellipse, { x:x+0.22, y:y+0.21, w:0.3, h:0.3, fill:{color:rk.c}, line:{type:"none"} });
  s.addText(rk.r, { x:x+0.72, y, w:rcw-2.35, h:rrh, fontFace:BFONT, fontSize:12, bold:true, color:"233A3F", valign:"middle", lineSpacing:13.5, margin:0 });
  s.addShape(p.ShapeType.roundRect, { x:x+rcw-1.5, y:y+0.19, w:1.3, h:0.34, rectRadius:0.05, fill:{color:rk.c}, line:{type:"none"} });
  s.addText(rk.lvl, { x:x+rcw-1.5, y:y+0.19, w:1.3, h:0.34, fontFace:BFONT, fontSize:10.5, bold:true, color:rk.c==AMBER?INK:WHITE, align:"center", valign:"middle", margin:0 });
});
footer(s,13);
s.addNotes("Priorisierte Risikosicht. Die drei roten Punkte sind allesamt pflegespezifisch und zulassungsnah — genau dort liegt der Unterschied zu einem generischen Distressed-Deal.");

// =====================================================================
// SLIDE 14 — NÄCHSTE SCHRITTE / CLOSING
// =====================================================================
s = p.addSlide(); bg(s, INK);
s.addShape(p.ShapeType.ellipse, { x:10.2, y:-1.4, w:4.6, h:4.6, fill:{type:"none"}, line:{color:PRIM, width:1.5, transparency:40} });
s.addShape(p.ShapeType.ellipse, { x:-1.2, y:4.9, w:3.8, h:3.8, fill:{type:"none"}, line:{color:AMBER, width:1.5, transparency:55} });

s.addText("NÄCHSTE SCHRITTE", { x:0.7, y:0.9, w:10, h:0.35, fontFace:BFONT, fontSize:13, bold:true, color:AMBER, charSpacing:3, margin:0 });
s.addText("So kommt der Prozess ins Rollen", { x:0.68, y:1.28, w:11.5, h:0.75, fontFace:HFONT, fontSize:34, bold:true, color:WHITE, margin:0 });

const steps = [
  { t:"Deal-Team & Berater aufsetzen", d:"Fachanwalt Insolvenz- + Medizinrecht, WP, ggf. externes M&A-Team" },
  { t:"Target-Screening starten", d:"Insolvenzbekanntmachungen + Netzwerk; Verfahrensart je Target erfassen" },
  { t:"Struktur-Entscheidung treffen", d:"Asset Deal vs. Insolvenzplan — Leitfrage: bleibt die Kassenzulassung?" },
  { t:"DD-Checkliste & Gate-Plan finalisieren", d:"Pflege-Prüfpunkte + Achievement-Gates als Steuerungsraster übernehmen" },
];
let stx = 0.7, sty = 2.5;
steps.forEach((st,i) => {
  const col = i%2, row = Math.floor(i/2);
  const x = stx + col*6.15, y = sty + row*1.55;
  s.addShape(p.ShapeType.roundRect, { x, y, w:5.9, h:1.35, rectRadius:0.08, fill:{color:INKSOFT}, line:{color:"1D5560", width:1} });
  iconCircle(s, x+0.28, y+0.4, 0.55, AMBER, String(i+1), 18);
  s.addText(String(i+1), { x:x+0.28, y:y+0.4, w:0.55, h:0.55, fontFace:BFONT, fontSize:18, bold:true, color:INK, align:"center", valign:"middle", margin:0 });
  s.addText(st.t, { x:x+1.05, y:y+0.2, w:4.7, h:0.45, fontFace:BFONT, fontSize:14.5, bold:true, color:WHITE, valign:"middle", margin:0 });
  s.addText(st.d, { x:x+1.05, y:y+0.66, w:4.7, h:0.55, fontFace:BFONT, fontSize:11, color:"C7D8DA", lineSpacing:13, margin:0 });
});

s.addText("Hinweis: Struktur- und Rechtsfragen (InsO, SGB V/XI, § 613a, Anfechtung) im Einzelfall anwaltlich absichern — dieser Leitfaden ist Orientierung, kein Rechtsrat.", {
  x:0.7, y:6.35, w:12, h:0.5, fontFace:BFONT, fontSize:11, italic:true, color:"9FBAC0", margin:0 });
s.addNotes("Abschluss mit vier konkreten Next Steps. Leitfrage der Struktur-Entscheidung wiederholen: Bleibt die Kassenzulassung erhalten? Disclaimer: kein Rechtsrat.");

// ---- write ----
p.writeFile({ fileName: "Kauf_insolventer_Pflegedienst.pptx" })
 .then(f => console.log("WROTE", f))
 .catch(e => { console.error(e); process.exit(1); });
