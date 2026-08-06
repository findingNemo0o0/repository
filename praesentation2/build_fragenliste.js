// Erzeugt die Fragenliste für das Gespräch mit der Eigentümerin von SKP.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  Header, Footer, PageNumber,
} = require("docx");

const BLUE = "1A70B8";
const AMBER = "E0861A";
const RED = "C0392B";
const OLIVE = "6E7D00";
const TXT = "3D3D3D";
const MUT = "777777";
const HF = "Quicksand";
const BF = "Open Sans";

const W_FRAGE = 6300;
const W_ANTWORT = 3338;
const TOTAL = W_FRAGE + W_ANTWORT;

const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const HAIR = { style: BorderStyle.SINGLE, size: 2, color: "DDDDDD" };

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 0, after: opts.after ?? 60, line: 260 },
    alignment: opts.align,
    children: [new TextRun({
      text,
      font: opts.font ?? BF,
      size: opts.size ?? 20,
      bold: opts.bold ?? false,
      italics: opts.italics ?? false,
      color: opts.color ?? TXT,
    })],
  });
}

function runs(parts, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 0, after: opts.after ?? 60, line: 260 },
    children: parts.map((r) => new TextRun({
      text: r.t,
      font: r.font ?? BF,
      size: r.size ?? 20,
      bold: r.bold ?? false,
      italics: r.italics ?? false,
      color: r.color ?? TXT,
    })),
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 340, after: 140 },
    children: [new TextRun({ text, font: HF, size: 26, bold: true, color: BLUE })],
  });
}

function h2(text, color = BLUE) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text, font: HF, size: 23, bold: true, color })],
  });
}

function cell(children, { width, shading, borders } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shading
      ? { type: ShadingType.CLEAR, fill: shading, color: "auto" }
      : undefined,
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    borders: borders ?? { top: HAIR, bottom: HAIR, left: NONE, right: NONE },
    children,
  });
}

// Eine Frageblock-Tabelle: links Frage + Begründung + Unterlage, rechts Platz für die Antwort
function frageTabelle(items) {
  const head = new TableRow({
    tableHeader: true,
    children: [
      cell([p("Frage", { bold: true, font: HF, size: 19, color: BLUE, after: 0 })],
        { width: W_FRAGE, shading: "EAF1F8" }),
      cell([p("Antwort", { bold: true, font: HF, size: 19, color: BLUE, after: 0 })],
        { width: W_ANTWORT, shading: "EAF1F8" }),
    ],
  });

  const rows = items.map((it) => {
    const left = [
      runs([
        { t: it.nr + "  ", bold: true, color: BLUE, font: HF, size: 20 },
        { t: it.frage, bold: true, size: 20 },
      ], { after: 60 }),
      runs([
        { t: "Warum: ", bold: true, size: 17, color: MUT },
        { t: it.warum, size: 17, color: MUT },
      ], { after: it.beleg ? 40 : 0 }),
    ];
    if (it.beleg) {
      left.push(runs([
        { t: "Unterlage: ", bold: true, size: 17, color: OLIVE },
        { t: it.beleg, size: 17, color: OLIVE, italics: true },
      ], { after: 0 }));
    }
    return new TableRow({
      cantSplit: true,
      children: [
        cell(left, { width: W_FRAGE }),
        cell([p("", { after: 0 }), p("", { after: 0 }), p("", { after: 0 })],
          { width: W_ANTWORT }),
      ],
    });
  });

  return new Table({
    columnWidths: [W_FRAGE, W_ANTWORT],
    width: { size: TOTAL, type: WidthType.DXA },
    rows: [head, ...rows],
  });
}

function box(title, lines, color = AMBER, fill = "FDF3E3") {
  const kids = [p(title, { bold: true, font: HF, size: 21, color, after: 80 })];
  lines.forEach((l, i) =>
    kids.push(runs([
      { t: "▪  ", color },
      { t: l, size: 19 },
    ], { after: i === lines.length - 1 ? 0 : 60 })));
  return new Table({
    columnWidths: [TOTAL],
    width: { size: TOTAL, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: TOTAL, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill, color: "auto" },
        margins: { top: 140, bottom: 140, left: 160, right: 160 },
        borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
        children: kids,
      })],
    })],
  });
}

// ---------------------------------------------------------------- Fragenblöcke
const A = [
  { nr: "A1", frage: "Ist der Kreis Pinneberg wirklich das einzige Einsatzgebiet?",
    warum: "Öffentlich belegbar ist ausschließlich der Kreis Pinneberg — s-k-p.net nennt kein weiteres Gebiet. Jede zusätzliche Region verändert die Bewertung erheblich, weil sie das Klumpenrisiko senkt.",
    beleg: "Standortliste je Schule mit Kreis und Beginn der Tätigkeit" },
  { nr: "A2", frage: "An wie vielen Schulen ist SKP heute tätig, und wie verteilt sich das Volumen darauf?",
    warum: "Im Poolmodell wird je Schule ein Träger beauftragt. Die Schulliste ist damit die eigentliche Umsatzliste — nicht die Zahl der Kinder.",
    beleg: "Je Schule: Schulart, Zahl der Kinder, Wochenstunden, Jahresumsatz" },
  { nr: "A3", frage: "Wie viele Kinder werden aktuell begleitet — und wie viele davon nach § 35a SGB VIII gegenüber § 112 SGB IX?",
    warum: "Die beiden Rechtskreise werden ab 2028 zusammengeführt. Wer heute in beiden arbeitet, kennt beide Verfahren und beide Ämter.",
    beleg: "Fallliste zum Stichtag, nach Rechtsgrundlage getrennt" },
  { nr: "A4", frage: "Wie verteilt sich das Volumen auf Grundschulen, weiterführende Schulen und Förderzentren?",
    warum: "Die Reform beginnt an Grundschulen. Weiterführende Schulen und Förderzentren bleiben am längsten im Einzelfall — das ist der stabile Teil des Geschäfts.",
    beleg: "Aufteilung nach Schulart, drei Schuljahre" },
  { nr: "A5", frage: "Wie hat sich die Zahl der begleiteten Kinder in den letzten drei Schuljahren entwickelt?",
    warum: "Landesweit steigt die Nachfrage deutlich. Bleibt SKP dahinter zurück, liegt es nicht am Markt, sondern am Unternehmen.",
    beleg: "Fallzahlen je Schuljahr, Zugänge und Abgänge" },
];

const B = [
  { nr: "B1", frage: "Welche Ihrer Schulen liegen in Stufe 1 oder Stufe 2 des Pinneberger Klassenassistenz-Plans?",
    warum: "Stufe 1 sind die Grundschulen in Elmshorn, Pinneberg und Barmstedt, Stufe 2 unter anderem Tornesch, Uetersen, Wedel, Schenefeld und Quickborn. Das ist das Volumen, das in einem Vergabeverfahren vollständig auf dem Spiel steht — nicht anteilig.",
    beleg: "Abgleich der Standortliste mit den Stufen des Kreiskonzepts" },
  { nr: "B2", frage: "Haben Sie am Teilnahmewettbewerb 2024 teilgenommen — und mit welchem Ergebnis?",
    warum: "Der Kreis hat die Klassenassistenz am 25.04.2024 EU-weit ausgeschrieben (TED 246925-2024). Wer die Eignungsprüfung damals bestanden hat, hat den schwierigsten Teil bereits hinter sich.",
    beleg: "Teilnahmeantrag, Eignungsmitteilung, Schriftverkehr mit der Vergabestelle" },
  { nr: "B3", frage: "Was ist Ihr letzter Kenntnisstand zum Zeitplan des Kreises?",
    warum: "Der Kreistag hat die Einführung am 11.12.2024 um zwei Jahre verschoben, eine Zuschlagsbekanntmachung ist bis heute nicht erschienen. Wer den Zeitplan aus erster Hand kennt, hat einen Informationsvorsprung, den wir nicht kaufen können.",
    beleg: "Schriftverkehr mit dem Kreis, Protokolle, Einladungen zu Trägerrunden" },
  { nr: "B4", frage: "Wurden Ihre Bewilligungsbescheide zuletzt einheitlich auf ein Schuljahresende befristet?",
    warum: "Das ist der verlässlichste Frühwarnindikator überhaupt. Der Kreis Pinneberg hat genau so umgesteuert — erst auf den 31.07.2024, dann auf den 31.07.2025. Wenn das passiert, ist die Umstellung beschlossen, auch wenn nichts verkündet wurde.",
    beleg: "Zehn aktuelle Bescheide mit Befristungsdatum" },
  { nr: "B5", frage: "Wie hat sich das Poolmodell in Tornesch/Uetersen auf SKP ausgewirkt?",
    warum: "Dort läuft seit 2018 ein Pool mit „Familienräume“ als Träger, an sieben Grundschulen. Wenn SKP dort Kinder verloren hat, kennen wir die Verlustquote bereits aus der Praxis.",
    beleg: "Fallzahlen in Tornesch und Uetersen vor und nach 2018" },
];

const C = [
  { nr: "C1", frage: "Welche Deckungssumme hat Ihre Betriebs- und Berufshaftpflicht?",
    warum: "Die Pinneberger Ausschreibung verlangt 10 Mio. €. Wer sie nicht nachweist, wird ausgeschlossen — ohne Nachverhandlung. Eine Aufstockung ist möglich, aber sie kostet Zeit.",
    beleg: "Versicherungsschein und Nachtrag" },
  { nr: "C2", frage: "Sind Sie im amtlichen Verzeichnis präqualifizierter Unternehmen eingetragen?",
    warum: "Mit PQ-Nummer (DIHK, pq-vol.de) genügt im Verfahren ein Verweis. Ohne sie muss jeder Nachweis einzeln beigebracht werden — Nachreichfrist sechs Kalendertage.",
    beleg: "PQ-Nummer oder Nachweis der Beantragung" },
  { nr: "C3", frage: "Können Sie eine Referenzliste über die letzten vier Jahre vorlegen?",
    warum: "Gefordert ist je Schulbegleitungsverhältnis Leistungsumfang, öffentlicher Auftraggeber und Ansprechpartner. Rückwirkend lässt sich das kaum aufbauen.",
    beleg: "Referenzliste in der geforderten Form" },
  { nr: "C4", frage: "Kann Ihre Buchhaltung den Umsatz der Leistungsart Schulbegleitung getrennt ausweisen — für drei Geschäftsjahre?",
    warum: "Gefordert sind Gesamtumsatz UND separat der Umsatz der Leistungsart. Wenn die Buchhaltung das nie getrennt hat, ist es nachträglich kaum rekonstruierbar.",
    beleg: "BWA oder Auswertung mit getrennter Umsatzausweisung" },
  { nr: "C5", frage: "Sind Sie als Träger der freien Jugendhilfe nach § 75 SGB VIII anerkannt?",
    warum: "Ohne Anerkennung ist der Zugang zu Vereinbarungsverfahren und zu Gremien beschränkt.",
    beleg: "Anerkennungsbescheid" },
  { nr: "C6", frage: "Liegen Schutzkonzept nach § 8a, die Regelung nach § 72a und der Zugang zu einer insoweit erfahrenen Fachkraft nach § 8b vor?",
    warum: "Mindeststandards des Kreiskonzepts. Erweiterte Führungszeugnisse sind bei Neueinstellung und danach alle fünf Jahre vorzulegen — die Praxis ist prüfbar.",
    beleg: "Schutzkonzept, Verfahrensanweisung, Nachweis der Führungszeugnispraxis" },
  { nr: "C7", frage: "SKP ist ein eingetragener Kaufmann. Ist eine Umwandlung geplant oder vorbereitet?",
    warum: "Die Rechtsform berührt Haftung, Vertragsübergang und die Frage, ob bestehende Vereinbarungen bei einem Verkauf fortgelten.",
    beleg: "Handelsregisterauszug, Gesellschaftsvertrag falls vorhanden" },
];

const D = [
  { nr: "D1", frage: "Wie viele Beschäftigte haben Sie, in Köpfen und in Vollzeitäquivalenten — und wie viele davon einen pädagogischen, therapeutischen oder pflegerischen Abschluss?",
    warum: "Bundesweit ist rund die Hälfte der Schulbegleitungen fachlich qualifiziert. Der Wert allein entscheidet nichts, aber er ist die Grundlage für D2.",
    beleg: "Personalliste zum Stichtag: Qualifikation, Stundenumfang, Vertragsart" },
  { nr: "D2", frage: "Wie viele Ihrer Beschäftigten erfüllen die Anforderung an eine Teamleitung — Sozialpädagogik, Pädagogik, Erzieher:in oder gleichwertig?",
    warum: "Das ist die entscheidende Frage der ganzen Liste. Das Kreiskonzept verlangt für die Teamleitung ausschließlich Fachkräfte, Schlüssel 1:15 bei 39 Wochenstunden. Ohne diese Rolle gibt es keinen Zuschlag — und sie ist die teuerste Lücke, wenn sie besteht.",
    beleg: "Namentliche Aufstellung mit Abschluss und Berufserfahrung" },
  { nr: "D3", frage: "Setzen Sie Honorarkräfte ein, und in welchem Umfang?",
    warum: "Das Kreiskonzept sagt wörtlich: „Es werden keine Honorarkräfte eingesetzt.“ Festanstellung ist Qualitätsmerkmal und Vertragsbedingung, nicht Verhandlungssache.",
    beleg: "Aufteilung Festanstellung gegenüber Honorar, Musterverträge beider Formen" },
  { nr: "D4", frage: "Gibt es ein schriftliches Fortbildungs- und Qualifizierungskonzept, Supervision und regelmäßige Teamsitzungen?",
    warum: "Alle drei sind gefordert und müssen zum Leistungsbeginn umsetzbar sein. Ein Qualifizierungskonzept ist zugleich eines der sieben Zuschlagskriterien — und der billigste Differenzierer im Wettbewerb.",
    beleg: "Konzept, Fortbildungsplan, Nachweis der Supervision" },
  { nr: "D5", frage: "Haben Sie Erfahrung damit, dass eine Kraft mehrere Kinder gleichzeitig begleitet?",
    warum: "Das ist die operative Kernkompetenz im Pool und in den Hamburger Kombinationsmaßnahmen. Erfahrung damit ist im Verfahren belegbar und im Betrieb sofort wirksam.",
    beleg: "Beispiele, Einsatzpläne, Rückmeldungen von Schulen" },
  { nr: "D6", frage: "Wie hoch sind Fluktuation, Vakanzquote und Krankenstand der letzten zwei Jahre?",
    warum: "Im Pool trägt der Träger Vertretung und Krankheit aus dem Budget. Ein hoher Krankenstand schlägt dann unmittelbar auf das Ergebnis durch.",
    beleg: "Personalkennzahlen 24 Monate" },
];

const E = [
  { nr: "E1", frage: "Sind Sie tarifgebunden — und wenn ja, an welchen Tarif?",
    warum: "§ 124 Abs. 1 SGB IX bestimmt, dass die Zahlung tariflicher Vergütung nicht als unwirtschaftlich abgelehnt werden darf. Dieser Schutz gilt nur für tarifgebundene Träger.",
    beleg: "Angabe zur Tarifbindung, gegebenenfalls Tarifwerk" },
  { nr: "E2", frage: "Wie hoch ist der tatsächliche Stundenlohn brutto, und wie weit reicht die Spanne?",
    warum: "Zum Vergleich: Mindestlohn 13,90 € (2026) und 14,60 € (2027). Der Kreis Pinneberg kalkuliert die Klassenassistenz nach TVöD SuE S2 Stufe 5, also rund 19,60 € je Stunde. Wer heute deutlich darunter zahlt, hat Spielraum — wer darüber liegt, hat ein Problem.",
    beleg: "Lohnjournal zwölf Monate, Musterarbeitsvertrag" },
  { nr: "E3", frage: "Wie sind die Schulferien vertraglich geregelt — durchbezahlt, Jahresarbeitszeitkonto oder unbezahlt?",
    warum: "Das ist der größte einzelne Hebel in der Personalkostenquote. Im Pool lässt er sich nicht mehr über Zusatzstunden ausgleichen.",
    beleg: "Musterarbeitsvertrag, Betriebsvereinbarung falls vorhanden" },
  { nr: "E4", frage: "Wie werden Vertretung, Krankheit, Klassenfahrten und Ausflüge heute vergütet?",
    warum: "Heute sind Fahrten als Zusatzstunden abrechenbar — in Pinneberg 8 Stunden je Tag ohne vorherigen Antrag. Im Pool zahlt der Träger das aus dem Budget. Der Unterschied ist bezifferbar.",
    beleg: "Abrechnungen von zwölf Monaten mit Zusatzstunden" },
  { nr: "E5", frage: "Welche Leistungs- und Vergütungsvereinbarung gilt, mit wem, seit wann — und wann wurde zuletzt verhandelt?",
    warum: "Laufzeit und Kündigungsfrist bestimmen, wie lange die heutige Preisbasis trägt. Das Ergebnis der letzten Verhandlung zeigt, wie durchsetzungsfähig SKP gegenüber dem Kreis ist.",
    beleg: "Vereinbarung mit Anlagen, Protokoll der letzten Verhandlung" },
];

const F = [
  { nr: "F1", frage: "Umsatz und Ergebnis der letzten drei Geschäftsjahre, getrennt nach Leistungsart?",
    warum: "Grundlage der Bewertung — und zugleich eine Eignungsanforderung im Vergabeverfahren (siehe C4).",
    beleg: "Jahresabschlüsse, BWA, Umsatzaufteilung" },
  { nr: "F2", frage: "Wie hoch ist die Personalkostenquote?",
    warum: "Sie bestimmt, wie stark ein Festpreis über bis zu neun Jahre auf das Ergebnis durchschlägt.",
    beleg: "GuV-Auswertung drei Jahre" },
  { nr: "F3", frage: "Wie sieht der Umsatzverlauf über ein Schuljahr aus?",
    warum: "Schulbegleitung ist ausgeprägt saisonal. Für die Liquiditätsplanung nach Übernahme brauchen wir den Monatsverlauf, nicht den Jahreswert.",
    beleg: "Monatsumsätze über 24 Monate" },
  { nr: "F4", frage: "Wie zahlt der Kreis — Zahlungsziel, Zahlungsverhalten, offener Forderungsbestand?",
    warum: "Verzögerte Bescheide und späte Zahlungen sind in der Schulbegleitung verbreitet und binden Working Capital.",
    beleg: "Offene-Posten-Liste, durchschnittliche Zahlungsdauer" },
  { nr: "F5", frage: "Gibt es ein zweites Standbein neben der Schulbegleitung?",
    warum: "Die Website nennt zusätzlich Beratung. Ein zweites Standbein senkt die Abhängigkeit von einer einzigen Vergabeentscheidung erheblich.",
    beleg: "Umsatz je Leistungsart, drei Jahre" },
];

const G = [
  { nr: "G1", frage: "Welche Aufgaben hängen persönlich an Ihnen?",
    warum: "Akquise, Schulkontakte, Personalauswahl, Abrechnung, Gremienarbeit — je mehr davon an einer Person hängt, desto größer das Übergangsrisiko.",
    beleg: "Aufgabenverteilung, Organigramm" },
  { nr: "G2", frage: "Gibt es eine zweite Führungsebene, und wer vertritt Sie?",
    warum: "Im Poolmodell wird ohnehin eine Teamleitungsebene gebraucht. Wer sie schon hat, muss sie nicht erst aufbauen.",
    beleg: "Organigramm, Stellenbeschreibungen" },
  { nr: "G3", frage: "Sind Ihre Beziehungen zu Schulleitungen und zum Jugendamt persönlich oder schriftlich verankert?",
    warum: "Ein Pool kommt nur „im Einvernehmen mit der Schulleitung“ zustande. Diese Beziehungen sind der eigentliche Vermögenswert — und genau der, der bei einem Eigentümerwechsel verloren gehen kann.",
    beleg: "Kooperationsvereinbarungen, Schulliste mit namentlichen Ansprechpartnern" },
  { nr: "G4", frage: "Wären Sie bereit, nach einem Verkauf für eine Übergangszeit an Bord zu bleiben?",
    warum: "Ergibt sich unmittelbar aus G1 und G3. Die Dauer ist verhandelbar, die Notwendigkeit meist nicht.",
    beleg: "—" },
  { nr: "G5", frage: "Warum verkaufen Sie, und warum jetzt?",
    warum: "Die ehrlichste Frage der Liste. Die Antwort sagt mehr über die Risiken aus als jede Kennzahl — besonders, wenn die Reform darin vorkommt.",
    beleg: "—" },
  { nr: "G6", frage: "Sprechen Sie mit weiteren Interessenten, etwa mit FiB oder Familienräume?",
    warum: "Beide sitzen im selben Kreis. Ein Zusammenschluss unter den drei Pinneberger Anbietern wäre im Vergabeverfahren eine ernstzunehmende Alternative zu einem Verkauf an uns.",
    beleg: "—" },
];

const H = [
  { nr: "H1", frage: "Wie schätzen Sie FiB und Familienräume ein — wo überschneiden sich die Gebiete?",
    warum: "Beide sind ebenfalls im Kreis Pinneberg tätig, FiB zusätzlich in Hamburg. Familienräume ist der Träger des Poolmodells in Tornesch/Uetersen. Wer im Verfahren gegen wen antritt, entscheidet über den Preis.",
    beleg: "Eigene Einschätzung, Schulen mit Überschneidung" },
  { nr: "H2", frage: "Sind Sie in Gremien vertreten — Arbeitsgemeinschaft nach § 78 SGB VIII, Arbeitskreise des Kreises, Trägerrunden?",
    warum: "Gremienpräsenz ist der Kanal, über den Kreise ihre Modelle vorbereiten. Wer dort sitzt, erfährt Zeitpläne, bevor sie veröffentlicht werden.",
    beleg: "Mitgliedschaften, Einladungen, Protokolle" },
  { nr: "H3", frage: "Gab es Anfragen von Schulen oder Ämtern außerhalb des Kreises Pinneberg?",
    warum: "Zeigt, ob eine Ausweitung realistisch ist — und ob der Name über die Kreisgrenze hinaus bekannt ist.",
    beleg: "Anfragen der letzten zwei Jahre" },
  { nr: "H4", frage: "Was müsste aus Ihrer Sicht passieren, damit SKP einen Standortvertrag gewinnt?",
    warum: "Offene Schlussfrage. Sie zeigt, wie gut die Eigentümerin die Reform verstanden hat — und ob sie einen Plan hat oder abwartet.",
    beleg: "—" },
];

// ------------------------------------------------------------------- Dokument
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: BF, size: 20, color: TXT } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },        // A4 hoch
        margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [runs([
          { t: "Fragenliste Eigentümerin SKP", size: 16, color: MUT, font: HF },
          { t: "   ·   vertraulich, nur für den internen Gebrauch", size: 16, color: MUT },
        ], { after: 0 })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            children: ["Seite ", PageNumber.CURRENT, " von ", PageNumber.TOTAL_PAGES],
            size: 16, color: MUT, font: BF,
          })],
        })],
      }),
    },
    children: [
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({
          text: "Fragenliste für das Gespräch mit der Eigentümerin von SKP",
          font: HF, size: 34, bold: true, color: BLUE,
        })],
      }),
      p("Soziale Kompetenzen Kreis Pinneberg e. K., Uetersen  ·  Stand August 2026",
        { color: MUT, size: 19, after: 200 }),

      p("Die Liste prüft nicht das Unternehmen im Allgemeinen, sondern genau die Punkte, an denen sich die Poolreform auf den Wert auswirkt. Jede Frage nennt darum mit, warum wir sie stellen und welche Unterlage die Antwort belegt. Die rechte Spalte bleibt frei für Notizen im Gespräch.",
        { after: 120 }),
      p("Acht Blöcke, 38 Fragen. Realistisch sind zwei Termine: Block A bis D im ersten Gespräch, E bis H im zweiten, nachdem die Unterlagen vorliegen.",
        { after: 200 }),

      box("Die fünf Antworten, an denen die Bewertung wirklich hängt", [
        "D2 — Gibt es Fachkräfte für die Teamleitung? Ohne sie kein Zuschlag, und die Lücke ist teuer.",
        "B1 — Welcher Umsatzanteil liegt an Schulen der Stufen 1 und 2? Das ist das Volumen, das ein einziges Verfahren entscheidet.",
        "A1 — Ist Pinneberg wirklich das einzige Gebiet? Jede weitere Region senkt das Klumpenrisiko spürbar.",
        "E2 und E3 — Ist-Stundenlohn und Ferienregelung. Sie entscheiden, ob unter einem Festpreis noch Marge bleibt.",
        "G3 — Sind die Schulbeziehungen schriftlich verankert oder hängen sie an der Eigentümerin?",
      ], BLUE, "EAF1F8"),

      h1("A · Gebiet und Volumen"),
      p("Was wir kaufen würden — und wie breit es aufgestellt ist.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(A),

      h1("B · Die Pinneberger Reform"),
      p("Wie stark trifft die Klassenassistenz genau dieses Unternehmen?", { color: MUT, size: 18, after: 140 }),
      frageTabelle(B),

      h1("C · Vergabefähigkeit"),
      p("Hier liegen die Ausschlusskriterien. Was fehlt, lässt sich im laufenden Verfahren nicht mehr heilen.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(C),

      h1("D · Personal und Qualifikation"),
      p("Die Bedingungen des Pool- und MPT-Modells — der wichtigste Block.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(D),

      h1("E · Löhne und Vergütung"),
      p("Ob unter einem Festpreis über bis zu neun Jahre noch Marge bleibt.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(E),

      h1("F · Wirtschaftliche Basis"),
      p("Die Zahlen, die wir ohnehin brauchen — mit Blick auf die Reform gelesen.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(F),

      h1("G · Eigentümerin und Übergang"),
      p("Was am Menschen hängt und was am Unternehmen.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(G),

      h1("H · Markt und Wettbewerb"),
      p("Was die Eigentümerin weiß und wir nicht recherchieren können.", { color: MUT, size: 18, after: 140 }),
      frageTabelle(H),

      h1("Antworten, die wir als Warnsignal werten"),
      box("Rot", [
        "Honorarkräfte in nennenswertem Umfang — Ausschlusskriterium, und die Umstellung auf Festanstellung verteuert sofort.",
        "Keine einzige Fachkraft, die die Teamleitung übernehmen könnte.",
        "Betriebshaftpflicht unter 10 Mio. € und keine Bereitschaft oder Möglichkeit zur Aufstockung.",
        "Umsatz lässt sich nicht nach Leistungsart trennen — dann fehlt ein Eignungsnachweis, der sich nicht nachbauen lässt.",
        "Bescheide sind bereits einheitlich auf ein Schuljahresende befristet, ohne dass es Kontakt zum Kreis gibt.",
        "Alle Schulbeziehungen laufen ausschließlich über die Eigentümerin, nichts ist schriftlich.",
      ], RED, "FDECEA"),

      new Paragraph({ spacing: { after: 160 }, children: [] }),

      box("Grün — Antworten, die den Preis rechtfertigen", [
        "Mehrere Fachkräfte, die eine Teamleitung übernehmen könnten, idealerweise mit Führungserfahrung.",
        "Schriftliche Kooperationsvereinbarungen mit Schulleitungen, nicht nur gute Kontakte.",
        "Erfahrung mit Gruppenbetreuung, also eine Kraft für mehrere Kinder.",
        "Teilnahme am Teilnahmewettbewerb 2024 mit bestätigter Eignung.",
        "Sitz in einer Arbeitsgemeinschaft nach § 78 SGB VIII oder einer Trägerrunde des Kreises.",
        "Ein zweites Standbein neben der Schulbegleitung.",
      ], OLIVE, "F1F5E0"),

      h1("Was wir ins Gespräch mitbringen"),
      p("Damit niemand am Tisch überrascht wird — das ist unser Kenntnisstand aus öffentlichen Quellen:",
        { after: 120 }),
      ...[
        "Der Kreis Pinneberg hat die „Klassenassistenz“ am 25.04.2024 EU-weit ausgeschrieben. Eine Zuschlagsbekanntmachung ist bis heute nicht erschienen — das Verfahren ist offen.",
        "Der Kreistag hat die Einführung am 11.12.2024 um zwei Jahre verschoben. Stufe 1 rechnerisch ab 01.08.2027.",
        "Vier Träger haben gegen das Verfahren geklagt. Das OVG Schleswig hat sie am 03.09.2024 unanfechtbar abgewiesen (5 MB 7/24). Es gibt keinen Rechtsweg und keinen Bestandsschutz.",
        "Das Kreiskonzept kalkuliert die Klassenassistenz nach TVöD SuE S2 Stufe 5 zuzüglich 10 % Sachkosten und 5 % Verwaltungskostenzuschlag — bei einem Festpreis, der Tarifsteigerungen ausdrücklich nicht ausgleicht.",
        "Im Kreis Pinneberg sind mit FiB und Familienräume zwei weitere Anbieter tätig; Familienräume führt das Poolmodell in Tornesch/Uetersen durch.",
      ].map((t) => runs([{ t: "▪  ", color: BLUE }, { t, size: 19 }], { after: 80 })),

      p("Quellenübersicht: quellen-und-offene-punkte.md im selben Ordner.",
        { color: MUT, size: 17, before: 160 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "Fragenliste-Eigentuemerin-SKP.docx");
  fs.writeFileSync(out, buf);
  console.log("geschrieben:", out);
});
