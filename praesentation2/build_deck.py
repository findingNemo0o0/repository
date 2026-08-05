#!/usr/bin/env python3
"""Baut die Vier-Folien-Präsentation zur Poolreform im Haus-Design.

Struktur nach Plan_1.docx:
  1  Alle Modelle
  2  Was sich verändert
  3  Was sich in der Firma verändern muss / Fragenliste
  4  Financial Model

Design kommt vollständig aus design-template.pptx (Master, Layouts, Theme,
Farbband, Verlaufslinie, Logo, Foliennummer, Datum).
"""

import os
import re
import shutil
import zipfile

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(HERE, "design-template.pptx")
OUT = os.path.join(HERE, "SKP-Poolreform-Modelle-und-Financial-Model.pptx")

# ---------------------------------------------------------------- Designsystem
BLUE = RGBColor(0x1A, 0x70, 0xB8)
AMBER = RGBColor(0xE0, 0x86, 0x1A)
OLIVE = RGBColor(0x6E, 0x7D, 0x00)
RED = RGBColor(0xC0, 0x39, 0x2B)
TXT = RGBColor(0x3D, 0x3D, 0x3D)
MUT = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RULE = RGBColor(0xDD, 0xDD, 0xDD)

C_BLUE = RGBColor(0xEA, 0xF1, 0xF8)
C_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
C_AMBER = RGBColor(0xFD, 0xF3, 0xE3)

HF, BF = "Quicksand", "Open Sans"

COL = [1.56, 5.30, 9.04]
CW, PAD = 3.55, 0.16
FULL_X, FULL_W = 1.56, 11.03
TWO_W = 7.29
TOP = 1.60
SRC_Y = 7.05

prs = Presentation(BASE)
LAYOUTS = {}
for m in prs.slide_masters:
    for lay in m.slide_layouts:
        LAYOUTS.setdefault(lay.name, lay)


# ------------------------------------------------------------------ Bausteine
def new_slide(name="Inhalt-1"):
    return prs.slides.add_slide(LAYOUTS[name])


def keep_placeholders(slide, keep):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep:
            ph._element.getparent().remove(ph._element)


def set_head(slide, eyebrow, title, size=19):
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text_frame.text = title
            r = ph.text_frame.paragraphs[0].runs[0]
            r.font.bold = True
            r.font.size = Pt(size)
        elif idx == 12:
            ph.text_frame.text = eyebrow


def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return tf


def para(tf, first=False, before=0, after=0, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    if line:
        p.line_spacing = line
    return p


def run(p, text, size, font=BF, color=TXT, bold=False, italic=False):
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return r


def colhead(slide, x, y, text, w=CW, color=BLUE):
    tf = tb(slide, x, y, w, 0.24)
    run(para(tf, True), text, 11, HF, color, bold=True)


def label(slide, x, y, text, w=CW, color=AMBER, size=8.4):
    tf = tb(slide, x, y, w, 0.20)
    run(para(tf, True), text, size, BF, color, bold=True)


def body(slide, x, y, w, h, text, size=8.6, color=TXT, bold=False, italic=False):
    tf = tb(slide, x, y, w, h)
    run(para(tf, True, line=1.15), text, size, BF, color, bold=bold, italic=italic)
    return tf


def bullets(slide, x, y, w, h, items, size=8.6, marker=AMBER):
    tf = tb(slide, x, y, w, h)
    for i, it in enumerate(items):
        p = para(tf, first=(i == 0), after=3, line=1.10)
        run(p, "▪ ", size, BF, marker)
        if isinstance(it, tuple):
            run(p, it[0], size, BF, TXT, bold=True)
            run(p, it[1], size, BF, TXT)
        else:
            run(p, it, size, BF, TXT)
    return tf


def card(slide, x, y, w, h, fill=C_GRAY):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.adjustments[0] = 0.06
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.text = ""
    return sh


def card_text(slide, x, y, w, h, fill, head, lines, head_color=BLUE, head_size=9.5):
    card(slide, x, y, w, h, fill)
    tf = tb(slide, x + PAD, y + 0.12, w - 2 * PAD, h - 0.24)
    first = True
    if head:
        run(para(tf, True), head, head_size, HF, head_color, bold=True)
        first = False
    for ln in lines:
        text, o = (ln, {}) if isinstance(ln, str) else ln
        p = para(tf, first=first, before=2, line=1.12)
        first = False
        run(p, text, o.get("size", 8.6), o.get("font", BF), o.get("color", TXT),
            o.get("bold", False), o.get("italic", False))
    return tf


def card_bullets(slide, x, y, w, h, fill, head, items, head_color=BLUE,
                 marker=None, size=8.4):
    card(slide, x, y, w, h, fill)
    tf = tb(slide, x + PAD, y + 0.12, w - 2 * PAD, h - 0.24)
    run(para(tf, True), head, 9.5, HF, head_color, bold=True)
    for it in items:
        p = para(tf, before=2, after=2, line=1.10)
        run(p, "▪ ", size, BF, marker or head_color)
        run(p, it, size, BF, TXT)
    return tf


def source(slide, text):
    tf = tb(slide, FULL_X, SRC_Y, 9.5, 0.30)
    run(para(tf, True, line=1.05), text, 6.8, BF, MUT)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def dot(slide, x, y, color, d=0.17):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                Inches(d), Inches(d))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.color.rgb = WHITE
    sh.line.width = Pt(1)
    sh.shadow.inherit = False
    return sh


def vrule(slide, x, y, h, color=RULE, w=0.04):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def style_chart(chart, axis=7.2, legend=False):
    chart.font.name = BF
    chart.font.size = Pt(axis)
    chart.font.color.rgb = TXT
    chart.has_title = False
    if legend:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(7)
        chart.legend.font.name = BF
        chart.legend.font.color.rgb = TXT
    else:
        chart.has_legend = False
    for ax in (chart.category_axis, chart.value_axis):
        ax.has_major_gridlines = False
        ax.tick_labels.font.size = Pt(axis)
        ax.tick_labels.font.name = BF
        ax.format.line.color.rgb = RULE
    chart.value_axis.has_major_gridlines = True
    gl = chart.value_axis.major_gridlines.format.line
    gl.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
    gl.width = Pt(0.75)
    chart.category_axis.tick_labels.font.color.rgb = TXT
    chart.value_axis.tick_labels.font.color.rgb = MUT


def fill_series(ser, rgb):
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = rgb
    ser.format.line.fill.background()


def data_labels(plot, size=7.0, fmt="0.00", color=TXT,
                pos=XL_LABEL_POSITION.OUTSIDE_END, bold=True):
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(size)
    dl.font.name = BF
    dl.font.bold = bold
    dl.font.color.rgb = color
    dl.number_format = fmt
    dl.number_format_is_linked = False
    dl.position = pos


# =============================================================================
# FOLIE 1 — Alle Modelle
# =============================================================================
s = new_slide()
keep_placeholders(s, {0, 12})
set_head(s, "Alle Modelle", "Der Zielzustand: Pool plus multiprofessionelles Team")

# --- Spalte 1: worauf es hinausläuft -----------------------------------------
colhead(s, COL[0], TOP, "Worauf es hinausläuft")
card(s, COL[0], 1.88, CW, 1.90, C_BLUE)
tf = tb(s, COL[0] + PAD, 2.00, CW - 2 * PAD, 1.66)
run(para(tf, True), "Pool + MPT", 9.5, HF, BLUE, bold=True)
run(para(tf, before=3, line=1.12),
    "Die Landesregierung nennt die Zusammenführung von Schulischer Assistenz und "
    "Schulbegleitung einen „wichtigen Baustein der multiprofessionellen "
    "Zusammenarbeit an Schulen“.", 8.4, BF, TXT)
for it in ["Pool: Budget je Schule statt Bewilligung je Kind",
           "MPT: Lehrkraft, Schulassistenz, Schulbegleitung, Schulsozialarbeit und "
           "Sonderpädagogik als ein Team"]:
    pp = para(tf, before=3, line=1.10)
    run(pp, "▪ ", 8.4, BF, AMBER)
    run(pp, it, 8.4, BF, TXT)

card_text(s, COL[0], 3.94, CW, 0.82, C_AMBER, None, [
    ("Ein Abschlusstermin ist offen: Die Landesregierung kann „einen Zeitpunkt für "
     "den Prozessabschluss nicht nennen“.",
     {"size": 8.4, "color": TXT, "bold": True}),
])

card_bullets(s, COL[0], 4.90, CW, 1.42, C_GRAY, "Wer sonst gewinnt", [
    "Im Kreis Pinneberg holten zuletzt DHB, Lebenshilfe, AWO, inab und die Diakonie "
    "die ausgeschriebenen Sozialleistungen",
    "In NRW gewannen auch gewerbliche Anbieter — GmbH & Co. KG und gGmbH",
])

# --- Spalte 2: Zeitstrahl ----------------------------------------------------
colhead(s, COL[1], TOP, "Wird bereits erprobt")
tl_x = COL[1] + 0.04
vrule(s, tl_x, 1.94, 4.24)
timeline = [
    ("2013/14", "Lübeck: Poolmodell flächendeckend, Budget aus Sozialdaten und "
                "Inklusionsquote", OLIVE),
    ("2018", "Kreis Pinneberg startet den Pool mit dem Testträger „Familienräume“ "
             "an 7 Grundschulen in Tornesch-Uetersen", RED),
    ("2020/21", "Ostholstein: 7 Modellschulen, u. a. Inselschule Fehmarn und "
                "Neustädter Bucht. Je Schule ein eigener Pool", AMBER),
    ("2022/23", "Flensburg: alle Grundschulen, keine Einzelanträge mehr", OLIVE),
    ("2024", "EU-Ausschreibung „Klassenassistenz“ — bis heute kein Zuschlag. "
             "Kreistag verschiebt am 11.12.2024 um zwei Jahre", RED),
    ("seit 2025", "Robert Bosch Stiftung und Deutsche Telekom Stiftung bewerten die "
                  "Poolmodelle und fördern die Weiterentwicklung", BLUE),
    ("2026/27", "Hamburg: Kombinationsmaßnahmen, eine Kraft für mehrere Kinder", BLUE),
]
y = 1.90
for date, text, c in timeline:
    dot(s, tl_x - 0.065, y, c)
    tfd = tb(s, COL[1] + 0.32, y - 0.03, 0.84, 0.22)
    run(para(tfd, True), date, 8.4, HF, c, bold=True)
    body(s, COL[1] + 1.20, y - 0.03, 2.35, 0.56, text, size=7.8)
    y += 0.62

# --- Spalte 3: Einzugsgebiete ------------------------------------------------
colhead(s, COL[2], TOP, "Einzugsgebiete Schulbegleitung")
body(s, COL[2], 1.86, CW, 0.24,
     "Kräftig = öffentlich belegte Schulbegleitung. Drei Anbieter teilen sich einen "
     "Kreis, zwei Verbände decken das Land ab.", size=7.4, color=MUT)

maps = [
    ("SKP", "karte_skp.png", BLUE,
     "Nur Kreis Pinneberg — s-k-p.net nennt kein weiteres Gebiet. Weitere Regionen "
     "wären im Datenraum zu belegen."),
    ("FiB — Familien im Blick", "karte_fib.png", RGBColor(0x96, 0x3A, 0x96),
     "Kreis Pinneberg und Hamburg, Sitz Pinneberg (fib-pinneberg.de)."),
    ("Familienräume", "karte_familienraeume.png", RED,
     "Ganzer Kreis Pinneberg, Pilot in Tornesch/Uetersen (familienraeume.de)."),
    ("AWO Schleswig-Holstein", "karte_awo.png", AMBER,
     "Neun Kreise und kreisfreie Städte (awo-sh.de, Standortliste)."),
    ("Lebenshilfe", "karte_lebenshilfe.png", OLIVE,
     "Vier Kreise belegt. Blass: Kreisverband vor Ort, Schulbegleitung dort nicht "
     "einzeln geprüft (lebenshilfe-sh.de)."),
]
my = 2.12
for name, fn, c, note in maps:
    s.shapes.add_picture(os.path.join(HERE, fn), Inches(COL[2]), Inches(my),
                         height=Inches(0.84))
    tfm = tb(s, 10.20, my + 0.06, 2.39, 0.20)
    run(para(tfm, True), name, 8.2, HF, c, bold=True)
    body(s, 10.20, my + 0.26, 2.39, 0.46, note, size=7.0, color=MUT)
    my += 0.94

source(s, "Quellen: Landtag SH Drs. 20/3271 (11.06.2025) und Drs. 20/2643(neu) · Kreis "
          "Pinneberg, Konzept „Klassenassistenz“ und PM zum Vertrag mit Familienräume · "
          "DISW-Gesamtevaluation Ostholstein 20.11.2023, S. 5 und 21 f. · TED 246925-2024 · Karten: "
          "s-k-p.net · fib-pinneberg.de · familienraeume.de · awo-sh.de · lebenshilfe-sh.de · "
          "Kartengrundlage GADM/deutschlandGeoJSON.")
notes(s, "Kernaussage: Es läuft nicht auf „Pool oder nicht“ hinaus, sondern auf Pool plus "
         "multiprofessionelles Team — das sagt die Landesregierung selbst, und einen "
         "Endtermin nennt sie nicht. Erprobt wird seit 2013; der Testträger direkt vor der "
         "Haustür heißt Familienräume. Wichtig in der Mitte: Die Pinneberger Ausschreibung "
         "ist seit über zwei Jahren offen, es gibt keinen Zuschlag — das Zeitfenster steht "
         "also noch offen. Rechts die eigentliche Wettbewerbslage: AWO und Lebenshilfe sind "
         "landesweit aufgestellt. Wichtig zur Karte: Bei SKP ist nur der Kreis Pinneberg "
         "eingefärbt, weil es für kein weiteres Gebiet eine öffentliche Quelle gibt — das ist "
         "eine Frage an das Management, keine Aussage über den tatsächlichen Umfang. Die drei "
         "kleinen Anbieter — SKP, FiB und Familienräume — "
         "sitzen alle im selben Kreis Pinneberg, also genau dort, wo die schärfste Reform läuft. "
         "FiB ist zusätzlich in Hamburg unterwegs. Wer hier einen Standortvertrag verliert, "
         "verliert ihn an einen Nachbarn von nebenan.")

# =============================================================================
# FOLIE 2 — Was sich verändert
# =============================================================================
s = new_slide()
keep_placeholders(s, {0, 12})
set_head(s, "Was sich verändert",
         "Drei Verschiebungen: Qualifikation, Vertrag, Machtverhältnis")

# Spalte 1 — Qualifikation
colhead(s, COL[0], TOP, "Qualifikation und Team")
card_bullets(s, COL[0], 1.92, CW, 1.62, C_GRAY, "Vorher", [
    "Eine Kraft, ein Kind, ein Bescheid",
    "Kein Fachkräftegebot — § 72 SGB VIII gilt für Schulbegleitung nicht",
    "Rund die Hälfte der Schulbegleitungen ist pädagogisch, therapeutisch oder "
    "pflegerisch qualifiziert, die andere Hälfte nicht",
], head_color=MUT, marker=MUT)
card_bullets(s, COL[0], 3.66, CW, 3.18, C_BLUE, "Nachher", [
    "Basiskraft bleibt formal ungelernt: „geeignete sozial erfahrene Kräfte“",
    "Neu und zwingend: Teamleitung ausschließlich Fachkraft, Schlüssel 1:15, "
    "39 Wochenstunden",
    "Keine Honorarkräfte, Festanstellung ist Vertragsbedingung",
    "Pflicht: Qualifizierungskonzept, Supervision, wöchentliche Teamsitzung",
])

# Spalte 2 — Verträge
colhead(s, COL[1], TOP, "Verträge")
card_bullets(s, COL[1], 1.92, CW, 1.62, C_GRAY, "Vorher", [
    "Bewilligungsbescheid je Kind, befristet auf ein Schuljahr",
    "Vergütung je Kind und Stunde",
    "Klassenfahrten und Ausflüge als Zusatzstunden abrechenbar",
    "Viele kleine Aufträge, kein Vergabeverfahren",
], head_color=MUT, marker=MUT)
card_bullets(s, COL[1], 3.66, CW, 3.18, C_BLUE, "Nachher", [
    "Ein Standortvertrag je Schule, vergeben im EU-Verfahren oder per "
    "Interessenbekundung",
    "Pinneberg: 01.11.2024 bis 31.07.2029, einseitig um vier Jahre verlängerbar — "
    "bis zu neun Jahre",
    "Festpreis: „Anpassungen durch Tarifsteigerungen erfolgen im Rahmen der "
    "Vertragslaufzeit nicht“",
    "Budget = Anzahl der Klassen, nicht der Fälle. Vertretung, Krankheit und "
    "Fahrten zahlt der Träger daraus",
])

# Spalte 3 — Machtverhältnisse
colhead(s, COL[2], TOP, "Politik und Machtverhältnisse")
card_bullets(s, COL[2], 1.92, CW, 1.50, C_GRAY, "Dafür", [
    "Landesregierung SH · Grüne: „Wir sind klar für jede Poollösung!“ · CDU-Antrag",
    "SPD in SH drängt auf Tempo, nicht auf Rücknahme",
    "Bund: Referentenentwurf 1. KJHSRG vom 23.03.2026",
    "Kommunale Kostenträger · Bosch- und Telekom-Stiftung",
], head_color=OLIVE, marker=OLIVE, size=8.0)
card_bullets(s, COL[2], 3.54, CW, 1.60, C_AMBER, "Dagegen", [
    "ver.di (24.04.2026): Fachkräftegebot werde durch einen Kompetenzansatz ersetzt",
    "SPD-Bundestagsabgeordnete: Streichliste „inakzeptabel“ (17.04.2026), über "
    "100.000 Unterschriften dagegen",
    "Lebenshilfe · AGJ · SoVD · Verfassungsblog: nicht verfassungskonform",
], head_color=AMBER, marker=AMBER, size=8.0)
card_bullets(s, COL[2], 5.22, CW, 1.62, C_BLUE, "Wer die Mehrheit hat", [
    "Land: CDU und Grüne, 48 von 69 Sitzen. Landtagswahl 2027",
    "Kreis Pinneberg: CDU stärkste Fraktion, 24 von 67. Landrätin Elfi Heesch "
    "(parteilos). Der Kreistag hat verschoben",
    "Bund: Der Entwurf kommt aus dem Haus von Karin Prien (CDU), bis 2025 "
    "Bildungsministerin in SH",
], size=8.0)

source(s, "Quellen: Kreis Pinneberg, Konzept „Klassenassistenz“, Kap. 4–6 · TED 246925-2024 · "
          "Landtag SH, Presseticker 28.02.2025 und Drs. 20/3271 (Frage 3 der SPD) · ver.di, "
          "24.04.2026 · SPD-Bundestagsabgeordnete Hostert, Klose, Heubach zur Streichliste, "
          "17.04.2026 · Sitzverteilung Landtag SH · Kreistag Pinneberg seit 14.05.2023.")
notes(s, "Drei Spalten, drei Verschiebungen. Erstens Qualifikation: Unten bleibt es ungelernt, "
         "aber es entsteht eine Fachkraftebene darüber — die ist der eigentliche Zugangsschlüssel. "
         "Zweitens Vertrag: aus hunderten Jahresbescheiden wird ein Festpreisvertrag über bis zu "
         "neun Jahre ohne Tarifausgleich. Das ist der ökonomische Kern. Drittens die Politik — und "
         "hier bitte genau unterscheiden: Die SPD taucht in beiden Spalten auf. In "
         "Schleswig-Holstein drängt sie auf schnellere Poollösungen, im Bund nennen ihre "
         "Abgeordneten die Streichliste inakzeptabel. Sie ist also gegen das Sparpaket, nicht "
         "gegen den Pool. Und unten die Mehrheiten: Land und Kreis werden beide CDU-geführt "
         "regiert, der Entwurf im Bund kommt aus dem Haus der früheren SH-Bildungsministerin. "
         "Dieselbe politische Linie auf drei Ebenen — Kritik gibt es laut, Hebel hat sie keinen.")

# =============================================================================
# FOLIE 3 — Fragenliste / Due Diligence
# =============================================================================
s = new_slide()
keep_placeholders(s, {0, 12})
set_head(s, "Fragenliste",
         "Fünf Fragen an SKP — und woran wir die Antwort messen")

questions = [
    ("1 · Fachkräftequote", C_BLUE, BLUE,
     "Wie viele Beschäftigte haben einen pädagogischen, therapeutischen oder "
     "pflegerischen Abschluss?",
     "Messlatte: bundesweit rund 50 %. Für den Pool zählt aber nur eines — gibt es "
     "Fachkräfte für die Teamleitung? Ohne sie kein Zuschlag.",
     "Anzufordern: Personalliste zum Stichtag mit Qualifikation, Stundenumfang und "
     "Vertragsart."),
    ("2 · Löhne und Tarif", C_BLUE, BLUE,
     "Zahlt SKP Tarif, angelehnt oder Mindestlohn? Wie hoch ist der Ist-Stundensatz, "
     "wie sind Ferien geregelt?",
     "Messlatte: Pinneberg kalkuliert mit TVöD SuE S2 Stufe 5 — rund 40 % über dem "
     "Mindestlohn. § 124 SGB IX schützt nur tarifgebundene Träger.",
     "Anzufordern: Musterarbeitsvertrag, Lohnjournal zwölf Monate, Angabe zur "
     "Tarifbindung und zum Honorarkräfteanteil."),
    ("3 · Pilotschulen", C_BLUE, BLUE,
     "Bestehen Kooperationen an den Schulen, die zuerst umgestellt werden?",
     "Messlatte: Zwei der sieben Zuschlagskriterien sind Kommunikation mit der "
     "Schulleitung und verpflichtende Zusammenarbeit.",
     "Anzufordern: Kooperationsvereinbarungen und Schulliste mit namentlichen "
     "Ansprechpartnern der Schulleitungen."),
    ("4 · Präqualifizierung und Konzepte", C_AMBER, AMBER,
     "Liegen PQ-Eintrag, 10 Mio. € Betriebshaftpflicht, Referenzliste über vier "
     "Jahre und getrennte Umsatzausweisung vor?",
     "Dazu die Konzepte: Qualifizierung, Kooperation, Krisen- und Schutzkonzept nach "
     "§§ 8a/8b/72a SGB VIII. Was fehlt, ist im laufenden Verfahren nicht mehr "
     "heilbar.",
     "Anzufordern: PQ-Nummer, Versicherungsschein, Referenzliste, Konzeptordner."),
    ("5 · Wie viele Schulen", C_AMBER, AMBER,
     "Standortliste: Schule, Schulart, Kreis, Kinder, Wochenstunden, Jahresumsatz.",
     "Die entscheidende Auswertung: Welcher Umsatzanteil liegt an Schulen, die in "
     "Stufe 1 oder 2 des Pinneberger Plans stehen? Das ist das Volumen, das in "
     "einem Verfahren auf dem Spiel steht.",
     "Anzufordern: Standort- und Umsatzliste je Schule über drei Geschäftsjahre."),
]
for i, (head, fill, hc, q, why, ask) in enumerate(questions):
    x = COL[i % 3]
    y = 1.75 + (i // 3) * 2.30
    card(s, x, y, CW, 2.00, fill)
    tf = tb(s, x + PAD, y + 0.12, CW - 2 * PAD, 1.76)
    run(para(tf, True), head, 9.5, HF, hc, bold=True)
    run(para(tf, before=3, line=1.12), q, 8.6, BF, TXT, bold=True)
    run(para(tf, before=4, line=1.12), why, 8.2, BF, MUT)
    run(para(tf, before=5, line=1.12), ask, 8.0, BF, hc, italic=True)

card_text(s, COL[2], 4.05, CW, 2.00, C_GRAY, "Was wir selbst nicht beantworten können", [
    ("Kreisscharfe Marktzahlen zur Schulbegleitung sind nicht öffentlich. SKPs eigene "
     "Website nennt als Einsatzgebiet ausschließlich den Kreis Pinneberg — für jede "
     "weitere Region fehlt bislang jeder Beleg.", {"size": 8.4}),
    ("Diese fünf Antworten entscheiden, ob SKP im Poolmarkt überhaupt antreten kann.",
     {"size": 8.4, "bold": True, "color": RED}),
])

source(s, "Quellen: TED 246925-2024, Eignungs- und Zuschlagskriterien · Kreis Pinneberg, "
          "Konzept „Klassenassistenz“, Kap. 4–6 · TVöD SuE, Tabelle ab 01.05.2026 · "
          "§ 124 Abs. 1 SGB IX · Dworschak u. a., Qualifikation von Schulbegleitungen · "
          "s-k-p.net.")
notes(s, "Das ist die Arbeitsliste für das Management-Gespräch. Jede Frage hat eine Messlatte — "
         "wir fragen nicht ins Blaue, wir prüfen gegen die Kriterien, die in der Ausschreibung "
         "wirklich stehen. Frage 4 ist die harte: Was dort fehlt, lässt sich nicht mehr "
         "nachholen, sobald ein Verfahren läuft. Frage 5 liefert die Zahl, die wir für die "
         "Bewertung brauchen.")

# =============================================================================
# FOLIE 4 — Financial Model
# =============================================================================
s = new_slide()
keep_placeholders(s, {0, 12})
set_head(s, "Financial Model",
         "Sechs Stellschrauben — die Menge stützt, der Preis je Fall drückt")

colhead(s, COL[0], TOP, "Neue Löhne — je Stunde brutto, in €")
cd = CategoryChartData()
cd.categories = ["Mindest-\nlohn 2026", "Mindest-\nlohn 2027",
                 "TVöD SuE\nS2/5", "TVöD SuE\nS3/5"]
cd.add_series("Bruttolohn je Stunde", (13.90, 14.60, 19.64, 22.15))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1.46), Inches(1.90),
                        Inches(3.75), Inches(2.45), cd)
ch = gf.chart
style_chart(ch)
ch.value_axis.maximum_scale = 24.0
ch.value_axis.major_unit = 8.0
fill_series(ch.plots[0].series[0], AMBER)
ch.plots[0].gap_width = 70
data_labels(ch.plots[0], size=7.5, fmt="0.00")
body(s, COL[0], 4.42, CW, 0.44,
     "Der Tarifsatz, mit dem Pinneberg kalkuliert, liegt rund 40 % über dem "
     "Mindestlohn (S2 Stufe 5 = 3.330,92 €, 39 Wochenstunden).", size=7.8, color=MUT)

colhead(s, COL[1], TOP, "Nachfrage — begleitete Kinder")
cd = CategoryChartData()
cd.categories = ["Schleswig-Holstein", "Hamburg"]
cd.add_series("2014", (2700, 1574))
cd.add_series("zuletzt", (7000, 4011))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5.20), Inches(1.90),
                        Inches(3.75), Inches(2.45), cd)
ch = gf.chart
style_chart(ch, legend=True)
ch.value_axis.maximum_scale = 8000.0
ch.value_axis.major_unit = 2000.0
for ser, col_ in zip(ch.plots[0].series, (MUT, BLUE)):
    fill_series(ser, col_)
ch.plots[0].gap_width = 80
data_labels(ch.plots[0], size=7.2, fmt="#,##0")
body(s, COL[1], 4.42, CW, 0.44,
     "SH 2014 → 2022 (126 Mio. € Kreisausgaben). Hamburg 2014/15 → 2025/26 "
     "(42,15 Mio. €), Budget 2027 nur +4,4 %.", size=7.8, color=MUT)

colhead(s, COL[2], TOP, "Vergütung und Personalschlüssel")
card_text(s, COL[2], 1.90, CW, 1.72, C_BLUE, "So rechnet der Kreis", [
    ("Festbetrag aus TVöD SuE S2 Stufe 5, plus 10 % Sachkosten, plus 5 % "
     "Verwaltungskostenzuschlag auf die Personalkosten.", {"size": 8.2}),
    ("TVöD SuE = Tarifvertrag für den öffentlichen Dienst, Sozial- und "
     "Erziehungsdienst. S2 ist die unterste Entgeltgruppe für Helfertätigkeiten, "
     "die Stufe steht für Berufserfahrung.", {"size": 7.8, "color": MUT}),
])
card_text(s, COL[2], 3.74, CW, 1.05, C_GRAY, "Personalschlüssel", [
    ("Eine Klassenassistenz je Klasse — 24 Wochenstunden in Klasse 1–2, 30 in "
     "Klasse 3–4. Teamleitung 1:15.", {"size": 8.2}),
])

# untere Reihe
row = [
    ("Vertragsdauer", C_AMBER, AMBER,
     "01.11.2024 bis 31.07.2029, einseitig um vier Jahre verlängerbar. Bis zu neun "
     "Jahre Festpreis ohne Tarifausgleich — bei einem Mindestlohn, der allein bis "
     "2027 um 5 % steigt."),
    ("Gerichtsverfahren Pinneberg", C_GRAY, BLUE,
     "Vier Träger klagten, das OVG Schleswig wies sie am 03.09.2024 unanfechtbar ab "
     "(5 MB 7/24). Kein Bestandsschutz, kein Abwehrrecht. Das Verfahren selbst ist "
     "seit 2024 ohne Zuschlag."),
    ("Unterm Strich", C_BLUE, BLUE,
     "Die Menge wächst, der Preis je Fall wird gedeckelt. Der Ertrag entscheidet sich "
     "nicht am Markt, sondern an einer einzigen Frage: Gewinnt SKP den Standortvertrag "
     "oder nicht?"),
]
for i, (head, fill, hc, text) in enumerate(row):
    card_text(s, COL[i], 5.02, CW, 1.42, fill, head, [(text, {"size": 8.4})],
              head_color=hc)

source(s, "Quellen: Mindestlohnkommission/BMAS · TVöD SuE, Tabelle 01.05.2026–31.03.2027 (VKA, "
          "39 Wochenstunden), Stundensätze eigene Umrechnung · Kreis Pinneberg, Konzept "
          "„Klassenassistenz“, Kap. 5–6 · Landtag SH, Rede der Ministerin 28.02.2025 · BSFB "
          "Hamburg, PM 26.06.2026 · TED 246925-2024 · OVG Schleswig 5 MB 7/24.")
notes(s, "Vier Zahlen zum Merken. Erstens: Der Tarifsatz, mit dem der Kreis kalkuliert, liegt "
         "rund 40 Prozent über dem Mindestlohn — wer heute Mindestlohn zahlt, hat im Pool "
         "zunächst Luft, aber weniger als man denkt. Zweitens die Nachfrage: In "
         "Schleswig-Holstein von 2.700 auf rund 7.000 begleitete Kinder, in Hamburg von 1.574 "
         "auf 4.011 — die Hamburger Zahl ist aus Juni 2026 und damit die aktuellste, die "
         "öffentlich vorliegt. Bundesweit erhielten 2024 rund 324.570 Kinder unter 18 "
         "Eingliederungshilfe, die Ausgaben stiegen um 12,9 Prozent. Drittens: Das Budget hängt "
         "an Klassen, nicht an Fällen — die Menge ist damit planbar. Viertens, und das ist der "
         "Haken: Der Preis steht bis zu neun Jahre fest, Tarifsteigerungen trägt der Träger "
         "allein. Deshalb ist die Kalkulation beim Angebot die wichtigste einzelne Entscheidung.")

prs.save(OUT)

# ---- leere Review-Metadaten der Vorlage entfernen ---------------------------
tmp = OUT + ".tmp"
zin = zipfile.ZipFile(OUT)
zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    if it.filename == "ppt/authors.xml":
        continue
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml":
        d = re.sub(rb'<Override PartName="/ppt/authors\.xml"[^>]*/>', b"", d)
    if it.filename == "ppt/_rels/presentation.xml.rels":
        d = re.sub(rb'<Relationship[^>]*Target="authors\.xml"[^>]*/>', b"", d)
    zout.writestr(it, d)
zout.close()
zin.close()
shutil.move(tmp, OUT)
print("geschrieben:", OUT)
