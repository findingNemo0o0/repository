#!/usr/bin/env python3
"""Baut die SKP-Marktanalyse im Standard-Design des Hauses.

Grundlage ist die Design-Vorlage (skp-design-template.pptx). Aus ihr stammen
ausschliesslich Master, Layouts, Theme, Schriften, Farben und die
Seitenmoebel (Farbband, Verlaufslinie, Logo, Foliennummer, Datum).
Es wird kein einziger Inhalt aus der Vorlage uebernommen.
"""

import copy
import os

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt, Emu

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(HERE, "skp-design-template.pptx")
OUT = os.path.join(HERE, "SKP-Marktanalyse-Schulbegleitung.pptx")

# ---------------------------------------------------------------- Designsystem
BLUE = RGBColor(0x1A, 0x70, 0xB8)   # opseo-Blau, Hauptakzent
AMBER = RGBColor(0xE0, 0x86, 0x1A)  # Hervorhebung
OLIVE = RGBColor(0x6E, 0x7D, 0x00)  # positive Kennzahl
RED = RGBColor(0xC0, 0x39, 0x2B)    # Warnung
TXT = RGBColor(0x3D, 0x3D, 0x3D)    # Fliesstext
MUT = RGBColor(0x88, 0x88, 0x88)    # Sekundaertext
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RULE = RGBColor(0xDD, 0xDD, 0xDD)

C_BLUE = RGBColor(0xEA, 0xF1, 0xF8)   # Kartenfuellung blau
C_GRAY = RGBColor(0xF2, 0xF2, 0xF2)   # Kartenfuellung neutral
C_AMBER = RGBColor(0xFD, 0xF3, 0xE3)  # Kartenfuellung Hinweis

HF = "Quicksand"   # Ueberschriften
BF = "Open Sans"   # Fliesstext

# Raster der Vorlage
COL = [1.56, 5.30, 9.04]
CW = 3.55
PAD = 0.16
INNER = CW - 2 * PAD           # 3.23
FULL_X, FULL_W = 1.56, 11.03
TWO_W = 7.29                   # Spalte 1+2
HALF_X, HALF_W = [1.56, 7.17], 5.42
TOP = 1.60                     # erste Inhaltszeile
SRC_Y = 7.05

LAY_TITLE = "Titel1-mit-Logo"
LAY_BODY = "Inhalt-1"

prs = Presentation(BASE)

# vorhandene Folien der Vorlage restlos entfernen -----------------------------
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)

LAYOUTS = {}
for master in prs.slide_masters:
    for lay in master.slide_layouts:
        LAYOUTS.setdefault(lay.name, lay)


# ------------------------------------------------------------------- Bausteine
def new_slide(layout_name):
    return prs.slides.add_slide(LAYOUTS[layout_name])


def drop(slide, *names):
    """Nicht benoetigte Layout-Platzhalter von der Folie nehmen."""
    for sh in list(slide.shapes):
        if sh.name in names or (sh.is_placeholder and sh.placeholder_format.idx in names):
            sh._element.getparent().remove(sh._element)


def strip_unused_placeholders(slide, keep_idx):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep_idx:
            ph._element.getparent().remove(ph._element)


def set_title(slide, eyebrow, title):
    """Kopfzeile (Farbband) und Folientitel der Vorlage fuellen."""
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            tf = ph.text_frame
            tf.text = title
            p = tf.paragraphs[0]
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(20)
        elif idx == 12:
            ph.text_frame.text = eyebrow


def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return tf


def para(tf, first=False, space_before=0, space_after=0, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
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


def body(slide, x, y, w, h, text, size=8.6, color=TXT, bold=False, italic=False):
    tf = tb(slide, x, y, w, h)
    run(para(tf, True, line=1.15), text, size, BF, color, bold=bold, italic=italic)
    return tf


def bullets(slide, x, y, w, h, items, size=8.8):
    tf = tb(slide, x, y, w, h)
    for i, it in enumerate(items):
        p = para(tf, first=(i == 0), space_after=4, line=1.1)
        run(p, "▪ ", size, BF, AMBER)
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
    """Karte mit Ueberschrift (Quicksand) und Textzeilen (Open Sans)."""
    card(slide, x, y, w, h, fill)
    tf = tb(slide, x + PAD, y + 0.12, w - 2 * PAD, h - 0.24)
    if head:
        run(para(tf, True), head, head_size, HF, head_color, bold=True)
        first = False
    else:
        first = True
    for ln in lines:
        text, opts = (ln, {}) if isinstance(ln, str) else ln
        p = para(tf, first=first, space_before=2, line=1.12)
        first = False
        run(p, text, opts.get("size", 8.6), opts.get("font", BF),
            opts.get("color", TXT), opts.get("bold", False), opts.get("italic", False))
    return tf


def source(slide, text):
    tf = tb(slide, FULL_X, SRC_Y, 9.5, 0.30)
    run(para(tf, True, line=1.05), text, 6.8, BF, MUT)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def bar(slide, x, y, w, value, maxv, color, width=0.12):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                Inches(max(w * value / maxv, 0.03)), Inches(width))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


# ------------------------------------------------------------------ Diagramme
def style_chart(chart, axis_size=7.5, legend=False, legend_size=7.5):
    chart.font.name = BF
    chart.font.size = Pt(axis_size)
    chart.font.color.rgb = TXT
    chart.has_title = False
    if legend:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(legend_size)
        chart.legend.font.name = BF
        chart.legend.font.color.rgb = TXT
    else:
        chart.has_legend = False
    for axis in (chart.category_axis, chart.value_axis):
        axis.has_major_gridlines = False
        axis.tick_labels.font.size = Pt(axis_size)
        axis.tick_labels.font.name = BF
        axis.format.line.color.rgb = RULE
    chart.value_axis.has_major_gridlines = True
    gl = chart.value_axis.major_gridlines.format.line
    gl.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
    gl.width = Pt(0.75)
    chart.category_axis.tick_labels.font.color.rgb = TXT
    chart.value_axis.tick_labels.font.color.rgb = MUT


def series_color(ser, rgb):
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = rgb
    ser.format.line.fill.background()


# =============================================================================
# 1 — Titel
# =============================================================================
s = new_slide(LAY_TITLE)
strip_unused_placeholders(s, {0})
set_title(s, None, "Schulbegleitung im Umbruch")
for ph in s.placeholders:
    if ph.placeholder_format.idx == 0:
        ph.text_frame.paragraphs[0].runs[0].font.size = Pt(26)

body(s, 1.56, 1.86, 7.40, 0.30,
     "Marktanalyse SKP · Kreis Steinburg · Hamburg · Segeberg · Ostholstein · Neumünster",
     size=10.5, color=MUT)

card_text(s, 1.56, 2.45, TWO_W, 1.05, C_BLUE, "Die Kernthese",
          [("Der Markt wächst — aber der Zugang zu ihm wird gerade neu vergeben.",
            {"size": 10.5, "bold": True})])

card_text(s, 1.56, 3.75, TWO_W, 1.25, C_GRAY, "Worum es auf sechs Folien geht",
          [("Wer zahlt, wer entscheidet, wie groß der Markt ist, was ein Kind kostet, "
            "welche Poolmodelle zur Debatte stehen — und was das für den Wert des "
            "Unternehmens bedeutet.", {})])

body(s, 1.56, 5.25, TWO_W, 0.30,
     "Stand: Juli 2026 · Folien 7 und 8 sind Anhang zum Nachschlagen",
     size=8.6, color=MUT)

notes(s, "Einstieg in 30 Sekunden: SKP verkauft Schulbegleitung an Kreise. Der Bedarf steigt "
         "seit zehn Jahren ununterbrochen. Gleichzeitig stellen die Kostenträger gerade das "
         "Vergabesystem um — von Einzelfallbewilligung auf Standortvertrag. Wer die Umstellung "
         "übersteht, hat ein größeres und planbareres Geschäft. Wer sie verpasst, verliert "
         "Standorte vollständig.")

# =============================================================================
# 2 — Das Geschaeft in 60 Sekunden
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Geschäftsmodell", "Drei Akteure — und nur einer ist Vertragspartner")

roles = [
    ("Kreis / kreisfreie Stadt", "DER VERTRAGSPARTNER", BLUE,
     "Rechtsträger, Zahler und Unterzeichner. In Schleswig-Holstein erstattet das Land rund "
     "84 % der SGB-IX-Kosten. In Hamburg tritt an seine Stelle die Schulbehörde BSFB."),
    ("Jugendamt / Amt für Teilhabe", "DIE FACHBEHÖRDE", BLUE,
     "Kein eigenes Rechtssubjekt, sondern ein Amt innerhalb des Kreises: Bedarfsfeststellung, "
     "Bewilligung, Trägerauswahl, Fachaufsicht."),
    ("Schule", "NIE VERTRAGSPARTNER", AMBER,
     "Aber Einsatzort mit faktischem Vetorecht: Ein Pool kommt nur „im Einvernehmen mit der "
     "Schulleitung“ zustande."),
]
for i, (name, label, color, text) in enumerate(roles):
    x = COL[i]
    colhead(s, x, TOP, name)
    body(s, x, TOP + 0.26, CW, 0.20, label, size=8.4, color=color, bold=True)
    body(s, x, TOP + 0.50, CW, 0.80, text)

card_text(s, FULL_X, 3.20, FULL_W, 1.50, C_BLUE, "Geltungsbereich: alle Schularten",
          ["Schulbegleitung gilt für Klasse 1 bis 13 — Regelschule wie Förderschule, "
           "berufsbildende Schulen und den offenen Ganztag (§ 112 SGB IX). Die Poolreform "
           "beginnt meist an Grundschulen, weil dort die Schulische Assistenz des Landes "
           "ansetzt; sie ist aber nicht darauf beschränkt."])

card_text(s, FULL_X, 4.95, FULL_W, 1.50, C_GRAY, "So entsteht heute Umsatz",
          ["Vergütet wird pro Kind und Stunde auf Basis eines Bewilligungsbescheids. "
           "Klassenfahrten und Ausflüge sind als Zusatzstunden abrechenbar — in Pinneberg "
           "8 Stunden je Tag, ohne vorherigen Antrag; das monatliche Stundensoll erhöht sich "
           "entsprechend. Im Pool entfällt diese Zusatzvergütung."])

source(s, "Quellen: SGB VIII, SGB IX · Landtag SH Drs. 20/2643 · Trägerschreiben Kreis "
          "Pinneberg 20.03.2024 · Empfehlungen Bayerischer Bezirketag/StMUK 2025.")
notes(s, "Drei Akteure, ein Vertragspartner. Wichtigster Punkt ist die dritte Spalte: Die Schule "
         "unterschreibt nie, entscheidet aber mit. Deshalb ist die Beziehung zu den Schulleitungen "
         "der eigentliche Vermögenswert. Zweiter Punkt: Schulbegleitung ist nicht auf die "
         "Grundschule beschränkt — nur die Poolreform startet dort.")

# =============================================================================
# 3 — Nachfrage und Standort
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Nachfrage und Standort",
          "Der Bedarf wächst — und SKPs Regionen sind besonders inklusiv")

colhead(s, COL[0], TOP, "Förderquote Deutschland, in %")
cd = CategoryChartData()
cd.categories = ["2008/09", "2021/22", "2022/23"]
cd.add_series("Inklusion (Regelschule)", (1.1, 3.5, 3.4))
cd.add_series("Exklusion (Förderschule)", (4.8, 4.3, 4.2))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(1.46), Inches(1.90),
                        Inches(3.75), Inches(2.55), cd)
ch = gf.chart
style_chart(ch, legend=True, legend_size=7.0)
ch.value_axis.maximum_scale = 9.0
ch.value_axis.major_unit = 3.0
for ser, col_ in zip(ch.plots[0].series, (BLUE, AMBER)):
    series_color(ser, col_)
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.font.size = Pt(7.5)
pl.data_labels.font.name = BF
pl.data_labels.font.bold = True
pl.data_labels.font.color.rgb = WHITE
pl.data_labels.number_format = "0.0"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.CENTER

colhead(s, COL[1], TOP, "Im Ländervergleich, 2022/23")
cd = CategoryChartData()
cd.categories = ["Deutschland", "Schleswig-Holstein", "Hamburg"]
cd.add_series("Inklusion (Regelschule)", (3.4, 4.4, 5.2))
cd.add_series("Exklusion (Förderschule)", (4.2, 2.3, 2.7))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(5.20), Inches(1.90),
                        Inches(3.75), Inches(2.55), cd)
ch = gf.chart
style_chart(ch, legend=True, legend_size=7.0)
ch.value_axis.maximum_scale = 9.0
ch.value_axis.major_unit = 3.0
for ser, col_ in zip(ch.plots[0].series, (BLUE, AMBER)):
    series_color(ser, col_)
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.font.size = Pt(7.5)
pl.data_labels.font.name = BF
pl.data_labels.font.bold = True
pl.data_labels.font.color.rgb = WHITE
pl.data_labels.number_format = "0.0"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.CENTER

card_text(s, FULL_X, 4.70, TWO_W, 1.40, C_GRAY, "Die Kehrseite",
          ["Beide Regionen haben die niedrigsten Förderschulanteile im Bund — die Reserve an "
           "Wechslern von der Förder- an die Regelschule ist weitgehend ausgeschöpft. Künftiges "
           "Wachstum muss aus neuen Diagnosen kommen, nicht aus Umschichtung."])

colhead(s, COL[2], TOP, "Was das für SKP bedeutet")
facts = [
    ("5,9 → 7,6 %", C_BLUE, BLUE,
     "Förderquote Deutschland seit 2008/09. Drei Viertel des Zuwachses sind neue Diagnosen — "
     "nicht Wechsler von der Förderschule."),
    ("4,4 / 5,2 %", C_BLUE, BLUE,
     "Inklusionsquote Schleswig-Holstein und Hamburg gegenüber 3,4 % im Bund. Schulbegleitung "
     "entsteht fast nur an Regelschulen."),
    ("≈ 2 Jahre", C_AMBER, AMBER,
     "Durchschnittliche Hilfedauer je Kind (§ 35a SGB VIII) — leicht steigend: 22 Monate "
     "(2013) auf 24,6 Monate (2023)."),
]
y = 1.90
for head, fill, hc, text in facts:
    card_text(s, COL[2], y, CW, 1.35, fill, head, [(text, {})], head_color=hc, head_size=12)
    y += 1.45

source(s, "Quellen: Bertelsmann Stiftung auf Grundlage KMK (2024) / Klemm, Schuljahre 2008/09–2022/23 "
          "(ohne Saarland) · Hilfedauer: AKJStat/HzE-Monitor TU Dortmund, beendete Hilfen nach "
          "§ 35a SGB VIII — alle Hilfearten, nicht nur Schulbegleitung.")
notes(s, "Zwei Botschaften auf einer Folie. Erstens: Der Nachfragesockel wächst strukturell — und "
         "der Zuwachs kommt zu drei Vierteln aus neuen Diagnosen, nicht aus dem Abbau der "
         "Förderschulen. Zweitens: SH und Hamburg liegen bei der Inklusionsquote deutlich über dem "
         "Bund, der adressierbare Markt ist hier also größer. Die Hilfedauer von rund zwei Jahren "
         "ist die Rechengröße für die nächste Folie.")

# =============================================================================
# 4 — Marktvolumen und Kosten
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Marktvolumen und Kosten",
          "Ausgaben wachsen schneller als Fälle — Löhne schneller als beide")

colhead(s, COL[0], TOP, "Schulbegleitung SH, Index 2014 = 100")
cd = CategoryChartData()
cd.categories = ["2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021"]
cd.add_series("Ausgaben", (100, 122, 148, 170, 199, 234, 286, 384))
cd.add_series("Leistungsempfänger", (100, 111, 131, 144, 162, 189, 204, 245))
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(1.46), Inches(1.90),
                        Inches(3.75), Inches(2.95), cd)
ch = gf.chart
style_chart(ch, legend=True, legend_size=7.0)
ch.value_axis.minimum_scale = 80.0
ch.value_axis.maximum_scale = 400.0
ch.value_axis.major_unit = 80.0
for ser, col_ in zip(ch.plots[0].series, (AMBER, BLUE)):
    ser.format.line.color.rgb = col_
    ser.format.line.width = Pt(2.25)
    ser.marker.format.fill.solid()
    ser.marker.format.fill.fore_color.rgb = col_
    ser.marker.format.line.color.rgb = col_
    ser.smooth = False

colhead(s, COL[1], TOP, "Gesetzlicher Mindestlohn, € je Stunde")
cd = CategoryChartData()
cd.categories = ["2015", "2017", "2019", "2021", "2023", "2024", "2025", "2026", "2027"]
cd.add_series("Mindestlohn", (8.50, 8.84, 9.19, 9.60, 12.00, 12.41, 12.82, 13.90, 14.60))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5.20), Inches(1.90),
                        Inches(3.75), Inches(2.95), cd)
ch = gf.chart
style_chart(ch, axis_size=7.0)
ch.value_axis.maximum_scale = 18.0
ch.value_axis.major_unit = 6.0
series_color(ch.plots[0].series[0], AMBER)
ch.plots[0].gap_width = 60
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.font.size = Pt(6.5)
pl.data_labels.font.name = BF
pl.data_labels.font.color.rgb = TXT
pl.data_labels.number_format = "0.00"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

card_text(s, FULL_X, 5.05, TWO_W, 1.05, C_GRAY,
          "Nicht die Menge treibt die Kosten, sondern der Preis je Fall",
          ["Die Ausgaben wachsen fast doppelt so schnell wie die Fallzahlen. Der Mindestlohn ist "
           "seit 2015 um 72 % gestiegen, allein 2026 um 8,4 %."])

colhead(s, COL[2], TOP, "Was ein Kind kostet")
card_text(s, COL[2], 1.90, CW, 1.75, C_BLUE, "Was der Kreis je Kind zahlt", [
    ("gerechnet auf die durchschnittliche Hilfedauer von 2,05 Jahren",
     {"size": 7.8, "color": MUT, "italic": True}),
    ("Hamburg, alle Schularten (2025/26)", {"size": 8.2, "color": MUT}),
    ("10.500 € je Jahr  →  21.500 € je Kind", {"size": 8.8, "bold": True}),
    ("Kreis Pinneberg, Grundschule (Ø 2017–2020)", {"size": 8.2, "color": MUT}),
    ("18.700 € je Jahr  →  38.300 € je Kind", {"size": 8.8, "bold": True}),
])
card_text(s, COL[2], 3.80, CW, 2.30, C_AMBER, "Was ein Kind an Personalkosten kostet", [
    ("Mindestlohn 13,90 € + 22 % Arbeitgeberanteil = 17,00 € je Stunde",
     {"size": 7.8, "color": MUT, "italic": True}),
    ("20 Wochenstunden × 39 Schulwochen = 780 h", {"size": 8.2, "color": MUT}),
    ("13.200 € je Jahr  →  27.100 € je Kind", {"size": 8.8, "bold": True, "color": AMBER}),
    ("Gegenprobe: Hamburgs Durchschnitt trägt rund 16 Wochenstunden, Pinnebergs Satz rund 28 — "
     "genau der Korridor der Klassenassistenz (24–30 h).", {"size": 8.2, "color": MUT}),
], head_color=AMBER)

source(s, "Quellen: Landtag SH Drs. 20/2643 · BSFB Hamburg · Konzept Klassenassistenz Kreis "
          "Pinneberg · Mindestlohnkommission/BMAS. Die Kosten je Kind sind eigene Berechnungen auf "
          "Basis der ausgewiesenen Annahmen. Der Mindestlohn ist die Untergrenze — das "
          "durchschnittliche Bruttogehalt in der Schulbegleitung liegt bei 2.700–3.100 € im Monat.")
notes(s, "Drei Aussagen. Erstens: Die Ausgaben wachsen fast doppelt so schnell wie die Fallzahlen — "
         "nicht die Menge treibt die Kosten, sondern der Preis je Fall. Zweitens: Der Mindestlohn ist "
         "seit 2015 um 72 Prozent gestiegen, allein 2026 um 8,4 Prozent. Drittens die Gegenüberstellung "
         "rechts. Die Gegenprobe ist wichtig: Pinnebergs historischer Satz entspricht 28 Wochenstunden "
         "und landet damit exakt im Korridor der geplanten Klassenassistenz — das bestätigt die Rechnung.")

# =============================================================================
# 5 — Die Modelle
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Die Modelle", "Vier Wege, die derzeit zur Debatte stehen")

models = [
    ("Klassenmodell", "Kreis Pinneberg · UPSIDE", BLUE,
     "Budget = Anzahl der Klassen, abzüglich vorhandener Landeskräfte. Eine Assistenz je Klasse, "
     "24 h (Kl. 1–2) bzw. 30 h (Kl. 3–4), Vergütung nach S2 Stufe 5.",
     "Stellenzahl steigt deutlich (223 → 431), aber Satz und Stunden sind gedeckelt. Nur "
     "Grundschulen. Bereits zweimal verschoben."),
    ("Poolmodell", "Kreis Ostholstein · BASISANNAHME", AMBER,
     "Budget = Summe der bisherigen Einzelbedarfe, schulscharf. Ein koordinierender Träger je "
     "Schule; Mitarbeitende wechseln zu ihm.",
     "Gleiches Volumen, mehr Kinder, weniger Stunden je Kind. „Konzentration der "
     "Leistungserbringer“ ist erklärtes Projektziel."),
    ("Kombinationsmaßnahmen", "Hamburg · LÄUFT SEIT 2026/27", AMBER,
     "Eine Kraft für mehrere Kinder einer Klasse oder Schule. Freiwilligendienstleistende werden "
     "vollständig darüber organisiert.",
     "Stunden und Qualifikation je Kind sinken. Höher qualifizierte Kräfte nur noch in "
     "begründeten Ausnahmefällen."),
]
for i, (name, sub, color, what, means) in enumerate(models):
    x = COL[i]
    colhead(s, x, TOP, name)
    body(s, x, TOP + 0.26, CW, 0.20, sub, size=8.4, color=color, bold=True)
    body(s, x, TOP + 0.50, CW, 1.20, what)
    card_text(s, x, 3.45, CW, 1.35, C_GRAY, None, [(means, {"color": MUT})])

card_text(s, FULL_X, 5.05, FULL_W, 1.10, C_AMBER,
          "Bildungsassistenz · Bundesentwurf 1. KJHSRG · OFFEN — AB 2028",
          ["§ 80a SGB VIII-E: reine Planungspflicht ohne Personalschlüssel. Individualanspruch nur "
           "noch, wenn ausschließlich eine 1:1-Begleitung hilft. Kein gesetzlicher Mindestumfang — "
           "ein Kreis kann ein deutlich dünneres Angebot aufstellen als Pinneberg."],
          head_color=AMBER)

source(s, "Quellen: Konzept „Klassenassistenz“ Kreis Pinneberg · DISW-Gesamtevaluation Ostholstein "
          "2023 · BSFB Hamburg · Referentenentwurf 1. KJHSRG (§§ 35d Abs. 4, 80a SGB VIII-E), "
          "Kabinettbefassung Sommer 2026.")
notes(s, "Vier Modelle, ein Muster: Alle ersetzen die Einzelfallbewilligung durch ein Budget. Der "
         "Unterschied liegt in der Bemessung — und die entscheidet über SKPs Volumen. Pinneberg "
         "rechnet nach Klassen und ist das großzügigste Modell; es ist Upside, nicht Basis. "
         "Ostholstein ist die realistische Annahme. Hamburg läuft bereits. Und der Bundesentwurf "
         "schreibt gar keinen Schlüssel vor — das ist die eigentliche Unsicherheit.")

# =============================================================================
# 6 — Chancen und Risiken
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Bewertung", "Chancen und Risiken auf einen Blick")

chancen = [
    ("Der Gesamtmarkt wächst weiter",
     "Fallzahlen und Ausgaben steigen in allen belastbaren Datenreihen. Kein Kreis erwartet einen "
     "Rückgang."),
    ("Größere, planbarere Aufträge",
     "Statt hunderter Einzelbewilligungen wenige Standortverträge über Jahre — mit geringerem "
     "Verwaltungsaufwand je Umsatzeuro."),
    ("Regionale Verankerung zählt formal",
     "Kenntnis der regionalen Strukturen und bestehende Kooperationen sind nachzuweisende "
     "Zuschlagskriterien."),
    ("Qualität schlägt Preis",
     "Kreis Düren gewichtet Qualität mit 60 von 100 Punkten. Ein dokumentiertes "
     "Qualifizierungskonzept ist der billigste Differenzierer."),
    ("Zeitfenster und stabile Segmente",
     "Steinburg und Neumünster haben noch kein Modell. Weiterführende Schulen und Förderzentren "
     "bleiben am längsten stabil."),
]
risiken = [
    ("Standortverlust ist total",
     "Wer den Zuschlag nicht bekommt, verliert das gesamte Volumen dort — zum Schuljahresbeginn, "
     "ohne Übergang."),
    ("Personal wandert mit",
     "In Ostholstein dokumentiert: Mitarbeitende werden an den koordinierenden Träger "
     "„übergeben“. Rechtlich freiwillig, faktisch die Regel."),
    ("Kein Rechtsschutz",
     "Das OVG Schleswig wies die Klage von vier Trägern gegen das Pinneberger Vergabeverfahren "
     "zurück — unanfechtbar."),
    ("Preisbindung ohne Tarifausgleich",
     "Pinneberg: Festpreis bis 31.07.2029, einseitig um vier Jahre verlängerbar. "
     "Tarifsteigerungen trägt allein der Träger."),
    ("Vergabefähigkeit fehlt oft",
     "10 Mio € Betriebshaftpflicht, Referenzliste über vier Jahre, getrennte Umsatzausweisung, "
     "Präqualifizierung — Details auf Folie 7."),
]

for ci, (head, items, accent) in enumerate((("Chancen", chancen, BLUE), ("Risiken", risiken, RED))):
    x = HALF_X[ci]
    colhead(s, x, TOP, head, w=HALF_W, color=accent)
    y = 1.94
    for title, text in items:
        tf = tb(s, x, y, HALF_W, 0.22)
        p = para(tf, True)
        run(p, "▪ ", 9, BF, accent)
        run(p, title, 9, HF, TXT, bold=True)
        body(s, x + 0.18, y + 0.22, HALF_W - 0.18, 0.52, text, size=8.4)
        y += 0.94

source(s, "Quellen: Landtag SH Drs. 20/2643 · TED-Bekanntmachungen Pinneberg, Düren, Euskirchen, "
          "Rhein-Erft, Essen · OVG Schleswig 5 MB 7/24 · DISW-Evaluation Ostholstein.")
notes(s, "Beide Spalten ehrlich vortragen. Alle Chancen hängen an einer Bedingung: Das Unternehmen "
         "muss vergabefähig werden. Bei den Risiken sind Punkt 3 und 4 die entscheidenden — es gibt "
         "keinen Rechtsweg gegen die Umstellung, und der Festpreis kann bis zu neun Jahre laufen. Wer "
         "diese beiden Punkte selbst vorträgt, gewinnt Vertrauen.")

# =============================================================================
# 7 — Anhang 1: Vergabefaehigkeit
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Anhang 1", "Vergabefähigkeit im Klartext — die Ausschlusskriterien")

items = [
    ("Betriebshaftpflicht 10 Mio €", C_AMBER, AMBER,
     "Mindestdeckungssumme im Pinneberger Verfahren. Bei Bietergemeinschaften muss jedes Mitglied "
     "sie nachweisen. Nicht kurzfristig heilbar."),
    ("Getrennte Umsatzausweisung", C_AMBER, AMBER,
     "Gefordert ist der Gesamtumsatz UND separat der Umsatz der Leistungsart Schulbegleitung, je "
     "drei Geschäftsjahre. Die Buchhaltung muss das trennen können — nachträglich kaum "
     "rekonstruierbar."),
    ("Präqualifizierung", C_BLUE, BLUE,
     "Eintrag im amtlichen Verzeichnis präqualifizierter Unternehmen (DIHK, pq-vol.de). Nachweise "
     "werden einmal zentral hinterlegt; im Verfahren genügt dann die PQ-Nummer. Ohne PQ: "
     "Formblatt je Verfahren, Nachreichfrist sechs Kalendertage."),
    ("Referenzliste über vier Jahre", C_BLUE, BLUE,
     "Je Schulbegleitungsverhältnis mit Leistungsumfang, öffentlichem Auftraggeber und "
     "Ansprechpartner. Im laufenden Verfahren nicht mehr aufzubauen."),
    ("Keine Honorarkräfte", C_GRAY, BLUE,
     "„Es werden keine Honorarkräfte eingesetzt“ — Festanstellung ist Qualitätsmerkmal und "
     "Vertragsbedingung."),
    ("Teamleitung als Fachkraft", C_GRAY, BLUE,
     "Ausschließlich Sozialpädagog:innen, Pädagog:innen, Erzieher:innen oder gleichwertig. "
     "Schlüssel 1:15. Ohne diese Funktion kein Zuschlag."),
]
for i, (head, fill, hc, text) in enumerate(items):
    x = COL[i % 3]
    y = 1.75 + (i // 3) * 2.25
    card_text(s, x, y, CW, 2.05, fill, head, [(text, {})], head_color=hc)

source(s, "Quelle: EU-Bekanntmachung TED 246925-2024 (Eignungskriterien) · Konzept "
          "„Klassenassistenz“ Kreis Pinneberg, Kap. 5a.")
notes(s, "Anhangfolie. Die ersten beiden Punkte (amber) sind die, die man nicht mehr heilen kann, "
         "sobald ein Verfahren läuft.")

# =============================================================================
# 8 — Anhang 2: Glossar
# =============================================================================
s = new_slide(LAY_BODY)
strip_unused_placeholders(s, {0, 12})
set_title(s, "Anhang 2", "Glossar")

terms = [
    ("Eingliederungshilfe",
     "Sozialleistung für Menschen mit Behinderung; Schulbegleitung ist eine Form davon."),
    ("Schulbegleitung",
     "Individuelle Unterstützung eines Kindes im Unterricht. Keine geschützte "
     "Berufsbezeichnung, kein Fachkräftegebot."),
    ("Schulische Assistenz",
     "Systemische Unterstützung aller Kinder, nur an Grundschulen, vom Land finanziert. Nicht "
     "dasselbe wie Schulbegleitung."),
    ("Poolmodell",
     "Mehrere Kinder werden gemeinsam begleitet; die Schule erhält ein Budget statt "
     "Einzelbewilligungen."),
    ("Infrastrukturangebot / Bildungsassistenz",
     "Pool außerhalb des individuellen Sozialrechts — kein Antrag, kein Bescheid. Im "
     "Bundesentwurf § 80a SGB VIII-E."),
    ("Leistungs- und Vergütungsvereinbarung",
     "Vertragsform des Sozialrechts (§§ 123 ff. SGB IX): Leistung, Qualität und Preis werden "
     "verhandelt, nicht ausgeschrieben. Bei Nichteinigung entscheidet die Schiedsstelle."),
    ("Teilnahmewettbewerb",
     "Erste Stufe einer EU-Ausschreibung: Eignungsprüfung. Erst danach werden Angebote "
     "eingeholt."),
    ("Los",
     "Teilpaket eines Auftrags, etwa eine Stadt. Kein Losentscheid — jedes Los wird nach "
     "Qualität und Preis vergeben."),
    ("Förder-, Inklusions-, Exklusionsquote",
     "Anteil aller Schüler:innen mit Förderbedarf / davon an Regelschulen / davon an "
     "Förderschulen."),
    ("Mengenrisiko",
     "Wer zahlt, wenn mehr Bedarf entsteht als kalkuliert. Im Einzelfall der Kreis, im Pool der "
     "Träger."),
]

# Umbruchschaetzung: Zeichen je Zeile bei 8.4 pt Open Sans auf 3.55"
CPL_TITLE, CPL_BODY = 32, 54


def est_height(title, text):
    tl = -(-len(title) // CPL_TITLE)
    bl = -(-len(text) // CPL_BODY)
    return tl * 0.18 + bl * 0.155 + 0.20


s_ = s
heights = [est_height(t, d) for t, d in terms]
total = sum(heights)
target = total / 3.0

# gleichmaessig auf drei Spalten verteilen
cols, cur, acc = [[], [], []], 0, 0.0
for (t, d), h in zip(terms, heights):
    if cur < 2 and acc >= target * (cur + 1) - h / 2:
        cur += 1
    cols[cur].append((t, d, h))
    acc += h

BOTTOM = 6.65
for ci, group in enumerate(cols):
    x = COL[ci]
    used = sum(h for _, _, h in group)
    lead = max((BOTTOM - TOP - used) / max(len(group), 1), 0.0)
    y = TOP
    for t, d, h in group:
        tf = tb(s_, x, y, CW, 0.20)
        run(para(tf, True), t, 9, HF, BLUE, bold=True)
        head_h = est_height(t, "") - 0.18
        body(s_, x, y + head_h, CW, h - head_h - 0.04, d, size=8.4, color=TXT)
        y += h + lead

source(s, "Definitionen nach SGB VIII / SGB IX · Landtag SH Drs. 20/2643 · BAGüS-Orientierungshilfe "
          "2019 · § 97 Abs. 4 GWB.")
notes(s, "Anhangfolie für Begriffsfragen.")

prs.save(OUT)
print("geschrieben:", OUT, "· Folien:", len(prs.slides.__iter__.__self__._sldIdLst))

# ---- Aufraeumen: leere Review-Metadaten aus der Vorlage entfernen -----------
import re as _re
import shutil as _shutil
import zipfile as _zipfile

_tmp = OUT + ".tmp"
_zin = _zipfile.ZipFile(OUT)
_zout = _zipfile.ZipFile(_tmp, "w", _zipfile.ZIP_DEFLATED)
for _it in _zin.infolist():
    if _it.filename == "ppt/authors.xml":
        continue
    _d = _zin.read(_it.filename)
    if _it.filename == "[Content_Types].xml":
        _d = _re.sub(rb'<Override PartName="/ppt/authors\.xml"[^>]*/>', b"", _d)
    if _it.filename == "ppt/_rels/presentation.xml.rels":
        _d = _re.sub(rb'<Relationship[^>]*Target="authors\.xml"[^>]*/>', b"", _d)
    _zout.writestr(_it, _d)
_zout.close()
_zin.close()
_shutil.move(_tmp, OUT)
print("bereinigt: ppt/authors.xml entfernt")
