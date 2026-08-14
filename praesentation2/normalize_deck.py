#!/usr/bin/env python3
"""Vereinheitlicht das Design der überarbeiteten Präsentation.

Eingabe:  upload_v2.pptx  (inhaltlich aktualisierte Fassung)
Ausgabe:  SKP-Research-geprueft-sortiert.pptx

Es wird ausschließlich Design angefasst — Schriftgrößen, Farben, Positionen,
fehlende Seitenmoebel. Inhalte bleiben erhalten; auf Folie 5 werden die drei
Blöcke der linken Spalte zu Bullets in 10 pt zusammengefasst, damit sie in
den farbigen Block passen.
"""
import os
import re
import shutil
import zipfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "upload_v2.pptx")
OUT = os.path.join(HERE, "SKP-Research-geprueft-sortiert.pptx")

BLUE = RGBColor(0x1A, 0x70, 0xB8)
AMBER = RGBColor(0xE0, 0x86, 0x1A)
OLIVE = RGBColor(0x6E, 0x7D, 0x00)
RED = RGBColor(0xC0, 0x39, 0x2B)
TXT = RGBColor(0x3D, 0x3D, 0x3D)
MUT = RGBColor(0x88, 0x88, 0x88)
GREY_HEAD = RGBColor(0x77, 0x77, 0x77)

C_BLUE = RGBColor(0xEA, 0xF1, 0xF8)
C_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
C_AMBER = RGBColor(0xFD, 0xF3, 0xE3)

HF, BF = "Quicksand", "Open Sans"

TITLE_PT = 20          # einheitliche Titelgröße
COLHEAD_PT = 11        # Spaltenkopf
CARDHEAD_PT = 9.5      # Kartenüberschrift
SRC_PT = 6.8           # Quellenzeile

COL = [1.56, 5.30, 9.04]
CW, PAD = 3.55, 0.16

prs = Presentation(SRC)
slides = list(prs.slides)


# ------------------------------------------------------------------ Helfer
def drop(shape):
    shape._element.getparent().remove(shape._element)


def set_geom(shape, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = (
        Inches(x), Inches(y), Inches(w), Inches(h))


def style_runs(tf, size=None, font=None, color=None, bold=None):
    for pa in tf.paragraphs:
        for r in pa.runs:
            if size is not None:
                r.font.size = Pt(size)
            if font is not None:
                r.font.name = font
            if color is not None:
                r.font.color.rgb = color
            if bold is not None:
                r.font.bold = bold


def first_text(shape):
    if not shape.has_text_frame:
        return ""
    for pa in shape.text_frame.paragraphs:
        t = "".join(r.text for r in pa.runs).strip()
        if t:
            return t
    return ""


def cards_of(slide):
    """Alle Kartenflächen (gefüllte AutoShapes ohne eigenen Text)."""
    out = []
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not first_text(sh):
            if sh.width and sh.width > Inches(1.0):
                out.append(sh)
    return out


def card_fill(shape):
    try:
        return shape.fill.fore_color.rgb
    except Exception:
        return None


def head_color_for(card):
    f = card_fill(card)
    if f == C_AMBER:
        return AMBER
    return BLUE


def enclosing_card(slide, box):
    """Die Karte, in der eine Textbox liegt."""
    if box.left is None:
        return None
    bx, by = box.left, box.top
    for c in cards_of(slide):
        if (c.left - Inches(0.05) <= bx and c.top - Inches(0.05) <= by
                and bx + box.width <= c.left + c.width + Inches(0.05)
                and by < c.top + c.height):
            return c
    return None


def tb(slide, x, y, w, h):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    t = box.text_frame
    t.word_wrap = True
    t.margin_left = t.margin_right = t.margin_top = t.margin_bottom = 0
    t.vertical_anchor = MSO_ANCHOR.TOP
    return t


def para(tf, first=False, before=0, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_before = Pt(before)
    p.space_after = Pt(0)
    if line:
        p.line_spacing = line
    return p


def run(p, text, size, font=BF, color=TXT, bold=False):
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return r


def card(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.adjustments[0] = 0.06
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.text = ""
    return sh



def rebuild_bullet(pa, full, accent, size=8.6):
    """Zeile neu aufbauen: Aufzaehlungszeichen im Akzent, Text in Grundfarbe."""
    for r in list(pa.runs):
        r._r.getparent().remove(r._r)
    marker, rest = "", full
    m = re.match(r"^(\s*▪\s*)(.*)$", full, re.S)
    if m:
        marker, rest = "▪ ", m.group(2)
    if marker:
        run(pa, marker, size, BF, accent)
    run(pa, rest, size, BF, TXT)


# ======================================================== 1. Titel vereinheitlichen
for s in slides:
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 0:
            set_geom(ph, 1.56, 1.00, 10.98, 0.57)
            style_runs(ph.text_frame, size=TITLE_PT, bold=True)
            ph.text_frame.word_wrap = True

# ======================================================== 2. Quellenzeilen
for s in slides:
    for sh in list(s.shapes):
        if sh.has_text_frame and first_text(sh).startswith("Quellen"):
            if sh.top is not None and sh.top > Inches(6.5):
                set_geom(sh, 1.56, 7.05, 9.50, 0.30)
                style_runs(sh.text_frame, size=SRC_PT, font=BF, color=MUT,
                           bold=False)
                for pa in sh.text_frame.paragraphs:
                    pa.line_spacing = 1.05
                sh.text_frame.word_wrap = True

# ======================================================== 3. Spaltenköpfe
for s in slides:
    for sh in s.shapes:
        if not sh.has_text_frame or sh.top is None:
            continue
        if abs(sh.top - Inches(1.60)) < Inches(0.04) and sh.height <= Inches(0.30):
            style_runs(sh.text_frame, size=COLHEAD_PT, font=HF, bold=True)

# ======================================================== 4. Kartenüberschriften
# Auf den nachträglich bearbeiteten Folien haben die Karten ihre
# Überschriftenformatierung verloren.
FIX_HEADS = {
    4: ["Vertragsdauer", "Unterm Strich"],
    6: ["Nachher", "Dafür", "Dagegen / Kritik", "Wer entscheidet"],
    9: ["Sofort / Sep. 2026", "Q4 2026", "Q4 2026 / vor Verfahren",
        "2026/27", "DD-Red Flag"],
}
for idx, heads in FIX_HEADS.items():
    s = slides[idx - 1]
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        ft = first_text(sh)
        if ft not in heads:
            continue
        c = enclosing_card(s, sh)
        col = head_color_for(c) if c is not None else BLUE
        head_done = False
        for pa in sh.text_frame.paragraphs:
            full = "".join(r.text for r in pa.runs)
            if not full.strip():
                continue
            if not head_done:
                for r in pa.runs:
                    r.font.size = Pt(CARDHEAD_PT)
                    r.font.name = HF
                    r.font.bold = True
                    r.font.color.rgb = col
                head_done = True
                continue
            # Fließtext: Aufzählungszeichen einfärben, Text neutral halten
            rebuild_bullet(pa, full, col)

# ======================================================== 5. Folie 5 · linke Spalte
s5 = slides[4]
for sh in list(s5.shapes):
    if sh.left is None or sh.left >= Inches(5.2):
        continue
    if sh.top is not None and Inches(1.80) <= sh.top <= Inches(6.60):
        drop(sh)

card(s5, COL[0], 1.88, CW, 4.78, C_BLUE)
tf = tb(s5, COL[0] + PAD, 2.00, CW - 2 * PAD, 4.54)

BLOCKS = [
    ("Pool + MPT", BLUE, [
        "Landesregierung: Zusammenführung ist „wichtiger Baustein der "
        "multiprofessionellen Zusammenarbeit an Schulen“",
        "Pool: Budget je Schule statt Bewilligung je Kind",
        "MPT: Lehrkraft, Schulassistenz, Schulbegleitung, Schulsozialarbeit "
        "und Sonderpädagogik als ein Team",
    ]),
    ("Update 13.08.2026", AMBER, [
        "Endzustand offen, Fahrplan konkret: Konzepte der Modellkommunen bis "
        "03.07.2026, Umsetzung 2026/27 bis Ende 2028/29, Evaluation "
        "Q4 2029/Q1 2030",
        "Landesweite Strategie danach offen",
    ]),
    ("Wer voraussichtlich gewinnt", BLUE, [
        "Große und regionale Träger besser positioniert: Personalmasse, "
        "Fachleitungen, Springer, Schulnähe",
        "Gemeinnützigkeit ist kein Automatismus: Dithmarschen öffnete auch für "
        "freie Leistungserbringer der Jugend- und Eingliederungshilfe",
        "Kleine private Anbieter haben eine Chance, wenn sie ganze Standorte "
        "bedienen können",
    ]),
]
first = True
for head, col, items in BLOCKS:
    p = para(tf, first=first, before=0 if first else 6, line=1.0)
    first = False
    run(p, head, 10, HF, col, bold=True)
    for it in items:
        p = para(tf, before=2, line=1.0)
        run(p, "▪ ", 10, BF, AMBER)
        run(p, it, 10, BF, TXT)

# ======================================================== 6. Folie 7 · Layout
s7 = slides[6]
target_layout = None
for m in prs.slide_masters:
    for lay in m.slide_layouts:
        if lay.name == "Inhalt-1":
            target_layout = lay
            break
    if target_layout:
        break
for rel in s7.part.rels.values():
    if rel.reltype == RT.SLIDE_LAYOUT:
        rel._target = target_layout.part
        break
for ph in s7.placeholders:
    if ph.placeholder_format.idx == 12:
        ph.text_frame.text = "Pilotphase"
    if ph.placeholder_format.idx == 0:
        set_geom(ph, 1.56, 1.00, 10.98, 0.57)
        style_runs(ph.text_frame, size=TITLE_PT, bold=True)
# Spaltenköpfe und Inhalt der Trennfolie auf das Raster bringen
for sh in s7.shapes:
    if not sh.has_text_frame or sh.left is None:
        continue
    ft = first_text(sh)
    if ft in ("Pinneberg", "Dithmarschen / Kiel"):
        set_geom(sh, 1.56 if ft == "Pinneberg" else 7.17, 1.60, 5.42, 0.24)
        style_runs(sh.text_frame, size=COLHEAD_PT, font=HF, bold=True)
    elif sh.top is not None and Inches(1.9) <= sh.top <= Inches(6.8):
        x = 1.56 if sh.left < Inches(6.5) else 7.17
        set_geom(sh, x, sh.top / 914400, 5.42, sh.height / 914400)
        sz = 9.0 if sh.height <= Inches(0.24) else 8.6
        style_runs(sh.text_frame, size=sz, font=BF)
        if sh.height <= Inches(0.24):
            style_runs(sh.text_frame, font=HF, bold=True, color=BLUE)
        else:
            style_runs(sh.text_frame, color=TXT, bold=False)

# ======================================================== 7. Folie 10 · Titel
s10 = slides[9]
tf10 = tb(s10, 1.56, 1.00, 10.98, 0.57)
run(para(tf10, True), "Kostenwirkung des Bundesentwurfs und Förderquoten der Länder",
    TITLE_PT, HF, GREY_HEAD, bold=True)
for ph in s10.placeholders:
    if ph.placeholder_format.idx == 12:
        ph.text_frame.text = "Anhang"
# Abbildungen unter den Titel setzen
pics = [sh for sh in s10.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
pics.sort(key=lambda sh: sh.left)
if len(pics) == 2:
    a, b = pics
    set_geom(a, 1.56, 1.80, 4.91, 3.14)
    set_geom(b, 7.30, 1.80, 5.29, 4.80)

# ======================================================== 8. Folie 11 · Kopf
s11 = slides[10]
for ph in list(s11.placeholders):
    if ph.placeholder_format.idx == 12:
        ph.text_frame.text = "Anhang · Quellen"
    if ph.placeholder_format.idx == 11:
        drop(ph)
for sh in s11.shapes:
    if first_text(sh).startswith("Quellen – Modelle"):
        set_geom(sh, 1.56, 1.00, 10.98, 0.57)
        style_runs(sh.text_frame, size=TITLE_PT, font=HF, color=GREY_HEAD,
                   bold=True)
    elif sh.has_text_frame and sh.top is not None and sh.top > Inches(1.4):
        x = 1.56 if sh.left < Inches(6.5) else 7.17
        set_geom(sh, x, sh.top / 914400, 5.42, sh.height / 914400)
        style_runs(sh.text_frame, size=8.6, font=BF, color=TXT)

# ======================================================== 9. Folie 8 · Spaltenköpfe
s8 = slides[7]
for sh in s8.shapes:
    if first_text(sh) in ("Chancen", "Risiken"):
        style_runs(sh.text_frame, size=COLHEAD_PT, font=HF, bold=True)


# ======================================================== 10. Feinschliff
def move_card_and_text(slide, card_shape, new_y, new_h):
    """Karte samt zugehoeriger Textbox verschieben und in der Hoehe anpassen."""
    old_y = card_shape.top
    for sh in slide.shapes:
        if sh is card_shape or not sh.has_text_frame or sh.left is None:
            continue
        if (card_shape.left - Inches(0.05) <= sh.left
                and sh.left + sh.width <= card_shape.left + card_shape.width + Inches(0.05)
                and old_y - Inches(0.05) <= sh.top < old_y + card_shape.height):
            sh.top = Inches(new_y + 0.12)
            sh.height = Inches(max(new_h - 0.24, 0.2))
    card_shape.top = Inches(new_y)
    card_shape.height = Inches(new_h)


# Folie 2: die beiden breiten Karten an ihren Inhalt anpassen
s2 = slides[1]
wide = sorted((c for c in cards_of(s2) if c.width > Inches(9)),
              key=lambda c: c.top)
for c, y in zip(wide, (3.20, 4.30)):
    move_card_and_text(s2, c, y, 0.90)

# Folie 9: Karten an ihren Inhalt anpassen
s9 = slides[8]
grid = sorted(cards_of(s9), key=lambda c: (round(c.top / 914400, 1),
                                           round(c.left / 914400, 1)))
for i, c in enumerate(grid):
    move_card_and_text(s9, c, 1.85 + (i // 3) * 1.75, 1.55)

# Folie 8: zu knappe Textbox der letzten Risikozeile
for sh in slides[7].shapes:
    if first_text(sh).startswith("Betreiber sind öffentlich nicht belastbar"):
        set_geom(sh, 7.35, 5.05, 5.24, 0.52)

# Folie 7: Spaltenkopf-Farbe angleichen
for sh in slides[6].shapes:
    if first_text(sh) == "Dithmarschen / Kiel":
        style_runs(sh.text_frame, color=BLUE)

# Fließtext ohne eigene Formatierung auf die Hausschrift bringen
for s in slides:
    for sh in s.shapes:
        if sh.is_placeholder or not sh.has_text_frame or sh.top is None:
            continue
        if sh.top < Inches(1.70) or first_text(sh).startswith("Quellen"):
            continue
        for pa in sh.text_frame.paragraphs:
            for r in pa.runs:
                if r.font.name is None:
                    r.font.name = BF
                if r.font.size is None:
                    r.font.size = Pt(8.6)
                if r.font.color is None or r.font.color.type is None:
                    r.font.color.rgb = TXT

prs.save(OUT)

# leere Review-Metadaten entfernen
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
