# -*- coding: utf-8 -*-
"""Baut den 14-Folien-Leitfaden 'Kauf insolventer Pflegedienste' in das
Corporate-Template InsolvenzProzess.pptx (Master/Logo/Theme bleiben erhalten)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "InsolvenzProzess.pptx"
OUT = "InsolvenzProzess.pptx"  # in place -> Ergebnisdatei

# ---- Markenpalette (aus Theme) ----
BLUE="1A70B8"; LIME="BBCF00"; GREEN="2A6A38"; ORANGE="F89929"; PURPLE="963A96"; RED="D73225"
GREY="777777"; WHITE="FFFFFF"; DARK="333333"
# Varianten/Tints (Designregel erlaubt Varianten für lesbaren Kontrast)
BLUETINT="EAF2F9"; GREYTINT="F1F4F6"; ORANGETINT="FDEFD9"; REDTINT="FBEAE7"; GREENTINT="E9F1EA"
BORDER="D8DEE2"; BLUEDK="14547F"
HFONT="Quicksand"; BFONT="Open Sans"

LEFT=1.56; RIGHT=12.54; CW=RIGHT-LEFT      # content frame (Inhalt-1)
CT=1.80; CB=6.70                            # content vertical band

prs=Presentation(SRC)
m1=prs.slide_masters[1]; m2=prs.slide_masters[2]
def lay(master,name): return next(L for L in master.slide_layouts if L.name==name)
L_TITLE=lay(m1,"Titel/Trenner-1")
L_CONTENT=lay(m1,"Inhalt-1")
L_END=lay(m2,"Endfolie-Logo")

# ---- vorhandene (leere) Folien entfernen (inkl. Parts/Rels) ----
from pptx.oxml.ns import qn
sldIdLst=prs.slides._sldIdLst
pres_part=prs.part
for sid in list(sldIdLst):
    rId=sid.get(qn('r:id'))
    sldIdLst.remove(sid)
    try: pres_part.drop_rel(rId)
    except Exception: pass

def rgb(h): return RGBColor.from_string(h)
def A(v): return Inches(v)
ALIGN={'l':PP_ALIGN.LEFT,'c':PP_ALIGN.CENTER,'r':PP_ALIGN.RIGHT}
ANCH={'t':MSO_ANCHOR.TOP,'m':MSO_ANCHOR.MIDDLE,'b':MSO_ANCHOR.BOTTOM}

def add_slide(layout):
    return prs.slides.add_slide(layout)

def shape(sl,kind,x,y,w,h,fill=None,line=None,lw=1.0,radius=0.10):
    sp=sl.shapes.add_shape(kind,A(x),A(y),A(w),A(h))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb=rgb(fill)
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=rgb(line); sp.line.width=Pt(lw)
    try: sp.shadow.inherit=False
    except: pass
    if kind==MSO_SHAPE.ROUNDED_RECTANGLE:
        try: sp.adjustments[0]=radius
        except: pass
    return sp

def _apply(tf,paras,anchor):
    tf.word_wrap=True
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    tf.vertical_anchor=ANCH[anchor]
    for pi,para in enumerate(paras):
        p=tf.paragraphs[0] if pi==0 else tf.add_paragraph()
        p.alignment=ALIGN[para.get('a','l')]
        if para.get('ls'): p.line_spacing=Pt(para['ls'])
        if para.get('sa') is not None: p.space_after=Pt(para['sa'])
        if para.get('sb') is not None: p.space_before=Pt(para['sb'])
        for run in para['r']:
            r=p.add_run(); r.text=run['t']
            r.font.name=run.get('f',BFONT); r.font.size=Pt(run.get('s',12))
            r.font.bold=run.get('b',False); r.font.italic=run.get('i',False)
            r.font.color.rgb=rgb(run.get('c',DARK))

def box(sl,x,y,w,h,paras,anchor='t'):
    tb=sl.shapes.add_textbox(A(x),A(y),A(w),A(h))
    _apply(tb.text_frame,paras,anchor); return tb

def label(sl,shp,text,size,color,font=BFONT,bold=True,anchor='m',align='c'):
    _apply(shp.text_frame,[{'a':align,'r':[{'t':text,'s':size,'c':color,'f':font,'b':bold}]}],anchor)

def para(text,size=12,color=DARK,font=BFONT,bold=False,italic=False,a='l',ls=None,sa=None,sb=None):
    return {'a':a,'ls':ls,'sa':sa,'sb':sb,'r':[{'t':text,'s':size,'c':color,'f':font,'b':bold,'i':italic}]}

def set_ph(sl,idx,text,size=None,color=None,bold=None):
    for ph in sl.placeholders:
        if ph.placeholder_format.idx==idx:
            ph.text_frame.word_wrap=True
            r=ph.text_frame.paragraphs[0].runs
            ph.text_frame.paragraphs[0].text=text if False else ph.text_frame.paragraphs[0].text
            # set text preserving inherited style
            ph.text=text
            if size or color or bold is not None:
                for p in ph.text_frame.paragraphs:
                    for rn in p.runs:
                        if size: rn.font.size=Pt(size)
                        if color: rn.font.color.rgb=rgb(color)
                        if bold is not None: rn.font.bold=bold
            return ph
    return None

def del_ph(sl,idxs):
    for ph in list(sl.placeholders):
        if ph.placeholder_format.idx in idxs:
            ph._element.getparent().remove(ph._element)

def content_slide(title,eyebrow=None):
    sl=add_slide(L_CONTENT)
    del_ph(sl,[1,12])        # OBJECT + BODY weg -> freie Fläche
    set_ph(sl,0,title)       # TITLE (Quicksand, Markenfarbe)
    return sl

def icon_circle(sl,x,y,d,fill,text,size,tcolor=WHITE):
    c=shape(sl,MSO_SHAPE.OVAL,x,y,d,d,fill=fill)
    label(sl,c,text,size,tcolor,font=HFONT,bold=True); return c

RR=MSO_SHAPE.ROUNDED_RECTANGLE; RECT=MSO_SHAPE.RECTANGLE; OV=MSO_SHAPE.OVAL; DIA=MSO_SHAPE.DIAMOND

# =====================================================================
# 1 — TITEL (Titel/Trenner-Layout)
# =====================================================================
s=add_slide(L_TITLE)
set_ph(s,0,"Kauf eines insolventen Pflegedienstes")
del_ph(s,[1,10])   # leere Namens-/Datumsplatzhalter entfernen
# Titelfolie: Zusatztexte als Boxen (klar positioniert)
box(s,0.71,5.15,11.6,0.5,[para("Prozess, Zeitplan & Achievement-Gates — Fokus außerklinische Intensivpflege (AKI, § 132l SGB V)",14,GREY,BFONT)])
box(s,0.71,2.95,11.0,0.7,[para("M&A-LEITFADEN · DISTRESSED HEALTHCARE",13,BLUE,HFONT,bold=True)])

# =====================================================================
# 2 — MANAGEMENT SUMMARY
# =====================================================================
s=content_slide("Warum insolvente AKI-Targets attraktiv sind","Management Summary")
items=[("1","Konsolidierungswelle läuft","Demografie, Personalmangel & IPReG treffen kleine Träger überproportional — 2024–2027 gilt als aktivstes Fenster.",BLUE),
       ("2","Massive Bewertungsrabatte","Distressed-Preise liegen weit unter regulären EBITDA-Multiples; dokumentiert bis ~90 % Wertverlust.",GREEN),
       ("3","Zwei klare Hebel","In der Pflege sind Mitarbeiter und Patienten die entscheidenden — oft die einzigen — Werttreiber.",ORANGE),
       ("4","Buy-and-Build-Fit","Zukauf statt Neugründung: Patienten, Personal und (idealerweise) Versorgungsverträge statt Aufbau bei null.",PURPLE)]
y=CT
for n,h,t,c in items:
    shape(s,RR,LEFT,y,6.35,1.05,fill=WHITE,line=BORDER,lw=1,radius=0.10)
    icon_circle(s,LEFT+0.2,y+0.25,0.55,c,n,16)
    box(s,LEFT+0.95,y+0.13,5.25,0.35,[para(h,13.5,c,HFONT,bold=True)])
    box(s,LEFT+0.95,y+0.5,5.3,0.5,[para(t,10.5,DARK,BFONT,ls=12.5)])
    y+=1.15
# rechtes Panel (Blau)
px=8.35
shape(s,RR,px,CT,4.19,4.6,fill=BLUE,radius=0.06)
box(s,px+0.3,CT+0.25,3.6,0.3,[para("KERNAUSSAGE",11,LIME,HFONT,bold=True)])
box(s,px+0.3,CT+0.6,3.6,1.15,[para("Der Versorgungsvertrag ist das eigentliche Asset — nicht das Inventar.",18,WHITE,HFONT,bold=True,ls=21)])
stats=[("2–3 Mon.","Insolvenzgeld-Fenster: Löhne getragen, Betrieb läuft weiter"),
       ("§ 132l","Versorgungsvertrag — Voraussetzung fürs Abrechnen"),
       ("< 4 Wo.","Zeitdruck der Distressed-DD statt Monaten")]
sy=CT+1.95
for n,l in stats:
    box(s,px+0.3,sy,1.55,0.6,[para(n,20,LIME,HFONT,bold=True)],anchor='m')
    box(s,px+1.95,sy,2.0,0.62,[para(l,10,"E9F1F8",BFONT,ls=11.5)],anchor='m')
    sy+=0.85

# =====================================================================
# 3 — ZWEI WELTEN (Vergleich)
# =====================================================================
s=content_slide("Zwei Welten: normaler M&A vs. Distressed","Grundverständnis")
c0x,c1x,c2x=LEFT,LEFT+2.35,LEFT+6.6
w0,w1,w2=2.25,4.15,4.3
shape(s,RR,c1x,CT,w1,0.45,fill=BLUE,radius=0.12); label(s,s.shapes[-1],"Normaler M&A",13,WHITE,HFONT,anchor='m')
shape(s,RR,c2x,CT,w2,0.45,fill=ORANGE,radius=0.12); label(s,s.shapes[-1],"Distressed / Insolvenz",13,WHITE,HFONT,anchor='m')
rows=[("Zeitrahmen","Monate, geordnet","Wochen — Verwalter drängt"),
      ("Due Diligence","Vollständiger Datenraum","Limitierte Infos, Käuferrisiko"),
      ("Verkäufer","Eigentümer","Insolvenzverwalter (+ Gläubiger)"),
      ("Gewährleistung","Garantien / W&I","Praktisch keine — „wie besichtigt“"),
      ("Preislogik","EBITDA-Multiple","Asset-/masseorientiert"),
      ("Personal","§ 613a voll","§ 613a, aber Kündigung max. 3 Mon."),
      ("Freigabe","Gesellschafter","Gericht + Gläubigerausschuss")]
y=CT+0.52; rh=0.55
for i,(a,b,c) in enumerate(rows):
    if i%2==0: shape(s,RECT,LEFT,y,CW,rh,fill=GREYTINT)
    box(s,c0x+0.05,y,w0,rh,[para(a,11.5,BLUE,BFONT,bold=True)],anchor='m')
    box(s,c1x+0.12,y,w1-0.2,rh,[para(b,11,DARK,BFONT)],anchor='m')
    box(s,c2x+0.12,y,w2-0.2,rh,[para(c,11,DARK,BFONT)],anchor='m')
    y+=rh
box(s,LEFT,y+0.12,CW,0.4,[para("Kernunterschied: Kauf unter Zeitdruck, ohne Garantien, von einem Dritten — Substanz selbst prüfen.",11.5,GREY,BFONT,italic=True)])

# =====================================================================
# 4 — VERFAHRENSKUNDE (3 Karten)
# =====================================================================
s=content_slide("Wer ist mein Gegenüber? Drei Verfahrensarten","Verfahrenskunde")
cards=[("Regelinsolvenz",BLUE,"Insolvenzverwalter",
        ["Schuldner verliert Verfügungsgewalt","Verwalter führt & verkauft die Masse","Neutral — aber kaum Gewährleistung"],"Partner: der Verwalter"),
       ("Eigenverwaltung",PURPLE,"Management + Sachwalter",
        ["Geschäftsführung bleibt (§ 270 InsO)","Sachwalter überwacht nur","Dieselben Leute wie in der Krise"],"Erhöhte Täuschungsgefahr"),
       ("Schutzschirm",GREEN,"Mgmt + vorl. Sachwalter",
        ["Sonderform vorläufige Eigenverwaltung","Nur bei drohender Insolvenz","Voraussetzung: Sanierungsfähigkeit"],"Frühe Krise, § 270b InsO")]
cw=(CW-0.44)/3; x=LEFT
for t,c,who,pts,tag in cards:
    shape(s,RR,x,CT,cw,4.55,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    shape(s,RR,x,CT,cw,0.8,fill=c,radius=0.06)
    shape(s,RECT,x,CT+0.4,cw,0.4,fill=c)
    box(s,x,CT+0.05,cw,0.75,[para(t,18,WHITE,HFONT,bold=True,a='c')],anchor='m')
    box(s,x+0.25,CT+0.95,cw-0.5,0.25,[para("VERHANDLUNGSPARTNER",9,GREY,HFONT,bold=True)])
    box(s,x+0.25,CT+1.2,cw-0.5,0.5,[para(who,14,c,HFONT,bold=True,ls=16)])
    yy=CT+1.85
    for pt in pts:
        shape(s,OV,x+0.28,yy+0.07,0.1,0.1,fill=ORANGE)
        box(s,x+0.5,yy-0.02,cw-0.75,0.55,[para(pt,11,DARK,BFONT,ls=13)])
        yy+=0.62
    shape(s,RR,x+0.25,CT+3.85,cw-0.5,0.45,fill=GREYTINT,radius=0.18)
    box(s,x+0.25,CT+3.85,cw-0.5,0.45,[para(tag,10.5,c,BFONT,bold=True,italic=True,a='c')],anchor='m')
    x+=cw+0.22

# =====================================================================
# 5 — ZWEI KAUFSTRUKTUREN + TWIST
# =====================================================================
s=content_slide("Zwei Kaufstrukturen — der Pflege-Twist entscheidet","Deal-Architektur")
def deal_card(x,letter,lcolor,title,sub,rows):
    shape(s,RR,x,CT,5.35,2.55,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    icon_circle(s,x+0.22,CT+0.2,0.55,lcolor,letter,18)
    box(s,x+0.95,CT+0.18,4.2,0.38,[para(title,16,BLUE,HFONT,bold=True)])
    box(s,x+0.95,CT+0.56,4.2,0.28,[para(sub,10.5,GREY,BFONT,italic=True)])
    yy=CT+1.0
    for sign,txt in rows:
        col=GREEN if sign=="+" else RED
        c=shape(s,OV,x+0.25,yy+0.02,0.24,0.24,fill=col); label(s,c,sign,12,WHITE,BFONT,anchor='m')
        box(s,x+0.6,yy-0.02,4.6,0.34,[para(txt,11.5,DARK,BFONT)],anchor='m')
        yy+=0.37
deal_card(LEFT,"A",BLUE,"Asset Deal","„Übertragende Sanierung“ — Normalfall",
          [("+","Keine Haftung für Altverbindlichkeiten"),("+","Keine Alt-Personalkosten vor Eröffnung"),
           ("−","Verträge gehen NICHT über (außer § 613a)"),("−","Jeder Kassen-/Lieferantenvertrag neu")])
deal_card(LEFT+5.63,"B",PURPLE,"Share Deal via Insolvenzplan","Rechtsträger überlebt entschuldet",
          [("+","Versorgungsverträge & Zulassungen bleiben"),("+","Kein Abrechnungsbruch, Verordnungen bestehen"),
           ("−","Mehr latente Haftung wird mitgekauft"),("−","Komplexer: Gläubigermehrheiten nötig")])
# Twist-Banner
ty=CT+2.75
shape(s,RR,LEFT,ty,CW,1.6,fill=BLUE,radius=0.05)
shape(s,RR,LEFT,ty,2.0,1.6,fill=ORANGE,radius=0.05); shape(s,RECT,LEFT+1.5,ty,0.5,1.6,fill=ORANGE)
box(s,LEFT,ty,2.0,1.6,[para("PFLEGE-",17,"3A2A00",HFONT,bold=True,a='c'),para("TWIST",17,"3A2A00",HFONT,bold=True,a='c')],anchor='m')
box(s,LEFT+2.25,ty+0.16,CW-2.4,1.3,[
    para("Der Versorgungsvertrag (§ 132l / § 132a SGB V, § 72 SGB XI) hängt am Rechtsträger — nicht am Betrieb.",13.5,WHITE,BFONT,bold=True,ls=16,sa=3),
    para("Im Asset Deal geht er NICHT automatisch über → drohende Abrechnungslücke. Betreiberwechsel vorab mit den Kassen klären; in der Intensivpflege ist der Insolvenzplan deshalb oft überlegen.",11.5,"E9F1F8",BFONT,ls=14)],anchor='m')

# =====================================================================
# 6 — PROZESS-ZEITSTRAHL + GATES (Hero)
# =====================================================================
s=content_slide("Insolvenz-Verfahrensachse über M&A-Phasen 0–6","Der Prozess auf einen Blick")
# Insolvenz-Lane
box(s,LEFT,CT-0.02,6,0.28,[para("INSOLVENZ-VERFAHREN",11,BLUE,HFONT,bold=True)])
scale=CW/12.33
def sx(v): return LEFT + v*scale
ins=[(0.0,2.15,"Krise / drohende ZU",GREY),(2.22,1.55,"Antrag",BLUE),
     (3.84,4.1,"Vorläufiges Verfahren · Insolvenzgeld (2–3 Mon.)",GREEN),
     (8.01,1.9,"Eröffnung",BLUE),(9.98,2.35,"Berichts- & Prüfungstermin",GREY)]
yl=CT+0.3
for vx,vw,t,c in ins:
    shape(s,RR,sx(vx),yl,vw*scale,0.62,fill=c,radius=0.10)
    box(s,sx(vx)+0.04,yl,vw*scale-0.08,0.62,[para(t,9.5,WHITE,BFONT,bold=True,a='c',ls=10.5)],anchor='m')
# Phasen-Lane
box(s,LEFT,CT+1.25,8,0.28,[para("M&A-PROZESS · PHASEN & ACHIEVEMENT-GATES",11,ORANGE,HFONT,bold=True)])
phases=[(0.0,2.15,"0","Sourcing & Ansprache"),(2.22,1.55,"1","NDA & Interesse"),
        (3.84,2.0,"2","Unterlagen & Konzept"),(5.92,2.0,"3","DD & Bewertung"),
        (8.01,1.9,"4","C-Level & Preis"),(9.98,1.15,"5","Verhandlung"),(11.21,1.12,"6","Closing")]
yp=CT+1.6
for vx,vw,n,t in phases:
    shape(s,RR,sx(vx),yp,vw*scale,0.85,fill=BLUETINT,line=BLUE,lw=1,radius=0.10)
    box(s,sx(vx)+0.07,yp+0.05,0.5,0.38,[para(n,18,BLUE,HFONT,bold=True)])
    box(s,sx(vx)+0.06,yp+0.42,vw*scale-0.1,0.4,[para(t,9,DARK,BFONT,bold=True,ls=10)])
# Gates (Rauten unter Phasenmitte)
gates=[(1.075,"G0","Target qualifiziert",1.6),(2.995,"G1","NDA + Datenraum",1.6),
       (4.84,"G2","Indik. Angebot / LOI",1.7),(6.92,"G3","DD & Bewertung fertig",1.7),
       (8.96,"G4","Freigabe Gläubiger­ausschuss",1.75),(10.555,"G5","Signing",1.0),(11.77,"G6","Closing",1.2)]
gy=CT+2.7
for cx,g,t,lw in gates:
    cxi=sx(cx)
    d=shape(s,DIA,cxi-0.30,gy,0.6,0.6,fill=ORANGE)
    label(s,d,g,11,"3A2A00",BFONT,anchor='m')
    lx=max(LEFT,min(cxi-lw/2, RIGHT-lw))
    box(s,lx,gy+0.64,lw,0.6,[para(t,8.5,GREY,BFONT,a='c',ls=9.5)])
box(s,LEFT,CB-0.35,CW,0.35,[para("Best-Practice-Kaufzeitpunkt: im vorläufigen Verfahren / rund um die Eröffnung — bevor Patienten & Fachkräfte abwandern und solange Insolvenzgeld die Löhne trägt.",11,ORANGE,BFONT,italic=True,bold=True,a='c')])

# =====================================================================
# 7 — PHASEN 0–3
# =====================================================================
def phase_rows(s,rows):
    y=CT; rh=1.13
    for n,t,d,g,gc in rows:
        shape(s,RR,LEFT,y,CW,rh-0.1,fill=WHITE,line=BORDER,lw=1,radius=0.05)
        icon_circle(s,LEFT+0.22,y+0.26,0.55,BLUE,n,18)
        box(s,LEFT+0.95,y+0.13,7.0,0.38,[para(t,14.5,BLUE,HFONT,bold=True)])
        box(s,LEFT+0.95,y+0.5,7.4,0.5,[para(d,10.5,DARK,BFONT,ls=12.5)])
        d2=shape(s,DIA,LEFT+8.75,y+0.32,0.4,0.4,fill=ORANGE)
        shape(s,RR,LEFT+9.25,y+0.28,2.2,0.5,fill=ORANGETINT,line=ORANGE,lw=1,radius=0.14)
        box(s,LEFT+9.3,y+0.28,2.1,0.5,[para(g,10,"8A5A10",BFONT,bold=True,a='c',ls=11)],anchor='m')
        y+=rh
s=content_slide("Phasen 0–3: Sourcing bis Bewertung","Ablauf im Detail · Teil 1")
phase_rows(s,[
 ("0","Sourcing & Ansprache","Targets identifizieren (Kaltakquise, Netzwerk, Verwalter). Insolvenzbekanntmachungen laufend monitoren.","G0 · Target",ORANGE),
 ("1","NDA & Verkaufsinteresse","Verkaufsbereitschaft klären, NDA zeichnen, Datenraum-Zugang. Verfahrensart & Ansprechpartner feststellen.","G1 · NDA + DR",ORANGE),
 ("2","Unterlagen & Konzept","§ 132l/132a-Vertrag, Anlage 2 (Vergütung), MD-Berichte, BWA. Integrations-/Skalierungskonzept skizzieren.","G2 · LOI",ORANGE),
 ("3","Due Diligence & Bewertung","Dokumente nachrechnen, Regress-/Abrechnungsrisiken prüfen. WP & ggf. externes M&A-Team. Bewertung + Sanierungsplan.","G3 · DD fertig",ORANGE)])

# =====================================================================
# 8 — PHASEN 4–6
# =====================================================================
s=content_slide("Phasen 4–6: Freigabe bis Integration","Ablauf im Detail · Teil 2")
rows=[("4","C-Level & interne Kaufpreislogik","Target intern vorstellen, Kaufpreiskomponenten festlegen (Assets, Patientenwert, Sanierungskosten, Risikoabschläge). Finanzierung sichern.","— intern",False),
      ("4a","Freigabe der Insolvenzorgane","Insolvenzspezifisch: Kaufvertrag steht unter der aufschiebenden Bedingung der Zustimmung von Gericht + Gläubigerausschuss.","G4 · Freigabe",True),
      ("5","Verhandlung & Einigung","Preis, Finanzierung, minimale Gewährleistungen, Betreiberwechsel/Zulassung mit Kassen, Personalübergang (§ 613a).","G5 · Signing",False),
      ("6","Closing & Integration","Vollzug, Übergang Patienten & Personal, Konzernanbindung, PDL/Fachkräfte binden, Kassenverträge aktivieren.","G6 · Closing",False)]
y=CT; rh=1.13
for n,t,d,g,hl in rows:
    cardfill=BLUE if hl else WHITE
    tcol=WHITE if hl else DARK; ttl=WHITE if hl else BLUE
    shape(s,RR,LEFT,y,CW,rh-0.1,fill=cardfill,line=(BLUE if hl else BORDER),lw=1,radius=0.05)
    icon_circle(s,LEFT+0.22,y+0.26,0.55,(ORANGE if hl else BLUE),n,(14 if len(n)>1 else 18),tcolor=("3A2A00" if hl else WHITE))
    box(s,LEFT+0.95,y+0.13,7.4,0.38,[para(t,14.5,ttl,HFONT,bold=True)])
    box(s,LEFT+0.95,y+0.5,7.5,0.55,[para(d,10.5,tcol,BFONT,ls=12.5)])
    d2=shape(s,DIA,LEFT+8.9,y+0.32,0.4,0.4,fill=ORANGE)
    shape(s,RR,LEFT+9.4,y+0.28,2.05,0.5,fill=(BLUEDK if hl else ORANGETINT),line=ORANGE,lw=1,radius=0.14)
    box(s,LEFT+9.4,y+0.28,2.05,0.5,[para(g,10,(ORANGE if hl else "8A5A10"),BFONT,bold=True,a='c')],anchor='m')
    y+=rh

# =====================================================================
# 9 — DUE DILIGENCE FOKUS (Grid 2x3)
# =====================================================================
s=content_slide("Due Diligence: die pflegespezifischen Pflichtchecks","Prüf-Fokus")
dd=[("1","Versorgungsverträge & Zulassung","§ 132l / § 132a SGB V, § 72 SGB XI gültig & übertragbar? Noch in der Kassenliste?",BLUE),
    ("2","Vergütung (Anlage 2)","Vergütungsvereinbarungen nachrechnen — tragfähig, marktüblich, verhandelbar?",GREEN),
    ("3","MD-Prüfung & Ruf","Qualitätsberichte, laufende Prüfverfahren, drohende Kündigung, Mängelbescheide.",PURPLE),
    ("4","Abrechnung & Regress","Leistungsnachweise vs. Abrechnung; latente Rückforderungen (Jahre rückwirkend).",RED),
    ("5","Personalqualifikation","Examinierte Fachkräfte, PDL vorhanden? Behandlungspflege fachgerecht erbracht?",ORANGE),
    ("6","Patientenstruktur","Fallzahlen vs. Verordnungen; Konzentrationsrisiko; Loyalität beim Wechsel.",BLUE)]
cw=(CW-0.4)/3; rh=2.25; x0=LEFT; y0=CT
for i,(n,t,d,c) in enumerate(dd):
    col=i%3; row=i//3
    x=x0+col*(cw+0.2); y=y0+row*(rh+0.2)
    shape(s,RR,x,y,cw,rh,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    icon_circle(s,x+0.22,y+0.22,0.5,c,n,15)
    box(s,x+0.85,y+0.2,cw-1.05,0.6,[para(t,12.5,c,HFONT,bold=True,ls=14)],anchor='m')
    box(s,x+0.25,y+0.95,cw-0.5,1.15,[para(d,10.5,DARK,BFONT,ls=13)])

# =====================================================================
# 10 — SCAM-RADAR
# =====================================================================
s=content_slide("Scam-Radar: typische Täuschungen","Rote Flaggen")
box(s,LEFT,CT-0.02,6,0.28,[para("ALLGEMEIN (INSOLVENZNÄHE)",11,BLUE,HFONT,bold=True)])
box(s,LEFT+5.63,CT-0.02,6,0.28,[para("PFLEGE-SPEZIFISCH (AM GEFÄHRLICHSTEN)",11,RED,HFONT,bold=True)])
gen=[("Geschönte Zahlen / verschwiegene Verluste","Aufklärungspflicht verletzt → arglistige Täuschung, Rückabwicklung"),
     ("Assets nicht frei","Sicherungsübereignung, Eigentumsvorbehalt, Leasing — weniger Substanz"),
     ("Cherry-Picking vor Insolvenz","„Gute“ Verträge ausgegliedert — du bekommst die Hülle")]
pfl=[("Scheinpatienten / Karteileichen","Aufgeblähte Fallzahlen → gegen Verordnungen & Abrechnungen prüfen"),
     ("Nicht erbrachte Leistungen","Latente Regresse der Kassen + § 263 StGB — beim Share Deal geerbt"),
     ("Qualifikations-Fake / Key-Person","Unqualifiziertes Personal als Fachleistung; PDL wandert ab → Zulassung wackelt")]
def flags(x,data,accent,tint):
    y=CT+0.32
    for h,t in data:
        shape(s,RR,x,y,5.35,1.05,fill=tint,line=accent,lw=1,radius=0.06)
        shape(s,RR,x,y,0.12,1.05,fill=accent,radius=0.3)
        box(s,x+0.28,y+0.13,4.95,0.35,[para(h,12.5,DARK,HFONT,bold=True)])
        box(s,x+0.28,y+0.5,4.95,0.5,[para(t,10.5,"444444",BFONT,ls=12.5)])
        y+=1.18
flags(LEFT,gen,BLUE,BLUETINT)
flags(LEFT+5.63,pfl,RED,REDTINT)
shape(s,RR,LEFT,CB-0.62,CW,0.6,fill=BLUE,radius=0.06)
box(s,LEFT+0.25,CB-0.62,CW-0.5,0.6,[{'a':'l','r':[
    {'t':"Gegenmittel:  ",'s':12.5,'c':LIME,'f':HFONT,'b':True},
    {'t':"Umsatz gegen dokumentierte Leistungen prüfen · Zulassung/§ 132l-Status bei den Kassen verifizieren · PDL & Fachkräfte vertraglich binden.",'s':11.5,'c':WHITE,'f':BFONT}]}],anchor='m')

# =====================================================================
# 11 — BEWERTUNG
# =====================================================================
s=content_slide("Kaufpreislogik im Distressed-Kontext","Bewertung")
box(s,LEFT,CT,CW,0.35,[para("Vom Substanzwert zum risikoadjustierten Angebotspreis:",13.5,BLUE,HFONT,bold=True)])
blocks=[("Substanz / Assets","Inventar, Fahrzeuge, IT — nur unbelastete Gegenstände",BLUE),
        ("+ Going-Concern","Patientenstamm, Personal & laufende Zulassung",GREEN),
        ("− Sanierungskosten","Personalaufbau, Zulassung/Wechsel, Nachqualifikation",ORANGE),
        ("− Risikoabschläge","Latente Regresse, DD-Lücken, fehlende Garantien",RED)]
bw=(CW-3*0.45)/4; x=LEFT; y=CT+0.5
for i,(t,d,c) in enumerate(blocks):
    shape(s,RR,x,y,bw,2.0,fill=WHITE,line=BORDER,lw=1,radius=0.06)
    shape(s,RR,x,y,bw,0.6,fill=c,radius=0.09); shape(s,RECT,x,y+0.3,bw,0.3,fill=c)
    box(s,x+0.1,y,bw-0.2,0.6,[para(t,12.5,WHITE,HFONT,bold=True,a='c',ls=13)],anchor='m')
    box(s,x+0.18,y+0.72,bw-0.36,1.15,[para(d,10.5,DARK,BFONT,ls=13)])
    if i<3: box(s,x+bw-0.02,y+0.65,0.5,0.6,[para("→",20,GREY,BFONT,bold=True,a='c')],anchor='m')
    x+=bw+0.45
shape(s,RR,LEFT,y+2.25,CW,1.35,fill=BLUE,radius=0.05)
box(s,LEFT+0.3,y+2.4,CW-0.6,0.4,[para("Ergebnis: risikoadjustierter Angebotspreis",17,WHITE,HFONT,bold=True)])
for j,t in enumerate(["Kein EBITDA-Multiple — der Verwalter verkauft substanz-/masseorientiert.",
                      "Fehlende Garantien preislich einkalkulieren — Risiko trägt der Käufer.",
                      "„Return in 3 Monaten“ nur bei gesunder Substanz + gesicherter Zulassung realistisch."]):
    yy=y+2.82+j*0.28
    shape(s,OV,LEFT+0.32,yy+0.06,0.11,0.11,fill=LIME)
    box(s,LEFT+0.55,yy-0.02,CW-0.9,0.3,[para(t,11,"E9F1F8",BFONT)])

# =====================================================================
# 12 — INTEGRATION (2 Hebel)
# =====================================================================
s=content_slide("Integration & Skalierung — die zwei Hebel","Nach dem Closing")
def lever(x,ic,title,num,color,pts):
    shape(s,RR,x,CT,5.35,4.55,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    icon_circle(s,x+0.3,CT+0.3,0.85,color,ic,20)
    box(s,x+1.35,CT+0.35,3.8,0.5,[para(title,22,color,HFONT,bold=True)],anchor='m')
    box(s,x+1.35,CT+0.85,3.8,0.3,[para("Werttreiber #"+num,11,GREY,BFONT,italic=True)])
    yy=CT+1.55
    for pt in pts:
        shape(s,OV,x+0.38,yy+0.05,0.14,0.14,fill=ORANGE)
        box(s,x+0.68,yy-0.05,4.5,0.6,[para(pt,12,DARK,BFONT,ls=14)])
        yy+=0.72
lever(LEFT,"MA","Mitarbeiter","1",BLUE,
      ["PDL & examinierte Fachkräfte sofort binden (Retention)","Nachqualifikation sichert Abrechenbarkeit","Personalschlüssel nach IPReG/AKI erfüllen","Recruiting-Pipeline des Konzerns nutzen"])
lever(LEFT+5.63,"PT","Patienten","2",GREEN,
      ["Versorgungskontinuität beim Wechsel sichern","Verordnungen/Genehmigungen mit Kassen überleiten","Auslastung der Intensiv-WGs erhöhen (Marge)","Zuweiser-/Kliniknetzwerk ausbauen"])

# =====================================================================
# 13 — RISIKO-AMPEL
# =====================================================================
s=content_slide("Risiko-Ampel: worauf das Deal-Team schaut","Steuerung")
risks=[("Versorgungsvertrag geht nicht über / Abrechnungslücke","HOCH",RED),
       ("Latente Regressforderungen der Kassen","HOCH",RED),
       ("Abwanderung von PDL / Fachkräften","HOCH",RED),
       ("Personalqualifikation nicht IPReG/AKI-konform","MITTEL",ORANGE),
       ("Patientenabwanderung beim Betreiberwechsel","MITTEL",ORANGE),
       ("Belastete / geleaste Assets","MITTEL",ORANGE),
       ("Zustimmung Gläubigerausschuss verzögert","NIEDRIG",GREEN),
       ("Altverbindlichkeiten (bei sauberem Asset Deal)","NIEDRIG",GREEN)]
cw2=5.35; rh=1.05
for i,(t,lvl,c) in enumerate(risks):
    col=i//4; row=i%4
    x=LEFT+col*(cw2+0.28); y=CT+row*(rh+0.1)
    shape(s,RR,x,y,cw2,rh,fill=WHITE,line=BORDER,lw=1,radius=0.06)
    shape(s,OV,x+0.22,y+0.36,0.32,0.32,fill=c)
    box(s,x+0.72,y,cw2-2.1,rh,[para(t,11.5,DARK,BFONT,bold=True,ls=13)],anchor='m')
    shape(s,RR,x+cw2-1.3,y+0.34,1.05,0.37,fill=c,radius=0.2)
    box(s,x+cw2-1.3,y+0.34,1.05,0.37,[para(lvl,10,WHITE,HFONT,bold=True,a='c')],anchor='m')

# =====================================================================
# 14 — NÄCHSTE SCHRITTE
# =====================================================================
s=content_slide("So kommt der Prozess ins Rollen")
steps=[("1","Deal-Team & Berater aufsetzen","Fachanwalt Insolvenz- + Medizinrecht, WP, ggf. M&A-Team"),
       ("2","Target-Screening starten","Insolvenzbekanntmachungen + Netzwerk; Verfahrensart erfassen"),
       ("3","Struktur-Entscheidung treffen","Asset Deal vs. Insolvenzplan — bleibt die Kassenzulassung?"),
       ("4","DD-Checkliste & Gate-Plan finalisieren","Pflege-Prüfpunkte + Achievement-Gates als Steuerraster")]
x0=LEFT; y0=CT+0.05
for i,(n,t,d) in enumerate(steps):
    col=i%2; row=i//2
    x=x0+col*5.63; y=y0+row*1.5
    shape(s,RR,x,y,5.35,1.32,fill=WHITE,line=BORDER,lw=1,radius=0.06)
    icon_circle(s,x+0.28,y+0.38,0.58,ORANGE,n,18,tcolor="3A2A00")
    box(s,x+1.05,y+0.22,4.1,0.42,[para(t,14,BLUE,HFONT,bold=True)],anchor='m')
    box(s,x+1.05,y+0.66,4.1,0.6,[para(d,10.5,DARK,BFONT,ls=12.5)])
shape(s,RR,LEFT,y0+3.15,CW,0.7,fill=BLUETINT,radius=0.06)
box(s,LEFT+0.3,y0+3.15,CW-0.6,0.7,[para("Hinweis: Struktur- und Rechtsfragen (InsO, SGB V/XI, § 613a, Anfechtung) im Einzelfall anwaltlich absichern — dieser Leitfaden ist Orientierung, kein Rechtsrat.",10.5,GREY,BFONT,italic=True)],anchor='m')

prs.save(OUT)
print("SAVED",OUT,"slides:",len(prs.slides._sldIdLst))
