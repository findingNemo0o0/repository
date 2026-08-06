# -*- coding: utf-8 -*-
"""Baut den Leitfaden 'Kauf insolventer Pflegedienste' (13 Folien) in das
Corporate-Template InsolvenzProzess.pptx (Master/Logo/Theme bleiben erhalten)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "InsolvenzProzess.pptx"
OUT = "InsolvenzProzess.pptx"

# ---- Markenpalette (aus Theme) ----
BLUE="1A70B8"; LIME="BBCF00"; GREEN="2A6A38"; ORANGE="F89929"; TEAL="0F8A8A"; RED="D73225"
GREY="777777"; WHITE="FFFFFF"; DARK="333333"
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

def add_slide(layout): return prs.slides.add_slide(layout)

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

def set_ph(sl,idx,text):
    for ph in sl.placeholders:
        if ph.placeholder_format.idx==idx:
            ph.text=text; return ph
    return None

def del_ph(sl,idxs):
    for ph in list(sl.placeholders):
        if ph.placeholder_format.idx in idxs:
            ph._element.getparent().remove(ph._element)

def content_slide(title,eyebrow=None):
    sl=add_slide(L_CONTENT)
    del_ph(sl,[1,12])
    set_ph(sl,0,title)
    return sl

def icon_circle(sl,x,y,d,fill,text,size,tcolor=WHITE):
    c=shape(sl,MSO_SHAPE.OVAL,x,y,d,d,fill=fill)
    label(sl,c,text,size,tcolor,font=HFONT,bold=True); return c

RR=MSO_SHAPE.ROUNDED_RECTANGLE; RECT=MSO_SHAPE.RECTANGLE; OV=MSO_SHAPE.OVAL; DIA=MSO_SHAPE.DIAMOND

# =====================================================================
# 1 — TITEL
# =====================================================================
s=add_slide(L_TITLE)
set_ph(s,0,"Kauf eines insolventen Pflegedienstes")
del_ph(s,[1,10])
box(s,0.71,5.15,11.6,0.5,[para("Prozess, Zeitplan & Achievement-Gates — Fokus außerklinische Intensivpflege (AKI, § 132l SGB V)",14,GREY,BFONT)])
box(s,0.71,2.95,11.0,0.7,[para("M&A-LEITFADEN · DISTRESSED HEALTHCARE",13,BLUE,HFONT,bold=True)])

# =====================================================================
# 2 — MANAGEMENT SUMMARY
# =====================================================================
s=content_slide("Warum insolvente AKI-Targets attraktiv sind")
items=[("1","Konsolidierungswelle läuft","Demografie, Personalmangel & IPReG treffen kleine Träger überproportional — 2024–2027 aktivstes Fenster.",BLUE),
       ("2","Massive Bewertungsrabatte","Distressed-Preise weit unter regulären EBITDA-Multiples; dokumentiert bis ~90 % Wertverlust.",GREEN),
       ("3","Zwei klare Hebel","In der Pflege sind Mitarbeiter und Patienten die entscheidenden — oft die einzigen — Werttreiber.",ORANGE),
       ("4","Buy-and-Build-Fit","Zukauf statt Neugründung: Patienten, Personal und (idealerweise) Versorgungsverträge statt Aufbau bei null.",TEAL)]
y=CT
for n,h,t,c in items:
    shape(s,RR,LEFT,y,6.35,1.05,fill=WHITE,line=BORDER,lw=1,radius=0.10)
    icon_circle(s,LEFT+0.2,y+0.25,0.55,c,n,16)
    box(s,LEFT+0.95,y+0.13,5.25,0.35,[para(h,13.5,c,HFONT,bold=True)])
    box(s,LEFT+0.95,y+0.5,5.3,0.5,[para(t,10.5,DARK,BFONT,ls=12.5)])
    y+=1.15
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
s=content_slide("Zwei Welten: normaler M&A vs. Distressed")
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
s=content_slide("Wer ist mein Gegenüber? Drei Verfahrensarten")
cards=[("Regelinsolvenz",BLUE,"Insolvenzverwalter",
        ["Schuldner verliert Verfügungsgewalt","Verwalter führt & verkauft die Masse","Neutral — aber kaum Gewährleistung"],"Partner: der Verwalter"),
       ("Eigenverwaltung",TEAL,"Management + Sachwalter",
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
s=content_slide("Zwei Kaufstrukturen — der Pflege-Twist entscheidet")
def deal_card(x,letter,lcolor,title,sub,rws):
    shape(s,RR,x,CT,5.35,2.55,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    icon_circle(s,x+0.22,CT+0.2,0.55,lcolor,letter,18)
    box(s,x+0.95,CT+0.18,4.2,0.38,[para(title,16,BLUE,HFONT,bold=True)])
    box(s,x+0.95,CT+0.56,4.2,0.28,[para(sub,10.5,GREY,BFONT,italic=True)])
    yy=CT+1.0
    for sign,txt in rws:
        col=GREEN if sign=="+" else RED
        c=shape(s,OV,x+0.25,yy+0.02,0.24,0.24,fill=col); label(s,c,sign,12,WHITE,BFONT,anchor='m')
        box(s,x+0.6,yy-0.02,4.6,0.34,[para(txt,11.5,DARK,BFONT)],anchor='m')
        yy+=0.37
deal_card(LEFT,"A",BLUE,"Asset Deal","„Übertragende Sanierung“ — Normalfall",
          [("+","Keine Haftung für Altverbindlichkeiten"),("+","Keine Alt-Personalkosten vor Eröffnung"),
           ("−","Verträge gehen NICHT über (außer § 613a)"),("−","Jeder Kassen-/Lieferantenvertrag neu")])
deal_card(LEFT+5.63,"B",BLUEDK,"Share Deal via Insolvenzplan","Rechtsträger überlebt entschuldet",
          [("+","Versorgungsverträge & Zulassungen bleiben"),("+","Kein Abrechnungsbruch, Verordnungen bestehen"),
           ("−","Mehr latente Haftung wird mitgekauft"),("−","Komplexer: Gläubigermehrheiten nötig")])
ty=CT+2.75
shape(s,RR,LEFT,ty,CW,1.6,fill=BLUE,radius=0.05)
shape(s,RR,LEFT,ty,2.0,1.6,fill=ORANGE,radius=0.05); shape(s,RECT,LEFT+1.5,ty,0.5,1.6,fill=ORANGE)
box(s,LEFT,ty,2.0,1.6,[para("PFLEGE-",17,"3A2A00",HFONT,bold=True,a='c'),para("TWIST",17,"3A2A00",HFONT,bold=True,a='c')],anchor='m')
box(s,LEFT+2.25,ty+0.16,CW-2.4,1.3,[
    para("Der Versorgungsvertrag (§ 132l / § 132a SGB V, § 72 SGB XI) hängt am Rechtsträger — nicht am Betrieb.",13.5,WHITE,BFONT,bold=True,ls=16,sa=3),
    para("Im Asset Deal geht er NICHT automatisch über → drohende Abrechnungslücke. Betreiberwechsel vorab mit den Kassen klären; in der Intensivpflege ist der Insolvenzplan deshalb oft überlegen.",11.5,"E9F1F8",BFONT,ls=14)],anchor='m')

# =====================================================================
# 6 — ZEITPLAN / GANTT (ab Antrag -> Closing; Milestones UNTER den Balken)
# =====================================================================
s=content_slide("Zeitplan: vom Insolvenzantrag bis zum Closing")
T0X=LEFT+2.05; TW=RIGHT-T0X; MMIN=0; MMAX=8; BARH=0.38
def tx(m): return T0X+(m-MMIN)/(MMAX-MMIN)*TW
# Krise nur als Vorlauf (nicht im Zeitplan)
box(s,LEFT,1.55,T0X-LEFT-0.12,0.3,[para("Krise / drohende ZU  →",8.5,GREY,BFONT,italic=True,a='r')],anchor='m')
shape(s,RECT,T0X,2.0,TW,0.014,fill=BORDER)
for m,lbl in [(0,"T0 · Antrag"),(3,"+3 · Eröffnung"),(6,"+6 Mon.")]:
    box(s,tx(m)+0.02,1.55,2.0,0.28,[para(lbl,8.5,GREY,HFONT,bold=True,a='l')])
for m in (3,6):
    shape(s,RECT,tx(m),2.0,0.012,3.9,fill=BORDER)
def lane_lab(y,text,color):
    box(s,LEFT,y-0.02,1.94,0.46,[para(text,8.5,color,HFONT,bold=True,ls=9.5)],anchor='m')
def seg(m1,m2,y,fill,txt=None,fs=8,tcol=WHITE):
    x=tx(m1); w=max(tx(m2)-tx(m1),0.15)
    shape(s,RR,x,y,w,BARH,fill=fill,radius=0.14)
    if txt: box(s,x+0.06,y,w-0.12,BARH,[para(txt,fs,tcol,BFONT,bold=True,a='c',ls=9)],anchor='m')
def ms(m,y,text,w=1.7):
    x=tx(m)
    shape(s,DIA,x-0.09,y+BARH+0.05,0.18,0.18,fill=ORANGE)
    lx=max(LEFT,min(x-w/2,RIGHT-w))
    box(s,lx,y+BARH+0.24,w,0.3,[para(text,7.5,DARK,BFONT,bold=True,a='c',ls=8.5)])
# Lane 1 — Insolvenzverfahren
y1=2.08
lane_lab(y1,"Insolvenzverfahren",BLUE)
seg(0,3,y1,GREEN,"Vorläufiges Verfahren · Insolvenzgeld (bis 3 Mon.)")
seg(3,8,y1,BLUEDK,"Eröffnetes Verfahren")
ms(0,y1,"Antrag (T0)"); ms(3,y1,"Eröffnung"); ms(5,y1,"Berichtstermin")
# Lane 2 — M&A-Arbeit (Käufer)
y2=3.04
lane_lab(y2,"M&A-Arbeit (Käufer)",ORANGE)
seg(0,3,y2,BLUE,"~6–10 Wochen: Sourcing · NDA · LOI · DD · Angebot")
ms(3,y2,"Signing¹")
# Lane 3 — Weg A: Regelinsolvenz -> Asset Deal
y3=4.00
lane_lab(y3,"Weg A · Regelinsolvenz → Asset Deal",BLUE)
seg(3,4.2,y3,BLUE,"Verkauf")
box(s,tx(4.5),y3+0.02,3.4,0.34,[para("schnell — Closing um die Eröffnung",8.5,GREEN,BFONT,bold=True,italic=True)],anchor='m')
ms(3.6,y3,"Gläubigerausschuss ✓ → Closing",2.1)
# Lane 4 — Weg B: Schutzschirm/Eigenverwaltung -> Insolvenzplan
y4=4.96
lane_lab(y4,"Weg B · Schutzschirm → Insolvenzplan",BLUEDK)
seg(0,3,y4,TEAL,"Schutzschirm / Planvorbereitung (max. 3 Mon.)")
seg(3,7,y4,BLUEDK,"Insolvenzplan-Verfahren")
ms(0,y4,"Bescheinigung § 270d"); ms(3,y4,"Planvorlage"); ms(5,y4,"Abstimmungstermin"); ms(7,y4,"Planbestätigung = Closing")
# Notes
box(s,LEFT,5.95,CW,0.26,[{'a':'l','r':[{'t':"Durchgehend:  ",'s':9,'c':ORANGE,'f':HFONT,'b':True},{'t':"C-Level-Mandat (ab LOI) · Kassen-/Zulassungsklärung · Personal (§ 613a)      ",'s':9,'c':DARK,'f':BFONT},{'t':"◆ = Meilenstein",'s':9,'c':GREY,'f':BFONT,'i':True}]}])
box(s,LEFT,6.21,CW,0.26,[{'a':'l','r':[{'t':"Bester Kaufzeitpunkt: im vorläufigen Verfahren (Insolvenzgeld trägt die Löhne).   ",'s':9,'c':GREEN,'f':BFONT,'b':True,'i':True},{'t':"¹ Signing unter Vorbehalt: Gläubigerausschuss + Gericht.",'s':9,'c':GREY,'f':BFONT,'i':True}]}])

# =====================================================================
# 7 — PHASEN kompakt (Tabelle)
# =====================================================================
s=content_slide("Phasen im Detail — von Sourcing bis Closing")
box(s,LEFT,1.68,CW,0.3,[{'a':'l','r':[{'t':"Durchgehend eingebunden: ",'s':10.5,'c':ORANGE,'f':HFONT,'b':True},{'t':"C-Level/Entscheider (Mandat ab LOI) · Kassen-/Zulassungsklärung · Personal (§ 613a)",'s':10.5,'c':GREY,'f':BFONT,'i':True}]}])
cA=LEFT; wA=2.1; cB=LEFT+2.15; wB=4.5; cC=LEFT+6.7; wC=1.75; cD=LEFT+8.5; wD=2.48
hy=2.06
shape(s,RR,LEFT,hy,CW,0.4,fill=BLUE,radius=0.06)
for cx,cw,t in [(cA,wA,"Phase"),(cB,wB,"Ablauf"),(cC,wC,"Gate / Ergebnis"),(cD,wD,"Wer / extern")]:
    box(s,cx+0.12,hy,cw-0.15,0.4,[para(t,10.5,WHITE,HFONT,bold=True)],anchor='m')
tbl=[("0  Sourcing & Ansprache","Targets via Insolvenzbekanntmachungen & Netzwerk; Verwalter/Berater kontaktieren","Target qualifiziert","intern + Verwalter"),
     ("1  NDA & Datenraum","NDA; Zugang zum (Verwalter-)Datenraum; Verfahrensart & Partner klären","Datenraum-Zugang","Verwalter"),
     ("2  Indik. Angebot / LOI","Info-Memo prüfen; nicht-bindendes Angebot; internes Mandat einholen","LOI","C-Level"),
     ("3  DD & Bewertung","§ 132l, Regress, MD, Personal, Patienten; Bewertung + Konzept","DD abgeschlossen","WP / Anwalt / M&A"),
     ("4  Verbindl. Angebot & Verhandlung","Binding Offer; Preis, Bedingungen, § 613a, Betreiberwechsel/Zulassung","Einigung","Verwalter / Kassen"),
     ("5  Signing (unter Vorbehalt)","Vertrag unter aufschiebender Bedingung: Gläubigerausschuss + Gericht","Signing","Gericht / Gläubiger"),
     ("6  Closing & Integration","Vollzug; Übergang Patienten & Personal; Konzernanbindung","Closing","intern")]
ry=hy+0.42; rh=0.57
for i,(a,b,c,d) in enumerate(tbl):
    if i%2==0: shape(s,RECT,LEFT,ry,CW,rh,fill=GREYTINT)
    box(s,cA+0.1,ry,wA-0.1,rh,[para(a,9.5,BLUE,BFONT,bold=True,ls=11)],anchor='m')
    box(s,cB+0.12,ry,wB-0.2,rh,[para(b,9,DARK,BFONT,ls=10.5)],anchor='m')
    box(s,cC+0.12,ry,wC-0.15,rh,[para(c,9,"8A5A10",BFONT,bold=True,ls=10.5)],anchor='m')
    box(s,cD+0.12,ry,wD-0.2,rh,[para(d,9,GREY,BFONT,ls=10.5)],anchor='m')
    ry+=rh

# =====================================================================
# 8 — DUE DILIGENCE & ROTE FLAGGEN (zusammengeführt aus alt 9+10)
# =====================================================================
s=content_slide("Due Diligence: Prüffokus & rote Flaggen")
LX=LEFT; RX=LEFT+5.63; colw=5.3
box(s,LX,1.72,colw,0.28,[para("PRÜFFOKUS (PFLEGE-DD)",11,BLUE,HFONT,bold=True)])
box(s,RX,1.72,colw,0.28,[para("ROTE FLAGGEN",11,RED,HFONT,bold=True)])
ddL=[("Versorgungsvertrag & Zulassung","§ 132l/§ 132a SGB V, § 72 SGB XI — gültig & übertragbar?",BLUE),
     ("Vergütung (Anlage 2)","nachrechnen: tragfähig, marktüblich, verhandelbar?",GREEN),
     ("MD-Prüfung & Ruf","Qualitätsberichte, laufende Prüfverfahren, Mängel",TEAL),
     ("Abrechnung & Regress","Nachweise vs. Abrechnung; latente Rückforderungen",ORANGE),
     ("Personalqualifikation & PDL","examinierte Fachkräfte? fachgerecht erbracht?",BLUE),
     ("Patientenstruktur","Fallzahlen vs. Verordnungen; Loyalität beim Wechsel",GREEN)]
ddR=[("Scheinpatienten / Karteileichen","Fallzahlen gegen Verordnungen & Abrechnungen prüfen"),
     ("Nicht erbrachte Leistungen","Regress der Kassen + § 263 StGB — im Share Deal geerbt"),
     ("Qualifikations-Fake / Key-Person","unqualifiziert als Fachleistung; PDL wandert ab"),
     ("Assets nicht frei","Sicherungsübereignung, Eigentumsvorbehalt, Leasing"),
     ("Geschönte Zahlen / Cherry-Picking","verschwiegene Verluste; „gute“ Verträge ausgegliedert")]
def itemrow(x,y,w,color,title,note):
    shape(s,OV,x,y+0.06,0.15,0.15,fill=color)
    box(s,x+0.28,y-0.02,w-0.3,0.54,[{'a':'l','ls':11,'r':[{'t':title+"  ",'s':10.5,'c':color,'f':HFONT,'b':True},{'t':note,'s':9,'c':DARK,'f':BFONT}]}])
yy=2.06
for t,n,c in ddL: itemrow(LX,yy,colw,c,t,n); yy+=0.6
yy=2.06
for t,n in ddR: itemrow(RX,yy,colw,RED,t,n); yy+=0.6
shape(s,RR,LEFT,5.9,CW,0.55,fill=BLUE,radius=0.06)
box(s,LEFT+0.25,5.9,CW-0.5,0.55,[{'a':'l','ls':12,'r':[{'t':"Gegenmittel:  ",'s':11,'c':LIME,'f':HFONT,'b':True},{'t':"Umsatz gegen dokumentierte Leistungen · Zulassung/§ 132l bei den Kassen verifizieren · PDL & Fachkräfte binden.",'s':10.5,'c':WHITE,'f':BFONT}]}],anchor='m')

# =====================================================================
# 9 — KAUFPREISLOGIK (Asset Deal vs. Insolvenzplan / Schutzschirm)
# =====================================================================
s=content_slide("Kaufpreislogik: Asset Deal vs. Insolvenzplan")
def logic_card(x,color,title,sub,lines,note):
    shape(s,RR,x,CT,5.3,3.9,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    shape(s,RR,x,CT,5.3,0.7,fill=color,radius=0.07); shape(s,RECT,x,CT+0.35,5.3,0.35,fill=color)
    box(s,x+0.2,CT+0.04,4.9,0.62,[para(title,13.5,WHITE,HFONT,bold=True,ls=15)],anchor='m')
    box(s,x+0.25,CT+0.78,4.85,0.3,[para(sub,10.5,color,BFONT,bold=True,italic=True)])
    yy=CT+1.15
    for sym,tx in lines:
        cc = GREEN if sym=="+" else (RED if sym=="−" else (BLUE if sym=="=" else GREY))
        if sym: box(s,x+0.25,yy,0.35,0.32,[para(sym,13,cc,HFONT,bold=True)])
        box(s,x+0.62,yy,4.5,0.36,[para(tx,10,DARK,BFONT,ls=11.5)],anchor='t')
        yy+=0.42
    shape(s,RR,x+0.2,CT+3.32,4.9,0.44,fill=GREYTINT,radius=0.12)
    box(s,x+0.3,CT+3.32,4.7,0.44,[para(note,9.5,color,BFONT,bold=True,italic=True,ls=11)],anchor='m')
logic_card(LEFT,BLUE,"Asset Deal — übertragende Sanierung","substanz-/masseorientiert",
   [("","Substanz: unbelastete Assets"),("+","Going-Concern (Patienten, Personal, Zulassung)"),
    ("−","Sanierungskosten (Personal, Zulassung)"),("−","Risikoabschläge (Regress, keine Garantien)"),
    ("=","risikoadjustierter Angebotspreis")],
   "Kein EBITDA-Multiple — zerschlagungsnah.")
logic_card(LEFT+5.68,BLUEDK,"Insolvenzplan — Eigenverwaltung / Schutzschirm","going-concern-/ertragswertorientiert",
   [("","Kein Zerschlagungspreis, sondern Planbeitrag:"),("+","frisches Geld / Debt-to-Equity"),
    ("+","Going Concern (Ertragswert / DCF)"),("−","Sanierungsbedarf & Planquote an Gläubiger"),
    ("=","Rechtsträger + Zulassung bleiben")],
   "Schutzschirm = Selbstsanierung, oft OHNE klassischen Kaufpreis.")
shape(s,RR,LEFT,CT+4.02,CW,0.55,fill=ORANGETINT,line=ORANGE,lw=1,radius=0.10)
box(s,LEFT+0.3,CT+4.02,CW-0.6,0.55,[{'a':'l','ls':12,'r':[{'t':"Merke:  ",'s':11,'c':"8A5A10",'f':HFONT,'b':True},{'t':"Schutzschirm/Eigenverwaltung zielt auf den Insolvenzplan → Preislogik = Going-Concern, nicht Zerschlagung.",'s':10.5,'c':DARK,'f':BFONT}]}],anchor='m')

# =====================================================================
# 10 — WAS PASSIERT MIT DEN SCHULDEN?
# =====================================================================
s=content_slide("Was passiert mit den Schulden?")
def debt_card(x,color,title,pts):
    shape(s,RR,x,CT,5.3,2.7,fill=WHITE,line=BORDER,lw=1,radius=0.05)
    shape(s,RR,x,CT,5.3,0.62,fill=color,radius=0.08); shape(s,RECT,x,CT+0.31,5.3,0.31,fill=color)
    box(s,x+0.2,CT,4.9,0.62,[para(title,14,WHITE,HFONT,bold=True)],anchor='m')
    yy=CT+0.78
    for p in pts:
        shape(s,OV,x+0.25,yy+0.05,0.13,0.13,fill=color)
        box(s,x+0.52,yy-0.02,4.6,0.44,[para(p,10.5,DARK,BFONT,ls=12)])
        yy+=0.47
debt_card(LEFT,BLUE,"Asset Deal",
   ["Schulden bleiben im insolventen Rechtsträger (Hülle)","Werden aus dem Verkaufserlös nach Quote bedient","Hülle wird liquidiert","Käufer erhält Assets schuldenfrei"])
debt_card(LEFT+5.68,BLUEDK,"Insolvenzplan (auch Schutzschirm)",
   ["Rechtsträger überlebt, Schulden werden restrukturiert","Gläubiger erhalten Quote (Cashflow + Planbeitrag)","Rest wird erlassen (§ 227 InsO) → Firma entschuldet","Zulassung & Verträge bleiben am Träger"])
by=CT+2.9
shape(s,RR,LEFT,by,CW,1.35,fill=GREYTINT,radius=0.06)
box(s,LEFT+0.25,by+0.12,CW-0.5,0.3,[para("WICHTIG — RANGFOLGE & RANDPUNKTE",10.5,RED,HFONT,bold=True)])
rl=[("1","Zuerst bedient: gesicherte Gläubiger (Absonderung) & Massegläubiger"),
    ("2","Dann: ungesicherte Gläubiger — reale Quote oft nur 3–5 %"),
    ("3","Alt-Gesellschafter verlieren Anteile (§ 225a); Löhne bis 3 Mon. via Insolvenzgeld")]
yy=by+0.46
for n,t in rl:
    box(s,LEFT+0.25,yy,0.3,0.28,[para(n+".",10.5,RED,HFONT,bold=True)])
    box(s,LEFT+0.62,yy,CW-1.0,0.3,[para(t,10.5,DARK,BFONT)])
    yy+=0.29

# =====================================================================
# 11 — INTEGRATION (2 Hebel)
# =====================================================================
s=content_slide("Integration & Skalierung — die zwei Hebel")
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
# 12 — RISIKO-AMPEL
# =====================================================================
s=content_slide("Risiko-Ampel: worauf das Deal-Team schaut")
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
# 13 — NÄCHSTE SCHRITTE
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
