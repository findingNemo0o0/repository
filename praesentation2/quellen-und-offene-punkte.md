# Vier-Folien-Präsentation — Quellen und offene Punkte

**Deck:** `SKP-Poolreform-Modelle-und-Financial-Model.pptx` · 4 Folien · Sprechernotizen in jeder Folie
**Struktur:** nach `plan.docx` · **Design:** Haus-Standard aus `design-template.pptx`

---

## Folie für Folie — woher die Zahlen kommen

### Folie 1 — Alle Modelle

| Aussage | Beleg |
|---|---|
| Pool + MPT ist der Zielzustand | Landtag SH, **Drs. 20/3271 vom 11.06.2025**, Antwort auf Frage 3: Die Zusammenführung sei „ein wichtiger Baustein der multiprofessionellen Zusammenarbeit an Schulen" |
| Kein Abschlusstermin | ebd.: „kann ein Zeitpunkt für den Prozessabschluss nicht genannt werden" |
| Rolle der beiden Stiftungen | Drs. 20/3271: Die bestehenden Poolmodelle würden „durch die Deutsche Telekom Stiftung sowie die Robert Bosch Stiftung **bewertet und die Weiterentwicklung unterstützt**". Die Ministerin (Rede 28.02.2025) beschreibt sie als „**engagierte Partner**", die „verschiedene Akteure zusammenbringen, den Prozess unterstützen und neue Wege der rechtskreisübergreifenden Zusammenarbeit für das gesamte Land entwickeln" — gemeinsam mit den kommunalen Landesverbänden, Kreisen, kreisfreien Städten, Schulträgern, Schulen und Schulaufsicht. **Prozesspartner, nicht externe Gutachter** |
| Testträger Kreis Pinneberg: **Familienräume Karin Struckmeier**, 7 Grundschulen Tornesch-Uetersen, Vertrag 26.06.2018 | Kreis Pinneberg, PM „Neues Modell für Schulbegleitungen wird eingeführt"; familienraeume.de |
| Ostholstein: 7 Modellschulen seit 11/2020 — **drei Einzelschulen** (Neustädter Bucht, Bad Schwartau, Schönwalde) und die **Modellregion Fehmarn mit vier Schulen** (u. a. Inselschule, Grundschule Burg, Montessori-Schule). **Je Schule ein eigener Pool** | DISW-Gesamtevaluation 20.11.2023, S. 5, 18, 21 f., 72 f.; Landtag SH Drs. 20/2643(neu) |
| Lübeck seit 2013/14, Flensburg seit 2022/23 flächendeckend | Drs. 20/2643(neu) |
| Hamburg: Kombinationsmaßnahmen ab 2026/27 | BSFB Hamburg |
| EU-Ausschreibung „Klassenassistenz" ohne Zuschlag | **TED 246925-2024** (25.04.2024). Eigene TED-Recherche über alle Bekanntmachungen des Kreises Pinneberg mit CPV 85xxx, Stand 04.08.2026: **keine Zuschlagsbekanntmachung** |
| Kreistag verschiebt am 11.12.2024 um zwei Jahre | Kreis Pinneberg, Seite „Modell Klassenassistenz" |
| Wer sonst gewinnt | TED-Zuschläge Kreis/Stadt Pinneberg 2024–2026: DHB gGmbH, Lebenshilfe gGmbH, AWO SH gGmbH, inab, Diakonisches Werk. NRW: lebe!zeit GmbH & Co. KG, Inclusio gGmbH |

**Karte.** Kreisgrenzen aus GADM über `deutschlandGeoJSON`, gerendert mit `make_maps.py`.
Zwei Sättigungsstufen: **kräftig = öffentlich belegt**, **blass = Angabe bzw. Verband vor Ort**.
Fünf Anbieter im Vergleich:

- **SKP:** **nur Kreis Pinneberg.** `s-k-p.net` nennt als Einsatzgebiet ausschließlich
  „Kreis Pinneberg"; die Firma heißt „SKP – Soziale Kompetenzen Kreis Pinneberg e. K.",
  Sitz Uetersen.

  > ⚠️ **Korrektur.** In den ersten Fassungen der Karte waren zusätzlich Steinburg, Segeberg,
  > Ostholstein, Neumünster und Hamburg blass eingefärbt und als „Angabe des Unternehmens"
  > beschriftet. Diese fünf Kreise stammten aus der **Aufgabenstellung zu Beginn dieses
  > Projekts**, nicht aus einer Quelle, auf die man in einer Due Diligence verweisen kann —
  > SKP selbst nennt sie nirgends. Sie sind aus der Karte entfernt. Ob SKP tatsächlich über
  > den Kreis Pinneberg hinaus tätig ist, ist über die Standortliste aus dem Datenraum zu
  > klären (Frage 5 auf Folie 3).
- **FiB:** „**Familien im Blick.Pinneberg GmbH**", gemeinnützig, Sitz Heinrich-Christiansen-Str. 43,
  25421 Pinneberg. Laut `fib-pinneberg.de` tätig **im Kreis Pinneberg und in Hamburg** —
  namentlich Appen, Elmshorn, Moorrege, Pinneberg, Quickborn, Rellingen, Schenefeld,
  Tangstedt, Tornesch, Uetersen und Wedel. Leistungen: Schulbegleitung und Integrationshilfe,
  sozialpädagogische Familienhilfe, Diabetesassistenz, Arbeitsassistenz. Nur ein Standort,
  keine Zweigbüros.
- **Familienräume:** „**Familienräume K. Struckmeier GmbH**", Sitz Pinneberg. Laut
  `familienraeume.de` im **gesamten Kreis Pinneberg** tätig; im Poolmodell der Region
  **Tornesch/Uetersen** ist sie der ausführende Leistungserbringer (Vertrag mit dem Kreis
  vom 26.06.2018).
- **AWO SH:** Quelle ist die Standortliste auf `awo-sh.de/schulische-unterstuetzungsangebote/schulbegleitung`.
  Sie nennt namentlich Bereichs-, Einrichtungs- und Teamleitungen für: Herzogtum Lauenburg
  Nord und Süd (Jahan Mortezai), Kreis Pinneberg (Anna Kindler), Lübeck (Julia Spiegel),
  Heide/Dithmarschen (Stefan Frahm) sowie Kiel, Rendsburg-Eckernförde, Neumünster, Plön und
  Bad Segeberg (Susanne Friederich) — zusammen neun Kreise und kreisfreie Städte.
- **Lebenshilfe:** blass eingefärbt sind hier — und nur hier — Kreise, in denen ein
  Kreisverband existiert, für die die Schulbegleitung aber nicht einzeln geprüft ist.
  Belegt sind die eigenen Schulbegleitungs-Angebotsseiten der Kreisverbände Steinburg
  (`lebenshilfe-steinburg.de`), Bad Segeberg (`lebenshilfe-segeberg.de`), Ostholstein
  (Bad Schwartau) und Pinneberg (`lebenshilfe-pi.de`, Elmshorn). Die Mitgliederliste des
  Landesverbands (`lebenshilfe-sh.de/verband/mitglieder-lhsh`) zeigt Kreisverbände fast
  landesweit — dort ist die Schulbegleitung aber nicht einzeln geprüft, daher blass.

> Der Hinweis „von Fehmarn bis Hamburg", der in Jobbörsen neben SKP auftaucht, gehört zu
> einer Anzeige des **Malteser Hilfsdienstes** — nicht zu SKP.

> 🔑 **Das Bild, das die fünf Karten ergeben:** SKP, FiB und Familienräume sitzen alle im
> **Kreis Pinneberg** — also genau dort, wo mit der „Klassenassistenz" die schärfste Reform
> Schleswig-Holsteins vorbereitet wird. FiB ist zusätzlich in Hamburg tätig, wo seit 2026/27
> die Kombinationsmaßnahmen laufen. Beide Heimatmärkte stellen also gleichzeitig um. AWO und
> Lebenshilfe sind dagegen landesweit aufgestellt und können den Verlust eines Standorts
> anderswo ausgleichen.

### Folie 2 — Was sich verändert

| Aussage | Beleg |
|---|---|
| Rund die Hälfte der Schulbegleitungen ist fachlich qualifiziert | Dworschak u. a.; QfI, „Kann das jeder?" |
| Kein Fachkräftegebot für Schulbegleitung | § 72 SGB VIII erfasst sie nicht |
| Teamleitung ausschließlich Fachkraft, Schlüssel 1:15, 39 Wochenstunden | Konzept „Klassenassistenz" Kreis Pinneberg, Kap. 6 |
| Keine Honorarkräfte | ebd., Kap. 4: „Der Einsatz von angestellten Mitarbeitern ist ein Qualitätsmerkmal" |
| Vertrag 01.11.2024–31.07.2029, einseitig +4 Jahre | TED 246925-2024 |
| Kein Tarifausgleich | Konzept, Kap. 6: „Anpassungen durch Tarifsteigerungen … erfolgen im Rahmen der Vertragslaufzeit nicht" |
| Grüne: „Wir sind klar für jede Poollösung!" | Malte Krüger, Landtag SH, Presseticker 28.02.2025 |
| CDU-Antrag „systemisch weiterentwickeln" | CDU-Landtagsfraktion SH, TOP 27+40 |
| **SPD in SH drängt auf Tempo, nicht auf Rücknahme** | Drs. 20/3271, Frage 3 stammt von Martin Habersaat (SPD): „**Zu wann** beabsichtigt die Landesregierung … Schulassistenz und Schulbegleitung zu Poollösungen zusammenzuführen?" |
| **SPD im Bund gegen die Streichliste** | Die SPD-Bundestagsabgeordneten Jasmina Hostert, Annika Klose und Heike Heubach nannten die Streichliste am **17.04.2026** „inakzeptabel" — „Wer bei Teilhabe und Jugendhilfe spart, spart am falschen Ende". Über 100.000 Unterschriften unter der Bundestagspetition dagegen |
| Land: CDU und Grüne, 48 von 69 Sitzen | Landtag SH, 20. Wahlperiode: CDU 34, Grüne 14; Opposition SPD 12, FDP 5, SSW 4. MP Daniel Günther (CDU), Bildungsministerin Dr. Dorit Stenke (CDU) seit Mai 2025. Nächste Wahl spätestens 18.04.2027 |
| Kreis Pinneberg: CDU stärkste Fraktion, 24 von 67 | Kreistagswahl 14.05.2023, CDU 35,7 %. Kreispräsident Helmuth Ahrens (CDU); Landrätin **Elfi Heesch (parteilos)**, seit 01.01.2021 vom Kreistag gewählt |
| Bund: Entwurf aus dem Haus von Karin Prien | Prien (CDU) war 2017–2025 Bildungsministerin in SH und leitet seit 07.05.2025 das BMBFSFJ, das den Referentenentwurf vorgelegt hat |
| ver.di-Kritik: Fachkräftegebot → Kompetenzansatz | ver.di, „Exklusion statt Inklusion", 24.04.2026 |
| Lebenshilfe Bundesverband, AGJ, SoVD, Verfassungsblog | Stellungnahmen 04/2026; Wrase, Verfassungsblog 28.05.2026 |
| OVG Schleswig weist Klage von vier Trägern ab | **5 MB 7/24**, Beschluss 03.09.2024, unanfechtbar |

### Folie 3 — Fragenliste

Alle Messlatten stammen aus den **Eignungs- und Zuschlagskriterien der TED-Bekanntmachung
246925-2024** und aus Kap. 4–6 des Pinneberger Konzepts: 10 Mio. € Betriebshaftpflicht,
Referenzliste über vier Jahre, getrennte Umsatzausweisung Schulbegleitung, PQ-Eintrag
(DIHK/pq-vol.de), sieben Zuschlagskriterien (darunter Kommunikation mit der Schulleitung und
verpflichtende Zusammenarbeit mit Kooperationspartnern), Schutzauftrag nach §§ 8a/8b/72a SGB VIII.

### Folie 4 — Financial Model

| Größe | Wert | Quelle |
|---|---|---|
| Mindestlohn 2026 / 2027 | 13,90 € / 14,60 € | Mindestlohnkommission, BMAS |
| TVöD SuE **S2 Stufe 5** | 3.330,92 €/Monat | Tabelle gültig 01.05.2026–31.03.2027 (VKA), gegengeprüft an zwei Quellen |
| TVöD SuE **S3 Stufe 5** | 3.755,52 €/Monat | ebd. |
| Umrechnung auf Stundensatz | ÷ 169,6 h (39 Wochenstunden × 4,348) | **eigene Rechnung** → 19,64 € bzw. 22,15 €; Abstand zum Mindestlohn 2026 rund **+41 %** |
| Schulbegleitungen SH | 2.700 (2014) → rund 7.000 (2022), rund **126 Mio. €** Kreisausgaben | Rede der Ministerin, Landtag SH, 28.02.2025 |
| Schulbegleitungen Hamburg | 1.574 (2014/15) → **4.011 (2025/26)**, 6,75 → **42,15 Mio. €** | BSFB Hamburg, PM **26.06.2026** — die aktuellste öffentlich vorliegende Reihe |
| Eingliederungshilfe bundesweit | 2024: **324.570 Kinder unter 18** (31,5 % aller Leistungsberechtigten), Ausgaben **+12,9 % auf 28,7 Mrd. €** | Destatis, Eingliederungshilfe nach SGB IX, Berichtsjahr 2024 |
| Vergütungsformel | S2/5 + 10 % Sachkosten + 5 % Verwaltungskostenzuschlag | Konzept, Kap. 6 |
| Personalschlüssel | 1 Klassenassistenz je Klasse, 24 h (Kl. 1–2) / 30 h (Kl. 3–4), Teamleitung 1:15 | ebd. |

> **TVöD SuE** = Tarifvertrag für den öffentlichen Dienst, Sparte **S**ozial- **u**nd
> **E**rziehungsdienst (kommunale Arbeitgeber, VKA). **S2** ist die unterste Entgeltgruppe
> für Helfertätigkeiten in Erziehung, Pflege und Betreuung; die **Stufe** steht für
> Berufserfahrung. Der Stundensatz enthält **keine** Arbeitgeberanteile und keine SuE-Zulage —
> beides erhöht die tatsächlichen Personalkosten weiter.

> ⚠️ **Korrektur gegenüber der ersten Fassung:** Dort standen 3.997,10 € (S2/5) und
> 4.506,62 € (S3/5) und daraus abgeleitet ein Abstand von 70 % zum Mindestlohn. Die Werte
> waren falsch. Richtig sind 3.330,92 € und 3.755,52 €, der Abstand beträgt rund **41 %**.

---

## Was wir nicht belegen konnten

1. **Der für Q1/2026 angekündigte Sachstandsbericht der Landesregierung** (zugesagt in
   Drs. 20/3271) war in den Landtagsdrucksachen nicht auffindbar. Vor dem Termin nachfassen —
   er ist die aktuellste Quelle zum Stand der Zusammenführung.
1a. **Ein Ergebnis der Stiftungsbegleitung ist nicht veröffentlicht.** Gesucht wurde auf
   `bosch-stiftung.de`, `telekom-stiftung.de`, `deutsches-schulportal.de`, im Jugendhilfeportal
   und in den Landtagsdrucksachen: **keine Studie, kein Gutachten, keine Handreichung, keine
   Empfehlung.** Die Stiftungen moderieren den Landesprozess, sie begutachten ihn nicht von
   außen. Wer den Zwischenstand kennen will, muss beim Ministerium oder bei den Stiftungen
   direkt fragen — das ist einer der wenigen Punkte, an denen ein Informationsvorsprung
   tatsächlich zu holen wäre.
2. **Ausgang des Pinneberger Vergabeverfahrens.** Belegt ist nur: keine veröffentlichte
   Zuschlagsbekanntmachung. Ob aufgehoben, ausgesetzt oder ruhend, muss beim Kreis erfragt
   werden (klassenassistenz@kreis-pinneberg.de).
3. **SKPs tatsächliche Einsatzgebiete, Fallzahlen und Umsätze.** Kreisscharfe Marktzahlen zur
   Schulbegleitung sind nicht öffentlich; SKP veröffentlicht keine Kennzahlen. Öffentlich
   belegbar ist ausschließlich der Kreis Pinneberg. Das ist der Grund für die Fragenliste auf
   Folie 3 — insbesondere Frage 5 (Standortliste je Schule).
4. **Die Testträger der Ostholsteiner Modellschulen** namentlich. Die Schulen sind inzwischen
   belegt (siehe Folie 1), die **Leistungserbringer nennt der DISW-Bericht aber bewusst nicht** —
   er spricht durchgängig nur von „dem Leistungserbringer". Die Pressemitteilung des Kreises
   Ostholstein bleibt aus dieser Umgebung nicht abrufbar (Verbindungsabbruch, vermutlich
   Geoblocking). Beim Kreis direkt anfragen.

> **Nachtrag zum Sachstandsbericht (Punkt 1).** Der über `ker-leipzig.de` verlinkte
> „Sachstandsbericht Schulbegleitung" vom 24.06.2022 stammt vom **Amt für Jugend und Familie
> der Stadt Leipzig** an deren Jugendhilfeausschuss (506 Fälle, davon 227 an Grundschulen)
> — er betrifft **Sachsen ohne Poolmodell** und ist nicht der hier gesuchte
> Sachstandsbericht der Landesregierung Schleswig-Holstein.

---

## Dateien

| Datei | Zweck |
|---|---|
| `build_deck.py` | erzeugt das Deck |
| `make_maps.py` | erzeugt die drei Karten aus GeoJSON |
| `check_layout.py` | prüft Rahmen, Textüberlauf und Überlappungen |
| `design-template.pptx` | Haus-Design ohne Folien und ohne Inhalte |
| `plan.docx` | Strukturvorgabe |
| `build_fragenliste.js` | erzeugt die Fragenliste für das Gespräch mit der Eigentümerin |
| `Fragenliste-Eigentuemerin-SKP.docx` | 38 Fragen in acht Blöcken, mit Spalte für Antworten |

Nachbauen: `python3 make_maps.py && python3 build_deck.py && python3 check_layout.py <deck>`
Die Kreisgrenzen liegen als `kreisgrenzen.geojson` im Ordner.
