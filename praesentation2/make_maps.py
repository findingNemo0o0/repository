#!/usr/bin/env python3
"""Erzeugt drei kleine Karten (Small Multiples) der Einzugsgebiete
Schulbegleitung: SKP, AWO Schleswig-Holstein, Lebenshilfe.

Geometrie: GADM-Kreisgrenzen aus deutschlandGeoJSON (isellsoap), gefiltert
auf Schleswig-Holstein und Hamburg.

Zwei Sättigungsstufen je Anbieter:
  belegt   = öffentlich nachweisbar (Anbieter-Website, Vergabe, Kreisquelle)
  Angabe   = plausibel/behauptet, nicht öffentlich belegt
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPolygon

HERE = os.path.dirname(os.path.abspath(__file__))
GEO = os.path.join(HERE, "kreisgrenzen.geojson")

BLUE = "#1A70B8"
AMBER = "#E0861A"
OLIVE = "#6E7D00"
VIOLET = "#963A96"
RED = "#C0392B"
EMPTY = "#EDEDED"
EDGE = "#FFFFFF"
LABEL = "#3D3D3D"

# Namen wie im GeoJSON (NAME_3)
ALL = ["Dithmarschen", "Flensburg Städte", "Kiel Städte", "Lauenburg",
       "Lübeck Städte", "Neumünster Städte", "Nordfriesland", "Ostholstein",
       "Pinneberg", "Plön", "Rendsburg-Eckernförde", "Schleswig-Flensburg",
       "Segeberg", "Steinburg", "Stormarn", "Hamburg Städte"]

SKP_BELEGT = ["Pinneberg"]
# Bewusst leer: Fuer weitere Einsatzgebiete gibt es keine oeffentliche Quelle.
# s-k-p.net nennt ausschliesslich den Kreis Pinneberg.
SKP_ANGABE = []

FIB_BELEGT = ["Pinneberg", "Hamburg Städte"]
FIB_ANGABE = []

FR_BELEGT = ["Pinneberg"]
FR_ANGABE = []

AWO_BELEGT = ["Lauenburg", "Pinneberg", "Lübeck Städte", "Dithmarschen",
              "Kiel Städte", "Rendsburg-Eckernförde", "Neumünster Städte",
              "Plön", "Segeberg"]
AWO_ANGABE = []

LH_BELEGT = ["Steinburg", "Segeberg", "Ostholstein", "Pinneberg"]
LH_ANGABE = ["Neumünster Städte", "Stormarn", "Lauenburg", "Nordfriesland",
             "Flensburg Städte", "Kiel Städte", "Lübeck Städte", "Plön",
             "Rendsburg-Eckernförde", "Dithmarschen"]

PANELS = [
    ("skp", BLUE, SKP_BELEGT, SKP_ANGABE),
    ("fib", VIOLET, FIB_BELEGT, FIB_ANGABE),
    ("familienraeume", RED, FR_BELEGT, FR_ANGABE),
    ("awo", AMBER, AWO_BELEGT, AWO_ANGABE),
    ("lebenshilfe", OLIVE, LH_BELEGT, LH_ANGABE),
]


def rings(feat):
    g = feat["geometry"]
    if g["type"] == "Polygon":
        return [g["coordinates"][0]]
    return [poly[0] for poly in g["coordinates"]]


def main():
    data = json.load(open(GEO))
    feats = {}
    for f in data["features"]:
        p = f["properties"]
        if p.get("NAME_1") in ("Schleswig-Holstein", "Hamburg"):
            feats[p["NAME_3"]] = f

    missing = [n for n in ALL if n not in feats]
    if missing:
        raise SystemExit(f"fehlende Geometrien: {missing}")

    for name, color, belegt, angabe in PANELS:
        fig, ax = plt.subplots(figsize=(1.20, 1.00), dpi=420)
        for kreis in ALL:
            if kreis in belegt:
                fc, alpha = color, 1.0
            elif kreis in angabe:
                fc, alpha = color, 0.32
            else:
                fc, alpha = EMPTY, 1.0
            for ring in rings(feats[kreis]):
                ax.add_patch(MplPolygon(ring, closed=True, facecolor=fc,
                                        alpha=alpha, edgecolor=EDGE,
                                        linewidth=0.4, zorder=2))
        ax.set_xlim(7.75, 11.45)
        ax.set_ylim(53.30, 55.10)
        ax.set_aspect(1 / 0.585)          # Breitengrad-Korrektur für ~54° N
        ax.axis("off")
        fig.subplots_adjust(0, 0, 1, 1)
        out = os.path.join(HERE, f"karte_{name}.png")
        fig.savefig(out, transparent=True, bbox_inches="tight", pad_inches=0.01)
        plt.close(fig)
        print("geschrieben:", out)


if __name__ == "__main__":
    main()
