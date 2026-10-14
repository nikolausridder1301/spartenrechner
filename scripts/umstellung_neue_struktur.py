#!/usr/bin/env python
"""Einmalige Umstellung der Betriebsparameter auf die neue Produktgruppen-Struktur.

Die Produktgruppen wurden 2026 reorganisiert (vor allem Pressbohrtechnik/HDD). Damit
die Spartenrechnung das Jahr durchgaengig in der NEUEN Struktur zeigt, braucht sie:

  1. die Umschluesselung alt -> neu je Artikel (aus der Reorganisations-Datei), um
     Buchungen unter alten Gruppen umzuschluesseln, und
  2. den Anfangsbestand zum 01.01. je NEUER Gruppe. Gepflegt war er je alter Gruppe;
     er wird hier anteilig aufgeteilt (Schluessel: Werkstattbestand zum 01.01. je
     Auftrag). Die Summe je alter Gruppe bleibt centgenau erhalten.

Beides wird in die Betriebsparameter-Datei geschrieben - die liegt im privaten Teil
(intern/), weil Anfangsbestaende Firmenzahlen sind. Die bisherige Datei bleibt in der
Versionsgeschichte des privaten Repositorys erhalten.

Aufruf:
    python scripts/umstellung_neue_struktur.py \\
        --reorg "rohdaten/2026/Q2/Alte vs Neue Produktgruppen Horizontalbohrtechnik.xlsx" \\
        --werkstattbestand-0101 rohdaten/2026/PWBS_mit_Produktgruppe/PWBS_Werkstattbestand_010126_Produktgruppe.xls \\
        --jahr 2026
"""
import argparse
import json
import os
import sys

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(SCRIPT_DIR, "..")
sys.path.insert(0, os.path.join(REPO_ROOT, "docs"))
import spartenrechnung_core as core  # noqa: E402


def lade_reorganisation(pfad):
    """Artikel -> [alte, neue Produktgruppe] aus der Reorganisations-Datei.

    Nur gueltige Artikel tragen eine neue Gruppe; ungueltige (ausgelaufene) werden
    nicht umgeschluesselt und bleiben in ihrer Gruppe."""
    df = pd.read_excel(pfad, sheet_name="Gültige_Artikel")
    sp_neu = next(c for c in df.columns if "NEU" in str(c).upper())
    karte = {}
    for art, alt, neu in zip(df["ARTIKEL_NR"], df["PRODUKTGRUPPE"], df[sp_neu]):
        a, alt_, neu_ = core._schluessel(art), core._schluessel(alt), core._schluessel(neu)
        if a and alt_ and neu_ and alt_ != neu_:
            if a in karte and karte[a] != [alt_, neu_]:
                raise ValueError(f"Artikel {a} widerspruechlich umgeschluesselt: {karte[a]} / {[alt_, neu_]}")
            karte[a] = [alt_, neu_]
    return karte


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--reorg", required=True, help="Reorganisations-Datei (Blatt 'Gültige_Artikel')")
    ap.add_argument("--werkstattbestand-0101", required=True,
                    help="Werkstattbestand zum 01.01. MIT Spalte Produktgruppe (Schluessel der Aufteilung)")
    ap.add_argument("--jahr", required=True, help="Geschaeftsjahr, dessen Anfangsbestand umgestellt wird")
    ap.add_argument("--parameter", default=os.path.join(REPO_ROOT, "intern", "betriebsparameter.xlsx"))
    args = ap.parse_args()

    with open(os.path.join(REPO_ROOT, "config", "produktgruppen.json"), encoding="utf-8") as f:
        aufgeloest = json.load(f).get("aufgeloest", {})

    alt = core.betriebsparameter_aus_excel(args.parameter)
    ab_alt = alt["anfangsbestand"].get(args.jahr)
    if not ab_alt:
        sys.exit(f"Kein Anfangsbestand {args.jahr} in {args.parameter}")

    karte = lade_reorganisation(args.reorg)
    ab_neu, anteile, naeherung = core.anfangsbestand_umschluesseln(
        ab_alt, args.werkstattbestand_0101, karte, aufgeloest)

    summe_alt, summe_neu = sum(ab_alt.values()), sum(ab_neu.values())
    if abs(summe_alt - summe_neu) > 0.005:
        sys.exit(f"Summe nicht erhalten: {summe_alt:.2f} -> {summe_neu:.2f}")

    print(f"Umschluesselung: {len(karte)} Artikel")
    print(f"Anfangsbestand {args.jahr}: {len(ab_alt)} alte -> {len(ab_neu)} neue Gruppen, "
          f"Summe unveraendert")
    for alt_pg, verteilung in sorted(anteile.items()):
        if set(verteilung) != {alt_pg}:
            print(f"   {alt_pg} -> " + ", ".join(f"{n} {a:.1%}" for n, a in
                                              sorted(verteilung.items(), key=lambda x: -x[1])))

    jahre = dict(alt["anfangsbestand"])
    jahre[args.jahr] = ab_neu
    core.schreibe_betriebsparameter_excel(jahre, args.parameter, parameter=alt.get("parameter"),
                                          umschluesselung=karte,
                                          naeherung={args.jahr: naeherung})
    print(f"Anteilig aufgeteilt (Naeherung): {', '.join(sorted(naeherung)) or '-'}")
    print(f"Geschrieben: {args.parameter}")


if __name__ == "__main__":
    main()
