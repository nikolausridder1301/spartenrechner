#!/usr/bin/env python
"""
Erzeugt automatisiert eine Spartenrechnung (Kostentraegerrechnung je Produktgruppe)
aus dem ERP-Rohdatenexport 'KPTM_Wertsummen_*.xls' (Penta/APplus).

Kommandozeilen-Wrapper um docs/spartenrechnung_core.py (dieselbe Logik nutzt auch
die Weboberflaeche unter docs/, dort per Pyodide direkt im Browser).

Aufruf:
    python scripts/spartenrechnung.py --rohdaten rohdaten/2026/Q1 --output output/2026-Q1-Spartenrechnung.xlsx
"""
import argparse
import glob
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.join(SCRIPT_DIR, "..")
CONFIG_DIR = os.path.join(REPO_ROOT, "config")
sys.path.insert(0, os.path.join(REPO_ROOT, "docs"))

import spartenrechnung_core as core  # noqa: E402


def find_kptm_file(rohdaten_dir):
    candidates = glob.glob(os.path.join(rohdaten_dir, "KPTM_Wertsummen*.xls*"))
    if not candidates:
        raise FileNotFoundError(
            f"Keine Datei 'KPTM_Wertsummen*.xls(x)' in {rohdaten_dir} gefunden. "
            "Diese Datei ist die Basis der Spartenrechnung."
        )
    if len(candidates) > 1:
        print(f"WARNUNG: mehrere KPTM-Dateien gefunden, verwende die erste: {candidates}", file=sys.stderr)
    return candidates[0]


def main():
    ap = argparse.ArgumentParser(description="Erzeugt automatisiert eine Spartenrechnung aus ERP-Rohdaten (KPTM_Wertsummen).")
    ap.add_argument("--rohdaten", required=True, help="Ordner mit den ERP-Rohdaten-Exporten fuer den Zeitraum")
    ap.add_argument("--zeitraum", help="Nur zum Ueberschreiben. Ohne Angabe wird der Zeitraum "
                    "aus den Periodenspalten des KPTM-Exports gelesen.")
    ap.add_argument("--output", required=True, help="Pfad der zu erzeugenden Ausgabedatei (.xlsx)")
    ap.add_argument("--bwa", help="Optional: Pfad zur BWA-Arbeitsmappe fuer den automatischen GuV-Kontroll-Check")
    ap.add_argument("--bwa-sheet", help="Sheet-Name in der BWA-Datei, z.B. 'BWA 03.2026'")
    ap.add_argument("--pwbs-anfang", help="Optional: PWBS_Werkstattbestand zum PERIODENBEGINN (fuer Bestandsveraenderung UFE)")
    ap.add_argument("--pwbs-ende", help="Optional: PWBS_Werkstattbestand zum PERIODENENDE")
    ap.add_argument("--pfak", nargs="*", default=None,
                    help="Optional: eine oder mehrere PFAK-Dateien fuer die Zuordnung Auftrag -> Produktgruppe. "
                         "Mehrere angeben, inkl. der FOLGE-Zeitraeume: am Stichtag noch laufende Auftraege werden "
                         "erst spaeter fertiggemeldet und fehlen sonst in der Zuordnung.")
    args = ap.parse_args()

    kptm_path = find_kptm_file(args.rohdaten)
    print(f"Lese {kptm_path} ...")

    if args.bwa and not args.bwa_sheet:
        print("FEHLER: --bwa-sheet muss zusammen mit --bwa angegeben werden (z.B. 'BWA 03.2026').", file=sys.stderr)
        sys.exit(1)

    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    summary = core.generate(
        kptm_path, CONFIG_DIR, args.output, zeitraum=args.zeitraum,
        bwa_path=args.bwa, bwa_sheet=args.bwa_sheet,
        pwbs_anfang=args.pwbs_anfang, pwbs_ende=args.pwbs_ende, pfak_pfade=args.pfak,
    )

    print(f"{summary['zeilen']} Zeilen geladen. Zeitraum: {summary['zeitraum']}")
    if summary["unbekannte_produktgruppen"]:
        print(f"WARNUNG: unbekannte Produktgruppen im Datensatz (nicht in config/produktgruppen.json): {summary['unbekannte_produktgruppen']}", file=sys.stderr)
    if summary["unbekannte_kostenarten"]:
        # Mit Betrag, nicht nur mit Kontonummer: erst der Betrag sagt, ob hier
        # 50 EUR oder 50.000 EUR aus dem Ergebnis herausfallen.
        betraege = summary.get("unmapped_betrag") or {}
        teile = [f"{k} ({betraege.get(k, 0):,.2f} EUR)" for k in summary["unbekannte_kostenarten"]]
        summe = sum(betraege.values())
        print(f"WARNUNG: Kostenarten ohne Mapping - diese Betraege fehlen im Ergebnis: "
              f"{', '.join(teile)}; zusammen {summe:,.2f} EUR", file=sys.stderr)
    if summary.get("fracht_ohne_ziel"):
        teile = [f"{k} ({v:,.2f} EUR)" for k, v in summary["fracht_ohne_ziel"].items()]
        print(f"WARNUNG: Ausgangsfracht konnte nicht auf ein Frachtkostenobjekt umgebucht "
              f"werden und blieb auf der operativen Sparte: {', '.join(teile)}. Es fehlt das "
              f"Kostenobjekt <Familie>0100. Der Deckungsbeitrag dieser Sparte ist dadurch zu "
              f"niedrig.", file=sys.stderr)
    stat = summary.get("ufe_statistik") or {}
    for h in stat.get("hinweise", []):
        print(f"HINWEIS (Bestandsveraenderung UFE): {h}", file=sys.stderr)
    for w in summary.get("ufe_warnungen", []):
        print(f"WARNUNG (Bestandsveraenderung UFE nicht berechnet): {w}", file=sys.stderr)
    if summary.get("per_regel_zugeordnet"):
        print("HINWEIS: nach Kontenregel zugeordnet (nicht einzeln hinterlegt): "
              + ", ".join(summary["per_regel_zugeordnet"]), file=sys.stderr)
    for w in summary.get("mgk_hinweise", []):
        print(f"WARNUNG: {w}", file=sys.stderr)
    saetze = summary.get("zuschlagssaetze") or {}
    if saetze:
        teile = [f"{k.upper()} {v:.2%}" for k, v in saetze.items()]
        print("Zuschlagssaetze (Probe, gegen die eigene Kalkulation pruefen): "
              + " | ".join(teile))
    if summary.get("unproduktive_gemeinkosten"):
        # Stunden nur, wenn ein Stundensatz hinterlegt ist - sonst waere "0 h" falsch.
        stunden = summary.get("unproduktive_stunden")
        zusatz = f" ({stunden:,.0f} h)" if stunden else ""
        print(f"Nicht auf Kostentraeger gebuchte Fertigungsstunden: "
              f"{summary['unproduktive_gemeinkosten']:,.2f} EUR{zusatz} - entspricht der "
              f"FEK-Unterdeckung, Verteilung auf die Sparten unbekannt")
    if summary["bwa_ergebnis"] is not None:
        print(f"BWA-Ergebnis (kumuliert, {args.bwa_sheet}): {summary['bwa_ergebnis']:,.2f}")
    print(f"Geschrieben: {args.output}")
    print(f"\nErloese (Summe der Sparten): {summary['erloese_summe']:,.2f}")
    print(f"DB III (Summe der Sparten):  {summary['db3_summe']:,.2f}")


if __name__ == "__main__":
    main()
