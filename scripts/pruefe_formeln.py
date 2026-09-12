#!/usr/bin/env python
"""Laesst Excel die erzeugte Mappe neu berechnen und vergleicht jede Formelzelle
mit dem Wert, den die Berechnungslogik ermittelt hat.

Noetig, weil openpyxl Formeln nur schreibt, aber nicht auswertet: ohne diesen
Schritt waere voellig offen, ob die Formeln dasselbe ergeben wie das Programm -
und eine Mappe, die etwas anderes rechnet als behauptet, waere schlimmer als gar
keine Herleitung.

Aufruf:  python scripts/pruefe_formeln.py output/Q1_herleitung.xlsx
"""
import os
import sys

import win32com.client


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    pfad = os.path.abspath(sys.argv[1])
    if not os.path.exists(pfad):
        print(f"Datei nicht gefunden: {pfad}")
        return 1

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(pfad)
        excel.CalculateFullRebuild()
        ws = wb.Worksheets("Spartenrechnung")
        kt = wb.Worksheets("Kontrolle")

        # 1. Wie viele Zellen sind ueberhaupt Formeln?
        formeln = werte = 0
        for r in range(1, 40):
            for c in range(2, 30):
                z = ws.Cells(r, c)
                if z.Value is None:
                    continue
                if str(z.Formula).startswith("="):
                    formeln += 1
                elif isinstance(z.Value, (int, float)):
                    werte += 1
        print(f"Blatt Spartenrechnung: {formeln} Formelzellen, {werte} feste Werte")

        # 2. Der Abweichungsblock auf dem Kontrollblatt - jede Zelle ist
        #    ABS(Formelergebnis - Wert der Berechnungslogik).
        start = None
        for r in range(1, 200):
            if str(kt.Cells(r, 1).Value or "").startswith("Abweichung Formel"):
                start = r + 2
                break
        if start is None:
            print("Abweichungsblock nicht gefunden")
            return 1
        maxd, geprueft, schlimmste = 0.0, 0, None
        for r in range(start, start + 40):
            key = kt.Cells(r, 1).Value
            if not key:
                break
            for c in range(2, 30):
                v = kt.Cells(r, c).Value
                if v is None or isinstance(v, str):
                    continue
                geprueft += 1
                if float(v) > maxd:
                    maxd, schlimmste = float(v), (str(key), kt.Cells(start - 1, c).Value)
        print(f"Verglichene Zellen: {geprueft}")
        print(f"Groesste Abweichung Formel ./. Berechnung: {maxd:,.4f} EUR")
        if schlimmste:
            print(f"  bei: {schlimmste[0]} / {schlimmste[1]}")

        # 3. Fehlerwerte irgendwo in der Mappe?
        fehlerhaft = []
        for blatt in wb.Worksheets:
            try:
                bereich = blatt.UsedRange.SpecialCells(-4123, 16)   # Formeln mit Fehler
                if bereich is not None and bereich.Count:
                    fehlerhaft.append((blatt.Name, bereich.Count))
            except Exception:
                pass       # SpecialCells wirft, wenn es nichts findet - das ist gut
        print("Fehlerwerte (#NV, #WERT! ...): "
              + (", ".join(f"{n}: {c}" for n, c in fehlerhaft) if fehlerhaft else "keine"))

        wb.Close(SaveChanges=False)
        return 0 if (maxd < 0.005 and not fehlerhaft) else 2
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())
