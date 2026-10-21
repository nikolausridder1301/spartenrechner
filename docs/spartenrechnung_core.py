"""
Kernlogik der Spartenrechnung - gemeinsam genutzt von der CLI (scripts/spartenrechnung.py)
und der Browser-Weboberflaeche (docs/app.js via Pyodide).

Bewusst ohne argparse/CLI-Code, damit dieses Modul unveraendert sowohl per
Kommandozeile importiert als auch im Browser (Pyodide) geladen werden kann.
Alle Funktionen arbeiten mit Datei-PFADEN (kompatibel mit Pyodides virtuellem
Dateisystem - im Browser schreibt JavaScript die hochgeladenen Dateien vorher
per pyodide.FS an genau diese Pfade, ohne dass irgendetwas das Geraet des
Nutzers verlaesst).
"""
import json
import os
import re
import warnings

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.utils import get_column_letter

FILL_AUTO = PatternFill("solid", fgColor="D9EAD3")       # gruen: automatisch, gut rekonziliert
FILL_AUTO_PRUEF = PatternFill("solid", fgColor="FFF2CC")  # gelb: automatisch, aber Naeherung/pruefen
FILL_MANUAL = PatternFill("solid", fgColor="F4CCCC")      # rot/rosa: manuell zu ergaenzen
FILL_FORMULA = PatternFill("solid", fgColor="D9D2E9")     # lila: Formel/Zwischensumme
FILL_HEADER = PatternFill("solid", fgColor="434343")
FONT_HEADER = Font(color="FFFFFF", bold=True)
FONT_BOLD = Font(bold=True)
THIN = Side(style="thin", color="AAAAAA")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ZEILEN_KLASSE = {
    "erloese": "auto", "bestand_fe": "auto", "bestand_ufe": "auto",
    "aktivierte_eigenleistung": "strukturell_null", "betriebsleistung": "formula",
    "material": "auto", "fremdleistungen": "auto", "personalaufwand": "strukturell_null",
    "sbe": "strukturell_null", "sba": "auto", "afa": "strukturell_null", "rohmarge_1": "formula",
    "fek": "auto", "fek_deckungsdifferenz": "manual", "fek_nach_dd": "formula", "db1": "formula",
    "mgk": "auto", "mgk_deckungsdifferenz": "manual", "mgk_nach_dd": "formula", "db2": "formula",
    "vvgk": "auto", "vvgk_deckungsdifferenz": "manual", "vvgk_nach_dd": "formula",
    # 'sonderposten' ist NICHT manuell zu ergaenzen: in der Referenzrechnung traegt
    # diese Zeile in jeder echten Sparte 0,00. Die Betraege (Photovoltaik, Grundstueck
    # Valluhn) stehen ausschliesslich auf Kostenobjekten, die keine Sparte sind - sie
    # koennen in einer Kostentraegerrechnung gar nicht je Sparte auftauchen. Rosa war
    # hier also eine Aufforderung, etwas nachzutragen, das es nicht gibt.
    "db3_vor_sonderposten": "formula", "sonderposten": "strukturell_null", "db3": "formula",
}

FILL_NULL = PatternFill("solid", fgColor="EFEFEF")       # grau: strukturell null

FILLS = {"auto": FILL_AUTO, "auto_pruef": FILL_AUTO_PRUEF, "manual": FILL_MANUAL,
         "formula": FILL_FORMULA, "strukturell_null": FILL_NULL}


def ermittle_stundensatz(df, mapping):
    """Rechnet den Verrechnungssatz der Fertigung aus den Rohdaten selbst aus.

    Fuer die Fertigungskostenstellen (Praefix 'K2') fuehrt KPTM zwei Wertarten:
    ISWF = Ist-Wert in EUR und ISMF = Ist-Menge in Stunden. Es gilt
    ISWF = ISMF x Satz, der Satz ist also schlicht ISWF / ISMF.

    Geprueft an Q1 2026: bei allen 14 Fertigungskostenstellen derselbe Satz, keine
    Abweichung ueber 1 Cent. Deshalb muss er nirgends hinterlegt werden - er kommt
    aus den Daten, die ohnehin hochgeladen werden. Das ist auch der sicherere Weg:
    der Satz ist Kalkulationsinnenleben, aus dem sich die Preisbildung zurueckrechnen
    laesst, und steht so an keiner Stelle gespeichert. (Der konkrete Wert gehoert
    deshalb auch nicht in diesen Kommentar - er steht in intern/HERLEITUNG.md.)
    """
    kst = df[df["Kostenart"].str.startswith(mapping["kostenstellen_praefix"])]
    wert = kst[kst["Wertart"] == mapping.get("kostenstellen_wertart", "ISWF")]["Wert/Menge"].sum()
    menge = kst[kst["Wertart"] == "ISMF"]["Wert/Menge"].sum()
    if not menge:
        return None
    return round(float(wert) / float(menge), 4)


def lade_betriebsparameter(config_dir):
    """Laedt die Betriebsparameter (Werkstattbestand zum Jahresanfang), falls vorhanden.

    Die Anfangsbestaende sind Firmenzahlen und liegen deshalb nicht im oeffentlichen
    Repository, sondern im privaten Teil (intern/) bzw. werden in der Weboberflaeche
    einmalig hinterlegt und dort im Browser gespeichert.

    Der Stundensatz stand hier frueher ebenfalls - er wird inzwischen aus den
    Rohdaten errechnet (siehe ermittle_stundensatz) und muss nicht mehr gepflegt
    werden.

    Fehlt die Datei, rechnet das Tool alle GuV-Zeilen unveraendert weiter; nur die
    Zeile Bestandsveraenderung UFE entfaellt.

    Gesucht wird an zwei Orten: im uebergebenen config-Verzeichnis (so schreibt die
    Weboberflaeche die Datei ins virtuelle Dateisystem) und daneben in 'intern/'
    (dort liegt das private Repository, wenn lokal per Kommandozeile gearbeitet wird).
    """
    namen = ["betriebsparameter.xlsx", "betriebsparameter.xls", "betriebsparameter.json"]
    orte = [config_dir, os.path.join(config_dir, os.pardir, "intern")]
    for ort in orte:
        for name in namen:
            pfad = os.path.join(ort, name)
            if not os.path.exists(pfad):
                continue
            if name.endswith(".json"):
                with open(pfad, encoding="utf-8") as f:
                    return json.load(f)
            return betriebsparameter_aus_excel(pfad)
    return {}


# Spaltenueberschriften, die als Sparte/Jahr/Wert erkannt werden. Bewusst grosszuegig:
# die Datei wird von Hand in Excel gepflegt, da heisst es mal "Sparte", mal "Produktgruppe".
_SPALTEN_SPARTE = {"sparte", "produktgruppe", "kostenobjekt", "pg"}
_SPALTEN_JAHR = {"geschaftsjahr", "geschaeftsjahr", "jahr"}
_SPALTEN_WERT = {"anfangsbestand", "werkstattbestand", "wert", "betrag", "eur"}


def _spalte_finden(spalten, erlaubt):
    for s in spalten:
        norm = str(s).strip().lower().replace("ä", "a").replace("ö", "o").replace("ü", "u")
        norm = norm.replace("ß", "ss").replace(" ", "").replace("(eur)", "").replace(".", "")
        if norm in erlaubt:
            return s
    return None


def betriebsparameter_aus_excel(pfad):
    """Liest die Betriebsparameter aus einer Excel-Datei.

    Erwartet ein Blatt mit den Spalten Sparte | Geschaeftsjahr | Anfangsbestand.
    Excel statt JSON, weil diese Datei im Controlling gepflegt wird und nicht von
    Entwicklern - eine Tabelle laesst sich dort oeffnen, pruefen und ergaenzen.

    Die Spaltennamen werden grosszuegig erkannt (Sparte/Produktgruppe, Jahr/
    Geschaeftsjahr, Anfangsbestand/Wert/Betrag). Fehlt die Jahresspalte, gilt die
    Datei fuer alle Jahre - dann traegt der Aufrufer die Verantwortung dafuer, dass
    sie zur Auswertung passt.
    """
    warnings.filterwarnings("ignore")
    # Die Kopfzeile wird gesucht, nicht vorausgesetzt: eine von Hand gepflegte Tabelle
    # traegt haeufig eine Titelzeile oder eine Leerzeile darueber.
    roh = sp_sparte = sp_wert = sp_jahr = None
    for kopfzeile in range(0, 10):
        kandidat = pd.read_excel(pfad, sheet_name=0, header=kopfzeile)
        kandidat.columns = [str(c).strip() for c in kandidat.columns]
        s = _spalte_finden(kandidat.columns, _SPALTEN_SPARTE)
        w = _spalte_finden(kandidat.columns, _SPALTEN_WERT)
        if s is not None and w is not None:
            roh, sp_sparte, sp_wert = kandidat, s, w
            sp_jahr = _spalte_finden(kandidat.columns, _SPALTEN_JAHR)
            break
    if roh is None:
        gefunden = pd.read_excel(pfad, sheet_name=0, nrows=5)
        raise ValueError(
            "Die Betriebsparameter-Datei braucht die Spalten 'Sparte' und 'Anfangsbestand' "
            f"(in einer der ersten zehn Zeilen). Gefunden wurde: {list(gefunden.columns)[:6]}"
        )

    # Zweite, optionale Tabelle auf demselben Blatt: freie Parameter als
    # Name/Wert-Paare. Hier gehoeren die erwarteten Zuschlagssaetze hin - sie sind
    # Kalkulationsinnenleben und duerfen nicht ins oeffentliche Repository, aber ohne
    # Sollwert kann eine hausweite Satzaenderung nicht auffallen.
    parameter = {}
    sp_name = _spalte_finden(roh.columns, {"parameter", "kennzahl", "name"})
    sp_wert2 = _spalte_finden(roh.columns, {"wert", "betrag"})
    if sp_name is not None and sp_wert2 is not None and sp_name != sp_sparte:
        for _, zeile in roh.iterrows():
            name = zeile[sp_name]
            if name is None or (isinstance(name, float) and pd.isna(name)):
                continue
            try:
                wert = float(zeile[sp_wert2])
            except (TypeError, ValueError):
                continue
            if pd.isna(wert):
                continue          # Erlaeuterungszeilen ohne Wert ueberspringen
            parameter[str(name).strip().lower()] = wert

    anfangsbestand = {}
    for _, zeile in roh.iterrows():
        sparte = _schluessel(zeile[sp_sparte])
        if not sparte:
            continue
        try:
            wert = float(zeile[sp_wert])
        except (TypeError, ValueError):
            continue
        if pd.isna(wert):
            continue
        jahr = "alle"
        if sp_jahr is not None and not pd.isna(zeile[sp_jahr]):
            jahr = str(int(zeile[sp_jahr]))
        anfangsbestand.setdefault(jahr, {})[sparte] = round(wert, 2)
    # Optional: Spalte 'Herkunft' - welche Anfangsbestaende bei der Umstellung auf die
    # neue Produktgruppen-Struktur nur anteilig aufgeteilt wurden. Deren UFE-Zellen
    # zeigt die Spartenrechnung gelb (Naeherung).
    naeherung = {}
    sp_herkunft = _spalte_finden(roh.columns, {"herkunft"})
    if sp_herkunft is not None:
        for _, zeile in roh.iterrows():
            text = str(zeile[sp_herkunft]).lower()
            sparte = _schluessel(zeile[sp_sparte])
            if sparte and ("aufgeteilt" in text or "naeherung" in text or "näherung" in text):
                jahr = (str(int(zeile[sp_jahr])) if sp_jahr is not None
                        and not pd.isna(zeile[sp_jahr]) else "alle")
                naeherung.setdefault(jahr, []).append(sparte)
    return {"anfangsbestand": anfangsbestand, "parameter": parameter,
            "umschluesselung": umschluesselung_aus_excel(pfad),
            "anfangsbestand_naeherung": naeherung}


def umschluesselung_aus_excel(pfad):
    """Liest das optionale Blatt 'Umschluesselung' der Betriebsparameter.

    Spalten: Artikel | alte Produktgruppe | neue Produktgruppe. Grundlage ist die
    Produktgruppen-Reorganisation 2026 (Datei 'Alte vs Neue Produktgruppen
    Horizontalbohrtechnik'). Gespeichert werden nur Artikel, deren Gruppe sich
    aendert. Liefert {artikel: [alt, neu]} oder {}, wenn das Blatt fehlt.
    """
    try:
        blaetter = pd.ExcelFile(pfad).sheet_names
    except Exception:
        return {}
    blatt = next((b for b in blaetter if "umschl" in b.lower()), None)
    if blatt is None:
        return {}
    for kopfzeile in range(0, 10):
        df = pd.read_excel(pfad, sheet_name=blatt, header=kopfzeile)
        df.columns = [str(c).strip() for c in df.columns]
        sp_art = _spalte_finden(df.columns, {"artikel", "artikelnummer", "artikel_nr", "artikelnr"})
        sp_alt = _spalte_finden(df.columns, {"alteproduktgruppe", "alt", "produktgruppealt"})
        sp_neu = _spalte_finden(df.columns, {"neueproduktgruppe", "neu", "produktgruppeneu"})
        if sp_art is not None and sp_alt is not None and sp_neu is not None:
            break
    else:
        raise ValueError(
            f"Das Blatt '{blatt}' braucht die Spalten 'Artikel', 'alte Produktgruppe' und "
            f"'neue Produktgruppe'.")
    karte = {}
    for art, alt, neu in zip(df[sp_art], df[sp_alt], df[sp_neu]):
        a, alt_, neu_ = _schluessel(art), _schluessel(alt), _schluessel(neu)
        if a and alt_ and neu_ and alt_ != neu_:
            karte[a] = [alt_, neu_]
    return karte


def umschluesseln(df, karte, aufgeloest=None):
    """Bringt Buchungen aus der alten Produktgruppen-Struktur in die neue.

    Seit April 2026 bucht das ERP auf die neuen Produktgruppen; aeltere Auftraege laufen
    aber unter ihrer alten Gruppe weiter. Ein kumulierter Export des Jahres mischt also
    beide Strukturen - dieselbe Produktlinie stuende in zwei Spalten. Umgeschluesselt
    wird je Buchung ueber den Artikel, und nur, wenn die gebuchte Gruppe die alte Gruppe
    dieses Artikels ist (sonst ist die Buchung schon neu, oder der Artikel wurde nie
    umgeschluesselt). Buchungen ohne Artikel bleiben in ihrer Gruppe - ausser die
    Gruppe ist aufgeloest (z.B. PB0003 Pilot Tubes), dann geht sie an den Nachfolger.

    Gemessen an Q1 2026: 93,7 % des Betrags auf PB0001-PB0005 artikelgenau.
    Die gebuchte Gruppe bleibt in der Spalte 'Produktgruppe (gebucht)' erhalten.
    """
    df = df.copy()
    alt = df["Produktgruppe"].copy()
    # Direkt neben die Produktgruppe, damit im Datenblatt beide nebeneinander stehen.
    df.insert(list(df.columns).index("Produktgruppe") + 1, "Produktgruppe (gebucht)", alt)
    neu = alt.copy()
    if "Artikelnummer" in df.columns:
        ziel = df["Artikelnummer"].map(_schluessel).map(karte.get)
        passt = ziel.notna() & (ziel.map(lambda z: z[0] if z else None) == alt)
        neu[passt] = ziel[passt].map(lambda z: z[1])
    for weg, nachfolger in (aufgeloest or {}).items():
        neu[neu == weg] = nachfolger
    geaendert = neu.fillna("") != alt.fillna("")
    df["Produktgruppe"] = neu
    return df, {
        "zeilen": int(geaendert.sum()),
        "betrag": round(float(df.loc[geaendert, "Wert/Menge"].abs().sum()), 2),
        "paare": {f"{a}->{n}": int(c) for (a, n), c in
                  pd.DataFrame({"a": alt[geaendert], "n": neu[geaendert]})
                  .value_counts().items()},
    }


def _centgenau_verteilen(betrag, anteile):
    """Verteilt einen Betrag nach Anteilen, so dass die gerundeten Teile exakt wieder
    den Betrag ergeben (der Rundungsrest geht an den groessten Teil)."""
    teile = {k: round(betrag * a, 2) for k, a in anteile.items()}
    rest = round(betrag - sum(teile.values()), 2)
    if teile and abs(rest) > 0:
        groesster = max(teile, key=lambda k: abs(teile[k]))
        teile[groesster] = round(teile[groesster] + rest, 2)
    return teile


def anfangsbestand_umschluesseln(anfangsbestand, pwbs_pfad, karte, aufgeloest=None):
    """Teilt einen Anfangsbestand der alten Struktur auf die neuen Gruppen auf.

    Einmalig fuer das Umstellungsjahr: Der Anfangsbestand liegt nur je ALTER Gruppe vor
    (er ist gegen die Referenzrechnung geprueft), die neue Struktur braucht ihn je NEUER
    Gruppe. Schluessel ist der Werkstattbestand zum selben Stichtag je Auftrag: Fuer jede
    alte Gruppe der Anteil des offenen Werts, der auf jede neue Gruppe entfaellt.

    Die Summe je alter Gruppe bleibt centgenau erhalten - es wird nur verteilt. Die
    Verteilung selbst ist eine Naeherung: Der verfuegbare Werkstattbestand zum 01.01.26
    wurde erst im Oktober nachtraeglich gezogen, seine Werte stimmen nicht mit dem
    Bilanzwert ueberein (siehe HERLEITUNG). Fuer die ANTEILE innerhalb einer Gruppe ist
    er trotzdem die beste verfuegbare Quelle.

    Liefert (neuer_anfangsbestand, anteile je alter Gruppe, naeherung) - naeherung sind
    die neuen Gruppen, die einen ANTEIL (< 100 %) einer alten Gruppe erhalten haben.
    """
    df = lies_penta(pwbs_pfad)
    df.columns = [str(c).strip() for c in df.columns]
    art = df["Artikelnummer"].map(_schluessel)
    spalte = df["Produktgruppe"].map(_schluessel)
    # Fuer Artikel der Umschluesselung gelten alt UND neu aus der Liste - genau wie bei
    # den Buchungen. Nur so landen Anfangsbestand und Buchungen in derselben Gruppe.
    alt = [karte[a][0] if a in karte else s for a, s in zip(art, spalte)]
    neu = [karte[a][1] if a in karte else s for a, s in zip(art, spalte)]
    neu = [(aufgeloest or {}).get(n, n) for n in neu]
    tmp = pd.DataFrame({"art": art, "alt": alt, "neu": neu, "wert": df["Offener Wert"]})
    tmp = tmp[tmp["art"].notna() & tmp["neu"].notna()]    # ohne Serviceauftraege

    anteile = {}
    for alt_pg, gruppe in tmp.groupby("alt"):
        summe = gruppe.groupby("neu")["wert"].sum()
        gesamt = float(summe.sum())
        if abs(gesamt) > 0.005:
            anteile[alt_pg] = {n: float(v) / gesamt for n, v in summe.items() if abs(v) > 0.005}

    neu_ab, naeherung = {}, set()
    # Neue Gruppen, die aus einer alten hervorgehen, am 01.01. aber keinen offenen
    # Auftrag hatten, bekommen ausdruecklich 0,00 - sonst meldet die UFE-Zeile sie als
    # "Anfangsbestand fehlt", obwohl er gemessen null ist.
    for alt_pg, ziel in karte.values():
        ziel = (aufgeloest or {}).get(ziel, ziel)
        if abs(float(anfangsbestand.get(alt_pg, 0.0))) > 0.005:
            neu_ab.setdefault(ziel, 0.0)
            naeherung.add(ziel)
    for alt_pg, betrag in anfangsbestand.items():
        verteilung = anteile.get(alt_pg) or {(aufgeloest or {}).get(alt_pg, alt_pg): 1.0}
        for ziel, teil in _centgenau_verteilen(float(betrag), verteilung).items():
            neu_ab[ziel] = round(neu_ab.get(ziel, 0.0) + teil, 2)
            if verteilung.get(ziel, 0) < 0.9999 and abs(betrag) > 0.005:
                naeherung.add(ziel)
    return neu_ab, anteile, naeherung


def aufteilungsanteile(df, mapping):
    """Wie verteilt sich jede GEBUCHTE Produktgruppe auf die neuen Gruppen?

    Gebraucht fuer Groessen, die nur je Produktgruppe vorliegen und keinen Artikel
    tragen - vor allem die Deckungsdifferenzen der Kostentraeger-Auswertung, die auf
    der alten Gruppe gebucht sind. Sie werden im Verhaeltnis ihrer Bezugsbasis
    aufgeteilt: FEK-Differenz nach FEK, MGK-Differenz nach Material + Fremdleistungen,
    VVGK-Differenz nach Herstellkosten. Das ist eine NAEHERUNG - exakt waere es nur,
    wenn die Buchhaltung die Differenzen schon in der neuen Struktur gebucht haette.
    Gruppen, die vollstaendig in EINE neue uebergehen (z.B. PB0003 -> PB0002), sind
    dagegen exakt.

    Liefert {dd_zeile: {alte_pg: {neue_pg: anteil}}}, nur fuer tatsaechlich geteilte
    oder umbenannte Gruppen.
    """
    if "Produktgruppe (gebucht)" not in df.columns:
        return {}
    kz = mapping["kostenart_zu_zeile"]
    praefix = mapping["kostenstellen_praefix"]
    kst_wertart = mapping.get("kostenstellen_wertart", "ISWF")

    def zeile(ka, wertart):
        ka = str(ka)
        if ka.startswith(praefix):
            return ("fek", 1) if wertart == kst_wertart else (None, 0)
        r = kz.get(ka)
        return (r["zeile"], r["sign"]) if r else (None, 0)

    zz = [zeile(k, w) for k, w in zip(df["Kostenart"], df["Wertart"])]
    tmp = pd.DataFrame({
        "alt": df["Produktgruppe (gebucht)"].values, "neu": df["Produktgruppe"].values,
        "zeile": [z[0] for z in zz],
        "w": df["Wert/Menge"].values * [z[1] for z in zz],
    }).dropna(subset=["alt", "neu", "zeile"])
    basen = {
        "fek_deckungsdifferenz": ["fek"],
        "mgk_deckungsdifferenz": ["material", "fremdleistungen"],
        "vvgk_deckungsdifferenz": ["material", "fremdleistungen", "mgk", "fek"],
    }
    anteile = {}
    for dd, zeilen in basen.items():
        summe = tmp[tmp["zeile"].isin(zeilen)].groupby(["alt", "neu"])["w"].sum()
        for alt_pg, gruppe in summe.groupby(level=0):
            ziele = {n: float(v) for (_, n), v in gruppe.items() if abs(v) > 0.005}
            if set(ziele) <= {alt_pg}:
                continue                      # nichts umgeschluesselt
            gesamt = sum(ziele.values())
            if abs(gesamt) < 0.01:
                continue
            anteile.setdefault(dd, {})[alt_pg] = {n: v / gesamt for n, v in ziele.items()}
    return anteile


def schreibe_betriebsparameter_excel(anfangsbestand_je_jahr, out_path, parameter=None,
                                     umschluesselung=None, naeherung=None):
    """Schreibt die Betriebsparameter als Excel-Datei - dasselbe Format, das
    betriebsparameter_aus_excel wieder einliest."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Betriebsparameter"
    ws["A1"] = "Werkstattbestand je Sparte zum 01.01. (Anfangsbestand für die Zeile Bestandsveränderung UFE)"
    ws["A1"].font = FONT_BOLD
    kopf = ["Sparte", "Geschäftsjahr", "Anfangsbestand"] + (["Herkunft"] if naeherung else [])
    for i, titel in enumerate(kopf, start=1):
        z = ws.cell(row=3, column=i, value=titel)
        z.font = FONT_HEADER
        z.fill = FILL_HEADER
        z.border = BORDER
    r = 4
    for jahr, werte in sorted(anfangsbestand_je_jahr.items()):
        for sparte, wert in sorted(werte.items()):
            ws.cell(row=r, column=1, value=sparte).border = BORDER
            ws.cell(row=r, column=2, value=int(jahr) if str(jahr).isdigit() else jahr).border = BORDER
            z = ws.cell(row=r, column=3, value=float(wert))
            z.number_format = "#,##0.00"
            z.border = BORDER
            if naeherung:
                aufgeteilt = sparte in naeherung.get(str(jahr), set())
                ws.cell(row=r, column=4, value=("aufgeteilt (Naeherung)" if aufgeteilt
                                                else "gepflegt")).border = BORDER
            r += 1
    # Zweite Tabelle: freie Parameter. Hier gehoeren die erwarteten Zuschlagssaetze
    # hin - damit eine hausweite Satzaenderung auffaellt, ohne dass der Sollwert im
    # oeffentlichen Repository steht.
    sp_par = 6 if naeherung else 5
    for i, titel in ((sp_par, "Parameter"), (sp_par + 1, "Wert")):
        z = ws.cell(row=3, column=i, value=titel)
        z.font = FONT_HEADER
        z.fill = FILL_HEADER
        z.border = BORDER
    for i, (name, wert) in enumerate(sorted((parameter or {}).items()), start=4):
        ws.cell(row=i, column=sp_par, value=name).border = BORDER
        z = ws.cell(row=i, column=sp_par + 1, value=float(wert))
        z.number_format = "0.0000"
        z.border = BORDER
    ws.cell(row=max(len(parameter or {}), 1) + 6, column=sp_par, value=(
        "Namen: mgk_zuschlagssatz und vvgk_zuschlagssatz, Wert als Dezimalzahl (z.B. 0,15 "
        "fuer 15 %). Weicht der aus den Daten zurueckgerechnete Satz davon ab, meldet das "
        "Tool es."
    )).font = Font(italic=True, size=9)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    for sp, b in ((("D", 22), ("E", 3), ("F", 24), ("G", 12)) if naeherung
                  else (("D", 3), ("E", 24), ("F", 12))):
        ws.column_dimensions[sp].width = b
    ws.freeze_panes = "A4"

    # Optional: Umschluesselung alte -> neue Produktgruppe je Artikel (Reorganisation
    # 2026). Gebraucht, solange im Jahr noch Buchungen unter alten Gruppen liegen.
    if umschluesselung:
        wu = wb.create_sheet("Umschluesselung")
        wu["A1"] = ("Produktgruppen-Reorganisation: alte -> neue Gruppe je Artikel. Nur Artikel, "
                    "deren Gruppe sich aendert.")
        wu["A1"].font = Font(italic=True, size=9)
        for i, titel in enumerate(("Artikel", "alte Produktgruppe", "neue Produktgruppe"), start=1):
            z = wu.cell(row=3, column=i, value=titel)
            z.font, z.fill, z.border = FONT_HEADER, FILL_HEADER, BORDER
        for r, (art, (alt, neu)) in enumerate(sorted(umschluesselung.items()), start=4):
            wu.cell(row=r, column=1, value=str(art))
            wu.cell(row=r, column=2, value=alt)
            wu.cell(row=r, column=3, value=neu)
        for sp, b in (("A", 14), ("B", 20), ("C", 20)):
            wu.column_dimensions[sp].width = b
        wu.freeze_panes = "A4"
    wb.save(out_path)
    return out_path


def betriebsparameter_fortschreiben(bisher, jahr, bestand, out_path):
    """Schreibt die Betriebsparameter-Datei mit einem neuen Anfangsbestand fort.

    'bisher' ist der Inhalt der hinterlegten Parameter (wie betriebsparameter_aus_excel
    ihn liefert). Alles darin bleibt erhalten - die Anfangsbestaende der uebrigen
    Jahre, die Zuschlagssaetze und die Umschluesselung. Frueher schrieb die Oberflaeche
    nur den neuen Bestand; wer die Datei dann wieder lud, hatte alles andere verloren.
    Der neue Bestand ist gemessen, nicht aufgeteilt - seine Naeherungs-Markierung entfaellt.
    """
    bisher = bisher or {}
    jahre = {j: dict(w) for j, w in (bisher.get("anfangsbestand") or {}).items()}
    jahre[str(jahr)] = dict(bestand)
    naeherung = {j: set(v) for j, v in (bisher.get("anfangsbestand_naeherung") or {}).items()
                 if j != str(jahr) and v}
    schreibe_betriebsparameter_excel(jahre, out_path, parameter=bisher.get("parameter"),
                                     umschluesselung=bisher.get("umschluesselung"),
                                     naeherung=naeherung or None)


def load_config(config_dir):
    with open(os.path.join(config_dir, "kostenart_mapping.json"), encoding="utf-8") as f:
        mapping = json.load(f)
    with open(os.path.join(config_dir, "produktgruppen.json"), encoding="utf-8") as f:
        pg = json.load(f)
    # Betriebsparameter ueberlagern das Mapping, damit der Rest des Codes sie wie
    # bisher unter mapping["standard_stundensatz"] findet.
    betrieb = lade_betriebsparameter(config_dir)
    if betrieb.get("standard_stundensatz"):
        mapping["standard_stundensatz"] = betrieb["standard_stundensatz"]
    # Erwartete Zuschlagssaetze, falls in den Betriebsparametern hinterlegt. Ohne sie
    # prueft zuschlagssaetze() nur, ob die Sparten untereinander uebereinstimmen - eine
    # hausweite Satzaenderung faellt dann nicht auf.
    for schluessel in ("mgk", "vvgk"):
        for quelle in (betrieb, betrieb.get("parameter", {})):
            wert = quelle.get(f"{schluessel}_zuschlagssatz") if quelle else None
            if wert:
                mapping[f"{schluessel}_zuschlagssatz_erwartet"] = float(wert)
                break
    return mapping, pg


def read_bwa_ergebnis(bwa_path, sheet_name):
    """Liest das 'Vorlaeufige Ergebnis' (Spalte 'kumuliert') aus einer BWA-Arbeitsmappe.

    Die BWA enthaelt pro Monat ein Sheet 'BWA MM.YYYY' mit den GuV-Zeilen firmenweit
    (KEINE Sparten-Aufschluesselung). Zeile 'Vorlaeufiges Ergebnis', Spalte 'kumuliert'
    entspricht exakt der Zeile 'Ergebnis lt. GuV vor Besserungsschein' in der bisherigen
    manuellen Spartenrechnung und dient dort als Kontroll-Anker fuer die bottom-up
    berechnete DB III.
    """
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(bwa_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' nicht gefunden. Verfuegbare Sheets (Auszug): {wb.sheetnames[:10]} ..."
        )
    ws = wb[sheet_name]
    header_row = 5
    label_col = 3  # C
    kum_col = None
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=header_row, column=c).value == "kumuliert":
            kum_col = c
            break
    if kum_col is None:
        raise ValueError(f"Spalte 'kumuliert' nicht in Zeile {header_row} von Sheet '{sheet_name}' gefunden.")
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(row=r, column=label_col).value == "Vorläufiges Ergebnis":
            return ws.cell(row=r, column=kum_col).value
    raise ValueError(f"Zeile 'Vorläufiges Ergebnis' nicht in Sheet '{sheet_name}' gefunden.")


# Welche Spalten eine Datei tragen muss, und woran man sie erkennt, wenn die
# falsche hochgeladen wurde. Ohne diese Pruefung quittiert das Tool eine
# verwechselte Datei mit "KeyError: 'Kostenart'" - fuer den Anwender unbrauchbar.
_PFLICHTSPALTEN = {
    "KPTM": ["Kostenart", "Produktgruppe", "Wertart", "Wert/Menge"],
    "PWBS": ["Rückmeldenummer", "Artikelnummer", "Offener Wert"],
    "PFAK": ["RUECKMELDE_NR", "PRODUKTGRUPPE"],
    "Kostenträger": ["Objekttyp", "Kostenobjekt", "Kostenart", "Wert/Menge"],
}


def lies_penta(pfad):
    """Liest das Blatt 'Penta' eines ERP-Exports - mit verstaendlicher Meldung, wenn es fehlt.

    Fehlt das Blatt, ist es fast immer keine Originaldatei: entweder in Excel neu
    gespeichert ('Tabelle1') oder ein Datenblatt aus einer vom Tool erzeugten
    Spartenrechnung, herauskopiert. Letzteres sieht aus wie ein Export, ist aber
    gekuerzt (Spalten ohne Informationsgehalt fehlen, darunter das Geschaeftsjahr; bei
    der Kostentraeger-Auswertung stehen nur die Umlagezeilen drin)."""
    try:
        return pd.read_excel(pfad, sheet_name="Penta")
    except ValueError as e:
        if "Penta" not in str(e):
            raise
    name = os.path.basename(str(pfad))
    try:
        blaetter = pd.ExcelFile(pfad).sheet_names
        erste = pd.read_excel(pfad, sheet_name=0, header=None, nrows=1).iloc[0, 0]
    except Exception:
        blaetter, erste = [], None
    aus_tool = isinstance(erste, str) and ("gelbe" in erste or "gelben" in erste)
    text = (f"'{name}' ist kein Original-Export aus dem ERP: das Tabellenblatt 'Penta' fehlt"
            + (f" (gefunden: {', '.join(map(str, blaetter[:4]))})" if blaetter else "") + ". ")
    if aus_tool:
        text += ("Die Datei ist ein Datenblatt aus einer vom Tool erzeugten Spartenrechnung - "
                 "das ist gekürzt und taugt nicht als Eingabe. Bitte die Originaldatei "
                 "hochladen, so wie sie aus Penta exportiert wurde.")
    else:
        text += ("Bitte die Datei unverändert hochladen, so wie sie aus Penta exportiert wurde "
                 "(nicht in Excel neu gespeichert oder umbenannt).")
    raise ValueError(text)


def _pruefe_datei(df, art, pfad):
    """Prueft eine eingelesene Datei auf Brauchbarkeit, bevor gerechnet wird."""
    name = os.path.basename(str(pfad))
    fehlend = [sp for sp in _PFLICHTSPALTEN[art] if sp not in df.columns]
    if fehlend:
        # Steckt vielleicht eine andere Exportart dahinter? Dann ist der Hinweis
        # "Sie haben die falsche Datei hochgeladen" hilfreicher als eine Spaltenliste.
        for andere, spalten in _PFLICHTSPALTEN.items():
            if andere != art and all(sp in df.columns for sp in spalten):
                raise ValueError(
                    f"'{name}' sieht aus wie ein {andere}-Export, erwartet wird aber "
                    f"{art}. Bitte die richtige Datei in das passende Feld laden."
                )
        raise ValueError(
            f"In '{name}' fehlen die Spalten: {', '.join(fehlend)}. Erwartet wird ein "
            f"{art}-Export mit den Spalten {', '.join(_PFLICHTSPALTEN[art])}. "
            f"Gefunden wurden: {', '.join(str(c) for c in list(df.columns)[:8])} ..."
        )
    if df.empty:
        raise ValueError(
            f"'{name}' enthaelt keine Datenzeilen (nur die Kopfzeile). Das Ergebnis "
            f"waere durchgaengig 0,00 - deshalb wird hier abgebrochen statt eine leere "
            f"Rechnung auszugeben."
        )
    return df


def load_kptm(path):
    df = lies_penta(path)
    df.columns = [c.strip() for c in df.columns]
    _pruefe_datei(df, "KPTM", path)
    # Zahlenspalte absichern: kommt sie als Text (z.B. mit Dezimalkomma), scheitert
    # die Summierung sonst mitten in der Rechnung mit einer Typmeldung aus numpy.
    if not pd.api.types.is_numeric_dtype(df["Wert/Menge"]):
        umgewandelt = pd.to_numeric(
            df["Wert/Menge"].astype(str).str.replace(".", "", regex=False)
                            .str.replace(",", ".", regex=False), errors="coerce")
        if umgewandelt.isna().all():
            raise ValueError(
                f"Die Spalte 'Wert/Menge' in '{os.path.basename(str(path))}' enthaelt "
                f"keine lesbaren Zahlen. Bitte den Export ohne Textformatierung ziehen."
            )
        df["Wert/Menge"] = umgewandelt.fillna(0.0)
    df["Kostenart"] = df["Kostenart"].astype(str).str.strip()
    df["Produktgruppe"] = df["Produktgruppe"].astype(str).str.strip()
    df.loc[df["Produktgruppe"].isin(["nan", "None", ""]), "Produktgruppe"] = None
    return df


def _jahr_aus_zeitraum(zeitraum):
    """Liest die Jahreszahl aus dem Zeitraum-Text, z.B. 'Jan-Maerz 2026' -> '2026'."""
    treffer = re.findall(r"(20\d{2})", str(zeitraum))
    return treffer[-1] if treffer else None


# Schreibweise wie in der Referenzmappe des Controllings ("Spartenrechnung
# Jan-März 2026"), damit die erzeugte Datei neben deren Blättern nicht auffaellt.
_MONATE = ["Jan", "Feb", "März", "Apr", "Mai", "Juni",
           "Juli", "Aug", "Sep", "Okt", "Nov", "Dez"]


def _zeitraum_aus_dateiname(kptm_path):
    """Rueckfall, wenn der Export keine Periodenspalten mitbringt: der ERP-Export
    traegt den Zeitraum im Namen, z.B. 'KPTM_Wertsummen_01-0326' -> 01 bis 03/2026."""
    treffer = re.search(r"(\d{2})-(\d{2})(\d{2})", os.path.basename(str(kptm_path or "")))
    if not treffer:
        return None
    von, bis, jj = (int(g) for g in treffer.groups())
    if not (1 <= von <= bis <= 12):
        return None
    return von, bis, 2000 + jj, 2000 + jj


def _perioden_aus_daten(df):
    """Welche Buchungsperioden stecken im Export? Gibt (von, bis, jahr_von, jahr_bis)."""
    if "Periode" not in df.columns or "Geschäftsjahr" not in df.columns:
        return None
    perioden = pd.to_numeric(df["Periode"], errors="coerce").dropna()
    jahre = pd.to_numeric(df["Geschäftsjahr"], errors="coerce").dropna()
    # Periode 0 und 13 sind Eroeffnungs- bzw. Abschlussbuchungen, keine Monate.
    perioden = perioden[(perioden >= 1) & (perioden <= 12)]
    if perioden.empty or jahre.empty:
        return None
    return (int(perioden.min()), int(perioden.max()),
            int(jahre.min()), int(jahre.max()))


def _stichtag_aus_dateiname(pfad):
    """Stichtag aus dem Dateinamen eines Werkstattbestands, z.B.
    'PWBS_Werkstattbestand_010426.xls' oder '..._010326_Produktgruppe.xls' -> 1.4.2026.
    Die Datei selbst traegt kein Datum - der Name ist die einzige Angabe."""
    import datetime
    treffer = re.search(r"Werkstattbestand[_ ]?(\d{2})(\d{2})(\d{2})(?!\d)",
                        os.path.basename(str(pfad)), re.IGNORECASE)
    if not treffer:
        return None
    tag, monat, jahr = (int(x) for x in treffer.groups())
    try:
        return datetime.date(2000 + jahr, monat, tag)
    except ValueError:
        return None


def pruefe_stichtag(pfad, kptm_df, art):
    """Passt der Stichtag des Werkstattbestands zum Zeitraum des KPTM-Exports?

    Der Bestand zum Periodenende gehoert auf den ersten Tag NACH dem Zeitraum (Jan-Maerz
    -> 01.04.), der Anfangsbestand auf den 01.01. Der Vortag wird ebenso akzeptiert
    (31.03. statt 01.04.). Ein Werkstattbestand mit falschem Stichtag rechnet still eine
    falsche UFE-Zeile - plausibel aussehend, um eine ganze Monatsbewegung daneben. Das ist
    passiert: Q1 mit dem Bestand zum 01.03. statt 01.04.

    Wirft ValueError bei Widerspruch. Liefert einen Hinweis-Text, wenn der Dateiname kein
    Datum traegt (dann ist keine Pruefung moeglich), sonst None.
    """
    import datetime
    spanne = _perioden_aus_daten(kptm_df)
    stichtag = _stichtag_aus_dateiname(pfad)
    name = os.path.basename(str(pfad))
    if spanne is None:
        return None
    von, bis, j_von, j_bis = spanne
    if art == "ende":
        erster = (datetime.date(j_bis + 1, 1, 1) if bis == 12
                  else datetime.date(j_bis, bis + 1, 1))
    else:
        erster = datetime.date(j_von, von, 1)
    erlaubt = {erster, erster - datetime.timedelta(days=1)}
    wofuer = "zum Periodenende" if art == "ende" else "zum Periodenbeginn (Anfangsbestand)"
    if stichtag is None:
        return (f"Der Stichtag des Werkstattbestands {wofuer} ist aus dem Dateinamen "
                f"'{name}' nicht ablesbar - bitte selbst pruefen, dass es der "
                f"{erster:%d.%m.%Y} ist.")
    if stichtag not in erlaubt:
        raise ValueError(
            f"Der Werkstattbestand '{name}' ist der Stand vom {stichtag:%d.%m.%Y}. "
            f"Der KPTM-Export reicht bis Periode {bis}/{j_bis} - {wofuer} gehoert deshalb der "
            f"Werkstattbestand zum {erster:%d.%m.%Y}. Mit dem falschen Stichtag waere die "
            f"Zeile 'Bestandsveraenderung UFE' um die Bewegung der fehlenden bzw. "
            f"ueberzaehligen Monate falsch - und mit ihr DB I bis DB III.")
    return None


def zeitraum_aus_daten(df, kptm_path=None):
    """Leitet den Auswertungszeitraum aus dem Export ab.

    Frueher musste er von Hand eingetippt werden - eine Fehlerquelle ohne Nutzen:
    die Rohdaten wissen selbst, welche Perioden sie enthalten, ein Tippfehler stand
    dagegen anschliessend als Ueberschrift in der Mappe.

    Rueckgabe: (Anzeigetext, Dateiname), z.B. ('Jan-März 2026',
    'Spartenrechnung_2026_01-03.xlsx').
    """
    spanne = _perioden_aus_daten(df) or _zeitraum_aus_dateiname(kptm_path)
    if not spanne:
        return "Zeitraum unbekannt", "Spartenrechnung.xlsx"
    von, bis, jahr_von, jahr_bis = spanne
    jahr = str(jahr_bis) if jahr_von == jahr_bis else f"{jahr_von}/{jahr_bis}"
    monate = _MONATE[von - 1] if von == bis else f"{_MONATE[von - 1]}-{_MONATE[bis - 1]}"
    kurz = f"{von:02d}" if von == bis else f"{von:02d}-{bis:02d}"
    return f"{monate} {jahr}", f"Spartenrechnung_{jahr_bis}_{kurz}.xlsx"


def geschaeftsjahr(df, zeitraum=None):
    """Bestimmt das Geschaeftsjahr der Auswertung.

    Massgeblich ist die Spalte 'Geschaeftsjahr' der Rohdaten - der Zeitraum-Text ist
    nur ein Anzeigefeld und enthaelt nicht zwingend eine Jahreszahl. Frueher wurde das
    Jahr aus diesem Freitext gelesen; stand dort z.B. nur 'Q1', fand das Tool den
    hinterlegten Anfangsbestand nicht und liess die UFE-Zeile grundlos aus.
    """
    if "Geschäftsjahr" in df.columns:
        werte = df["Geschäftsjahr"].dropna()
        if len(werte):
            haeufigstes = werte.astype(int).mode()
            if len(haeufigstes):
                return str(int(haeufigstes.iloc[0]))
    return _jahr_aus_zeitraum(zeitraum)


def lade_anfangsbestand(config_dir, jahr):
    """Laedt den Werkstattbestand zum Jahresanfang aus den Betriebsparametern.

    Traegt die Datei keine Jahresspalte, liegen die Werte unter 'alle' und gelten
    unabhaengig vom Geschaeftsjahr.
    """
    je_jahr = lade_betriebsparameter(config_dir).get("anfangsbestand", {})
    if jahr and str(jahr) in je_jahr:
        return je_jahr[str(jahr)]
    return je_jahr.get("alle")


def _schluessel(x):
    """Normalisiert Auftrags-/Artikelnummern auf eine einheitliche Textform.

    Noetig, weil dieselbe Nummer je nach Datei als '154081', 154081 oder '154081.0'
    ankommt. Ohne diese Normalisierung findet der Abgleich zwischen PWBS und PFAK
    keinen einzigen Treffer.
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if not s or s.lower() == "nan":
        return None
    if s.endswith(".0"):
        s = s[:-2]
    return s or None


def lade_produktgruppen_map(pfak_pfade=None, kptm_df=None):
    """Baut die Zuordnung Fertigungsauftrag/Artikel -> Produktgruppe.

    Zwei Quellen, die sich ergaenzen:

    1. KPTM (Spalte 'Kostenobjekt'): Das ist DIESELBE Nummer wie die Rueckmeldenummer
       in PWBS und PFAK - geprueft, und wo beide Quellen eine Produktgruppe kennen,
       stimmen sie zu 100 % ueberein (704 von 704 Auftraegen). Diese Quelle ist
       entscheidend, weil auf laufenden Auftraegen bereits Kosten gebucht sind: sie
       stehen deshalb schon im KPTM des LAUFENDEN Zeitraums.
    2. PFAK (Spalte 'RUECKMELDE_NR'): exakt, aber der Export filtert auf 'letzte
       Rueckmeldung im Zeitraum'. Auftraege, die am Stichtag noch laufen, erscheinen
       erst im Export eines spaeteren Zeitraums - PFAK allein reicht deshalb beim
       laufenden Abschluss nicht.

    Zusaetzlich eine Zuordnung ueber die Artikelnummer (nur wo eindeutig) als Ersatzweg.

    Bewusst NICHT verwendet wird der Kundenvorgang als Schluessel: er zieht
    Serviceauftraege (Monteureinsaetze, Inbetriebnahmen, Schulungen) in die Sparten,
    die in der manuellen Referenzrechnung nicht zum Werkstattbestand zaehlen.
    """
    nach_auftrag = {}
    artikel_kandidaten = {}
    widersprueche = []

    if kptm_df is not None:
        auftrag_kandidaten = {}
        # Nach einer Umschluesselung die GEBUCHTE Gruppe nehmen: PFAK kennt nur diese,
        # und der Abgleich unten waere sonst bei jedem umgeschluesselten Auftrag ein
        # scheinbarer Widerspruch. In die neue Struktur bringt werkstattbestand() sie.
        sp_pg = ("Produktgruppe (gebucht)" if "Produktgruppe (gebucht)" in kptm_df.columns
                 else "Produktgruppe")
        for kobj, art, pg in zip(kptm_df["Kostenobjekt"], kptm_df["Artikelnummer"],
                                 kptm_df[sp_pg]):
            pg = _schluessel(pg)
            if not pg:
                continue
            if _schluessel(kobj):
                auftrag_kandidaten.setdefault(_schluessel(kobj), set()).add(pg)
            if _schluessel(art) and _schluessel(art) != "0":
                artikel_kandidaten.setdefault(_schluessel(art), set()).add(pg)
        nach_auftrag.update({a: next(iter(v)) for a, v in auftrag_kandidaten.items() if len(v) == 1})

    for pfad in (pfak_pfade or []):
        df = pd.read_excel(pfad, sheet_name="PFAK_Fertigungsauftragskopf")
        df.columns = [c.strip() for c in df.columns]
        for rm, art, pg in zip(df["RUECKMELDE_NR"], df["ARTIKEL_NR"], df["PRODUKTGRUPPE"]):
            pg = _schluessel(pg)
            if not pg:
                continue
            if _schluessel(rm):
                # PFAK ist die genauere Quelle (Auftragskopf-Stammsatz) und ueberschreibt
                # KPTM. Dass beide je verschiedene Gruppen nennen, ist die stillschweigende
                # Annahme, auf der die ganze Zuordnung ruht - gemessen trifft sie zu
                # (0 Widersprueche bei 3.801 gemeinsamen Auftraegen ueber drei Zeitraeume).
                # Sie wird hier geprueft statt geglaubt: bricht sie, waeren Bestaende
                # zwischen Sparten verschoben, ohne dass irgendetwas auffiele.
                vorher = nach_auftrag.get(_schluessel(rm))
                if vorher is not None and vorher != pg:
                    widersprueche.append((_schluessel(rm), vorher, pg))
                nach_auftrag[_schluessel(rm)] = pg
            if _schluessel(art):
                artikel_kandidaten.setdefault(_schluessel(art), set()).add(pg)

    # Mehrdeutige Artikel (kommen bei mehreren Sparten vor) taugen nicht als
    # Ersatzschluessel und werden verworfen - konservativ richtig, aber nicht
    # folgenlos: ihr Anteil ist von 0,3 % auf 2,8 % gestiegen, weil mit jedem
    # weiteren Export Artikel ihre Eindeutigkeit verlieren koennen.
    nach_artikel = {a: next(iter(v)) for a, v in artikel_kandidaten.items() if len(v) == 1}
    mehrdeutig = sorted(a for a, v in artikel_kandidaten.items() if len(v) > 1)
    return nach_auftrag, nach_artikel, mehrdeutig, widersprueche


def werkstattbestand(pwbs_pfad, nach_auftrag, nach_artikel, manuelle_zuordnung=None,
                     karte=None, aufgeloest=None):
    """Summiert den offenen Werkstattbestand je Produktgruppe zu einem Stichtag.

    Massgeblich ist die Spalte 'Offener Wert' (= Istwert abzueglich bereits
    abgeliefertem Wert), nicht 'Istwert' - letztere verfehlt die Referenzwerte.

    Ueber `manuelle_zuordnung` ({Auftragsnummer: Produktgruppe}) lassen sich einzelne
    Auftraege von Hand zuordnen - gedacht fuer die wenigen Posten, die keine Quelle
    kennt.

    Reihenfolge: Hand > Auftragsnummer (gebuchte Gruppe lt. KPTM/PFAK) > Spalte
    'Produktgruppe' des Werkstattbestands > Artikelnummer. Dieselbe Reihenfolge bildet
    die Formel im Blatt 'Daten_PWBS_*' ab.

    Liefert (werte_je_produktgruppe, statistik). Die Statistik enthaelt auch die Liste
    der nicht zuordenbaren Auftraege, nach Wert absteigend - so lassen sich gezielt die
    groessten zuerst von Hand zuordnen.
    Nicht zuordenbare Zeilen werden danach unterschieden, ob sie eine Artikelnummer
    tragen: Serviceauftraege (Monteureinsatz, Inbetriebnahme, Schulung) haben keine,
    weil nichts gefertigt wird - sie zaehlen auch in der manuellen Referenzrechnung
    nicht zum Werkstattbestand. Ein nicht zuordenbarer Auftrag MIT Artikelnummer ist
    dagegen eine echte Luecke in der PFAK-Abdeckung.
    """
    df = lies_penta(pwbs_pfad)
    df.columns = [c.strip() for c in df.columns]
    _pruefe_datei(df, "PWBS", pwbs_pfad)
    je_pg = {}
    zugeordnet = service = luecke = 0.0
    ueber_artikel = service_mit_quelle = 0.0
    offene_auftraege = []
    service_verdacht = []
    manuell = manuelle_zuordnung or {}
    karte = karte or {}
    # Seit Oktober 2026 liefert die Buchhaltung den Werkstattbestand mit der Spalte
    # 'Produktgruppe' - die aktuelle Gruppe des Auftrags, also schon die neue Struktur.
    # Dann braucht es weder PFAK noch den Umweg ueber die KPTM-Kostenobjekte.
    spalte = df["Produktgruppe"] if "Produktgruppe" in df.columns else [None] * len(df)
    ueber_spalte = 0.0

    def neue_struktur(pg, art):
        # Aeltere Quellen (PFAK, alte PWBS-Exporte) kennen nur die alte Gruppe.
        ziel = karte.get(_schluessel(art))
        if ziel and ziel[0] == pg:
            pg = ziel[1]
        return (aufgeloest or {}).get(pg, pg) if karte else pg

    for rm, art, bez, wert, pg_spalte in zip(df["Rückmeldenummer"], df["Artikelnummer"],
                                             df["Bezeichnung"], df["Offener Wert"], spalte):
        wert = float(wert or 0.0)
        if _schluessel(art) is None:
            # Serviceauftrag (Monteureinsatz, Inbetriebnahme, Schulung): kein Artikel,
            # also nichts gefertigt. Zaehlt in der Referenzrechnung nicht zum
            # Werkstattbestand und wird deshalb GAR NICHT zugeordnet - auch dann nicht,
            # wenn eine Quelle eine Produktgruppe dafuer kennt. Ohne diese Sperre wandern
            # z.B. drei Monteureinsaetze nach PB0005 und verfaelschen die Sparte.
            service += wert
            # Ein Serviceauftrag, dessen Bezeichnung ein Fertigungsteil benennt, ist
            # ein Verdachtsfall: er wird still abgezogen, obwohl moeglicherweise doch
            # gefertigt wurde. Der Betrag waechst (0,2 % im April, 0,6 % im August),
            # deshalb wird er gezaehlt statt uebersehen.
            if _schluessel(rm) in nach_auftrag or _schluessel(rm) in manuell:
                service_mit_quelle += wert
                service_verdacht.append({
                    "auftrag": _schluessel(rm),
                    "bezeichnung": str(bez).strip() if bez is not None else "",
                    "wert": round(wert, 2),
                })
            continue
        von_hand = manuell.get(_schluessel(rm))
        aus_spalte = _schluessel(pg_spalte)
        # Die Gruppe, auf die der Auftrag GEBUCHT ist (Hand bzw. KPTM), geht der Spalte
        # 'Produktgruppe' des Werkstattbestands vor: die Spalte zeigt die heutige Gruppe
        # des Artikels, und die weicht waehrend der Reorganisation von der Buchung ab
        # (Jan-Sep 2026: 14 Auftraege). Die Spalte bleibt Ersatz fuer Auftraege ohne
        # Buchung im Zeitraum.
        gebucht = von_hand or nach_auftrag.get(_schluessel(rm))
        ueber_auftrag = gebucht or aus_spalte
        pg = ueber_auftrag or nach_artikel.get(_schluessel(art))
        if pg is not None:
            pg = neue_struktur(pg, art)
        if aus_spalte and not gebucht:
            ueber_spalte += wert
        if pg is not None:
            je_pg[pg] = je_pg.get(pg, 0.0) + wert
            zugeordnet += wert
            if not ueber_auftrag:
                # Nur ueber die Artikelnummer zugeordnet. Dieser Weg ist schwaecher:
                # wird ein Artikel in einem spaeteren Export mehrdeutig, faellt die
                # Zuordnung weg - mehr Quelldaten koennen das Ergebnis also
                # verschlechtern. Deshalb wird sein Gewicht ausgewiesen.
                ueber_artikel += wert
        else:
            luecke += wert
            offene_auftraege.append({
                "auftrag": _schluessel(rm),
                "artikel": _schluessel(art),
                "bezeichnung": str(bez).strip() if bez is not None else "",
                "wert": round(wert, 2),
            })
    gesamt = zugeordnet + service + luecke
    offene_auftraege.sort(key=lambda x: -x["wert"])
    statistik = {
        "gesamt": round(gesamt, 2),
        "zugeordnet": round(zugeordnet, 2),
        "service": round(service, 2),
        "offen": round(luecke, 2),
        "luecke_quote": (luecke / gesamt if gesamt else 0.0),
        "service_quote": (service / gesamt if gesamt else 0.0),
        "offene_auftraege": offene_auftraege,
        "ueber_artikel": round(ueber_artikel, 2),
        "service_mit_quelle": round(service_mit_quelle, 2),
        "service_verdacht": sorted(service_verdacht, key=lambda x: -x["wert"])[:20],
        "mit_spalte": "Produktgruppe" in df.columns,
        "ueber_spalte": round(ueber_spalte, 2),
    }
    return je_pg, statistik


def bestandsveraenderung_ufe(pwbs_ende, pfak_pfade=None, anfangsbestand=None, pwbs_anfang=None,
                             max_luecke=0.01, jahr=None, kptm_df=None, manuelle_zuordnung=None,
                             karte=None, aufgeloest=None):
    """Bestandsveraenderung unfertige Erzeugnisse = Endbestand minus Anfangsbestand.

    Der Anfangsbestand kommt entweder aus einem PWBS-Export zum Periodenbeginn
    (sauberste Quelle) oder aus den hinterlegten Jahres-Anfangsbestaenden
    (config/werkstattbestand_anfang.json). Da eine kumulierte Periode immer am 01.01.
    beginnt, ist der Anfangsbestand eine Jahreskonstante und muss nur einmal je Jahr
    bestimmt werden.

    Verifiziert an den Sparten PB0003, PB0004 und PB0005: die existierten vor 2025
    nicht, ihr Anfangsbestand 2026 muss daher exakt der kumulierten UFE-Veraenderung
    2025 entsprechen - er tut es centgenau.

    Liefert (werte_je_produktgruppe, warnungen). Bei zu geringer Zuordnungsquote wird
    KEIN Ergebnis geliefert, sondern nur eine Warnung - lieber keine Zahl als eine falsche.

    Die Schwelle von 1 % ist an der Referenzrechnung geeicht, nicht geraten:
      Restluecke 0,1 %  ->  13 von 13 Sparten centgenau
      Restluecke 4,7 %  ->   6 Sparten falsch, groesste Abweichung fuenfstellig
    Die nicht zuordenbaren Auftraege verteilen sich also keineswegs gleichmaessig -
    schon wenige Prozent Luecke koennen eine einzelne Sparte deutlich verfaelschen.
    """
    nach_auftrag, nach_artikel, mehrdeutige_artikel, widersprueche = lade_produktgruppen_map(
        pfak_pfade, kptm_df=kptm_df)
    ende, statistik = werkstattbestand(pwbs_ende, nach_auftrag, nach_artikel, manuelle_zuordnung,
                                       karte=karte, aufgeloest=aufgeloest)
    # Kennzahlen zur Zuordnungsqualitaet sofort setzen, nicht erst am Ende: bei
    # blockierter Zeile wird frueh zurueckgekehrt, und dann fehlen genau die Angaben,
    # mit denen man die Blockade erklaeren koennte.
    statistik["mehrdeutige_artikel"] = len(mehrdeutige_artikel)
    statistik["widersprueche"] = len(widersprueche)
    statistik["hinweise"] = []
    luecke = statistik["luecke_quote"]

    warnungen = []
    if luecke > max_luecke and not statistik["mit_spalte"] and not pfak_pfade:
        # Ohne Produktgruppen-Spalte und ohne PFAK fehlt die Quelle - dann ist die
        # Abhilfe nicht "neuerer PFAK", sondern die richtige Auswertung.
        warnungen.append(
            f"{luecke:.1%} des Werkstattbestands sind keiner Sparte zuzuordnen: Die Datei "
            f"enthaelt keine Spalte 'Produktgruppe'. Bitte den Werkstattbestand in der "
            f"Form mit Produktgruppe verwenden, wie ihn die Buchhaltung seit Oktober 2026 "
            f"liefert - und zum Stichtag gezogen, nicht nachtraeglich.")
    elif luecke > max_luecke:
        warnungen.append(
            f"{luecke:.1%} des Werkstattbestands entfallen auf {len(statistik['offene_auftraege'])} "
            f"Fertigungsauftraege, die keiner Sparte zugeordnet werden konnten "
            f"(zulaessig: {max_luecke:.0%}). Grund: ein PFAK-Export enthaelt fast nur "
            f"fertiggemeldete Auftraege - wer am Stichtag noch laeuft, steht nicht darin. "
            f"Abhilfe, in dieser Reihenfolge: (1) den neuesten vorliegenden PFAK-Export "
            f"verwenden, (2) die groessten offenen Auftraege unten von Hand einer Sparte "
            f"zuordnen, (3) die Berechnung in ein bis zwei Monaten wiederholen - bis dahin "
            f"sind die Auftraege fertiggemeldet und damit eindeutig zugeordnet."
        )

    if pwbs_anfang:
        anfang, stat_anfang = werkstattbestand(pwbs_anfang, nach_auftrag, nach_artikel,
                                               manuelle_zuordnung, karte=karte,
                                               aufgeloest=aufgeloest)
        luecke_a = stat_anfang["luecke_quote"]
        if luecke_a > max_luecke:
            warnungen.append(
                f"Anfangsbestand: {luecke_a:.1%} nicht zuordenbar (siehe Hinweis oben)."
            )
    elif anfangsbestand:
        anfang = anfangsbestand
        # Die mitgelieferte Excel-Vorlage hat dieselben Spalten wie eine gepflegte
        # Datei, nur mit Nullen. Ungefuellt hochgeladen sieht sie wie ein
        # Anfangsbestand von 0,00 aus: die Zeile zeigt dann den vollen Endbestand
        # statt der Veraenderung - eine plausibel aussehende Zahl, die um den
        # gesamten Anfangsbestand danebenliegt. Deshalb hier abbrechen.
        if not any(abs(w) > 0.005 for w in anfang.values()):
            warnungen.append(
                f"Der hinterlegte Anfangsbestand ist durchgehend 0,00 EUR - das ist die "
                f"unausgefuellte Vorlage, keine gepflegten Werte. Gerechnet wuerde sonst "
                f"{sum(ende.values()):,.2f} EUR, also der volle Endbestand statt der "
                f"Veraenderung. Abhilfe: in der Vorlage je Sparte den Werkstattbestand "
                f"zum 01.01. eintragen und die Datei erneut laden."
            )
    else:
        warnungen.append(
            f"Kein Anfangsbestand fuer das Geschaeftsjahr {jahr or '(unbekannt)'} in den "
            f"Betriebsparametern. Der Anfangsbestand ist der Werkstattbestand zum "
            f"01.01.{jahr or 'JJJJ'}: Er ergibt sich aus dem Abschluss des Vorjahres "
            f"(Stichtag 01.01.) - dort bietet das Tool nach der Berechnung an, ihn als "
            f"betriebsparameter.xlsx zu sichern. Alternativ in der Kommandozeile mit "
            f"--pwbs-anfang den Werkstattbestand zum 01.01. angeben."
        )
        return None, warnungen, statistik

    if warnungen:
        return None, warnungen, statistik

    hinweise = statistik["hinweise"]
    werte = {pg: ende.get(pg, 0.0) - anfang.get(pg, 0.0) for pg in set(anfang) | set(ende)}

    # Derselbe Fehler in klein: ist der Anfangsbestand nur fuer einen Teil der
    # Sparten gepflegt, zeigen die uebrigen den vollen Endbestand statt der
    # Veraenderung. Das faellt in der Summe nicht auf, verfaelscht aber genau die
    # betroffenen Sparten - deshalb namentlich nennen.
    # Ein ausdruecklich hinterlegtes 0,00 ist eine Aussage ("hatte keinen Bestand"),
    # kein fehlender Wert - gemeldet wird nur, wo die Sparte gar nicht vorkommt.
    ohne_anfang = sorted(pg for pg, wert in ende.items()
                         if abs(wert) > 0.005 and pg not in anfang)
    if ohne_anfang:
        hinweise.append(
            f"Für {len(ohne_anfang)} Sparte(n) ist kein Anfangsbestand hinterlegt: "
            f"{', '.join(ohne_anfang[:6])}{' ...' if len(ohne_anfang) > 6 else ''}. "
            f"Deren Zeile zeigt deshalb den vollen Endbestand statt der Veränderung. "
            f"Richtig ist das nur, wenn die Sparte am 01.01. tatsächlich keinen "
            f"Werkstattbestand hatte (z.B. erst danach eingeführt)."
        )

    # Die Prozentschwelle schuetzt die Summe, nicht die einzelne Sparte. Die nicht
    # zuordenbaren Auftraege verteilen sich nicht gleichmaessig - erfahrungsgemaess
    # landet rund die Haelfte der Luecke in EINER Sparte, im Extremfall alles.
    # Ist die Luecke groesser als der kleinste ausgewiesene Wert, kann diese Sparte
    # also komplett falsch sein, obwohl die Quote eingehalten ist.
    betroffen = sorted((abs(v), pg) for pg, v in werte.items() if abs(v) > 0.005)
    if betroffen and statistik["offen"] > betroffen[0][0]:
        kleinere = [pg for betrag, pg in betroffen if betrag < statistik["offen"]]
        hinweise.append(
            f"Die Zuordnungslücke von {_de(statistik['offen'])} EUR ist größer als der "
            f"ausgewiesene Wert von {len(kleinere)} Sparte(n): {', '.join(kleinere[:6])}"
            f"{' ...' if len(kleinere) > 6 else ''}. Die Quote von "
            f"{_de(statistik['luecke_quote'] * 100)[:-1].rstrip(',')} % ist zwar eingehalten, "
            f"aber sie misst am Gesamtbestand - diese Sparten können dennoch vollständig "
            f"falsch sein. Für belastbare Werte bei diesen Sparten die größten offenen "
            f"Aufträge von Hand zuordnen.", )
        statistik["kleiner_als_luecke"] = kleinere

    if widersprueche:
        beispiele = ", ".join(f"{a} ({x} vs. {y})" for a, x, y in widersprueche[:4])
        hinweise.append(
            f"{len(widersprueche)} Fertigungsauftrag/-auftraege werden von KPTM und PFAK "
            f"unterschiedlichen Sparten zugeordnet: {beispiele}"
            f"{' ...' if len(widersprueche) > 4 else ''}. Das Tool folgt PFAK. Bisher trat "
            f"dieser Fall nie auf - tritt er nun auf, verschieben sich Bestaende zwischen "
            f"Sparten. Bitte im ERP klaeren, welche Zuordnung gilt."
        )


    # Der Bestand zum Stichtag ist zugleich der Anfangsbestand der Folgeperiode.
    # Faellt der Stichtag auf den 01.01., ist es der des naechsten Geschaeftsjahres -
    # damit schreibt sich die Jahreskonstante aus den Daten selbst fort.
    statistik["bestand_stichtag"] = {pg: round(v, 2) for pg, v in sorted(ende.items())}
    return werte, warnungen, statistik


def build_spartenrechnung(df, mapping, pg_config):
    kz = mapping["kostenart_zu_zeile"]
    kst_prefix = mapping["kostenstellen_praefix"]
    kst_wertart = mapping.get("kostenstellen_wertart", "ISWF")
    fracht_suffix = mapping.get("umbuchung_fracht_suffix", "0100")

    known_order = list(pg_config["reihenfolge"])
    seen_pg = sorted(x for x in df["Produktgruppe"].dropna().unique())
    extra_pg = [p for p in seen_pg if p not in known_order]
    produktgruppen = known_order + extra_pg

    zeilen = list(mapping["zeilen_reihenfolge"])
    result = pd.DataFrame(0.0, index=zeilen, columns=produktgruppen)
    unmapped_kostenarten = set()
    # Auffangregel nach Kontenpraefix. Die Einzelliste kennt nur die Konten, die im
    # Kalibrierungszeitraum vorkamen; ohne diese Regel fallen spaeter auftauchende
    # Konten still heraus. Per Regel zugeordnete werden gesondert gemeldet.
    praefix_regel = mapping.get("kostenart_praefix_regel", {})
    per_regel = set()
    # Eine Fracht-Umbuchung, die ihr Ziel nicht findet, bucht still auf die operative
    # Sparte zurueck und verfaelscht deren Deckungsbeitrag. Heute existiert fuer jede
    # vorkommende Familie ein *0100-Objekt - bei einer neuen Familie waere das anders,
    # und ohne diese Erfassung faellt es niemandem auf.
    fracht_ohne_ziel = {}
    # Betraege je nicht zugeordneter Kostenart. Der blosse Kontocode sagt nicht, ob
    # 50 EUR oder 50.000 EUR fehlen - genau das entscheidet aber, wie dringend es ist.
    unmapped_betrag = {}

    for _, row in df.iterrows():
        ka = row["Kostenart"]
        pg = row["Produktgruppe"]
        wert = row["Wert/Menge"]
        if pg is None or pg not in produktgruppen:
            continue
        if ka in kz:
            regel = kz[ka]
            ziel_pg = pg
            if regel.get("umbuchung") == "fracht":
                ziel_pg = pg[:2] + fracht_suffix
                if ziel_pg not in produktgruppen:
                    # Kein Fracht-Kostenobjekt vorhanden: auf der Sparte lassen und merken.
                    fracht_ohne_ziel[pg] = fracht_ohne_ziel.get(pg, 0.0) + float(wert)
                    ziel_pg = pg
            result.loc[regel["zeile"], ziel_pg] += regel["sign"] * wert
        elif ka.startswith(kst_prefix):
            # Nur Wertart ISWF (Ist-WERT in EUR). ISMF waere die Ist-MENGE in Stunden
            # und darf nicht in den Eurobetrag addiert werden.
            if row["Wertart"] == kst_wertart:
                result.loc["fek", pg] += wert
        elif ka[:1] in praefix_regel:
            regel = praefix_regel[ka[:1]]
            result.loc[regel["zeile"], pg] += regel["sign"] * wert
            per_regel.add(ka)
        else:
            unmapped_kostenarten.add(ka)
            unmapped_betrag[ka] = unmapped_betrag.get(ka, 0.0) + float(wert)

    result["Summe"] = result.sum(axis=1)
    result = result[["Summe"] + produktgruppen]
    result = neu_berechnen(result)

    befunde = {
        "unmapped_betrag": {k: round(v, 2) for k, v in sorted(unmapped_betrag.items())},
        "fracht_ohne_ziel": {k: round(v, 2) for k, v in sorted(fracht_ohne_ziel.items())},
    }
    return result, produktgruppen, unmapped_kostenarten, sorted(per_regel), befunde


def fertigungsstunden(df, mapping):
    """Baut die Auswertung der Fertigungsauftraege je Kostenstelle.

    Entspricht der bisher manuell gepflegten Zusatzdatei "Auswertung
    Fertigungsauftraege - geleistete Stunden". Alle Spalten wurden gegen diese
    Datei geprueft (14/14 Kostenstellen centgenau):
      FEK inkl. Gemeinkosten = Kostenart K2*, Wertart ISWF
      davon unproduktive Gemeinkosten = zusaetzlich Auftragsart '5'
        (Kostenstellen-Sammelauftraege)
      produktive FEK = Differenz der beiden
      davon Lager-/Kundenauftraege = Aufteilung nach Auftragsart
      Stunden = Wertart ISMF (Ist-Menge), direkt aus den Daten

    Die Stunden werden bewusst NICHT als Betrag/Stundensatz gerechnet: ISMF fuehrt
    sie bereits, das ist exakt und spart die Division durch einen Satz, der selbst
    nur abgeleitet ist.
    """
    kst_prefix = mapping["kostenstellen_praefix"]
    kst_wertart = mapping.get("kostenstellen_wertart", "ISWF")
    satz = mapping.get("standard_stundensatz")
    art_gemein = set(mapping.get("auftragsart_gemeinkosten", ["5"]))
    art_lager = set(mapping.get("auftragsart_lager", ["LAG", "DPL", "PBL", "INN", "ITL", "SRL"]))

    kst = df[df["Kostenart"].str.startswith(kst_prefix)]
    stunden_je_kst = kst[kst["Wertart"] == "ISMF"].groupby("Kostenart")["Wert/Menge"].sum()
    sub = kst[kst["Wertart"] == kst_wertart]
    zeilen = []
    for ka, grp in sub.groupby("Kostenart"):
        bez = grp["KOSTENART_BEZ"].iloc[0] if len(grp) else ""
        gesamt = grp["Wert/Menge"].sum()
        gemein = grp[grp["Auftragsart"].isin(art_gemein)]["Wert/Menge"].sum()
        produktiv = gesamt - gemein
        lager = grp[grp["Auftragsart"].isin(art_lager)]["Wert/Menge"].sum()
        kunde = produktiv - lager
        zeilen.append({
            "Kostenstelle": ka,
            "Bezeichnung": str(bez).strip(),
            "FEK inkl. Gemeinkosten": gesamt,
            "davon unproduktive Gemeinkosten": gemein,
            "produktive FEK": produktiv,
            "davon Lageraufträge": lager,
            "davon Kundenaufträge": kunde,
            # ISMF fuehrt die Stunden direkt; der Satz ist nur der Rueckfall,
            # falls ein Export die Mengen-Wertart einmal nicht mitliefert.
            "Stunden gesamt": float(stunden_je_kst.get(ka, gesamt / satz if satz else 0.0)),
        })
    return pd.DataFrame(zeilen).sort_values("Kostenstelle")


def unproduktive_gemeinkosten(df, mapping):
    """Summe der Fertigungsstunden, die geleistet, aber auf keinen Kostentraeger
    gebucht wurden (Auftragsart '5' = Kostenstellen-Sammelauftraege).

    Das ist inhaltlich die FEK-Unterdeckung: Stunden, die bezahlt, aber ueber den
    Stundensatz nicht auf Produkte verrechnet wurden. Nachgewiesen an der
    Referenzrechnung Q1 2026:
      verrechnete FEK + unproduktive Gemeinkosten stimmen mit der Zeile
      'FEK nach Deckungsdifferenz (IST)' auf 0,25 % ueberein.

    Der Betrag wird hier NICHT auf die Sparten verteilt - das tut die Buchhaltung im
    Monatsabschluss je Kostenstelle; das Ergebnis kommt ueber die Kostentraeger-
    Auswertung herein (siehe lade_deckungsdifferenzen). Aus dem KPTM allein liess sich
    kein Schluessel finden, der die Referenz reproduziert. Die Zahl bleibt als
    Information stehen: Sie zeigt, WORAUS die FEK-Unterdeckung besteht.
    """
    kst_prefix = mapping["kostenstellen_praefix"]
    kst_wertart = mapping.get("kostenstellen_wertart", "ISWF")
    arten = set(mapping.get("auftragsart_gemeinkosten", ["5"]))
    satz = mapping.get("standard_stundensatz")
    sub = df[(df["Kostenart"].str.startswith(kst_prefix))
             & (df["Wertart"] == kst_wertart)
             & (df["Auftragsart"].isin(arten))]
    betrag = float(sub["Wert/Menge"].sum())
    return betrag, (betrag / satz if satz else 0.0)


DD_ZEILEN = ("fek_deckungsdifferenz", "mgk_deckungsdifferenz", "vvgk_deckungsdifferenz")


def umlage_regel(kostenart, regeln):
    """Zu welcher Zeile gehoert eine Umlage-Kostenart ('U' + Kostenstellennummer)?

    Es gilt der laengste passende Praefix der Kostenstellennummer: '218' (QS, gehoert
    laut Referenz zur MGK-Differenz) muss vor '21' (uebrige Fertigung) greifen.
    Liefert {"zeile", "sign"} oder None, wenn keine Regel passt - der Aufrufer meldet
    das dann.
    """
    nummer = str(kostenart)[1:]
    treffer = [p for p in regeln if nummer.startswith(p)]
    return regeln[max(treffer, key=len)] if treffer else None


def lade_deckungsdifferenzen(pfad, mapping):
    """Liest die Deckungsdifferenzen je Sparte aus der Kostentraeger-Auswertung.

    Die Buchhaltung verteilt im Monatsabschluss die Ueber-/Unterdeckung jeder
    Kostenstelle ('Umlage DDIFF <Kostenstelle>', Kostenart 'U' + Kostenstellennummer)
    auf die Kostentraeger. Die Kostentraeger-Auswertung zeigt das Ergebnis bereits je
    Produktgruppe - es muss also nichts geschaetzt oder verteilt werden, nur jede
    Kostenstelle dem richtigen der drei Toepfe zugeordnet werden (siehe
    'deckungsdifferenz_umlage' in der Konfiguration).

    Vorher war das die groesste Luecke der Rechnung: der KPTM-Export enthaelt nur die
    VERRECHNETE Seite, die Ist-Kosten der Kostenstellen kommen darin nicht vor. Mehrere
    Verteilungsmodelle liessen sich nicht gegen die Referenz bestaetigen - kein Wunder,
    die Verteilung laeuft je Kostenstelle ueber deren eigene Bezugsgroessen.

    Verifiziert an Q1 2026: FEK- und MGK-Deckungsdifferenz 16 von 16 Kostentraegern
    centgenau, VVGK 13 von 16. Die drei uebrigen widersprechen
    in der Referenz ihrer eigenen Zeile 'VVGK nach Deckungsdifferenz' - dort, wo es ins
    DB III einfliesst, stimmt die Referenz mit diesem Wert ueberein.

    Liefert (werte, befund):
      werte  = {zeile: {kostenobjekt: betrag}}   - auch Nicht-Sparten wie 712000
      befund = {"unbekannt": {kostenart: betrag}, "zeilen_umlage": n,
                "zeilen_gesamt": n, "perioden": (von, bis, jahr_von, jahr_bis),
                "umlage_df": DataFrame der Umlagezeilen (fuer das Datenblatt)}
    """
    df = lies_penta(pfad)
    df.columns = [str(c).strip() for c in df.columns]
    _pruefe_datei(df, "Kostenträger", pfad)

    # Die Kostenstellen-Auswertung hat exakt dieselben Spalten und kommt im selben
    # Paket von der Buchhaltung - eine Verwechslung ist naheliegend und fiele an den
    # Spalten nicht auf. Erkennbar ist sie nur am Objekttyp.
    objekttypen = set(df["Objekttyp"].dropna().astype(str).str.strip())
    if objekttypen and objekttypen != {"Kostenträger"}:
        name = os.path.basename(str(pfad))
        raise ValueError(
            f"'{name}' enthaelt Objekttyp {', '.join(sorted(objekttypen))} - gebraucht wird "
            f"die Auswertung nach KOSTENTRAEGERN, nicht nach Kostenstellen. Beide Dateien "
            f"kommen gemeinsam von der Buchhaltung und haben dieselben Spalten.")

    df["Kostenart"] = df["Kostenart"].astype(str).str.strip()
    df["Kostenobjekt"] = (df["Kostenobjekt"].astype(str).str.strip()
                          .str.replace(r"\.0$", "", regex=True))
    df["Wert/Menge"] = pd.to_numeric(df["Wert/Menge"], errors="coerce").fillna(0.0)

    regeln = mapping.get("deckungsdifferenz_umlage", {})
    umlage = df[df["Kostenart"].str.match(r"^U\d")].copy()
    treffer = umlage["Kostenart"].map(lambda ka: umlage_regel(ka, regeln))
    umlage["Zeile"] = treffer.map(lambda r: r["zeile"] if r else "")
    umlage["Vorzeichen"] = treffer.map(lambda r: r["sign"] if r else 0)

    umlage["Ziel"] = umlage["Kostenobjekt"]
    umlage["Anteil"] = 1.0
    werte = _umlage_werte(umlage)
    unbekannt = {}
    for ka, betrag in umlage[umlage["Zeile"] == ""].groupby("Kostenart")["Wert/Menge"].sum().items():
        unbekannt[ka] = float(betrag)

    return werte, {
        "unbekannt": {k: round(v, 2) for k, v in unbekannt.items() if abs(v) > 0.005},
        "zeilen_umlage": len(umlage),
        "zeilen_gesamt": len(df),
        "perioden": _perioden_aus_daten(df),
        "umlage_df": umlage,
    }


def _umlage_werte(umlage):
    """{zeile: {ziel: summe}} aus den Umlagezeilen - Wert x Vorzeichen x Anteil.
    Dieselbe Rechnung steht als Formel im Blatt 'Daten_Kostentraeger'."""
    werte = {}
    gueltig = umlage[umlage["Zeile"] != ""]
    betrag = gueltig["Wert/Menge"] * gueltig["Vorzeichen"] * gueltig["Anteil"]
    for (zeile, ziel), summe in betrag.groupby([gueltig["Zeile"], gueltig["Ziel"]]).sum().items():
        werte.setdefault(zeile, {})[ziel] = float(summe)
    return werte


def _umlage_aufteilen(umlage, anteile):
    """Teilt Umlagezeilen, die auf einer alten (geteilten) Produktgruppe gebucht sind,
    auf die neuen Gruppen auf: aus einer Zeile werden mehrere mit je einem Anteil.
    Liefert (aufgeteilte Zeilen, Menge der (zeile, neue_pg) mit Anteil < 100 %)."""
    teile, naeherung = [], set()
    for _, z in umlage.iterrows():
        aufteilung = anteile.get(z["Zeile"], {}).get(z["Kostenobjekt"])
        if not aufteilung:
            teile.append(z)
            continue
        for ziel, anteil in aufteilung.items():
            neu = z.copy()
            neu["Ziel"], neu["Anteil"] = ziel, anteil
            teile.append(neu)
            if anteil < 0.9999:
                naeherung.add((z["Zeile"], ziel))
    return pd.DataFrame(teile).reset_index(drop=True), naeherung


def zuschlagssaetze(result, mapping):
    """Rechnet die Zuschlagssaetze der Kalkulation aus dem Ergebnis zurueck.

    Die Kalkulation des Hauses wurde vollstaendig an der Referenzrechnung Q1 2026
    verifiziert:
      FEK  = Fertigungsstunden x Standardstundensatz (Betriebsparameter)
      MGK  = fester Satz auf (Material + Fremdleistungen)       11/14 Sparten < 50 Cent
      VVGK = fester Satz auf (Material + Fremdleistungen + MGK + FEK), also auf die
             Herstellkosten                                     14/15 Sparten < 1 EUR

    Die Saetze werden NICHT zur Berechnung verwendet - alle Werte kommen direkt aus dem
    ERP. Sie dienen als Plausibilitaetsprobe: weicht ein Satz ab, wurde er entweder im
    ERP geaendert oder die Kontenzuordnung stimmt nicht mehr.
    """
    def wert(zeile):
        return float(result.loc[zeile, "Summe"])

    saetze, hinweise = {}, []

    basis_mgk = wert("material") + wert("fremdleistungen")
    if abs(basis_mgk) >= 1:
        saetze["mgk"] = wert("mgk") / basis_mgk
    basis_vvgk = basis_mgk + wert("mgk") + wert("fek")
    if abs(basis_vvgk) >= 1:
        saetze["vvgk"] = wert("vvgk") / basis_vvgk

    # Plausibilitaetsprobe OHNE hinterlegten Sollwert. Frueher wurde gegen eine
    # erwartete Quote in der Konfiguration geprueft; die wurde entfernt, weil der
    # Satz Kalkulationsinnenleben ist und das Repository oeffentlich - damit lief
    # die Probe unbemerkt leer. Der Ersatz braucht keine Geheimzahl: Das Haus
    # verrechnet EINEN Satz fuer alle Sparten, also muessen alle Sparten denselben
    # Satz zeigen. Streut er, stimmt die Kontenzuordnung nicht mehr oder der Satz
    # wurde unterjaehrig geaendert - beides faellt so auf, ohne Sollwert.
    sparten = [c for c in result.columns if c != "Summe"]
    for schluessel, bez, grundlage, zaehler, nenner in (
        ("mgk", "MGK", "Material + Fremdleistungen",
         "mgk", ["material", "fremdleistungen"]),
        ("vvgk", "VVGK", "Herstellkosten = Material + Fremdleistungen + MGK + FEK",
         "vvgk", ["material", "fremdleistungen", "mgk", "fek"]),
    ):
        if schluessel not in saetze:
            continue
        je_sparte = {}
        for pg in sparten:
            basis = sum(float(result.loc[z, pg]) for z in nenner)
            if abs(basis) < 1000:      # Kleinstsparten verrauschen die Quote
                continue
            je_sparte[pg] = float(result.loc[zaehler, pg]) / basis
        if len(je_sparte) < 3:
            continue
        toleranz = mapping.get(f"{schluessel}_zuschlagssatz_toleranz", 0.005)
        erwartet = mapping.get(f"{schluessel}_zuschlagssatz_erwartet")
        if erwartet and abs(saetze[schluessel] - erwartet) > toleranz:
            hinweise.append(
                f"Der rechnerische {bez}-Zuschlagssatz betraegt {saetze[schluessel]:.2%}, "
                f"hinterlegt sind {erwartet:.2%} (Bemessungsgrundlage: {grundlage}). Entweder "
                f"wurde der Satz im ERP geaendert - dann bitte in den Betriebsparametern "
                f"nachziehen - oder die Kontenzuordnung stimmt nicht mehr."
            )
        tief = min(je_sparte, key=je_sparte.get)
        hoch = max(je_sparte, key=je_sparte.get)
        spanne = je_sparte[hoch] - je_sparte[tief]
        if spanne > toleranz:
            hinweise.append(
                f"Der rechnerische {bez}-Zuschlagssatz ist nicht einheitlich: "
                f"{je_sparte[tief]:.2%} bei {tief} gegen {je_sparte[hoch]:.2%} bei {hoch} "
                f"(Bemessungsgrundlage: {grundlage}). Das Haus verrechnet einen einzigen "
                f"Satz - eine Streuung bedeutet also, dass die Kontenzuordnung nicht mehr "
                f"passt oder der Satz unterjaehrig geaendert wurde."
            )
    return saetze, hinweise


def neu_berechnen(result):
    """Berechnet alle abgeleiteten Zeilen (Zwischensummen) neu.

    Als eigene Funktion, damit sie auch nach nachtraeglichem Einsetzen von Werten
    (z.B. Bestandsveraenderung UFE aus dem Werkstattbestand) erneut laufen kann.
    """
    sparten = [c for c in result.columns if c != "Summe"]
    result["Summe"] = result[sparten].sum(axis=1)

    result.loc["betriebsleistung"] = (
        result.loc["erloese"] + result.loc["bestand_fe"] + result.loc["bestand_ufe"] + result.loc["aktivierte_eigenleistung"]
    )
    result.loc["rohmarge_1"] = (
        result.loc["betriebsleistung"] - result.loc["material"] - result.loc["fremdleistungen"]
        - result.loc["personalaufwand"] - result.loc["sbe"] - result.loc["sba"] - result.loc["afa"]
    )
    result.loc["fek_nach_dd"] = result.loc["fek"] + result.loc["fek_deckungsdifferenz"]
    result.loc["db1"] = result.loc["rohmarge_1"] - result.loc["fek_nach_dd"]
    result.loc["mgk_nach_dd"] = result.loc["mgk"] + result.loc["mgk_deckungsdifferenz"]
    result.loc["db2"] = result.loc["db1"] - result.loc["mgk_nach_dd"]
    result.loc["vvgk_nach_dd"] = result.loc["vvgk"] + result.loc["vvgk_deckungsdifferenz"]
    result.loc["db3_vor_sonderposten"] = result.loc["db2"] - result.loc["vvgk_nach_dd"]
    result.loc["db3"] = result.loc["db3_vor_sonderposten"] + result.loc["sonderposten"]
    return result



# Farben des Diagramms: die Reihenfolge des klassischen Office-Themas, also genau
# die Toene, die auch das Diagramm der Referenzmappe zeigt (accent1..accent6, dann
# das aufgehellte Rot fuer VVGK), und Gold fuer den Deckungsbeitrag. So sieht das
# erzeugte Diagramm aus wie die Auswertungen, die im Haus ohnehin herumliegen.
DIAGRAMM_SERIEN = [
    ("material", "Material", "4F81BD"),
    ("fremdleistungen", "Fremdleistungen", "C0504D"),
    ("personalaufwand", "Personalaufwand", "9BBB59"),
    ("sba", "sonst. betriebl. Aufwand", "8064A2"),
    ("fek_nach_dd", "Fertigungseinzelkosten", "4BACC6"),
    ("mgk_nach_dd", "Materialgemeinkosten", "F79646"),
    ("vvgk_nach_dd", "Verwaltung & Vertrieb", "D99694"),
    ("db3", "Deckungsbeitrag III", "FFC000"),
]


def _ist_sparte(code):
    """Echte Produktgruppe (zwei Buchstaben + vier Ziffern)? Die Referenzrechnung
    fuehrt daneben Pseudo-Kostenobjekte wie 910000 fuer Abgrenzungen - die gehoeren
    nicht ins Diagramm, weil sie keine Sparte sind."""
    c = str(code)
    return len(c) == 6 and c[:2].isalpha() and c[2:].isdigit()


def _diagramm(ws, result, mapping, zeitraum, zeilen_nr, cols, start_zeile):
    """Gestapelte Balken je Sparte: woraus sich die Betriebsleistung zusammensetzt.

    Entspricht dem Diagramm, das bisher von Hand gepflegt wurde. Die Datenreihen
    verweisen auf die Tabelle darueber, sind also keine Kopie: wer oben eine
    Deckungsdifferenz nachtraegt, sieht sie sofort im Bild.
    """
    # Nur Sparten mit Bewegung. Leere Spalten (noch nicht bebuchte oder ausgelaufene
    # Produktgruppen) machen das Bild schmal und unlesbar, ohne etwas auszusagen.
    sparten = [c for c in cols if _ist_sparte(c)
               and abs(float(result.loc["betriebsleistung", c])) > 0.005]
    if not sparten or not zeilen_nr:
        return start_zeile

    kopf = start_zeile + 1
    ws.cell(row=kopf - 1, column=1,
            value="Datengrundlage des Diagramms (verweist auf die Tabelle oben)").font = Font(
                italic=True, size=9, color="808080")
    ws.cell(row=kopf, column=1, value="Position").font = FONT_BOLD
    for j, pg in enumerate(sparten, start=2):
        z = ws.cell(row=kopf, column=j, value=str(ws.cell(row=4, column=cols.index(pg) + 2).value or pg))
        z.font = Font(bold=True, size=9)
        z.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[kopf].height = 30

    r = kopf + 1
    serien_zeilen = []
    for key, beschriftung, _farbe in DIAGRAMM_SERIEN:
        if key not in zeilen_nr:
            continue
        ws.cell(row=r, column=1, value=beschriftung)
        for j, pg in enumerate(sparten, start=2):
            sp = get_column_letter(cols.index(pg) + 2)
            z = ws.cell(row=r, column=j, value=f"={sp}{zeilen_nr[key]}")
            z.number_format = "#,##0"
        serien_zeilen.append((r, key))
        r += 1

    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.title = f"Spartenrechnung {zeitraum} – Kostenstruktur und Ergebnis je Sparte"
    # Kein Achsentitel: openpyxl setzt ihn waagerecht, dadurch lag das "EUR" quer
    # ueber den Zahlen der Achse. Die Einheit steht ohnehin in der Kopfzeile des
    # Blattes ("abs. Zahlen in EUR").
    ch.y_axis.numFmt = "#,##0"
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E8EAED"))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    # gapWidth ist der Abstand ZWISCHEN den Balken in Prozent der Balkenbreite:
    # je groesser, desto schmaler der Balken. 150 wie in der Referenzmappe.
    ch.gapWidth = 150
    ch.height = 11
    ch.width = max(20, 2.0 * len(sparten))
    ch.legend.position = "r"
    ch.legend.overlay = False

    kategorien = Reference(ws, min_col=2, max_col=len(sparten) + 1, min_row=kopf)
    farben = {k: f for k, _b, f in DIAGRAMM_SERIEN}
    for zeile, key in serien_zeilen:
        daten = Reference(ws, min_col=1, max_col=len(sparten) + 1, min_row=zeile, max_row=zeile)
        ch.add_data(daten, titles_from_data=True, from_rows=True)
        ch.series[-1].graphicalProperties = GraphicalProperties(solidFill=farben[key])
        ch.series[-1].graphicalProperties.ln = LineProperties(noFill=True)
    ch.set_categories(kategorien)
    ws.add_chart(ch, f"A{r + 2}")
    return r + 24


def write_output(result, mapping, zeitraum, out_path, bwa_ergebnis=None, bwa_sheet=None,
                 stunden_df=None, unproduktiv=None, herleitung=None, pg_bezeichnungen=None,
                 ufe_warnungen=None, zeilen_klasse=None, deckungsdifferenzen=None,
                 naeherung_zellen=None, struktur=None, abstimmung=None, guv=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Spartenrechnung"

    labels = mapping["zeilen_labels"]
    cols = list(result.columns)

    ws.cell(row=1, column=1, value=f"Perforator – Spartenrechnung {zeitraum} (automatisiert erzeugt)").font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value="abs. Zahlen in EUR")
    ws.cell(row=3, column=1, value="Kostenobjekt")
    for j, c in enumerate(cols, start=2):
        cell = ws.cell(row=3, column=j, value=c)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center")

    # Klartextnamen unter die Codes. Ohne sie liest niemand ausser dem Controlling
    # die Tabelle - 'DP0002' sagt nichts, 'DTH < 139,7' schon.
    bezeichnungen = pg_bezeichnungen or {}
    ws.cell(row=4, column=1, value="Bezeichnung").font = Font(italic=True, size=9)
    for j, c in enumerate(cols, start=2):
        cell = ws.cell(row=4, column=j, value=bezeichnungen.get(c, "" if c == "Summe" else c))
        cell.font = Font(italic=True, size=9)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[4].height = 26

    row_i = 5
    zeilen_nr = {}
    for key in mapping["zeilen_reihenfolge"] + [
        "betriebsleistung", "rohmarge_1", "fek_nach_dd", "db1", "mgk_nach_dd", "db2",
        "vvgk_nach_dd", "db3_vor_sonderposten", "db3",
    ]:
        if key not in result.index:
            continue
        zeilen_nr[key] = row_i
        label_cell = ws.cell(row=row_i, column=1, value=labels.get(key, key))
        klass = (zeilen_klasse or {}).get(key) or ZEILEN_KLASSE.get(key, "manual")
        # Eine zurueckgehaltene UFE-Zeile darf NICHT gruen erscheinen. Gruen heisst
        # laut Legende "automatisch berechnet, gegen die Referenz geprueft" - hier
        # steht aber eine 0,00, weil das Tool die Zahl bewusst verweigert. Wer nur
        # die Datei bekommt, laese sonst eine geprueft aussehende Null, wo eine
        # Zeile mit Millionenwirkung fehlt.
        if key == "bestand_ufe" and ufe_warnungen:
            klass = "manual"
            label_cell.value = labels.get(key, key) + "  (nicht berechenbar – siehe Hinweise)"
            label_cell.font = Font(bold=True, color="B3261E")
        if klass == "formula":
            label_cell.font = FONT_BOLD
        for j, c in enumerate(cols, start=2):
            val = float(result.loc[key, c])
            cell = ws.cell(row=row_i, column=j, value=round(val, 2))
            # Einzelne Zellen einer gruenen Zeile koennen eine Naeherung sein (anteilig
            # auf die neuen Produktgruppen aufgeteilt) - die werden gelb.
            cell.fill = (FILL_AUTO_PRUEF if klass == "auto" and (key, c) in (naeherung_zellen or ())
                         else FILLS[klass])
            cell.border = BORDER
            cell.number_format = "#,##0.00"
        row_i += 1

    if ufe_warnungen:
        z = ws.cell(row=2, column=3, value="⚠ Bestandsveränderung UFE wurde NICHT berechnet – "
                                           "die Zeile steht auf 0,00. Siehe Blatt Hinweise.")
        z.font = Font(bold=True, color="B3261E")

    row_i = _diagramm(ws, result, mapping, zeitraum, zeilen_nr, cols, row_i + 2)

    if bwa_ergebnis is not None:
        db3_summe = float(result.loc["db3", "Summe"])
        delta = db3_summe - bwa_ergebnis
        row_i += 1
        c = ws.cell(row=row_i, column=1, value=f"Ergebnis lt. GuV, kumuliert (BWA {bwa_sheet})")
        c.font = FONT_BOLD
        cell = ws.cell(row=row_i, column=2, value=round(bwa_ergebnis, 2))
        cell.fill = FILL_AUTO
        cell.border = BORDER
        cell.number_format = "#,##0.00"
        row_i += 1
        c = ws.cell(row=row_i, column=1, value="Δ DB III (Summe der Sparten) ./. Ergebnis lt. GuV")
        c.font = FONT_BOLD
        cell = ws.cell(row=row_i, column=2, value=round(delta, 2))
        cell.fill = FILL_FORMULA
        cell.border = BORDER
        cell.number_format = "#,##0.00"
        row_i += 1
        ws.cell(row=row_i, column=1, value=(
            "Enthaelt u.a. Sachkonto-Positionen ohne Produktgruppe (Grundstuecke, Abgrenzungen, "
            "periodenfremde Ertraege) - siehe Hinweise-Blatt 'Was die Spalte Summe nicht enthaelt'."
        )).font = Font(italic=True, size=9, color="808080")
        row_i += 1
        ws.cell(row=row_i, column=1,
                value="  Hinweis: Δ entspricht der Differenz, die bisher manuell erklaert wurde "
                      "(Sonderposten, 'keine Abrechnung BAB', 'fehlende Abbildung SuSa') - siehe Hinweise-Sheet.")

    # Information zur FEK-Unterdeckung: Hoehe bekannt, Verteilung auf die Sparten nicht.
    if unproduktiv:
        betrag, stunden = unproduktiv
        row_i += 2
        c = ws.cell(row=row_i, column=1, value="Nicht auf Kostenträger gebuchte Fertigungsstunden")
        c.font = FONT_BOLD
        cell = ws.cell(row=row_i, column=2, value=round(betrag, 2))
        cell.fill = FILL_AUTO_PRUEF
        cell.border = BORDER
        cell.number_format = "#,##0.00"
        row_i += 1
        ws.cell(row=row_i, column=1,
                value=f"  entspricht {stunden:,.0f} Stunden zu {mapping.get('standard_stundensatz')} EUR/h "
                      f"(Auftragsart 5, Kostenstellen-Sammelauftraege)")
        row_i += 1
        ws.cell(row=row_i, column=1,
                value="  Das ist inhaltlich die FEK-Unterdeckung. Die Hoehe ist belegt, die Verteilung "
                      "auf die einzelnen Sparten nicht - siehe Hinweise.")

    ws.column_dimensions["A"].width = 34
    for j in range(2, len(cols) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = "B4"

    leg_row = row_i + 2
    ws.cell(row=leg_row, column=1, value="Legende:").font = FONT_BOLD
    legende = [
        (FILL_AUTO, "automatisch aus Rohdaten berechnet (zellgenau gegen die Referenz geprueft)"),
        (FILL_AUTO_PRUEF, "automatisch berechnet, aber Naeherung – bitte stichprobenartig pruefen"),
        (FILL_NULL, "strukturell null – gehoert nicht auf die Sparten (siehe Hinweise)"),
        (FILL_MANUAL, "in den Rohdaten NICHT enthalten – manuell zu ergaenzen"),
        (FILL_FORMULA, "reine Zwischensumme (Formel aus anderen Zeilen)"),
    ]
    for i, (fill, text) in enumerate(legende, start=1):
        c = ws.cell(row=leg_row + i, column=1, value="  " + text)
        c.fill = fill

    if stunden_df is not None and len(stunden_df):
        ws3 = wb.create_sheet("Fertigungsstunden")
        ws3.cell(row=1, column=1,
                 value=f"Auswertung Fertigungsauftraege je Kostenstelle – {zeitraum}").font = Font(bold=True, size=12)
        # Die Stunden kommen aus der Wertart ISMF, nicht aus Betrag/Satz - der Satz
        # ist hier nur die Umrechnung, die das ERP selbst verwendet (ISWF = ISMF x Satz).
        satz = mapping.get("standard_stundensatz")
        ws3.cell(row=2, column=1, value=(
            "Euro-Spalten: Wertart ISWF. Spalte 'Stunden gesamt': Wertart ISMF, direkt aus "
            "den Rohdaten - nicht aus Betrag geteilt durch Satz."
            + (f" (Das ERP bewertet die Stunden mit {satz} EUR/h.)" if satz else "")))
        spalten = list(stunden_df.columns)
        for j, name in enumerate(spalten, start=1):
            c = ws3.cell(row=4, column=j, value=name)
            c.fill = FILL_HEADER
            c.font = FONT_HEADER
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        for i, (_, zeile) in enumerate(stunden_df.iterrows(), start=5):
            for j, name in enumerate(spalten, start=1):
                wert = zeile[name]
                c = ws3.cell(row=i, column=j, value=wert)
                c.border = BORDER
                if isinstance(wert, (int, float)):
                    c.number_format = "#,##0.00"
                    c.fill = FILL_AUTO
        summe_row = 5 + len(stunden_df)
        ws3.cell(row=summe_row, column=1, value="SUMME").font = FONT_BOLD
        for j, name in enumerate(spalten, start=1):
            if j <= 2:
                continue
            c = ws3.cell(row=summe_row, column=j, value=float(stunden_df[name].sum()))
            c.font = FONT_BOLD
            c.number_format = "#,##0.00"
            c.border = BORDER
        ws3.column_dimensions["A"].width = 14
        ws3.column_dimensions["B"].width = 22
        for j in range(3, len(spalten) + 1):
            ws3.column_dimensions[get_column_letter(j)].width = 17

    ws2 = wb.create_sheet("Hinweise")
    dd = deckungsdifferenzen or {}
    if dd.get("geladen"):
        dd_abschnitt = [
            "AUS DER KOSTENTRAEGER-AUSWERTUNG (gruen) - FEK-/MGK-/VVGK-Deckungsdifferenz:",
            "- Die Buchhaltung verteilt im Monatsabschluss die Ueber-/Unterdeckung jeder Kostenstelle",
            "  auf die Kostentraeger ('Umlage DDIFF <Kostenstelle>', Kostenart U + Kostenstellennummer).",
            "  Die Kostentraeger-Auswertung zeigt das Ergebnis bereits je Produktgruppe. Zugeordnet wird",
            "  nur, welche Kostenstelle in welchen Topf gehoert (Regel im Blatt 'Mapping'):",
            "     Einkauf/Lager (11...) und QS (218...)   ->  MGK-Deckungsdifferenz",
            "     uebrige Fertigung (21...)               ->  FEK-Deckungsdifferenz",
            "     Verwaltung (3...) und Vertrieb (4...)   ->  VVGK-Deckungsdifferenz",
            "  Geprueft an Q1 2026: FEK und MGK 16 von 16 Sparten centgenau, VVGK 13 von 16. Die drei",
            "  uebrigen widersprechen in der Referenz ihrer eigenen Zeile 'VVGK nach Deckungsdifferenz' -",
            "  dort, wo es ins DB III einfliesst, stimmt die Referenz mit diesem Wert ueberein.",
            "  Die einzelnen Umlagezeilen stehen im Blatt 'Daten_Kostentraeger'.",
        ] + (["", "  ACHTUNG:"] + ["  - " + h for h in dd.get("hinweise", [])]
             if dd.get("hinweise") else [])
    else:
        dd_abschnitt = [
            "MANUELL ZU ERGAENZEN (rosa) - FEK-/MGK-/VVGK-Deckungsdifferenz:",
            "- Gegenueberstellung der ueber die Kostentraeger VERRECHNETEN Betraege (die gruenen Zeilen",
            "  oben) mit den TATSAECHLICHEN Kostenstellenkosten. Der KPTM-Export enthaelt nur die",
            "  verrechnete Seite. Die Differenzen kommen aus der Kostentraeger-Auswertung der Buchhaltung",
            "  (Penta, Objekttyp Kostentraeger) - die wurde fuer diese Datei nicht mitgegeben.",
        ] + ["  - " + h for h in dd.get("hinweise", [])]
    struktur_abschnitt = []
    if struktur and struktur.get("fehlt"):
        struktur_abschnitt = [
            "!!! PRODUKTGRUPPEN: ALTE UND NEUE STRUKTUR GEMISCHT !!!",
            f"- Der Export enthaelt Buchungen auf inzwischen aufgeloesten Produktgruppen "
            f"({', '.join(struktur['alte'])}), es ist aber keine Umschluesselung hinterlegt.",
            "  Dieselbe Produktlinie steht dadurch in zwei Spalten. Abhilfe: die aktuelle",
            "  betriebsparameter.xlsx (mit Blatt 'Umschluesselung') laden.",
            "",
        ]
    elif struktur:
        struktur_abschnitt = [
            "NEUE PRODUKTGRUPPEN-STRUKTUR (Reorganisation 2026):",
            "- Seit Periode 4/2026 bucht das ERP auf die neuen Produktgruppen. Aeltere Buchungen des",
            "  Jahres stehen noch unter der alten Gruppe; sie werden je Artikel in die neue Struktur",
            "  umgeschluesselt (Liste 'Alte vs Neue Produktgruppen', Blatt 'Umschluesselung').",
            f"  Hier betroffen: {struktur['zeilen']} Buchungszeilen. Im Blatt 'Daten_KPTM' steht die",
            "  urspruenglich gebuchte Gruppe in der Spalte 'Produktgruppe (gebucht)' daneben.",
            "- Aufgeloeste Gruppen gehen an ihren Nachfolger (PB0003 Pilot Tubes -> PB0002).",
            "- Drill Pipes (DP) und Injektionstechnik (IT) sind von der Reorganisation nicht betroffen.",
            "",
            "NAEHERUNG (gelb) - was nur je Produktgruppe vorliegt, nicht je Artikel:",
            "- Deckungsdifferenzen, die die Buchhaltung noch auf einer alten, inzwischen geteilten Gruppe",
            "  gebucht hat, werden nach ihrer Bezugsbasis anteilig verteilt: FEK-Differenz nach FEK,",
            "  MGK-Differenz nach Material + Fremdleistungen, VVGK-Differenz nach Herstellkosten.",
            "- Der Anfangsbestand zum 01.01. der geteilten Gruppen ist im Verhaeltnis des",
            "  Werkstattbestands zum 01.01. je Auftrag aufgeteilt"
            + (f" ({', '.join(struktur.get('naeherung_ab', []))})." if struktur.get("naeherung_ab") else "."),
            "  Ueber die betroffenen Sparten zusammen sind die Werte exakt - geschaetzt ist nur die",
            "  Verteilung zwischen ihnen.",
            "",
        ]
    hinweise = [
        "Diese Datei wurde automatisiert erzeugt (lokal im Browser, ERP-Export KPTM_Wertsummen / Penta).",
        "",
        "AUTOMATISCH BERECHNET (gruen) - alle Zeilen wurden gegen die manuelle Referenzrechnung",
        "Q1 2026 ZELLGENAU je Sparte geprueft (nicht nur die Summe):",
        "- Erloese: Konten 4440*/4460*/4470* + Fracht-Erloese 4521*/4531*/4535*, Vorzeichen gedreht -> 16/16 Sparten centgenau",
        "- Bestandsveraenderung FE: Konten 4825* -> 11 centgenau, 1 Abweichung 20 EUR",
        "- Material: Konten 5300100 + 5320100 + 5800000 -> 13 centgenau, 1 Abweichung 30 EUR",
        "- Fremdleistungen: Konten 5900000 + 5901000 -> 8/8 centgenau",
        "- sbA: Konten 6420000/6651000/6666000/6681000/6825000 auf der Original-Produktgruppe,",
        "  Ausgangsfrachten 6740000/6741000 UMGEBUCHT auf das Fracht-Kostenobjekt derselben Sparten-",
        "  Familie (DP*->DP0100, PB*->PB0100, IT*->IT0100) -> 8/8 centgenau. Diese Umbuchung entspricht",
        "  dem Vermerk 'nach Umbuchung Frachtkosten' in der manuellen Referenzrechnung.",
        "- FEK: Ist-Werte der Fertigungskostenstellen (Kostenart K2*, NUR Wertart ISWF) je Produktgruppe",
        "  -> 11 centgenau, 3 Abweichungen unter 410 EUR. Hinweis: das ERP bewertet Fertigungsstunden",
        "  mit dem Standardstundensatz, es gilt ISWF = ISMF (Stunden) x Satz.",
        "- MGK: Konto C599199 -> 13 centgenau, 1 Abweichung 3 EUR",
        "- overhead Kosten (VVGK vor Deckungsdifferenz): Konto Q500100 -> 12 centgenau, 3 unter 122 EUR",
        "",
        "- Bestandsveraenderung UFE: Werkstattbestand(Stichtag) minus Werkstattbestand(01.01.), je Sparte.",
        "  Die Sparte je Auftrag steht im Werkstattbestand selbst (Spalte 'Produktgruppe'); aeltere",
        "  Exporte ohne diese Spalte werden ueber PFAK.RUECKMELDE_NR zugeordnet -> 13/13 Sparten centgenau.",
        "",
        *struktur_abschnitt,
        "STRUKTURELL NULL (grau) - diese Zeilen gehoeren nicht auf die Sparten:",
        "- Personalaufwand, sbE, Afa: In der Kostentraegerrechnung stecken die Personalkosten bereits in",
        "  den Verrechnungssaetzen (Stundensatz, MGK-Zuschlag, Overhead-Pool). Ein zusaetzlicher",
        "  Ausweis je Sparte waere Doppelzaehlung. In der manuellen Referenzrechnung stehen Werte daher",
        "  ausschliesslich auf den Sonder-Kostenobjekten 910000 'Abgrenzungen Walkenried' und 920000",
        "  'Weiterberechnung Valluhn' - ueber 7 geprueften Perioden gibt es genau eine Ausnahme auf einer",
        "  echten Sparte. Afa ist in allen Perioden null. sbE laut BWA Q1 2026 ebenfalls 0,00 EUR.",
        "- aktivierte Eigenleistung: in der gesamten Historie genau eine Buchung (Feb. 2024, im Dez. 2024",
        "  wieder aufgeloest); im ERP existiert dafuer gar kein Konto.",
        "",
        "WAS DIE SPALTE 'SUMME' NICHT ENTHAELT:",
        "Die manuelle Referenzrechnung fuehrt neben den echten Sparten (Produktgruppen) fuenf weitere",
        "Spalten fuer Sachkonto-Buchungen ohne Produktgruppe: 712000 'Photovoltaik Anlage', 721000",
        "'Grundstueck Valluhn', 910000 'Abgrenzungen Walkenried', 920000 'Weiterberechnung Valluhn',",
        "999999 'periodenfremder Ertrag'. Keiner dieser fuenf Codes kommt in KPTM als Produktgruppe vor -",
        "das Tool kann sie also strukturell nicht erzeugen, gleich welcher Zeitraum hochgeladen wird.",
        "Die Spalte 'Summe' ist deshalb die Summe UEBER DIE SPARTEN, nicht das Gesamtunternehmensergebnis.",
        "",
        "Der Vergleich mit der BWA weiter unten (falls hochgeladen) zeigt die Groessenordnung dieser und",
        "aehnlicher Posten fuer den jeweils aktuellen Zeitraum - das ist der eigentliche Zweck der Zeile",
        "'Δ zu erklaeren'.",
        "",
        *dd_abschnitt,
        "",
        "STRUKTURELL NULL (grau) - die Zeile gehoert dazu, hat je Sparte aber keinen Inhalt:",
        "- Sonderposten/Einmaleffekte: In der Referenzrechnung traegt diese Zeile in JEDER echten Sparte",
        "  0,00. Die Betraege (Photovoltaik, Grundstueck Valluhn) stehen ausschliesslich auf",
        "  Kostenobjekten, die keine Sparte sind. Eine Kostentraegerrechnung kann sie je Sparte gar",
        "  nicht ausweisen - hier fehlt also nichts, es ist nichts zu ergaenzen.",
        "",
        "Quelle: " + zeitraum,
    ]
    if ufe_warnungen:
        hinweise = [
            "!!! ACHTUNG: Bestandsveraenderung UFE wurde NICHT berechnet !!!",
            "",
            "Die Zeile steht auf 0,00 - das ist KEIN Ergebnis, sondern eine fehlende Zahl.",
            "Betriebsleistung, Rohmarge I und alle Deckungsbeitraege darunter sind dadurch",
            "um diesen Betrag zu niedrig. Der Grund:",
            "",
        ] + ["  " + w for w in ufe_warnungen] + ["", "-" * 100, ""] + hinweise

    for i, line in enumerate(hinweise, start=1):
        ws2.cell(row=i, column=1, value=line)
    ws2.column_dimensions["A"].width = 140

    if abstimmung:
        _abstimmung_blatt(wb, abstimmung, zeitraum)
    if guv:
        guv_pos = _guv_blatt(wb, guv, zeitraum)
        if herleitung is not None:
            # Die Herleitung setzt im Blatt 'Abgleich GuV' Formeln auf die Datenblaetter.
            herleitung = dict(herleitung, guv={"pos": guv_pos, "susa_pfad": guv["susa_pfad"]})

    # Rohdaten und Formeln anhaengen, damit jede Zahl rueckverfolgbar ist.
    # Schlaegt das fehl, bleibt die Datei mit den reinen Werten erhalten - die
    # Nachvollziehbarkeit ist ein Zusatz, kein Grund, das Ergebnis zu verlieren.
    if herleitung:
        try:
            import herleitung_excel
            herleitung_excel.baue_herleitungsmappe(wb, result, mapping, **herleitung)
        except Exception as e:
            ws2.cell(row=len(hinweise) + 2, column=1,
                     value=f"Hinweis: Die Datenblaetter zur Herleitung konnten nicht "
                           f"erzeugt werden ({e}). Die Werte oben sind davon unberuehrt.")

    wb.save(out_path)


def deckungsdifferenzen_einsetzen(result, kostentraeger_pfad, mapping, kptm_df, anteile=None):
    """Setzt die Deckungsdifferenzen (und Sonderposten) aus der Kostentraeger-Auswertung
    in das Ergebnis ein. Liefert (result, befund) - befund ist None, wenn die Datei
    nichts Verwertbares enthielt; dann bleiben die Zeilen wie sie waren (rosa).

    'anteile' (siehe aufteilungsanteile): Differenzen, die noch auf einer alten,
    inzwischen geteilten Produktgruppe gebucht sind, werden anteilig auf die neuen
    Gruppen verteilt. befund["naeherung"] nennt die Zellen, die so entstanden sind.
    """
    werte, befund = lade_deckungsdifferenzen(kostentraeger_pfad, mapping)
    befund["naeherung"] = set()
    if anteile:
        befund["umlage_df"], befund["naeherung"] = _umlage_aufteilen(befund["umlage_df"], anteile)
        werte = _umlage_werte(befund["umlage_df"])

    # Beide Dateien muessen denselben Zeitraum abdecken. Eine Kostentraeger-Auswertung
    # zu Q1 neben einem KPTM-Export Jan-Juni ergaebe Deckungsdifferenzen, die nicht zu
    # den uebrigen Zeilen passen - plausibel aussehend und still falsch.
    kptm_spanne, kt_spanne = _perioden_aus_daten(kptm_df), befund["perioden"]
    if kptm_spanne and kt_spanne and kptm_spanne != kt_spanne:
        def text(s):
            von, bis, j1, j2 = s
            return f"Periode {von}-{bis}/{j2}" if j1 == j2 else f"Periode {von}/{j1}-{bis}/{j2}"
        raise ValueError(
            f"Die Kostentraeger-Auswertung umfasst {text(kt_spanne)}, der KPTM-Export aber "
            f"{text(kptm_spanne)}. Beide muessen denselben Zeitraum abdecken - sonst passen "
            f"die Deckungsdifferenzen nicht zu den uebrigen Zeilen.")

    befund["hinweise"] = []
    if not befund["zeilen_umlage"]:
        # Die Umlage laeuft erst im Monatsabschluss. Ein frueher gezogener Export hat
        # die Zeilen noch nicht - dann waeren die Differenzen 0,00, und das saehe
        # genauso aus wie "keine Differenz".
        befund["hinweise"].append(
            "Die Kostentraeger-Auswertung enthaelt keine einzige Umlagezeile ('Umlage DDIFF', "
            "Kostenart U...). Vermutlich wurde sie vor der Abschlussbuchung der "
            "Deckungsdifferenzen gezogen. Die drei Zeilen bleiben deshalb leer - bitte den "
            "Export nach dem Monatsabschluss erneut ziehen.")
        befund["geladen"] = False
        return result, befund

    sparten = [c for c in result.columns if c != "Summe"]
    ohne_spalte = {}
    for zeile, je_objekt in werte.items():
        if zeile not in result.index:
            continue
        for objekt, betrag in je_objekt.items():
            if objekt in sparten:
                result.loc[zeile, objekt] = betrag
            elif _ist_sparte(objekt) and abs(betrag) > 0.005:
                ohne_spalte[objekt] = ohne_spalte.get(objekt, 0.0) + betrag
    if ohne_spalte:
        befund["hinweise"].append(
            "Deckungsdifferenzen auf Sparten, die im KPTM-Export gar nicht vorkommen und "
            "deshalb keine Spalte haben - diese Betraege fehlen in der Rechnung: "
            + ", ".join(f"{k} ({v:,.2f} EUR)" for k, v in sorted(ohne_spalte.items())))
    if befund["unbekannt"]:
        befund["hinweise"].append(
            "Umlage-Kostenarten ohne Zuordnung zu einer der drei Zeilen - diese Betraege "
            "fehlen in der Rechnung: "
            + ", ".join(f"{k} ({v:,.2f} EUR)" for k, v in sorted(befund["unbekannt"].items()))
            + ". Abhilfe: Kostenstellenbereich in 'deckungsdifferenz_umlage' "
              "(config/kostenart_mapping.json) ergaenzen.")
    befund["geladen"] = True
    return neu_berechnen(result), befund


# ---------------------------------------------------------------------------
# Abstimmung mit der Buchhaltung
# ---------------------------------------------------------------------------
# Eine Referenzrechnung gibt es nur fuer Q1 2026. Damit auch jede kuenftige Rechnung
# geprueft ist, stimmt sich das Tool selbst gegen zwei Quellen ab, die UNABHAENGIG vom
# KPTM-Export entstehen:
#   1. die Kostentraeger-Auswertung (Kostentraegerrechnung der Buchhaltung) - daraus
#      wird jede Zeile je Sparte ein zweites Mal gerechnet;
#   2. die Summen- und Saldenliste der FiBu - Erloese, Werkstattbestand, UFE.
# Geeicht an Q1 2026: Die Kostentraeger-Auswertung zeigt dort genau die Abweichungen,
# die auch die Referenz zeigt (Buchungsstand des KPTM) - sie taugt also als Referenz.
# Material und Fremdleistungen sind gegen die FiBu NICHT abstimmbar: die Kostenrechnung
# bewertet anders (strukturelle Differenz) - deshalb wird das gar nicht erst versucht.

def _de(betrag, vorzeichen=False):
    """Betrag im deutschen Format (1.234,56) - die Texte erscheinen so im Browser."""
    text = f"{float(betrag):{'+' if vorzeichen else ''},.2f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


ABSTIMMUNG_KT_ZEILEN = ["erloese", "bestand_fe", "material", "fremdleistungen", "sba",
                        "fek", "mgk", "vvgk"]
# Konten der unfertigen Erzeugnisse in der FiBu (Bilanz).
UFE_KONTEN = ["1040000", "1041000"]


def _pruefung(bereich, text, status, soll=None, ist=None, details=None):
    return {"bereich": bereich, "text": text, "status": status,
            "soll": None if soll is None else round(float(soll), 2),
            "ist": None if ist is None else round(float(ist), 2),
            "delta": None if soll is None or ist is None else round(float(ist) - float(soll), 2),
            "details": details or []}


def abstimmung_kostentraeger(kptm_df, kostentraeger_pfad, mapping, pg_config):
    """Zweite Rechnung aus der Kostentraeger-Auswertung, Zeile fuer Zeile je Sparte.

    Verglichen wird in der GEBUCHTEN Struktur: die Kostentraeger-Auswertung kennt die
    Umschluesselung nicht. Toleranz wie in der Validierung gegen die Referenz:
    unter 1 % bzw. 1 EUR gilt als Buchungsstand (die beiden Auswertungen werden nicht
    in derselben Sekunde gezogen), alles darueber als Abweichung.
    """
    df = kptm_df
    if "Produktgruppe (gebucht)" in df.columns:
        df = df.copy()
        df["Produktgruppe"] = df["Produktgruppe (gebucht)"]
    r_kptm, sparten, _, _, _ = build_spartenrechnung(df, mapping, pg_config)

    kt = lies_penta(kostentraeger_pfad)
    kt.columns = [str(c).strip() for c in kt.columns]
    kt = kt[~kt["Kostenart"].astype(str).str.strip().str.startswith("U")].copy()
    kt["Kostenart"] = kt["Kostenart"].astype(str).str.strip()
    kt["Produktgruppe"] = kt["Kostenobjekt"].map(_schluessel)
    if "Wertart" not in kt.columns:
        kt["Wertart"] = mapping.get("kostenstellen_wertart", "ISWF")
    r_kt, _, _, _, _ = build_spartenrechnung(kt, mapping, pg_config)

    spalten = sorted({p for p in list(r_kptm.columns) + list(r_kt.columns)
                      if p != "Summe" and _ist_sparte(p)})
    pruefungen = []
    for zeile in ABSTIMMUNG_KT_ZEILEN:
        exakt = nah = 0
        ab = []
        for p in spalten:
            x = float(r_kptm.loc[zeile, p]) if p in r_kptm.columns else 0.0
            y = float(r_kt.loc[zeile, p]) if p in r_kt.columns else 0.0
            if abs(x) < 0.005 and abs(y) < 0.005:
                continue
            d = x - y
            if abs(d) < 0.005:
                exakt += 1
            elif abs(d) < max(1.0, abs(y) * 0.01):
                nah += 1
            else:
                ab.append(f"{p}: Tool {_de(x)} / Kostenträger {_de(y)} (Δ {_de(d, True)})")
        summe_x = float(sum(r_kptm.loc[zeile, p] for p in spalten if p in r_kptm.columns))
        summe_y = float(sum(r_kt.loc[zeile, p] for p in spalten if p in r_kt.columns))
        text = (f"{mapping['zeilen_labels'].get(zeile, zeile)}: {exakt} Sparten centgenau"
                + (f", {nah} im Rundungsbereich (Buchungsstand)" if nah else ""))
        if ab:
            text += f", {len(ab)} weichen ab"
        pruefungen.append(_pruefung("Kostenträger-Auswertung", text,
                                    "abweichung" if ab else "ok", summe_y, summe_x, ab))
    return pruefungen


def lade_susa(pfad):
    """Summen- und Saldenliste der FiBu (CSV mit ';' und deutschen Zahlen, oder Excel).
    Liefert (DataFrame mit Index Kontonummer, Spalten anfang/saldo/beschreibung, bis-Datum)."""
    name = os.path.basename(str(pfad))
    try:
        if str(pfad).lower().endswith(".csv"):
            roh = pd.read_csv(pfad, sep=";", dtype=str, encoding="utf-8-sig")
        else:
            roh = pd.read_excel(pfad, dtype=str)
    except Exception as e:
        raise ValueError(f"'{name}' ist keine lesbare Summen- und Saldenliste ({e}).")
    roh.columns = [str(c).strip() for c in roh.columns]
    noetig = ["Code", "Anf.-Saldo", "Saldo"]
    if not all(sp in roh.columns for sp in noetig):
        raise ValueError(
            f"'{name}' sieht nicht wie eine Summen- und Saldenliste (SuSa) aus - es fehlen die "
            f"Spalten {', '.join(sp for sp in noetig if sp not in roh.columns)}. Erwartet wird "
            f"der SuSa-Export der FiBu (z.B. 'SuSa 01-09.2026.csv').")

    def zahl(s):
        return pd.to_numeric(s.astype(str).str.strip().str.replace(".", "", regex=False)
                             .str.replace(",", ".", regex=False), errors="coerce").fillna(0.0)

    df = pd.DataFrame({
        "konto": roh["Code"].map(_schluessel),
        "beschreibung": roh.get("Beschreibung", pd.Series([""] * len(roh))).fillna(""),
        "anfang": zahl(roh["Anf.-Saldo"]),
        "saldo": zahl(roh["Saldo"]),
        "typ": roh.get("Typ", pd.Series([""] * len(roh))).fillna("").astype(str).str.strip(),
    }).dropna(subset=["konto"]).set_index("konto")
    bis = None
    if "Jahr bis" in roh.columns:
        bis = pd.to_datetime(roh["Jahr bis"].dropna().iloc[0], dayfirst=True,
                             errors="coerce") if roh["Jahr bis"].notna().any() else None
    return df, bis


def abstimmung_susa(susa_pfad, kptm_df, mapping, result, ufe_statistik, ufe_berechnet):
    """FiBu-Abstimmung: Erlöse, Werkstattbestand zum Stichtag, Bestandsveränderung UFE."""
    import datetime
    susa, bis = lade_susa(susa_pfad)
    name = os.path.basename(str(susa_pfad))

    # Gleicher Zeitraum wie der KPTM-Export? Sonst ist jede Abstimmung wertlos.
    spanne = _perioden_aus_daten(kptm_df)
    if spanne and bis is not None and not pd.isna(bis):
        _, b, _, j = spanne
        ende = (datetime.date(j + 1, 1, 1) if b == 12 else datetime.date(j, b + 1, 1)) \
            - datetime.timedelta(days=1)
        if bis.date() != ende:
            raise ValueError(
                f"Die Summen- und Saldenliste '{name}' reicht bis {bis:%d.%m.%Y}, der "
                f"KPTM-Export bis {ende:%d.%m.%Y}. Beide muessen denselben Zeitraum abdecken.")

    pruefungen = []
    # 1. Erloese: alle Konten, die das Mapping als Erloes kennt, PLUS jedes weitere
    #    44er-Konto der FiBu. So faellt ein neues Erloeskonto auf, das das Mapping
    #    noch nicht kennt (z.B. 4480000 Erloese Schneidrollen, erstmals 09/2026).
    kz = mapping["kostenart_zu_zeile"]
    erloes_konten = {k for k, v in kz.items() if v["zeile"] == "erloese"}
    erloes_konten |= {k for k in susa.index if k.startswith("44")}
    fibu = -float(susa.loc[susa.index.isin(erloes_konten), "saldo"].sum())
    tool = float(result.loc["erloese", "Summe"])
    fehlend = sorted(k for k in erloes_konten if k in susa.index and k not in kz
                     and abs(susa.loc[k, "saldo"]) > 0.005)
    details = [f"{k} {susa.loc[k, 'beschreibung']}: {_de(-susa.loc[k, 'saldo'])} - Konto fehlt "
               f"im Mapping" for k in fehlend]
    pruefungen.append(_pruefung(
        "FiBu (SuSa)", "Erlöse gesamt gegen Erlöskonten der FiBu",
        "ok" if abs(tool - fibu) < 1.0 else "abweichung", fibu, tool, details))

    # 2. Werkstattbestand zum Stichtag = Bilanzkonten der unfertigen Erzeugnisse.
    #    Faengt einen Werkstattbestand mit falschem Stichtag ebenso ab wie einen, der
    #    nachtraeglich gezogen wurde (dessen Werte stimmen nicht mehr mit der FiBu).
    st = ufe_statistik or {}
    if st:
        fibu_ende = float(susa.loc[susa.index.isin(UFE_KONTEN), "saldo"].sum())
        fibu_anfang = float(susa.loc[susa.index.isin(UFE_KONTEN), "anfang"].sum())
        pwbs = float(st.get("zugeordnet", 0.0)) + float(st.get("offen", 0.0))
        pruefungen.append(_pruefung(
            "FiBu (SuSa)",
            "Werkstattbestand zum Stichtag gegen FiBu-Konten " + " + ".join(UFE_KONTEN),
            "ok" if abs(pwbs - fibu_ende) < 1.0 else "abweichung", fibu_ende, pwbs,
            [] if abs(pwbs - fibu_ende) < 1.0 else [
                "Häufigste Ursachen: Werkstattbestand nicht am Stichtag gezogen (später "
                "gezogene Exporte zeigen den heutigen Stand), oder Stichtag passt nicht "
                "zum Zeitraum."]))

        # 3. Der Anfangsbestand ist der einzige Wert der UFE-Zeile, der nicht aus einer
        #    Datei des Zeitraums kommt, sondern hinterlegt ist. Er wird gegen die FiBu
        #    zum 01.01. gehalten. Ein Unterschied ist kein Fehler an sich (der hinterlegte
        #    Wert kann Bestand ohne Sparte bewusst auslassen), aber er muss bekannt und
        #    ueber das Jahr konstant sein - und er erklaert, um wie viel die UFE-Summe der
        #    Sparten von der Buchung der FiBu abweicht.
        if ufe_berechnet:
            sparten = [c for c in result.columns if c != "Summe"]
            ufe_tool = float(result.loc["bestand_ufe", sparten].sum())
            fibu_bewegung = fibu_ende - fibu_anfang
            offen = float(st.get("offen", 0.0))
            ab_tool = float(st.get("zugeordnet", 0.0)) - ufe_tool
            ab_delta = fibu_anfang - ab_tool
            details = [
                f"Bestandsveränderung UFE: Summe der Sparten {_de(ufe_tool)}, Buchung der "
                f"FiBu {_de(fibu_bewegung)} (Δ {_de(ufe_tool - fibu_bewegung, True)}).",
                f"Der Unterschied setzt sich zusammen aus dem Anfangsbestand-Unterschied "
                f"({_de(ab_delta, True)}) und dem nicht zuordenbaren Werkstattbestand ({_de(offen)}).",
            ]
            if abs(ab_delta) >= 1.0:
                details.append(
                    "Der hinterlegte Anfangsbestand weicht von der FiBu zum 01.01. ab. Ist der "
                    "Unterschied jeden Monat derselbe, ist das bekannter Bestand ohne Sparte; "
                    "ändert er sich, wurden die Betriebsparameter verändert.")
            pruefungen.append(_pruefung(
                "FiBu (SuSa)", "Hinterlegter Anfangsbestand gegen FiBu zum 01.01.",
                "ok" if abs(ab_delta) < 1.0 else "hinweis", fibu_anfang, ab_tool, details))
    return pruefungen


# ---------------------------------------------------------------------------
# Abgleich mit der GuV
# ---------------------------------------------------------------------------
# Bruecke vom DB III der Sparten zum Ergebnis der GuV, gerechnet aus der SuSa. Die BWA
# der Buchhaltung ist aus derselben SuSa abgeleitet - geprueft fuer Jan-Sep 2026:
# 'Vorlaeufiges Ergebnis' und die Erloese je Bereich stimmen centgenau. Deshalb wird
# die BWA-Datei selbst nicht gebraucht.
#
# Jedes GuV-Konto der SuSa gehoert genau einem Posten an. Ein Rest am Ende ist deshalb
# ein echter Befund (z.B. ein Konto, das keinem Posten zugeordnet werden konnte), keine
# Rechenluecke.

# Kontenbereiche wie in der BWA der Buchhaltung (dort 'Umsatzerloese').
GUV_ERLOES_PRAEFIXE = ("42", "44", "45", "47", "483")
GUV_FE_PRAEFIXE = ("4825",)          # Bestandsveraenderung fertige Erzeugnisse
GUV_UFE_PRAEFIXE = ("4814",)         # Bestandsveraenderung unfertige Erzeugnisse
GUV_KOSTEN_PRAEFIXE = ("5", "6", "7")  # Material, Personal, sbA, Abschreibungen, Steuern

# Erloese je Bereich der BWA -> Spartenfamilie. Frachterloese (452/453...) tragen den
# Bereich in den letzten drei Stellen.
GUV_BEREICHE = [
    ("Bohrstangen", "DP", ("4440",), "400"),
    ("Horizontalbohrtechnik", "PB", ("4460",), "710"),
    ("Injektionstechnik", "IT", ("4470",), "060"),
    ("Schneidrollen", "SR", ("4480",), None),
]


def _bereich_des_kontos(konto):
    for name, familie, praefixe, fracht in GUV_BEREICHE:
        if konto.startswith(praefixe):
            return familie
        if fracht and konto[:3] in ("452", "453") and konto.endswith(fracht):
            return familie
    return None


def guv_abgleich(susa_pfad, kptm_df, mapping, result):
    """Bruecke DB III (Sparten) -> Ergebnis lt. GuV, dazu Erloese je Bereich und Konten.

    Liefert ein dict fuer Oberflaeche und Excel-Blatt:
      bruecke:  [(text, betrag, art)]  art = 'summe' | 'posten' | 'rest'
      bereiche: [(name, familie, guv, sparten, delta)]
      abweichende_buchungen: [dict]  Erloesbuchungen, deren Konto zu einem anderen
                                     Bereich gehoert als ihre Produktgruppe
      konten:   [dict]  je GuV-Konto: Posten, GuV, direkt auf Sparten gebucht, Rest
    """
    susa, _ = lade_susa(susa_pfad)
    if (susa["typ"] != "").any():
        guv = susa[susa["typ"].isin(["Ertrag", "Aufwand"])].copy()
    else:   # Export ohne Spalte 'Typ': Kontenklassen 4-8
        guv = susa[susa.index.str[:1].isin(list("45678"))].copy()
    # Ertrag positiv, Aufwand negativ - wie in der BWA
    guv["betrag"] = -guv["saldo"]

    cfg = mapping.get("guv_abgleich", {})
    erloes_praefixe = tuple(cfg.get("erloese_praefixe", GUV_ERLOES_PRAEFIXE))
    sbe = set(cfg.get("sonstige_ertraege", []))
    neutral = set(cfg.get("neutral", []))
    zins = tuple(cfg.get("zins_praefixe", ["73", "87"]))
    steuern = set(cfg.get("steuern", []))

    def posten(konto):
        if konto.startswith(GUV_FE_PRAEFIXE):
            return "fe"
        if konto.startswith(GUV_UFE_PRAEFIXE):
            return "ufe"
        if konto in neutral:
            return "neutral"
        if konto in sbe:
            return "sbe"
        if konto in steuern:
            return "steuern"
        if konto.startswith(zins):
            return "zins"
        if konto.startswith(erloes_praefixe):
            return "erloese"
        # Alles Uebrige ist Betriebsaufwand bzw. -ertrag: Material, Personal, sbA,
        # Abschreibungen, Erstattungen, Konzernumlagen. Eine eigene Restgruppe gibt es
        # bewusst nicht - so geht jedes Konto in genau einen Posten.
        return "kosten"
    guv["posten"] = [posten(k) for k in guv.index]
    summe = guv.groupby("posten")["betrag"].sum()
    g = lambda k: float(summe.get(k, 0.0))

    sparten = [c for c in result.columns if c != "Summe"]
    S = lambda z: float(result.loc[z, sparten].sum()) if z in result.index else 0.0
    db3 = S("db3")
    # Was die Kostenrechnung den Sparten an Kosten belastet - alles zwischen
    # Betriebsleistung und DB III (Material, Fremdleistungen, sbA, FEK/MGK/VVGK nach
    # Deckungsdifferenz, Sonderposten).
    kosten_kore = S("betriebsleistung") - db3
    ergebnis_guv = float(guv["betrag"].sum())

    betrieb = [
        ("Erlöse ohne Sparte (Abfallverwertung, Sonstiges, Skonti)", g("erloese") - S("erloese"), "erloese"),
        ("Bestandsveränderung FE: FiBu gegen Kostenrechnung", g("fe") - S("bestand_fe"), "fe"),
        ("Bestandsveränderung UFE: FiBu gegen Sparten", g("ufe") - S("bestand_ufe"), "ufe"),
        ("Sonstige betriebliche Erträge (nicht auf Sparten)", g("sbe"), "sbe"),
        ("Kosten (Material, Personal, sbA, Abschreibungen): FiBu gegen Kostenrechnung",
         g("kosten") + kosten_kore, "kosten"),
    ]
    betriebsergebnis = db3 + sum(b for _, b, _ in betrieb)
    bruecke = [("DB III (Summe der Sparten)", db3, "summe", "db3")]
    bruecke += [(t, b, "posten", k) for t, b, k in betrieb]
    bruecke.append(("Betriebsergebnis lt. GuV (BWA)", betriebsergebnis, "summe", "betriebsergebnis"))
    bruecke += [
        ("Neutrales Ergebnis (Anlagenverkauf, Kursdifferenzen, periodenfremd)", g("neutral"),
         "posten", "neutral"),
        ("Zinsergebnis", g("zins"), "posten", "zins"),
        ("Steuern", g("steuern"), "posten", "steuern"),
    ]
    rest = ergebnis_guv - betriebsergebnis - g("neutral") - g("zins") - g("steuern")
    # Bliebe hier etwas, waere ein Konto in keinem Posten - z.B. aktivierte
    # Eigenleistung, die in der Betriebsleistung steckt, aber kein GuV-Konto ist. Die
    # Zeile steht immer da (im Excel-Blatt als Formel, die 0,00 zeigen muss); die
    # Oberflaeche blendet sie aus, solange sie null ist.
    bruecke.append(("Nicht zugeordneter Rest (muss 0,00 sein)", rest, "rest", "rest"))
    bruecke.append(("Ergebnis lt. GuV (entspricht „Vorläufiges Ergebnis“ der BWA)",
                    ergebnis_guv, "summe", "ergebnis"))

    # Erloese je Bereich: GuV nach Konto, Sparten nach Produktgruppe
    guv["bereich"] = [_bereich_des_kontos(k) for k in guv.index]
    bereiche = []
    for name, familie, _, _ in GUV_BEREICHE:
        guv_wert = float(guv.loc[(guv["posten"] == "erloese") & (guv["bereich"] == familie), "betrag"].sum())
        sp_wert = float(sum(result.loc["erloese", p] for p in sparten if p.startswith(familie)))
        if abs(guv_wert) > 0.005 or abs(sp_wert) > 0.005:
            bereiche.append({"name": name, "familie": familie, "guv": round(guv_wert, 2),
                             "sparten": round(sp_wert, 2), "delta": round(guv_wert - sp_wert, 2)})
    sonst = float(guv.loc[(guv["posten"] == "erloese") & guv["bereich"].isna(), "betrag"].sum())
    bereiche.append({"name": "Sonstiges und Skonti", "familie": "", "guv": round(sonst, 2),
                     "sparten": 0.0, "delta": round(sonst, 2)})

    # Erloesbuchungen, deren Konto zu einem anderen Bereich gehoert als ihre Produktgruppe
    kz = mapping["kostenart_zu_zeile"]
    e = kptm_df[kptm_df["Kostenart"].map(lambda k: kz.get(k, {}).get("zeile") == "erloese")].copy()
    e["kb"] = e["Kostenart"].map(_bereich_des_kontos)
    e["fam"] = e["Produktgruppe"].fillna("").str[:2]
    e = e[e["kb"].notna() & (e["kb"] != e["fam"])]
    abweichend = []
    for (konto, pg), wert in e.groupby(["Kostenart", "Produktgruppe"])["Wert/Menge"].sum().items():
        abweichend.append({"konto": konto, "bezeichnung": str(susa["beschreibung"].get(konto, "")),
                           "produktgruppe": pg, "betrag": round(-float(wert), 2)})
    abweichend.sort(key=lambda x: -abs(x["betrag"]))

    # Konto fuer Konto: was davon direkt auf Sparten gebucht ist (KPTM)
    direkt = kptm_df[kptm_df["Produktgruppe"].notna()].groupby("Kostenart")["Wert/Menge"].sum()
    namen = {"erloese": "Erlöse", "fe": "Bestandsveränderung FE", "ufe": "Bestandsveränderung UFE",
             "sbe": "Sonstige betriebliche Erträge", "kosten": "Kosten", "neutral": "Neutral",
             "zins": "Zinsen", "steuern": "Steuern"}
    konten = []
    for konto, z in guv.sort_values("betrag", key=abs, ascending=False).iterrows():
        if abs(z["betrag"]) < 0.005:
            continue
        d = -float(direkt.get(konto, 0.0))
        konten.append({"konto": konto, "bezeichnung": str(z["beschreibung"]),
                       "posten": namen[z["posten"]], "guv": round(float(z["betrag"]), 2),
                       "direkt": round(d, 2), "rest": round(float(z["betrag"]) - d, 2)})

    return {"bruecke": [{"text": t, "betrag": round(b, 2), "art": a, "key": k}
                        for t, b, a, k in bruecke],
            "susa_pfad": susa_pfad,
            "ergebnis_guv": round(ergebnis_guv, 2), "db3": round(db3, 2),
            "betriebsergebnis": round(betriebsergebnis, 2),
            "bereiche": bereiche, "abweichende_buchungen": abweichend[:40],
            "konten": konten}


def _guv_blatt(wb, guv, zeitraum):
    """Blatt 'Abgleich GuV': Bruecke, Erloese je Bereich, Konten im Einzelnen."""
    ws = wb.create_sheet("Abgleich GuV", 2 if "Abstimmung" in wb.sheetnames else 1)
    ws.cell(row=1, column=1, value=f"Abgleich mit der GuV – {zeitraum}").font = Font(bold=True, size=12)
    ws.cell(row=2, column=1, value=(
        "Vom DB III der Sparten zum Ergebnis der GuV. Grundlage ist die Summen- und "
        "Saldenliste; das Ergebnis entspricht dem 'Vorläufigen Ergebnis' der BWA. Jedes "
        "GuV-Konto gehört genau einem Posten an.")).font = Font(italic=True, size=9)
    pos = {"bruecke": {}, "bereiche": [], "abweichend": [], "konten": []}
    r = 4
    for b in guv["bruecke"]:
        pos["bruecke"][b["key"]] = r
        fett = b["art"] == "summe"
        z1 = ws.cell(row=r, column=1, value=("" if fett else "  ") + b["text"])
        z2 = ws.cell(row=r, column=2, value=b["betrag"])
        z2.number_format = "#,##0.00"
        for z in (z1, z2):
            z.border = BORDER
            if fett:
                z.font = FONT_BOLD
                z.fill = FILL_FORMULA
            if b["art"] == "rest" and abs(b["betrag"]) >= 0.01:
                z.fill = FILL_MANUAL
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="Erlöse je Bereich – GuV nach Konto, Sparten nach Produktgruppe").font = FONT_BOLD
    r += 1
    for j, t in enumerate(("Bereich", "GuV", "Sparten", "Δ"), start=1):
        z = ws.cell(row=r, column=j, value=t)
        z.fill, z.font = FILL_HEADER, FONT_HEADER
    r += 1
    for b in guv["bereiche"]:
        pos["bereiche"].append((r, b["familie"]))
        for j, w in enumerate((b["name"] + (f" ({b['familie']}*)" if b["familie"] else ""),
                               b["guv"], b["sparten"], b["delta"]), start=1):
            z = ws.cell(row=r, column=j, value=w)
            z.border = BORDER
            if j > 1:
                z.number_format = "#,##0.00"
        r += 1
    if guv["abweichende_buchungen"]:
        r += 1
        ws.cell(row=r, column=1, value=(
            "Ursache der Verschiebungen zwischen den Bereichen: Erlösbuchungen, deren Konto zu "
            "einem anderen Bereich gehört als ihre Produktgruppe")).font = Font(italic=True, size=9)
        r += 1
        for j, t in enumerate(("Konto", "Bezeichnung", "Produktgruppe", "Erlös"), start=1):
            z = ws.cell(row=r, column=j, value=t)
            z.fill, z.font = FILL_HEADER, FONT_HEADER
        r += 1
        for a in guv["abweichende_buchungen"]:
            pos["abweichend"].append((r, a["konto"], a["produktgruppe"]))
            for j, w in enumerate((a["konto"], a["bezeichnung"], a["produktgruppe"], a["betrag"]), start=1):
                z = ws.cell(row=r, column=j, value=w)
                z.border = BORDER
                if j == 4:
                    z.number_format = "#,##0.00"
            r += 1

    r += 2
    ws.cell(row=r, column=1, value=(
        "Konten im Einzelnen – 'direkt auf Sparten' ist, was im KPTM-Export auf einer "
        "Produktgruppe gebucht ist. Der Rest läuft über die Kostenstellen (Stundensatz, "
        "Zuschläge, Deckungsdifferenzen) oder gehört keiner Sparte.")).font = Font(italic=True, size=9)
    r += 1
    for j, t in enumerate(("Konto", "Bezeichnung", "Posten", "GuV", "direkt auf Sparten", "Rest"), start=1):
        z = ws.cell(row=r, column=j, value=t)
        z.fill, z.font = FILL_HEADER, FONT_HEADER
    r += 1
    for k in guv["konten"]:
        pos["konten"].append((r, k["konto"]))
        for j, w in enumerate((k["konto"], k["bezeichnung"], k["posten"], k["guv"], k["direkt"], k["rest"]), start=1):
            z = ws.cell(row=r, column=j, value=w)
            z.border = BORDER
            if j >= 4:
                z.number_format = "#,##0.00"
        r += 1
    for sp, b in (("A", 70), ("B", 34), ("C", 22), ("D", 16), ("E", 18), ("F", 16)):
        ws.column_dimensions[sp].width = b
    return pos


def _abstimmung_blatt(wb, pruefungen, zeitraum):
    """Blatt 'Abstimmung': die Prüfungen gegen die Buchhaltung, fuer jeden Leser der Datei."""
    ws = wb.create_sheet("Abstimmung", 1)
    ws.cell(row=1, column=1, value=f"Abstimmung mit der Buchhaltung – {zeitraum}").font = Font(
        bold=True, size=12)
    ws.cell(row=2, column=1, value=(
        "Das Tool rechnet aus dem KPTM-Export. Hier wird das Ergebnis gegen zwei Quellen "
        "gehalten, die unabhaengig davon entstehen: die Kostenträger-Auswertung (zweite "
        "Rechnung je Sparte) und die Summen- und Saldenliste der FiBu.")).font = Font(
        italic=True, size=9)
    kopf = ["Quelle", "Prüfung", "Status", "Buchhaltung", "Tool", "Δ"]
    for j, t in enumerate(kopf, start=1):
        z = ws.cell(row=4, column=j, value=t)
        z.fill, z.font = FILL_HEADER, FONT_HEADER
    r = 5
    status_text = {"ok": "✓ stimmt", "hinweis": "ℹ Hinweis", "abweichung": "⚠ prüfen"}
    status_fill = {"ok": FILL_AUTO, "hinweis": FILL_AUTO_PRUEF, "abweichung": FILL_MANUAL}
    for p in pruefungen:
        werte = [p["bereich"], p["text"], status_text[p["status"]],
                 p["soll"], p["ist"], p["delta"]]
        for j, w in enumerate(werte, start=1):
            z = ws.cell(row=r, column=j, value=w)
            z.border = BORDER
            if j >= 4:
                z.number_format = "#,##0.00"
            if j == 3:
                z.fill = status_fill[p["status"]]
        r += 1
        for d in p["details"][:12]:
            ws.cell(row=r, column=2, value="   " + d).font = Font(size=9, color="555555")
            r += 1
    for sp, b in (("A", 24), ("B", 90), ("C", 12), ("D", 16), ("E", 16), ("F", 12)):
        ws.column_dimensions[sp].width = b


def pruefe_werkstattbestand_buchung(kostentraeger_pfad, ufe, ufe_statistik, schwelle=0.02):
    """Passt der Werkstattbestand zur gebuchten Bestandsveraenderung?

    Die Kostentraeger-Auswertung enthaelt die Buchung der Bestandsveraenderung auf den
    Bilanzkonten der unfertigen Erzeugnisse (1041000, auf Kostenobjekt 910000). Daraus
    folgt der zu erwartende Endbestand: Anfangsbestand + Buchung. Ein Werkstattbestand,
    der nachtraeglich gezogen wurde, zeigt die Auftraege mit ihrem HEUTIGEN Stand und
    liegt deutlich daneben - gemessen an Q1 2026: am Stichtag gezogen 0,41 %, im Oktober
    rueckwirkend gezogen 11,5 %. Die verbleibende kleine Differenz ist der konstante
    Unterschied zwischen hinterlegtem Anfangsbestand und FiBu (Bestand ohne Sparte);
    sie war im Q1 und im September centgenau derselbe Betrag.

    Liefert eine Pruefung (siehe _pruefung) oder None, wenn nichts zu pruefen ist.
    """
    if not ufe or not ufe_statistik:
        return None
    kt = lies_penta(kostentraeger_pfad)
    kt.columns = [str(c).strip() for c in kt.columns]
    gebucht = kt["Kostenart"].astype(str).str.strip().isin(UFE_KONTEN)
    if not gebucht.any():
        return None
    bewegung = float(kt.loc[gebucht, "Wert/Menge"].sum())
    zugeordnet = float(ufe_statistik.get("zugeordnet", 0.0))
    ende = zugeordnet + float(ufe_statistik.get("offen", 0.0))
    anfang = zugeordnet - float(sum(ufe.values()))
    erwartet = anfang + bewegung
    delta = ende - erwartet
    quote = abs(delta) / abs(erwartet) if erwartet else 0.0
    details = [
        f"Werkstattbestand laut Datei (ohne Serviceaufträge): {_de(ende)}",
        f"erwartet: Anfangsbestand {_de(anfang)} + gebuchte Bestandsveränderung "
        f"{_de(bewegung, True)} = {_de(erwartet)}",
        f"Unterschied {_de(delta, True)} ({f'{quote * 100:.1f}'.replace('.', ',')} %)",
    ]
    if quote > schwelle:
        details.append(
            "Der Werkstattbestand passt nicht zur Buchung. Fast immer wurde er nicht am "
            "Stichtag, sondern später gezogen - dann zeigt er die Aufträge mit ihrem "
            "heutigen Stand. Bitte den am Stichtag gezogenen Werkstattbestand verwenden.")
    return _pruefung(
        "Kostenträger-Auswertung",
        "Werkstattbestand gegen gebuchte Bestandsveränderung (Konto 1041000)",
        "abweichung" if quote > schwelle else "ok", erwartet, ende, details)


def generate(kptm_path, config_dir, out_path, zeitraum=None, bwa_path=None, bwa_sheet=None,
             pwbs_anfang=None, pwbs_ende=None, pfak_pfade=None, manuelle_zuordnung=None,
             kostentraeger_pfad=None, susa_pfad=None):
    """Ein-Funktions-Einstieg fuer die Weboberflaeche: liest KPTM (+optional BWA,
    +optional Werkstattbestand fuer die Bestandsveraenderung UFE), berechnet die
    Spartenrechnung und schreibt die Ausgabedatei. Gibt eine kurze Zusammenfassung
    (dict) zurueck, u.a. fuer Warnungen zu unbekannten Kostenarten/Produktgruppen.

    'zeitraum' ist nur noch ein Ueberschreiben von Hand; ohne Angabe wird er aus den
    Periodenspalten des Exports gelesen (siehe zeitraum_aus_daten)."""
    mapping, pg_config = load_config(config_dir)
    df = load_kptm(kptm_path)
    abgeleitet, dateiname = zeitraum_aus_daten(df, kptm_path)
    zeitraum = zeitraum or abgeleitet
    jahr = geschaeftsjahr(df, zeitraum)

    # Neue Produktgruppen-Struktur (Reorganisation 2026): Buchungen unter alten Gruppen
    # werden artikelgenau umgeschluesselt, damit das ganze Jahr in EINER Struktur steht.
    # Nur ab dem Umstellungsjahr - aeltere Zeitraeume bleiben, wie sie gebucht wurden.
    params = lade_betriebsparameter(config_dir)
    aufgeloest = pg_config.get("aufgeloest", {})
    ab_jahr = pg_config.get("umstellung_ab_jahr")
    karte = (params.get("umschluesselung") or {}) if (
        jahr and ab_jahr and int(jahr) >= int(ab_jahr)) else {}
    struktur = None
    if karte:
        df, struktur = umschluesseln(df, karte, aufgeloest)
        struktur["naeherung_ab"] = sorted(
            (params.get("anfangsbestand_naeherung") or {}).get(str(jahr), []))
    else:
        aufgeloest = {}
        # Ohne Umschluesselung stuenden alte und neue Gruppen nebeneinander - dieselbe
        # Produktlinie in zwei Spalten. Das faellt in der Summe nicht auf, deshalb melden.
        alte = sorted(set(df["Produktgruppe"].dropna()) & set(pg_config.get("aufgeloest", {})))
        if jahr and ab_jahr and int(jahr) >= int(ab_jahr) and alte:
            struktur = {"fehlt": True, "alte": alte}
    # Der Stundensatz wird aus den Rohdaten errechnet, nicht gepflegt (siehe
    # ermittle_stundensatz). Ein hinterlegter Wert dient nur noch als Rueckfall.
    satz = ermittle_stundensatz(df, mapping)
    if satz:
        mapping["standard_stundensatz"] = satz
    result, produktgruppen, unmapped, per_regel, befunde = build_spartenrechnung(
        df, mapping, pg_config)

    dd_befund = None
    if kostentraeger_pfad:
        result, dd_befund = deckungsdifferenzen_einsetzen(
            result, kostentraeger_pfad, mapping, df,
            anteile=aufteilungsanteile(df, mapping) if karte else None)
    dd_geladen = bool(dd_befund and dd_befund.get("geladen"))
    # Geladene Deckungsdifferenzen sind gruen, nicht mehr rosa. Die Sonderposten-Zeile
    # nur dann, wenn tatsaechlich etwas auf einer Sparte landet - sonst bleibt sie grau.
    zeilen_klasse = {}
    if dd_geladen:
        zeilen_klasse.update({z: "auto" for z in DD_ZEILEN})
        sparten = [c for c in result.columns if c != "Summe"]
        if (result.loc["sonderposten", sparten].abs() > 0.005).any():
            zeilen_klasse["sonderposten"] = "auto"

    ufe_warnungen = []
    ufe_statistik = None
    werkstatt_pruefung = None
    gespeicherter_ab = lade_anfangsbestand(config_dir, jahr)
    # Vor allem anderen: passen die Stichtage zum Zeitraum? Sonst ist jede weitere
    # Pruefung der UFE-Zeile wertlos.
    stichtag_hinweise = [h for h in (
        pruefe_stichtag(pwbs_ende, df, "ende") if pwbs_ende else None,
        pruefe_stichtag(pwbs_anfang, df, "anfang") if pwbs_anfang else None) if h]
    nach_auftrag = nach_artikel = None
    if pwbs_ende:
        # PFAK ist optional: die Zuordnung Auftrag -> Produktgruppe steckt bereits in KPTM.
        nach_auftrag, nach_artikel, _, _ = lade_produktgruppen_map(pfak_pfade, kptm_df=df)
        ufe, ufe_warnungen, ufe_statistik = bestandsveraenderung_ufe(
            pwbs_ende, pfak_pfade, anfangsbestand=gespeicherter_ab, pwbs_anfang=pwbs_anfang,
            jahr=jahr, kptm_df=df, manuelle_zuordnung=manuelle_zuordnung,
            karte=karte, aufgeloest=aufgeloest)
        if ufe_statistik is not None:
            ufe_statistik["hinweise"] = stichtag_hinweise + ufe_statistik.get("hinweise", [])
        if ufe and not karte and (params.get("anfangsbestand_naeherung") or {}).get(str(jahr)):
            # Anfangsbestand in der neuen Produktgruppen-Struktur (bei der Umstellung
            # aufgeteilt), gerechnet wird aber wie gebucht: PB-Sparten waeren je Sparte
            # falsch, nur ihre Summe stimmte. Deshalb keine Zahl.
            ufe_warnungen.append(
                "Die hinterlegten Betriebsparameter enthalten den Anfangsbestand in der neuen "
                "Produktgruppen-Struktur, die Spartenrechnung folgt aber der gebuchten "
                "Produktgruppe. Bitte die aktuelle betriebsparameter.xlsx (gebuchte Struktur, "
                "ohne Spalte 'Herkunft') laden.")
            ufe = None
        if ufe and kostentraeger_pfad:
            werkstatt_pruefung = pruefe_werkstattbestand_buchung(
                kostentraeger_pfad, ufe, ufe_statistik)
            if werkstatt_pruefung and werkstatt_pruefung["status"] == "abweichung":
                # Wie bei zu grosser Zuordnungsluecke: lieber keine Zahl als eine falsche.
                ufe_warnungen.append(
                    "Der Werkstattbestand passt nicht zur gebuchten Bestandsveränderung "
                    "(Kostenträger-Auswertung, Konto 1041000). Fast immer wurde er nicht am "
                    "Stichtag, sondern später gezogen - dann zeigt er die Aufträge mit ihrem "
                    "heutigen Stand. Bitte den am Stichtag gezogenen Werkstattbestand verwenden.")
                ufe_warnungen.extend(werkstatt_pruefung["details"][:3])
                ufe = None
        if ufe:
            for pg, wert in ufe.items():
                if pg not in result.columns and abs(wert) > 0.005:
                    # Sparte mit Werkstattbestand, aber ohne Buchung im Zeitraum -
                    # eigene Spalte statt den Betrag still fallen zu lassen.
                    result[pg] = 0.0
                if pg in result.columns:
                    result.loc["bestand_ufe", pg] = wert
            result = neu_berechnen(result)

    # Aufgeloeste Gruppen (z.B. PB0003) sind nach der Umschluesselung leer - die
    # Spalte bleibt dann weg. Traegt sie doch noch etwas, bleibt sie sichtbar.
    for weg in aufgeloest:
        if weg in result.columns and not (result[weg].abs() > 0.005).any():
            result = result.drop(columns=weg)
            produktgruppen = [p for p in produktgruppen if p != weg]

    # Gelb statt gruen: Zellen, die aus einer anteiligen Aufteilung stammen.
    naeherung_zellen = set((dd_befund or {}).get("naeherung", set())) if dd_geladen else set()
    if struktur and not struktur.get("fehlt") and pwbs_ende:
        naeherung_zellen |= {("bestand_ufe", pg) for pg in struktur["naeherung_ab"]
                             if pg in result.columns}

    bwa_ergebnis = None
    if bwa_path and bwa_sheet:
        bwa_ergebnis = read_bwa_ergebnis(bwa_path, bwa_sheet)

    # Abstimmung mit der Buchhaltung - damit auch ohne Referenzrechnung jede Rechnung
    # gegen unabhaengige Quellen geprueft ist (siehe abstimmung_kostentraeger/_susa).
    abstimmung = []
    if kostentraeger_pfad:
        abstimmung += abstimmung_kostentraeger(df, kostentraeger_pfad, mapping, pg_config)
    if werkstatt_pruefung:
        abstimmung.append(werkstatt_pruefung)
    guv = None
    if susa_pfad:
        abstimmung += abstimmung_susa(susa_pfad, df, mapping, result, ufe_statistik,
                                      bool(pwbs_ende and not ufe_warnungen))
        guv = guv_abgleich(susa_pfad, df, mapping, result)

    saetze, satz_hinweise = zuschlagssaetze(result, mapping)
    unprod_betrag, unprod_stunden = unproduktive_gemeinkosten(df, mapping)
    stunden_df = fertigungsstunden(df, mapping)
    write_output(result, mapping, zeitraum, out_path, bwa_ergebnis=bwa_ergebnis, bwa_sheet=bwa_sheet,
                 stunden_df=stunden_df, unproduktiv=(unprod_betrag, unprod_stunden),
                 pg_bezeichnungen=pg_config.get("bezeichnung", {}),
                 ufe_warnungen=ufe_warnungen if pwbs_ende else None,
                 zeilen_klasse=zeilen_klasse,
                 deckungsdifferenzen=dd_befund,
                 naeherung_zellen=naeherung_zellen,
                 struktur=struktur,
                 abstimmung=abstimmung,
                 guv=guv,
                 herleitung={
                     "kostentraeger": ({"umlage": dd_befund["umlage_df"],
                                        "zeilen_gesamt": dd_befund["zeilen_gesamt"]}
                                       if dd_geladen else None),
                     "produktgruppen": produktgruppen,
                     "kptm_df": df,
                     "umschluesselung": karte or None,
                     "aufgeloest": aufgeloest,
                     "anfangsbestand_naeherung": (struktur or {}).get("naeherung_ab"),
                     "pwbs_ende": pwbs_ende,
                     "pwbs_anfang": pwbs_anfang,
                     "nach_auftrag": nach_auftrag,
                     "nach_artikel": nach_artikel,
                     "anfangsbestand": gespeicherter_ab,
                     "manuelle_zuordnung": manuelle_zuordnung,
                     # Nur wenn die Zeile tatsaechlich gerechnet wurde, darf in der
                     # Mappe eine Formel stehen. Sonst zeigte sie eine Zahl, die das
                     # Tool bewusst zurueckhaelt, weil die Zuordnungsluecke zu gross ist.
                     "ufe_berechnet": bool(pwbs_ende and not ufe_warnungen),
                 })

    known_pg = set(pg_config["reihenfolge"])
    unbekannte_pg = [p for p in produktgruppen if p not in known_pg]

    return {
        "zeilen": len(df),
        "zeitraum": zeitraum,
        "dateiname": dateiname,
        # Aus den Rohdaten, nicht aus dem Zeitraum-Text gelesen: die Oberflaeche
        # braucht das Jahr, um den Werkstattbestand als Jahreskonstante abzulegen.
        "jahr": jahr,
        "produktgruppen": produktgruppen,
        "unbekannte_produktgruppen": unbekannte_pg,
        "unbekannte_kostenarten": sorted(unmapped),
        "per_regel_zugeordnet": per_regel,
        "unmapped_betrag": befunde["unmapped_betrag"],
        "fracht_ohne_ziel": befunde["fracht_ohne_ziel"],
        "bwa_ergebnis": bwa_ergebnis,
        "ufe_warnungen": ufe_warnungen,
        "ufe_statistik": ufe_statistik,
        "produktgruppen_auswahl": list(pg_config["reihenfolge"]),
        "zuschlagssaetze": saetze,
        "mgk_hinweise": satz_hinweise,
        "unproduktive_gemeinkosten": unprod_betrag,
        "unproduktive_stunden": unprod_stunden,
        "deckungsdifferenzen": {
            "geladen": dd_geladen,
            "hinweise": (dd_befund or {}).get("hinweise", []),
            "summen": {z: float(result.loc[z, "Summe"]) for z in DD_ZEILEN} if dd_geladen else {},
        },
        "struktur": ({k: v for k, v in struktur.items() if k != "paare"} if struktur else None),
        "naeherung_zellen": len(naeherung_zellen),
        "abstimmung": abstimmung,
        "guv": ({k: v for k, v in guv.items() if k not in ("konten", "susa_pfad")}
                | {"konten": guv["konten"][:25]} if guv else None),
        "susa_geladen": bool(susa_pfad),
        "db3_summe": float(result.loc["db3", "Summe"]),
        "erloese_summe": float(result.loc["erloese", "Summe"]),
    }
