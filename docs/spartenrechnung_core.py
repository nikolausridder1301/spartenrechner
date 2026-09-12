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
    "db3_vor_sonderposten": "formula", "sonderposten": "manual", "db3": "formula",
}

FILL_NULL = PatternFill("solid", fgColor="EFEFEF")       # grau: strukturell null

FILLS = {"auto": FILL_AUTO, "auto_pruef": FILL_AUTO_PRUEF, "manual": FILL_MANUAL,
         "formula": FILL_FORMULA, "strukturell_null": FILL_NULL}


def lade_betriebsparameter(config_dir):
    """Laedt die Betriebsparameter (Stundensatz, Anfangsbestaende), falls vorhanden.

    Diese Werte sind Kalkulationsinnenleben und liegen deshalb bewusst NICHT im
    Repository: aus Stundensatz und Zuschlagssaetzen liesse sich die Preisbildung
    zurueckrechnen. Die Datei bleibt lokal (siehe .gitignore) bzw. wird in der
    Weboberflaeche einmalig hochgeladen und dort im Browser gespeichert.

    Fehlt sie, rechnet das Tool alle GuV-Zeilen unveraendert weiter - nur die
    UFE-Zeile und die Umrechnung von Euro in Stunden entfallen.

    Gesucht wird an zwei Orten: im uebergebenen config-Verzeichnis (so schreibt die
    Weboberflaeche die Datei ins virtuelle Dateisystem) und daneben in 'intern/'
    (dort liegt das private Repository, wenn lokal per Kommandozeile gearbeitet wird).
    """
    kandidaten = [
        os.path.join(config_dir, "betriebsparameter.json"),
        os.path.join(config_dir, os.pardir, "intern", "betriebsparameter.json"),
    ]
    for pfad in kandidaten:
        if os.path.exists(pfad):
            with open(pfad, encoding="utf-8") as f:
                return json.load(f)
    return {}


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


def load_kptm(path):
    df = pd.read_excel(path, sheet_name="Penta")
    df.columns = [c.strip() for c in df.columns]
    df["Kostenart"] = df["Kostenart"].astype(str).str.strip()
    df["Produktgruppe"] = df["Produktgruppe"].astype(str).str.strip()
    df.loc[df["Produktgruppe"].isin(["nan", "None", ""]), "Produktgruppe"] = None
    return df


def _jahr_aus_zeitraum(zeitraum):
    """Liest die Jahreszahl aus dem Zeitraum-Text, z.B. 'Jan-Maerz 2026' -> '2026'."""
    treffer = re.findall(r"(20\d{2})", str(zeitraum))
    return treffer[-1] if treffer else None


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
    """Laedt den Werkstattbestand zum Jahresanfang aus den Betriebsparametern."""
    if not jahr:
        return None
    return lade_betriebsparameter(config_dir).get("anfangsbestand", {}).get(str(jahr))


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

    if kptm_df is not None:
        auftrag_kandidaten = {}
        for kobj, art, pg in zip(kptm_df["Kostenobjekt"], kptm_df["Artikelnummer"],
                                 kptm_df["Produktgruppe"]):
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
                nach_auftrag[_schluessel(rm)] = pg  # PFAK ist die genauere Quelle
            if _schluessel(art):
                artikel_kandidaten.setdefault(_schluessel(art), set()).add(pg)

    nach_artikel = {a: next(iter(v)) for a, v in artikel_kandidaten.items() if len(v) == 1}
    return nach_auftrag, nach_artikel


def werkstattbestand(pwbs_pfad, nach_auftrag, nach_artikel, manuelle_zuordnung=None):
    """Summiert den offenen Werkstattbestand je Produktgruppe zu einem Stichtag.

    Massgeblich ist die Spalte 'Offener Wert' (= Istwert abzueglich bereits
    abgeliefertem Wert), nicht 'Istwert' - letztere verfehlt die Referenzwerte.

    Ueber `manuelle_zuordnung` ({Auftragsnummer: Produktgruppe}) lassen sich einzelne
    Auftraege von Hand zuordnen - gedacht fuer die wenigen Posten, die keine Quelle
    kennt. Diese Zuordnung hat Vorrang vor den automatischen Quellen.

    Liefert (werte_je_produktgruppe, statistik). Die Statistik enthaelt auch die Liste
    der nicht zuordenbaren Auftraege, nach Wert absteigend - so lassen sich gezielt die
    groessten zuerst von Hand zuordnen.
    Nicht zuordenbare Zeilen werden danach unterschieden, ob sie eine Artikelnummer
    tragen: Serviceauftraege (Monteureinsatz, Inbetriebnahme, Schulung) haben keine,
    weil nichts gefertigt wird - sie zaehlen auch in der manuellen Referenzrechnung
    nicht zum Werkstattbestand. Ein nicht zuordenbarer Auftrag MIT Artikelnummer ist
    dagegen eine echte Luecke in der PFAK-Abdeckung.
    """
    df = pd.read_excel(pwbs_pfad, sheet_name="Penta")
    df.columns = [c.strip() for c in df.columns]
    je_pg = {}
    zugeordnet = service = luecke = 0.0
    offene_auftraege = []
    manuell = manuelle_zuordnung or {}
    for rm, art, bez, wert in zip(df["Rückmeldenummer"], df["Artikelnummer"],
                                  df["Bezeichnung"], df["Offener Wert"]):
        wert = float(wert or 0.0)
        if _schluessel(art) is None:
            # Serviceauftrag (Monteureinsatz, Inbetriebnahme, Schulung): kein Artikel,
            # also nichts gefertigt. Zaehlt in der Referenzrechnung nicht zum
            # Werkstattbestand und wird deshalb GAR NICHT zugeordnet - auch dann nicht,
            # wenn eine Quelle eine Produktgruppe dafuer kennt. Ohne diese Sperre wandern
            # z.B. drei Monteureinsaetze nach PB0005 und verfaelschen die Sparte.
            service += wert
            continue
        pg = (manuell.get(_schluessel(rm))
              or nach_auftrag.get(_schluessel(rm))
              or nach_artikel.get(_schluessel(art)))
        if pg is not None:
            je_pg[pg] = je_pg.get(pg, 0.0) + wert
            zugeordnet += wert
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
    }
    return je_pg, statistik


def bestandsveraenderung_ufe(pwbs_ende, pfak_pfade=None, anfangsbestand=None, pwbs_anfang=None,
                             max_luecke=0.01, jahr=None, kptm_df=None, manuelle_zuordnung=None):
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
      Restluecke 4,7 %  ->   6 Sparten falsch, groesste Abweichung 27.957 EUR
    Die nicht zuordenbaren Auftraege verteilen sich also keineswegs gleichmaessig -
    schon wenige Prozent Luecke koennen eine einzelne Sparte deutlich verfaelschen.
    """
    nach_auftrag, nach_artikel = lade_produktgruppen_map(pfak_pfade, kptm_df=kptm_df)
    ende, statistik = werkstattbestand(pwbs_ende, nach_auftrag, nach_artikel, manuelle_zuordnung)
    luecke = statistik["luecke_quote"]

    warnungen = []
    if luecke > max_luecke:
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
                                               manuelle_zuordnung)
        luecke_a = stat_anfang["luecke_quote"]
        if luecke_a > max_luecke:
            warnungen.append(
                f"Anfangsbestand: {luecke_a:.1%} nicht zuordenbar (siehe Hinweis oben)."
            )
    elif anfangsbestand:
        anfang = anfangsbestand
    else:
        warnungen.append(
            f"Kein Anfangsbestand fuer das Geschaeftsjahr {jahr or '(unbekannt)'} hinterlegt. "
            f"Abhilfe: entweder einen PWBS-Export zum Periodenbeginn (Stichtag 01.01.) zusaetzlich "
            f"hochladen, oder den Anfangsbestand einmalig in config/werkstattbestand_anfang.json "
            f"ergaenzen. Fuer 2026 ist er bereits hinterlegt."
        )
        return None, warnungen, statistik

    if warnungen:
        return None, warnungen, statistik

    werte = {pg: ende.get(pg, 0.0) - anfang.get(pg, 0.0) for pg in set(anfang) | set(ende)}
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
                    ziel_pg = pg  # kein Fracht-Kostenobjekt vorhanden: unveraendert lassen
            result.loc[regel["zeile"], ziel_pg] += regel["sign"] * wert
        elif ka.startswith(kst_prefix):
            # Nur Wertart ISWF (Ist-WERT in EUR). ISMF waere die Ist-MENGE in Stunden
            # und darf nicht in den Eurobetrag addiert werden.
            if row["Wertart"] == kst_wertart:
                result.loc["fek", pg] += wert
        else:
            unmapped_kostenarten.add(ka)

    result["Summe"] = result.sum(axis=1)
    result = result[["Summe"] + produktgruppen]
    result = neu_berechnen(result)

    return result, produktgruppen, unmapped_kostenarten


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
      Stunden = Betrag / Standardstundensatz
    """
    kst_prefix = mapping["kostenstellen_praefix"]
    kst_wertart = mapping.get("kostenstellen_wertart", "ISWF")
    satz = mapping.get("standard_stundensatz")
    art_gemein = set(mapping.get("auftragsart_gemeinkosten", ["5"]))
    art_lager = set(mapping.get("auftragsart_lager", ["LAG", "DPL", "PBL", "INN", "ITL", "SRL"]))

    sub = df[(df["Kostenart"].str.startswith(kst_prefix)) & (df["Wertart"] == kst_wertart)]
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
            "Stunden gesamt": gesamt / satz if satz else 0.0,
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

    Der Betrag wird bewusst NICHT auf die Sparten verteilt: Welcher Schluessel die
    Referenzrechnung reproduziert, liess sich nicht ermitteln (bestes falsifizierbares
    Modell verfehlt um 21.396 EUR). Er wird nur als Information ausgewiesen - damit ist
    bekannt, welcher Betrag zu verteilen waere, sobald der BAB vorliegt.
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

    for schluessel, bez, grundlage in (
        ("mgk", "MGK", "Material + Fremdleistungen"),
        ("vvgk", "VVGK", "Herstellkosten = Material + Fremdleistungen + MGK + FEK"),
    ):
        if schluessel not in saetze:
            continue
        erwartet = mapping.get(f"{schluessel}_zuschlagssatz_erwartet")
        toleranz = mapping.get(f"{schluessel}_zuschlagssatz_toleranz", 0.005)
        if erwartet is None:
            continue
        if abs(saetze[schluessel] - erwartet) > toleranz:
            hinweise.append(
                f"Der rechnerische {bez}-Zuschlagssatz betraegt {saetze[schluessel]:.2%} statt der "
                f"erwarteten {erwartet:.1%} (Bemessungsgrundlage: {grundlage}). Entweder wurde der "
                f"Satz im ERP geaendert - dann bitte in config/kostenart_mapping.json nachziehen - "
                f"oder die Kontenzuordnung stimmt nicht mehr."
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


def write_output(result, mapping, zeitraum, out_path, bwa_ergebnis=None, bwa_sheet=None,
                 stunden_df=None, unproduktiv=None):
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

    row_i = 4
    for key in mapping["zeilen_reihenfolge"] + [
        "betriebsleistung", "rohmarge_1", "fek_nach_dd", "db1", "mgk_nach_dd", "db2",
        "vvgk_nach_dd", "db3_vor_sonderposten", "db3",
    ]:
        if key not in result.index:
            continue
        label_cell = ws.cell(row=row_i, column=1, value=labels.get(key, key))
        klass = ZEILEN_KLASSE.get(key, "manual")
        if klass == "formula":
            label_cell.font = FONT_BOLD
        for j, c in enumerate(cols, start=2):
            val = float(result.loc[key, c])
            cell = ws.cell(row=row_i, column=j, value=round(val, 2))
            cell.fill = FILLS[klass]
            cell.border = BORDER
            cell.number_format = "#,##0.00"
        row_i += 1

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
        c = ws.cell(row=row_i, column=1, value="Δ DB III (bottom-up) ./. Ergebnis lt. GuV")
        c.font = FONT_BOLD
        cell = ws.cell(row=row_i, column=2, value=round(delta, 2))
        cell.fill = FILL_FORMULA
        cell.border = BORDER
        cell.number_format = "#,##0.00"
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
        ws3.cell(row=2, column=1,
                 value=f"Stundenbewertung mit {mapping.get('standard_stundensatz')} EUR/h "
                       f"(Standardkostensatz des ERP)")
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
        "- Bestandsveraenderung UFE: Werkstattbestand(Stichtag) minus Werkstattbestand(01.01.), je Sparte",
        "  ueber PWBS.Rueckmeldenummer = PFAK.RUECKMELDE_NR zugeordnet -> 13/13 Sparten centgenau.",
        "  Wird nur berechnet, wenn PWBS- und PFAK-Dateien mitgegeben werden.",
        "",
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
        "MANUELL ZU ERGAENZEN (rosa) - aus den vorliegenden Rohdaten nachweislich NICHT herleitbar:",
        "- FEK-/MGK-/VVGK-Deckungsdifferenz: Gegenueberstellung der ueber die Kostentraeger VERRECHNETEN",
        "  Betraege (= die gruenen Zeilen oben) mit den TATSAECHLICHEN Kostenstellenkosten. Die Ist-Seite",
        "  stammt aus dem BAB (Betriebsabrechnungsbogen). Keiner der sechs ERP-Exporte ist eine",
        "  Kostenstellenrechnung: die Ist-Kosten der Kostenstellen Lager und Einkauf kommen darin gar",
        "  nicht vor, QS nur als Stundenbewertung. Ueber 100 Herleitungshypothesen wurden numerisch",
        "  ausgeschlossen (siehe HERLEITUNG.md im Projekt).",
        "",
        "  Bei der FEK-Unterdeckung ist immerhin geklaert, WORAUS sie besteht: aus den Fertigungsstunden,",
        "  die geleistet und bezahlt, aber auf keinen Kostentraeger gebucht wurden (Auftragsart 5).",
        "  Rechnerisch: verrechnete FEK + diese unproduktiven Gemeinkosten ergibt den Wert, den die",
        "  manuelle Referenzrechnung als 'FEK nach Deckungsdifferenz (IST)' ausweist - auf 0,25 % genau.",
        "  Die Hoehe steht deshalb oben im Blatt. Offen ist nur, nach welchem Schluessel der Betrag auf",
        "  die einzelnen Sparten verteilt wird; das beste ueberpruefbare Modell verfehlt um 21.396 EUR.",
        "- Sonderposten/Einmaleffekte: bewusste manuelle Korrekturen, jede Periode neu zu beurteilen.",
        "",
        "HINWEIS ZUM ERP-EXPORT: Die Datei KPTM_Wertsummen ist auf 'Datenbestand = Istdaten' gefiltert",
        "(alle 15.045 Zeilen tragen diesen Wert) und enthaelt nur die Wertarten ISWV/ISWF/ISMF. Falls die",
        "gleiche Abfrage ohne diesen Filter auch Soll-/Plandaten liefert, koennten damit die drei",
        "Deckungsdifferenzen automatisiert werden - das waere der naechste sinnvolle Ausbauschritt.",
        "",
        "Quelle: " + zeitraum,
    ]
    for i, line in enumerate(hinweise, start=1):
        ws2.cell(row=i, column=1, value=line)
    ws2.column_dimensions["A"].width = 140

    wb.save(out_path)


def generate(kptm_path, config_dir, zeitraum, out_path, bwa_path=None, bwa_sheet=None,
             pwbs_anfang=None, pwbs_ende=None, pfak_pfade=None, manuelle_zuordnung=None):
    """Ein-Funktions-Einstieg fuer die Weboberflaeche: liest KPTM (+optional BWA,
    +optional Werkstattbestand fuer die Bestandsveraenderung UFE), berechnet die
    Spartenrechnung und schreibt die Ausgabedatei. Gibt eine kurze Zusammenfassung
    (dict) zurueck, u.a. fuer Warnungen zu unbekannten Kostenarten/Produktgruppen."""
    mapping, pg_config = load_config(config_dir)
    df = load_kptm(kptm_path)
    result, produktgruppen, unmapped = build_spartenrechnung(df, mapping, pg_config)

    ufe_warnungen = []
    ufe_statistik = None
    if pwbs_ende:
        # PFAK ist optional: die Zuordnung Auftrag -> Produktgruppe steckt bereits in KPTM.
        jahr = geschaeftsjahr(df, zeitraum)
        gespeicherter_ab = lade_anfangsbestand(config_dir, jahr)
        ufe, ufe_warnungen, ufe_statistik = bestandsveraenderung_ufe(
            pwbs_ende, pfak_pfade, anfangsbestand=gespeicherter_ab, pwbs_anfang=pwbs_anfang,
            jahr=jahr, kptm_df=df, manuelle_zuordnung=manuelle_zuordnung)
        if ufe:
            for pg, wert in ufe.items():
                if pg in result.columns:
                    result.loc["bestand_ufe", pg] = wert
            result = neu_berechnen(result)

    bwa_ergebnis = None
    if bwa_path and bwa_sheet:
        bwa_ergebnis = read_bwa_ergebnis(bwa_path, bwa_sheet)

    saetze, satz_hinweise = zuschlagssaetze(result, mapping)
    unprod_betrag, unprod_stunden = unproduktive_gemeinkosten(df, mapping)
    stunden_df = fertigungsstunden(df, mapping)
    write_output(result, mapping, zeitraum, out_path, bwa_ergebnis=bwa_ergebnis, bwa_sheet=bwa_sheet,
                 stunden_df=stunden_df, unproduktiv=(unprod_betrag, unprod_stunden))

    known_pg = set(pg_config["reihenfolge"])
    unbekannte_pg = [p for p in produktgruppen if p not in known_pg]

    return {
        "zeilen": len(df),
        "produktgruppen": produktgruppen,
        "unbekannte_produktgruppen": unbekannte_pg,
        "unbekannte_kostenarten": sorted(unmapped),
        "bwa_ergebnis": bwa_ergebnis,
        "ufe_warnungen": ufe_warnungen,
        "ufe_statistik": ufe_statistik,
        "produktgruppen_auswahl": list(pg_config["reihenfolge"]),
        "zuschlagssaetze": saetze,
        "mgk_hinweise": satz_hinweise,
        "unproduktive_gemeinkosten": unprod_betrag,
        "unproduktive_stunden": unprod_stunden,
        "db3_summe": float(result.loc["db3", "Summe"]),
        "erloese_summe": float(result.loc["erloese", "Summe"]),
    }
