"""Baut in die Ausgabedatei die Rohdaten UND die Formeln ein, die aus ihnen die
Spartenrechnung ergeben.

Der Zweck ist Nachvollziehbarkeit: Wer die Datei bekommt, soll jede Zahl anklicken
und sehen koennen, woraus sie entsteht - ohne dieses Programm zu kennen und ohne
uns fragen zu muessen. Excel rechnet die Zellen selbst, die Werte stehen also nicht
nur da, sie entstehen vor den Augen des Lesers neu.

Aufbau der erzeugten Mappe:

    Spartenrechnung        jede Zelle eine SUMMEWENNS-Formel auf 'Daten_KPTM'
    Daten_KPTM             der hochgeladene KPTM-Export, plus vier Hilfsspalten,
                           die selbst Formeln sind (Zeile, Vorzeichen, Ziel-Sparte,
                           Betrag) - die Zuordnung ist damit Zeile fuer Zeile sichtbar
    Mapping                Kostenart -> GuV-Zeile, Vorzeichen, Umbuchung.
                           Quelle der Hilfsspalten, hier nachschlagbar
    Daten_PWBS_Ende        Werkstattbestand zum Stichtag, mit Sparten-Zuordnung
    Daten_PWBS_Anfang      dito zum Periodenbeginn, falls hochgeladen
    Zuordnung_Auftraege    Fertigungsauftrag -> Sparte samt Quelle der Zuordnung
    Betriebsparameter      Anfangsbestaende je Sparte
    Kontrolle              die von Python gerechneten Werte. Die Spartenrechnung
                           zeigt oben die groesste Abweichung zwischen Formel und
                           Kontrollwert an - steht dort nicht 0,00, stimmt etwas nicht

Bewusst NICHT verformelt sind die drei Deckungsdifferenzen (sie kommen aus dem BAB,
nicht aus diesen Daten) und die Zeile Bestandsveraenderung UFE, deren Sparten-
Zuordnung ueber mehrere Quellen laeuft; fuer sie verweist die Mappe auf die
Datenblaetter, statt eine Formel vorzutaeuschen, die die Herkunft verschleiert.
"""
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Zeilen, die aus KPTM per SUMMEWENNS entstehen
KPTM_ZEILEN = {"erloese", "bestand_fe", "material", "fremdleistungen", "personalaufwand",
               "sbe", "sba", "afa", "fek", "mgk", "vvgk", "aktivierte_eigenleistung"}

# Zeilen, die sich aus anderen Zeilen ergeben: Zielzeile -> (Summanden, Subtrahenden)
ABGELEITET = {
    "betriebsleistung": (["erloese", "bestand_fe", "bestand_ufe", "aktivierte_eigenleistung"], []),
    "rohmarge_1": (["betriebsleistung"],
                   ["material", "fremdleistungen", "personalaufwand", "sbe", "sba", "afa"]),
    "fek_nach_dd": (["fek", "fek_deckungsdifferenz"], []),
    "db1": (["rohmarge_1"], ["fek_nach_dd"]),
    "mgk_nach_dd": (["mgk", "mgk_deckungsdifferenz"], []),
    "db2": (["db1"], ["mgk_nach_dd"]),
    "vvgk_nach_dd": (["vvgk", "vvgk_deckungsdifferenz"], []),
    "db3_vor_sonderposten": (["db2"], ["vvgk_nach_dd"]),
    "db3": (["db3_vor_sonderposten", "sonderposten"], []),
}

FILL_DATEN = PatternFill("solid", fgColor="1F3864")
FONT_WEISS = Font(color="FFFFFF", bold=True)
FILL_HILFS = PatternFill("solid", fgColor="FFF2CC")


def _blatt_mit_daten(wb, titel, df, hinweis=None):
    """Schreibt einen DataFrame als Blatt. Gibt (Blatt, erste Datenzeile) zurueck."""
    ws = wb.create_sheet(titel)
    start = 1
    if hinweis:
        ws.cell(row=1, column=1, value=hinweis).font = Font(italic=True, size=9)
        start = 3
    for j, spalte in enumerate(df.columns, start=1):
        z = ws.cell(row=start, column=j, value=str(spalte))
        z.font = FONT_WEISS
        z.fill = FILL_DATEN
    for i, (_, zeile) in enumerate(df.iterrows(), start=start + 1):
        for j, spalte in enumerate(df.columns, start=1):
            wert = zeile[spalte]
            if pd.isna(wert):
                continue
            if isinstance(wert, pd.Timestamp):
                wert = wert.to_pydatetime()
            ws.cell(row=i, column=j, value=wert)
    ws.freeze_panes = ws.cell(row=start + 1, column=1)
    return ws, start + 1


def _mapping_blatt(wb, mapping, produktgruppen):
    """Kostenart -> Zeile/Vorzeichen/Umbuchung, als nachschlagbare Tabelle."""
    kz = mapping["kostenart_zu_zeile"]
    labels = mapping["zeilen_labels"]
    zeilen = []
    for ka, regel in sorted(kz.items()):
        zeilen.append({
            "Kostenart": str(ka),
            "Zeile": regel["zeile"],
            "Zeilenbezeichnung": labels.get(regel["zeile"], regel["zeile"]),
            "Vorzeichen": regel["sign"],
            "Umbuchung": regel.get("umbuchung", ""),
        })
    df = pd.DataFrame(zeilen)
    ws, _ = _blatt_mit_daten(
        wb, "Mapping", df,
        "Welches Sachkonto auf welche GuV-Zeile geht. Quelle der Hilfsspalten in "
        "'Daten_KPTM'. 'Umbuchung = fracht' bedeutet: der Betrag wird auf das "
        "Frachtkostenobjekt der Sparten-Familie gebucht, nicht auf die Sparte selbst.")
    for sp, breite in (("A", 12), ("B", 22), ("C", 34), ("D", 11), ("E", 12)):
        ws.column_dimensions[sp].width = breite

    # Die gueltigen Sparten als sichtbarer Bereich. Die Fracht-Umbuchung in
    # 'Daten_KPTM' schlaegt hier nach, ob es das Frachtkostenobjekt ueberhaupt gibt.
    # (Als Array-Konstante in der Formel verweigert Excel die Datei.)
    ws.cell(row=1, column=7, value="Gültige Sparten").font = Font(bold=True)
    for i, pg in enumerate(produktgruppen, start=2):
        ws.cell(row=i, column=7, value=pg)
    ws.column_dimensions["G"].width = 16

    # Auffangregel nach Kontenpraefix. Sie muss hier stehen, damit die Hilfsspalten
    # in 'Daten_KPTM' dasselbe rechnen wie das Programm - sonst zeigt die Mappe
    # andere Zahlen als das Tool, und zwar unbemerkt.
    regel = mapping.get("kostenart_praefix_regel", {})
    ws.cell(row=1, column=9, value="Auffangregel: Konto beginnt mit").font = Font(bold=True)
    ws.cell(row=1, column=10, value="Zeile").font = Font(bold=True)
    ws.cell(row=1, column=11, value="Vorzeichen").font = Font(bold=True)
    for i, (praefix, r) in enumerate(sorted(regel.items()), start=2):
        ws.cell(row=i, column=9, value=praefix)
        ws.cell(row=i, column=10, value=r["zeile"])
        ws.cell(row=i, column=11, value=r["sign"])
    ws.cell(row=len(regel) + 3, column=9, value=(
        "Greift nur, wenn links kein Einzeleintrag passt. Gilt fuer Konten, die in der "
        "Einzelliste fehlen, weil sie im Kalibrierungszeitraum nicht vorkamen."
    )).font = Font(italic=True, size=9)
    for sp, br in (("I", 30), ("J", 14), ("K", 11)):
        ws.column_dimensions[sp].width = br
    n = max(len(regel), 1)
    return ws, (f"Mapping!$G$2:$G${len(produktgruppen) + 1}", f"Mapping!$I$2:$K${n + 1}")


def _kptm_blatt(wb, df, mapping, bereiche):
    """Der KPTM-Export mit vier Hilfsspalten, die selbst Formeln sind.

    Dadurch ist die Zuordnung nicht das Ergebnis eines unsichtbaren Programmschritts,
    sondern in jeder Zeile nachvollziehbar und im Zweifel korrigierbar.
    """
    ws, erste = _blatt_mit_daten(
        wb, "Daten_KPTM", df,
        "Der hochgeladene KPTM-Export, unveraendert. Die vier gelben Spalten rechts "
        "sind Formeln: sie schlagen im Blatt 'Mapping' nach, auf welche GuV-Zeile eine "
        "Kostenart geht. Das Blatt 'Spartenrechnung' summiert dann ueber diese Spalten.")
    letzte = erste + len(df) - 1
    spalten = {str(c): get_column_letter(i) for i, c in enumerate(df.columns, start=1)}
    s_ka, s_pg = spalten["Kostenart"], spalten["Produktgruppe"]
    s_wa, s_wert = spalten["Wertart"], spalten["Wert/Menge"]

    kopf = erste - 1
    basis = len(df.columns)
    titel = ["Zeile (Formel)", "Vorzeichen (Formel)", "Ziel-Sparte (Formel)", "Betrag (Formel)"]
    for k, t in enumerate(titel):
        z = ws.cell(row=kopf, column=basis + 1 + k, value=t)
        z.font = Font(bold=True)
        z.fill = FILL_HILFS
    c_zeile = get_column_letter(basis + 1)
    c_vz = get_column_letter(basis + 2)
    c_ziel = get_column_letter(basis + 3)
    c_betrag = get_column_letter(basis + 4)

    prefix = mapping["kostenstellen_praefix"]
    wertart = mapping.get("kostenstellen_wertart", "ISWF")
    fracht = mapping.get("umbuchung_fracht_suffix", "0100")

    pg_bereich, regel_bereich = bereiche
    for r in range(erste, letzte + 1):
        # Reihenfolge wie im Programm: Einzeleintrag, dann Kostenstelle, dann Auffangregel.
        ws.cell(row=r, column=basis + 1, value=(
            f'=IFERROR(VLOOKUP(TEXT({s_ka}{r},"@"),Mapping!$A:$E,2,FALSE),'
            f'IF(AND(LEFT(TEXT({s_ka}{r},"@"),{len(prefix)})="{prefix}",'
            f'{s_wa}{r}="{wertart}"),"fek",'
            f'IFERROR(VLOOKUP(LEFT(TEXT({s_ka}{r},"@"),1),{regel_bereich},2,FALSE),"")))'))
        ws.cell(row=r, column=basis + 2, value=(
            f'=IFERROR(VLOOKUP(TEXT({s_ka}{r},"@"),Mapping!$A:$E,4,FALSE),'
            f'IFERROR(VLOOKUP(LEFT(TEXT({s_ka}{r},"@"),1),{regel_bereich},3,FALSE),1))'))
        ws.cell(row=r, column=basis + 3, value=(
            f'=IF(IFERROR(VLOOKUP(TEXT({s_ka}{r},"@"),Mapping!$A:$E,5,FALSE),"")="fracht",'
            f'IF(COUNTIF({pg_bereich},LEFT({s_pg}{r},2)&"{fracht}")>0,'
            f'LEFT({s_pg}{r},2)&"{fracht}",{s_pg}{r}),{s_pg}{r})'))
        z = ws.cell(row=r, column=basis + 4, value=f'={c_vz}{r}*{s_wert}{r}')
        z.number_format = "#,##0.00"

    ws.auto_filter.ref = f"A{kopf}:{c_betrag}{letzte}"
    for c in (c_zeile, c_vz, c_ziel, c_betrag):
        ws.column_dimensions[c].width = 19
    return {"blatt": "Daten_KPTM", "erste": erste, "letzte": letzte,
            "zeile": c_zeile, "ziel": c_ziel, "betrag": c_betrag}


def _pwbs_blatt(wb, pfad, titel, hinweis):
    """Werkstattbestand, bei dem die Sparte per INDEX/VERGLEICH ermittelt wird.

    Die Zuordnung ist damit nicht das Ergebnis eines unsichtbaren Programmschritts,
    sondern in der Zelle selbst nachlesbar. Sie bildet die Logik des Programms genau
    ab, in dieser Reihenfolge:
      1. keine Artikelnummer -> Serviceauftrag, zaehlt gar nicht mit (Monteureinsatz,
         Inbetriebnahme, Schulung: es wird nichts gefertigt)
      2. Treffer ueber die Auftragsnummer in 'Zuordnung_Auftraege'
      3. sonst Treffer ueber die Artikelnummer in 'Zuordnung_Artikel'
      4. sonst nicht zuordenbar
    """
    roh = pd.read_excel(pfad, sheet_name="Penta")
    roh.columns = [str(c).strip() for c in roh.columns]
    ws, erste = _blatt_mit_daten(wb, titel, roh, hinweis)
    letzte = erste + len(roh) - 1
    spalten = {str(c): get_column_letter(i) for i, c in enumerate(roh.columns, start=1)}
    s_rm, s_art = spalten["Rückmeldenummer"], spalten["Artikelnummer"]
    s_wert = spalten["Offener Wert"]

    kopf = erste - 1
    c_sparte = get_column_letter(len(roh.columns) + 1)
    z = ws.cell(row=kopf, column=len(roh.columns) + 1, value="Sparte (Formel)")
    z.font = Font(bold=True)
    z.fill = FILL_HILFS
    for r in range(erste, letzte + 1):
        ws.cell(row=r, column=len(roh.columns) + 1, value=(
            f'=IF({s_art}{r}="","(Serviceauftrag)",'
            f'IFERROR(INDEX(Zuordnung_Auftraege!$B:$B,'
            f'MATCH(TEXT({s_rm}{r},"@"),Zuordnung_Auftraege!$A:$A,0)),'
            f'IFERROR(INDEX(Zuordnung_Artikel!$B:$B,'
            f'MATCH(TEXT({s_art}{r},"@"),Zuordnung_Artikel!$A:$A,0)),'
            f'"(nicht zuordenbar)")))'))
    ws.column_dimensions[c_sparte].width = 20
    ws.auto_filter.ref = f"A{kopf}:{c_sparte}{letzte}"
    return {"blatt": titel, "erste": erste, "letzte": letzte,
            "sparte": c_sparte, "wert": s_wert}


def _kontrollblatt(wb, result, produktgruppen):
    """Die von Python gerechneten Werte, damit die Formeln dagegen pruefbar sind."""
    ws = wb.create_sheet("Kontrolle")
    ws.cell(row=1, column=1, value=(
        "Von der Berechnungslogik ermittelte Werte. Das Blatt 'Spartenrechnung' "
        "vergleicht seine Formelergebnisse damit und zeigt oben die groesste "
        "Abweichung. Steht dort nicht 0,00, weichen Formel und Programm voneinander "
        "ab - dann bitte melden.")).font = Font(italic=True, size=9)
    ws.cell(row=3, column=1, value="Zeile").font = Font(bold=True)
    for j, pg in enumerate(produktgruppen, start=2):
        ws.cell(row=3, column=j, value=pg).font = Font(bold=True)
    for i, key in enumerate(result.index, start=4):
        ws.cell(row=i, column=1, value=key)
        for j, pg in enumerate(produktgruppen, start=2):
            z = ws.cell(row=i, column=j, value=float(result.loc[key, pg]))
            z.number_format = "#,##0.00"
    return {"erste": 4, "zeilen": {k: 4 + i for i, k in enumerate(result.index)}}


def baue_herleitungsmappe(wb, result, mapping, produktgruppen, kptm_df,
                          pwbs_ende=None, pwbs_anfang=None, nach_auftrag=None,
                          nach_artikel=None, anfangsbestand=None,
                          manuelle_zuordnung=None, ufe_berechnet=False):
    """Haengt an eine bestehende Mappe die Datenblaetter an und ersetzt die Werte
    im Blatt 'Spartenrechnung' durch Formeln, die auf diese Blaetter zeigen.

    Gibt die Anzahl verformelter Zellen zurueck.
    """
    _, bereiche = _mapping_blatt(wb, mapping, produktgruppen)
    kptm = _kptm_blatt(wb, kptm_df, mapping, bereiche)

    # Die beiden Nachschlagetabellen muessen VOR den PWBS-Blaettern stehen,
    # weil deren Sparten-Formel hier hineinschlaegt.
    if nach_auftrag:
        eintraege = dict(nach_auftrag)
        quelle = {a: "Auftragsnummer (KPTM/PFAK)" for a in eintraege}
        for a, pg in (manuelle_zuordnung or {}).items():
            eintraege[a] = pg          # Handzuordnung hat Vorrang
            quelle[a] = "von Hand zugeordnet"
        df_z = pd.DataFrame([{"Fertigungsauftrag": a, "Sparte": p, "Quelle": quelle[a]}
                             for a, p in sorted(eintraege.items())])
        ws, _ = _blatt_mit_daten(
            wb, "Zuordnung_Auftraege", df_z,
            "Welcher Fertigungsauftrag zu welcher Sparte gehoert. Gewonnen aus der Spalte "
            "'Kostenobjekt' des KPTM-Exports und aus 'RUECKMELDE_NR' der PFAK-Datei; wo beide "
            "Quellen etwas wissen, stimmen sie ueberein. Von Hand vorgenommene Zuordnungen "
            "haben Vorrang und sind in der Spalte 'Quelle' als solche gekennzeichnet. "
            "Die Blaetter 'Daten_PWBS_*' schlagen hier nach.")
        for sp, br in (("A", 20), ("B", 12), ("C", 26)):
            ws.column_dimensions[sp].width = br

    if nach_artikel:
        df_a = pd.DataFrame([{"Artikelnummer": a, "Sparte": p}
                             for a, p in sorted(nach_artikel.items())])
        ws, _ = _blatt_mit_daten(
            wb, "Zuordnung_Artikel", df_a,
            "Ersatzschluessel: Artikelnummern, die eindeutig zu genau einer Sparte gehoeren. "
            "Wird nur herangezogen, wenn die Auftragsnummer keinen Treffer liefert.")
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 12

    pwbs_e = pwbs_a = None
    if pwbs_ende:
        pwbs_e = _pwbs_blatt(
            wb, pwbs_ende, "Daten_PWBS_Ende",
            "Werkstattbestand zum Stichtag. Massgeblich ist die Spalte 'Offener Wert'. "
            "Die gelbe Spalte rechts ermittelt die Sparte per INDEX/VERGLEICH - erst ueber "
            "die Auftragsnummer, ersatzweise ueber die Artikelnummer. Serviceauftraege "
            "(ohne Artikelnummer) zaehlen bewusst nicht mit.")
    if pwbs_anfang:
        pwbs_a = _pwbs_blatt(
            wb, pwbs_anfang, "Daten_PWBS_Anfang",
            "Werkstattbestand zum Periodenbeginn (01.01.), gleiche Systematik wie das "
            "Blatt 'Daten_PWBS_Ende'.")

    if anfangsbestand:
        df_b = pd.DataFrame([{"Sparte": p, "Anfangsbestand": v}
                             for p, v in sorted(anfangsbestand.items())])
        ws, erste = _blatt_mit_daten(
            wb, "Betriebsparameter", df_b,
            "Werkstattbestand je Sparte zum 01.01. Grundlage der Zeile "
            "'Bestandsveraenderung UFE': UFE = Bestand(Stichtag) - Anfangsbestand.")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 18
        for r in range(erste, erste + len(df_b)):
            ws.cell(row=r, column=2).number_format = "#,##0.00"

    kontrolle = _kontrollblatt(wb, result, produktgruppen)
    anzahl = _verformele_spartenrechnung(
        wb, result, mapping, produktgruppen, kptm, kontrolle,
        pwbs_e=pwbs_e, pwbs_a=pwbs_a,
        anfangsbestand_blatt=bool(anfangsbestand), ufe_berechnet=ufe_berechnet)

    if "Hinweise" in wb.sheetnames:
        ws = wb["Hinweise"]
        r = ws.max_row + 2
        for zeile in [
            "NACHVOLLZIEHBARKEIT",
            f"Im Blatt 'Spartenrechnung' sind {anzahl} Zellen Formeln, die auf die "
            "Datenblaetter zeigen. Klicken Sie eine Zahl an, um zu sehen, woraus sie entsteht.",
            "",
            "  Zeilen MIT Formel:  Erloese, Bestandsveraenderung FE, Material, Fremdleistungen,",
            "                      sbA, FEK, MGK, overhead Kosten - sie summieren ueber "
            "'Daten_KPTM'.",
            "                      Dazu alle Zwischensummen (Betriebsleistung, Rohmarge I, DB I-III).",
            "",
            "  Bestandsveraenderung UFE hat ebenfalls eine Formel:",
            "    = Bestand(Stichtag) aus 'Daten_PWBS_Ende' minus Anfangsbestand.",
            "    Die Sparte je Auftrag ermittelt dort eine INDEX/VERGLEICH-Formel - erst ueber die",
            "    Auftragsnummer ('Zuordnung_Auftraege'), ersatzweise ueber die Artikelnummer",
            "    ('Zuordnung_Artikel'). Auftraege ohne Artikelnummer sind Serviceauftraege und",
            "    zaehlen nicht mit.",
            "",
            "  Zeilen OHNE Formel - und warum:",
            "    Die drei Deckungsdifferenzen  stammen aus dem BAB, nicht aus diesen Exporten.",
            "                               Sie sind hier leer und von Hand zu ergaenzen.",
            "    Sonderposten               bewusste Einzelfall-Entscheidung je Periode.",
            "",
            "  Selbstpruefung: Oben links im Blatt 'Spartenrechnung' steht die groesste "
            "Abweichung zwischen",
            "  dem, was die Formeln rechnen, und dem, was das Programm gerechnet hat. Dort "
            "gehoert 0,00 zu stehen.",
        ]:
            ws.cell(row=r, column=1, value=zeile)
            r += 1
    return anzahl


def _verformele_spartenrechnung(wb, result, mapping, produktgruppen, kptm, kontrolle,
                                pwbs_e=None, pwbs_a=None, anfangsbestand_blatt=False,
                                ufe_berechnet=False):
    """Ersetzt die Zahlen im Blatt 'Spartenrechnung' durch Formeln auf die Rohdaten."""
    ws = wb["Spartenrechnung"]

    # Kopfzeile und Zeilenschluessel im bestehenden Blatt wiederfinden
    kopfzeile = spalte_erste_sparte = None
    for r in range(1, 12):
        werte = [ws.cell(row=r, column=c).value for c in range(1, 40)]
        if any(w in produktgruppen for w in werte if w):
            kopfzeile = r
            break
    if kopfzeile is None:
        return 0
    pg_spalte = {}
    for c in range(1, 40):
        w = ws.cell(row=kopfzeile, column=c).value
        if w in produktgruppen:
            pg_spalte[w] = c
        if w == "Summe":
            summen_spalte = c
    # In Spalte A steht die Bezeichnung, nicht der interne Schluessel - also
    # ueber die Label-Tabelle zurueckuebersetzen.
    label_zu_key = {}
    for key in result.index:
        label_zu_key[str(mapping["zeilen_labels"].get(key, key)).strip()] = key

    def key_zu_beschriftung(text):
        """Findet den Zeilenschluessel zur Beschriftung in Spalte A.

        Der Vergleich darf nicht exakt sein: Eine zurueckgehaltene Zeile traegt einen
        Zusatz wie '  (nicht berechenbar - siehe Hinweise)'. Wird sie dadurch nicht
        wiedererkannt, faellt sie aus den Summenformeln heraus - und ein spaeter von
        Hand nachgetragener Wert bliebe wirkungslos, ohne dass das jemand bemerkt.
        """
        t = str(text).strip()
        if t in label_zu_key:
            return label_zu_key[t]
        for trenner in ("  (", " ("):
            if trenner in t and t.split(trenner)[0].strip() in label_zu_key:
                return label_zu_key[t.split(trenner)[0].strip()]
        return None
    # Nur der zusammenhaengende Tabellenblock direkt unter der Kopfzeile. Weiter
    # unten steht die Datengrundlage des Diagramms, deren Zeilen genauso heissen
    # ("Material", "Fremdleistungen"). Ohne diese Grenze wuerden sie mitverformelt -
    # und zwar mit den Spalten der Haupttabelle, die dort nicht gelten.
    zeilen_nr = {}
    for r in range(kopfzeile + 1, kopfzeile + 60):
        beschriftung = ws.cell(row=r, column=1).value
        if beschriftung is None or not str(beschriftung).strip():
            break
        key = key_zu_beschriftung(beschriftung)
        if key:
            zeilen_nr[key] = r

    q_blatt, q_zeile = kptm["blatt"], kptm["zeile"]
    q_ziel, q_betrag = kptm["ziel"], kptm["betrag"]
    von, bis = kptm["erste"], kptm["letzte"]
    bereich_zeile = f"'{q_blatt}'!${q_zeile}${von}:${q_zeile}${bis}"
    bereich_ziel = f"'{q_blatt}'!${q_ziel}${von}:${q_ziel}${bis}"
    bereich_betrag = f"'{q_blatt}'!${q_betrag}${von}:${q_betrag}${bis}"

    anzahl = 0
    for key, r in zeilen_nr.items():
        for pg, c in pg_spalte.items():
            spalte = get_column_letter(c)
            if key == "bestand_ufe" and not ufe_berechnet:
                continue        # bleibt 0,00 - aber die Zeile zaehlt in den Summen mit
            if key in KPTM_ZEILEN:
                formel = (f'=SUMIFS({bereich_betrag},{bereich_zeile},"{key}",'
                          f'{bereich_ziel},{spalte}${kopfzeile})')
            elif key == "bestand_ufe" and ufe_berechnet and pwbs_e:
                # UFE = Bestand(Stichtag) - Anfangsbestand. Der Anfangsbestand kommt
                # entweder aus einem zweiten Werkstattbestand oder aus den hinterlegten
                # Betriebsparametern - je nachdem, was vorliegt.
                ende = (f"SUMIFS('{pwbs_e['blatt']}'!${pwbs_e['wert']}${pwbs_e['erste']}:"
                        f"${pwbs_e['wert']}${pwbs_e['letzte']},"
                        f"'{pwbs_e['blatt']}'!${pwbs_e['sparte']}${pwbs_e['erste']}:"
                        f"${pwbs_e['sparte']}${pwbs_e['letzte']},{spalte}${kopfzeile})")
                if pwbs_a:
                    anfang = (f"SUMIFS('{pwbs_a['blatt']}'!${pwbs_a['wert']}${pwbs_a['erste']}:"
                              f"${pwbs_a['wert']}${pwbs_a['letzte']},"
                              f"'{pwbs_a['blatt']}'!${pwbs_a['sparte']}${pwbs_a['erste']}:"
                              f"${pwbs_a['sparte']}${pwbs_a['letzte']},{spalte}${kopfzeile})")
                elif anfangsbestand_blatt:
                    anfang = (f"SUMIFS(Betriebsparameter!$B:$B,Betriebsparameter!$A:$A,"
                              f"{spalte}${kopfzeile})")
                else:
                    continue
                formel = f"={ende}-{anfang}"
            elif key in ABGELEITET:
                plus, minus = ABGELEITET[key]
                teile = [f"{spalte}{zeilen_nr[k]}" for k in plus if k in zeilen_nr]
                teile += [f"-{spalte}{zeilen_nr[k]}" for k in minus if k in zeilen_nr]
                formel = "=" + "+".join(teile).replace("+-", "-")
            else:
                continue   # UFE, Deckungsdifferenzen, Sonderposten: bleiben Werte
            z = ws.cell(row=r, column=c, value=formel)
            z.number_format = "#,##0.00"
            anzahl += 1
        # Summenspalte ueber die Sparten
        if zeilen_nr and summen_spalte:
            erste_pg = get_column_letter(min(pg_spalte.values()))
            letzte_pg = get_column_letter(max(pg_spalte.values()))
            z = ws.cell(row=r, column=summen_spalte, value=f"=SUM({erste_pg}{r}:{letzte_pg}{r})")
            z.number_format = "#,##0.00"

    # Selbstpruefung. Die Einzelvergleiche stehen als Block auf dem Kontrollblatt;
    # ein MAX ueber diesen Bereich genuegt dann in der Spartenrechnung. (MAX direkt
    # ueber Hunderte Einzelterme scheitert an der Grenze von 255 Argumenten.)
    kt = wb["Kontrolle"]
    kz = kontrolle["zeilen"]
    block = max(kz.values()) + 3
    kt.cell(row=block - 1, column=1,
            value="Abweichung Formel ./. Berechnung (je Zelle)").font = Font(bold=True, size=9)
    for j, pg in enumerate(produktgruppen, start=2):
        kt.cell(row=block, column=j, value=pg).font = Font(bold=True)
    kt.cell(row=block, column=1, value="Zeile").font = Font(bold=True)
    zeile_im_block = {}
    r_block = block + 1
    for key, r in zeilen_nr.items():
        if key not in kz:
            continue
        kt.cell(row=r_block, column=1, value=key)
        for pg, c in pg_spalte.items():
            sp = get_column_letter(c)
            kc = get_column_letter(list(produktgruppen).index(pg) + 2)
            z = kt.cell(row=r_block, column=list(produktgruppen).index(pg) + 2,
                        value=f"=ABS(Spartenrechnung!{sp}{r}-Kontrolle!{kc}{kz[key]})")
            z.number_format = "#,##0.00"
        zeile_im_block[key] = r_block
        r_block += 1

    if zeile_im_block:
        letzte_sp = get_column_letter(len(produktgruppen) + 1)
        ws.cell(row=kopfzeile - 1, column=1,
                value="Selbstprüfung – größte Abweichung Formel ./. Berechnung:").font = Font(
                    bold=True, size=9)
        z = ws.cell(row=kopfzeile - 1, column=2,
                    value=f"=MAX(Kontrolle!$B${block + 1}:${letzte_sp}${r_block - 1})")
        z.number_format = "#,##0.00"
        z.font = Font(bold=True)
    return anzahl
