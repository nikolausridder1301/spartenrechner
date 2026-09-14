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
    Zuordnung              Fertigungsauftrag -> Sparte UND Artikelnummer -> Sparte
                           (Ersatzschluessel), gemeinsam. Keine Kopie einer
                           hochgeladenen Datei, sondern eine aus KPTM + PFAK + evtl.
                           Handzuordnung zusammengefuehrte Tabelle - deshalb ein
                           eigenes Blatt statt einer Spalte in den Rohdaten
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

    # Die Nebentabellen beginnen in derselben Zeile wie die Haupttabelle (Kopf in 3,
    # Daten ab 4) und tragen dieselbe Kopfformatierung - sonst steht rechts daneben
    # eine zweite Tabelle, die um zwei Zeilen versetzt und anders formatiert ist.
    KOPF = 3
    def kopfzelle(spalte, text):
        z = ws.cell(row=KOPF, column=spalte, value=text)
        z.font, z.fill = FONT_WEISS, FILL_DATEN
        return z

    # Die gueltigen Sparten als sichtbarer Bereich. Die Fracht-Umbuchung in
    # 'Daten_KPTM' schlaegt hier nach, ob es das Frachtkostenobjekt ueberhaupt gibt.
    # (Als Array-Konstante in der Formel verweigert Excel die Datei.)
    kopfzelle(7, "Gültige Sparten")
    for i, pg in enumerate(produktgruppen, start=KOPF + 1):
        ws.cell(row=i, column=7, value=pg)
    ws.column_dimensions["G"].width = 16

    # Auffangregel nach Kontenpraefix. Sie muss hier stehen, damit die Hilfsspalten
    # in 'Daten_KPTM' dasselbe rechnen wie das Programm - sonst zeigt die Mappe
    # andere Zahlen als das Tool, und zwar unbemerkt.
    regel = mapping.get("kostenart_praefix_regel", {})
    kopfzelle(9, "Auffangregel: Konto beginnt mit")
    kopfzelle(10, "Zeile")
    kopfzelle(11, "Vorzeichen")
    for i, (praefix, r) in enumerate(sorted(regel.items()), start=KOPF + 1):
        ws.cell(row=i, column=9, value=praefix)
        ws.cell(row=i, column=10, value=r["zeile"])
        ws.cell(row=i, column=11, value=r["sign"])
    ws.cell(row=KOPF + len(regel) + 2, column=9, value=(
        "Greift nur, wenn links kein Einzeleintrag passt. Gilt fuer Konten, die in der "
        "Einzelliste fehlen, weil sie im Kalibrierungszeitraum nicht vorkamen."
    )).font = Font(italic=True, size=9)
    for sp, br in (("I", 30), ("J", 14), ("K", 11)):
        ws.column_dimensions[sp].width = br
    n = max(len(regel), 1)
    # Begrenzte Bereiche statt ganzer Spalten: bei 60.000 Formeln, die hier
    # nachschlagen, ist der Unterschied beim Neuberechnen deutlich spuerbar.
    letzte_mapping = 3 + len(zeilen)
    return ws, (f"Mapping!$G${KOPF + 1}:$G${KOPF + len(produktgruppen)}",
                f"Mapping!$I${KOPF + 1}:$K${KOPF + n}",
                f"Mapping!$A$4:$E${letzte_mapping}")


def _kptm_blatt(wb, df, mapping, bereiche):
    """Der KPTM-Export mit vier Hilfsspalten, die selbst Formeln sind.

    Dadurch ist die Zuordnung nicht das Ergebnis eines unsichtbaren Programmschritts,
    sondern in jeder Zeile nachvollziehbar und im Zweifel korrigierbar.
    """
    # Spalten ohne Informationsgehalt weglassen. Ein KPTM-Export fuehrt rund 36
    # Spalten, von denen etwa 20 entweder komplett leer sind oder in allen 15.000
    # Zeilen denselben Wert tragen (z.B. 'Datenbestand = Istdaten'). Sie machen das
    # Blatt gut doppelt so gross, ohne irgendetwas auszusagen - der konstante Wert
    # steht einmal als Notiz ueber der Tabelle, was lesbarer ist als 15.000
    # Wiederholungen. Spalten mit echtem Inhalt bleiben alle erhalten, auch die, die
    # das Werkzeug selbst nicht braucht.
    leer, konstant = [], {}
    for spalte in df.columns:
        werte = df[spalte].dropna().unique()
        if len(werte) == 0:
            leer.append(str(spalte))
        elif len(werte) == 1:
            konstant[str(spalte)] = werte[0]
    df = df.drop(columns=[c for c in df.columns if str(c) in leer or str(c) in konstant])

    hinweis = ("Der hochgeladene KPTM-Export. Die vier gelben Spalten rechts sind Formeln: "
               "sie schlagen im Blatt 'Mapping' nach, auf welche GuV-Zeile eine Kostenart "
               "geht. Das Blatt 'Spartenrechnung' summiert dann ueber diese Spalten.")
    if leer or konstant:
        teile = []
        if konstant:
            teile.append("in allen Zeilen gleich (" +
                         "; ".join(f"{k} = {v}" for k, v in list(konstant.items())[:6]) +
                         (" ..." if len(konstant) > 6 else "") + ")")
        if leer:
            teile.append(f"durchgehend leer ({', '.join(leer[:6])}"
                         + (" ..." if len(leer) > 6 else "") + ")")
        hinweis += (f"  |  Nicht abgebildet sind {len(leer) + len(konstant)} Spalten ohne "
                    f"Informationsgehalt: " + " sowie ".join(teile) + ".")

    ws, erste = _blatt_mit_daten(wb, "Daten_KPTM", df, hinweis)
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

    pg_bereich, regel_bereich, map_bereich = bereiche
    # Ohne TEXT(...,"@"): die Kostenart steht in beiden Blaettern bereits als Text,
    # die Umwandlung war wirkungslos und hat rund eine Million Zeichen Formeltext
    # gekostet.
    for r in range(erste, letzte + 1):
        # Reihenfolge wie im Programm: Einzeleintrag, dann Kostenstelle, dann Auffangregel.
        ws.cell(row=r, column=basis + 1, value=(
            f'=IFERROR(VLOOKUP({s_ka}{r},{map_bereich},2,FALSE),'
            f'IF(AND(LEFT({s_ka}{r},{len(prefix)})="{prefix}",'
            f'{s_wa}{r}="{wertart}"),"fek",'
            f'IFERROR(VLOOKUP(LEFT({s_ka}{r},1),{regel_bereich},2,FALSE),"")))'))
        ws.cell(row=r, column=basis + 2, value=(
            f'=IFERROR(VLOOKUP({s_ka}{r},{map_bereich},4,FALSE),'
            f'IFERROR(VLOOKUP(LEFT({s_ka}{r},1),{regel_bereich},3,FALSE),1))'))
        ws.cell(row=r, column=basis + 3, value=(
            f'=IF(IFERROR(VLOOKUP({s_ka}{r},{map_bereich},5,FALSE),"")="fracht",'
            f'IF(COUNTIF({pg_bereich},LEFT({s_pg}{r},2)&"{fracht}")>0,'
            f'LEFT({s_pg}{r},2)&"{fracht}",{s_pg}{r}),{s_pg}{r})'))
        z = ws.cell(row=r, column=basis + 4, value=f'={c_vz}{r}*{s_wert}{r}')
        z.number_format = "#,##0.00"

    ws.auto_filter.ref = f"A{kopf}:{c_betrag}{letzte}"
    for c in (c_zeile, c_vz, c_ziel, c_betrag):
        ws.column_dimensions[c].width = 19
    return {"blatt": "Daten_KPTM", "erste": erste, "letzte": letzte,
            "zeile": c_zeile, "ziel": c_ziel, "betrag": c_betrag,
            # Rohspalten - das Blatt 'Fertigungsstunden' summiert direkt darueber,
            # nicht ueber die Hilfsspalten: dort geht es um ISWF/ISMF der
            # Kostenstellen, nicht um die Zuordnung auf GuV-Zeilen.
            "kostenart": s_ka, "wertart": s_wa, "wert": s_wert,
            "auftragsart": spalten.get("Auftragsart")}


def _zuordnung_blatt(wb, nach_auftrag, nach_artikel, manuelle_zuordnung=None):
    """Ein gemeinsames Blatt fuer beide Zuordnungstabellen.

    Beide sind keine Kopie einer hochgeladenen Datei, sondern aus KPTM + PFAK (und
    evtl. Handzuordnung) zusammengefuehrte Nachschlagetabellen - deshalb ein eigenes
    Blatt statt einer Spalte in den Rohdaten. Zwei Tabellen statt zwei Blaetter, weil
    sie demselben Zweck dienen (Sparte zu einem Schluessel finden) und nur die
    Zuordnungslogik es rechtfertigt, sie ueberhaupt auszulagern - nicht ihre Groesse.
    """
    ws = wb.create_sheet("Zuordnung")
    ws.cell(row=1, column=1, value=(
        "Wie Fertigungsauftraege und Artikelnummern einer Sparte zugeordnet werden. "
        "Keine Kopie einer hochgeladenen Datei: beide Tabellen fuehren KPTM- und "
        "PFAK-Angaben zusammen. Die Blaetter 'Daten_PWBS_*' schlagen hier nach."
    )).font = Font(italic=True, size=9)

    ws.cell(row=2, column=1, value="Fertigungsauftrag → Sparte").font = Font(bold=True, size=9)
    for j, titel in enumerate(["Fertigungsauftrag", "Sparte", "Quelle"], start=1):
        z = ws.cell(row=3, column=j, value=titel)
        z.font, z.fill = FONT_WEISS, FILL_DATEN
    eintraege = dict(nach_auftrag or {})
    quelle = {a: "Auftragsnummer (KPTM/PFAK)" for a in eintraege}
    for a, pg in (manuelle_zuordnung or {}).items():
        eintraege[a] = pg          # Handzuordnung hat Vorrang
        quelle[a] = "von Hand zugeordnet"
    for i, (a, pg) in enumerate(sorted(eintraege.items()), start=4):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=pg)
        ws.cell(row=i, column=3, value=quelle[a])

    ws.cell(row=2, column=5, value="Artikelnummer → Sparte (Ersatzschlüssel)").font = Font(
        bold=True, size=9)
    for j, titel in ((5, "Artikelnummer"), (6, "Sparte")):
        z = ws.cell(row=3, column=j, value=titel)
        z.font, z.fill = FONT_WEISS, FILL_DATEN
    for i, (a, pg) in enumerate(sorted((nach_artikel or {}).items()), start=4):
        ws.cell(row=i, column=5, value=a)
        ws.cell(row=i, column=6, value=pg)

    ws.freeze_panes = "A4"
    for sp, br in (("A", 20), ("B", 12), ("C", 26), ("D", 3), ("E", 22), ("F", 12)):
        ws.column_dimensions[sp].width = br


def _pwbs_blatt(wb, pfad, titel, hinweis):
    """Werkstattbestand, bei dem die Sparte per INDEX/VERGLEICH ermittelt wird.

    Die Zuordnung ist damit nicht das Ergebnis eines unsichtbaren Programmschritts,
    sondern in der Zelle selbst nachlesbar. Sie bildet die Logik des Programms genau
    ab, in dieser Reihenfolge:
      1. keine Artikelnummer -> Serviceauftrag, zaehlt gar nicht mit (Monteureinsatz,
         Inbetriebnahme, Schulung: es wird nichts gefertigt)
      2. Treffer ueber die Auftragsnummer im Blatt 'Zuordnung' (Tabelle links)
      3. sonst Treffer ueber die Artikelnummer im Blatt 'Zuordnung' (Tabelle rechts)
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
            f'IFERROR(INDEX(Zuordnung!$B:$B,'
            f'MATCH(TEXT({s_rm}{r},"@"),Zuordnung!$A:$A,0)),'
            f'IFERROR(INDEX(Zuordnung!$F:$F,'
            f'MATCH(TEXT({s_art}{r},"@"),Zuordnung!$E:$E,0)),'
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

    # Muss VOR den PWBS-Blaettern stehen, weil deren Sparten-Formel hier hineinschlaegt.
    if nach_auftrag is not None or nach_artikel is not None:
        _zuordnung_blatt(wb, nach_auftrag, nach_artikel, manuelle_zuordnung)

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

    # Das Stundenblatt rechnet aus denselben Rohdaten - also gehoert es genauso
    # verformelt, sonst steht neben einer nachvollziehbaren Tabelle eine, die man
    # glauben muss.
    anzahl_stunden = verformele_fertigungsstunden(wb, mapping, kptm)

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
            f"Im Blatt 'Fertigungsstunden' sind es {anzahl_stunden} Zellen - dort summieren "
            "SUMMEWENNS ueber die Kostenstellen (Kostenart K2*) in 'Daten_KPTM', getrennt nach "
            "Wertart ISWF (Euro) und ISMF (Stunden) sowie nach Auftragsart.",
            "",
            "  Zeilen MIT Formel:  Erloese, Bestandsveraenderung FE, Material, Fremdleistungen,",
            "                      sbA, FEK, MGK, overhead Kosten - sie summieren ueber "
            "'Daten_KPTM'.",
            "                      Dazu alle Zwischensummen (Betriebsleistung, Rohmarge I, DB I-III).",
            "",
            "  Bestandsveraenderung UFE hat ebenfalls eine Formel:",
            "    = Bestand(Stichtag) aus 'Daten_PWBS_Ende' minus Anfangsbestand.",
            "    Die Sparte je Auftrag ermittelt dort eine INDEX/VERGLEICH-Formel - erst ueber die",
            "    Auftragsnummer, ersatzweise ueber die Artikelnummer (beide Tabellen im Blatt",
            "    'Zuordnung'). Auftraege ohne Artikelnummer sind Serviceauftraege und zaehlen",
            "    nicht mit.",
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


def verformele_fertigungsstunden(wb, mapping, kptm):
    """Ersetzt die Zahlen im Blatt 'Fertigungsstunden' durch Formeln auf 'Daten_KPTM'.

    Dasselbe Prinzip wie bei der Spartenrechnung: Das Blatt rechnet aus denselben
    Rohdaten, also soll man auch hier jede Zahl anklicken und ihre Herkunft sehen.
    Die Spaltenlogik entspricht fertigungsstunden() im Kernmodul:
      FEK inkl. Gemeinkosten          Kostenart = Kostenstelle, Wertart = ISWF
      davon unproduktive Gemeinkosten  zusaetzlich Auftragsart 5 (Sammelauftraege)
      produktive FEK                   Differenz der beiden
      davon Lagerauftraege             Auftragsarten LAG/DPL/PBL/INN/ITL/SRL
      davon Kundenauftraege            produktive FEK minus Lagerauftraege
      Stunden gesamt                   dieselbe Kostenstelle, aber Wertart ISMF

    Mehrere Auftragsarten werden als Summe einzelner SUMMEWENNS geschrieben, nicht
    als Array-Konstante - die laesst Excel in einer von openpyxl erzeugten Datei
    nicht zu (die Datei gilt dann als beschaedigt).
    """
    if "Fertigungsstunden" not in wb.sheetnames:
        return 0
    ws = wb["Fertigungsstunden"]
    b, v, bis = kptm["blatt"], kptm["erste"], kptm["letzte"]
    r_ka, r_wa, r_wert = kptm["kostenart"], kptm["wertart"], kptm["wert"]
    r_aa = kptm.get("auftragsart")
    if not r_aa:
        return 0

    def bereich(sp):
        return f"'{b}'!${sp}${v}:${sp}${bis}"

    wertart = mapping.get("kostenstellen_wertart", "ISWF")
    gemein = list(mapping.get("auftragsart_gemeinkosten", ["5"]))
    lager = list(mapping.get("auftragsart_lager", []))

    def summe_ueber_arten(zeile, arten):
        """Summe je Auftragsart, einzeln addiert statt als Array-Konstante."""
        return "+".join(
            f'SUMIFS({bereich(r_wert)},{bereich(r_ka)},$A{zeile},'
            f'{bereich(r_wa)},"{wertart}",{bereich(r_aa)},"{a}")'
            for a in arten)

    anzahl = 0
    zeile = 5
    while ws.cell(row=zeile, column=1).value and str(ws.cell(row=zeile, column=1).value) != "SUMME":
        formeln = {
            3: f'=SUMIFS({bereich(r_wert)},{bereich(r_ka)},$A{zeile},{bereich(r_wa)},"{wertart}")',
            4: "=" + summe_ueber_arten(zeile, gemein),
            5: f"=C{zeile}-D{zeile}",
            6: "=" + summe_ueber_arten(zeile, lager) if lager else None,
            7: f"=E{zeile}-F{zeile}",
            8: f'=SUMIFS({bereich(r_wert)},{bereich(r_ka)},$A{zeile},{bereich(r_wa)},"ISMF")',
        }
        for spalte, formel in formeln.items():
            if formel is None:
                continue
            z = ws.cell(row=zeile, column=spalte, value=formel)
            z.number_format = "#,##0.00"
            anzahl += 1
        zeile += 1

    # Summenzeile ebenfalls als Formel
    if str(ws.cell(row=zeile, column=1).value) == "SUMME":
        for spalte in range(3, 9):
            sp = get_column_letter(spalte)
            z = ws.cell(row=zeile, column=spalte, value=f"=SUM({sp}5:{sp}{zeile - 1})")
            z.font = Font(bold=True)
            z.number_format = "#,##0.00"
            anzahl += 1
    return anzahl
