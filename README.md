# Spartenrechner – Perforator

Erzeugt die monatliche/quartalsweise Spartenrechnung (Kostenträgerrechnung je
Produktgruppe/Sparte) automatisch aus dem ERP-Rohdatenexport (Penta/APplus,
Datei `KPTM_Wertsummen_*.xls`).

## 🌐 Weboberfläche (empfohlen)

**[→ Spartenrechner öffnen](https://nikolausridder1301.github.io/spartenrechner/)**

Dateien per Drag & Drop hochladen (KPTM-Rohdaten, Kostenträger-Auswertung, SuSa, Werkstattbestand),
"Spartenrechnung erstellen" klicken – fertige Excel-Datei zum
Download. Genau wie ein klassischer Online-Dateikonverter, mit einem
entscheidenden Unterschied: **es wird nichts hochgeladen.** Die komplette
Verarbeitung läuft direkt im Browser (Python als WebAssembly via
[Pyodide](https://pyodide.org)) – die Firmenzahlen verlassen zu keinem
Zeitpunkt Ihren Rechner, es gibt keinen Server, der etwas sieht oder
speichert. Der Quellcode der Seite liegt unter `docs/`.

## Kommandozeile (alternativ)

1. Neuen Rohdaten-Export ablegen unter `rohdaten/<Jahr>/<Zeitraum>/`,
   z.B. `rohdaten/2026/Q1/` oder `rohdaten/2026/02/` für Februar.
   Mindestens benötigt: die Datei `KPTM_Wertsummen_*.xls` aus der SQL-Abfrage
   ans ERP-System (kumuliert seit Jahresbeginn bis zum gewünschten Stichtag).
   Für die drei Deckungsdifferenzen zusätzlich die **Kostenträger-Auswertung** der
   Buchhaltung (`Kostenträger <Zeitraum>.xls`, Penta, Objekttyp „Kostenträger",
   nach dem Monatsabschluss gezogen) – sie wird im Ordner automatisch gefunden.
2. Skript ausführen:

   ```bash
   python scripts/spartenrechnung.py --rohdaten rohdaten/2026/Q1 --output output/2026-Q1-Spartenrechnung.xlsx
   ```

   Der Zeitraum wird nicht angegeben, sondern aus den Spalten `Geschäftsjahr`
   und `Periode` des Exports gelesen (`--zeitraum` überschreibt ihn nur).

   Optional mit automatischem GuV-Kontroll-Check gegen die BWA (siehe unten):

   ```bash
   python scripts/spartenrechnung.py --rohdaten rohdaten/2026/Q1 --output output/2026-Q1-Spartenrechnung.xlsx --bwa "C:\Users\nikol\Downloads\BWA Perforator 7.2026.xlsx" --bwa-sheet "BWA 03.2026"
   ```

3. Ergebnis liegt in `output/` – eine schlanke Excel-Datei mit der
   Spartenrechnung (Sparten als Spalten, GuV-Zeilen als Zeilen) plus einem
   Hinweise-Sheet.

## Was ist automatisiert, was nicht?

Die vollständige Herleitung jeder einzelnen Zeile – inklusive der Belege und der
widerlegten Hypothesen – steht in `intern/HERLEITUNG.md` (privates Repository). Kurzfassung:

**Automatisch, zellgenau gegen die Referenzrechnung Q1 2026 geprüft** (grün):
Erlöse, Bestandsveränderung FE, **Bestandsveränderung UFE**, Material,
Fremdleistungen, sbA, FEK, MGK und overhead Kosten – und damit die komplette
obere GuV bis einschließlich **Rohmarge I**. In der alten Produktgruppen-Struktur
(wie die Referenz) stimmen von 266 vergleichbaren Zellen **239 auf den Cent**, 26
liegen im Rundungsbereich (Buchungsstand des KPTM-Exports), **eine** weicht ab –
dort widerspricht die Referenz ihrer eigenen Folgezeile. In der neuen Struktur
(siehe unten) sind Drill Pipes und Injektionstechnik weiter je Sparte vergleichbar,
die reorganisierten PB-Gruppen als Summe: 179 exakt, 20 nah, dieselbe eine
Abweichung. Zusätzlich erzeugt das Tool das Blatt
**Fertigungsstunden** – die bisher separat manuell gepflegte Auswertung der
Fertigungsaufträge je Kostenstelle (70 von 70 Zellen centgenau).

**Strukturell null** (grau): **Personalaufwand, sbE, Afa** und **aktivierte
Eigenleistung** gehören nicht auf die Sparten – die Personalkosten stecken bereits
in den Verrechnungssätzen (Stundensatz, MGK-Zuschlag, Overhead-Pool), ein
zusätzlicher Ausweis wäre Doppelzählung. In der Referenzrechnung stehen Werte
ausschließlich auf Sonder-Kostenobjekten („Abgrenzungen Walkenried",
„Weiterberechnung Valluhn"). Belege in `intern/HERLEITUNG.md`.

**Aus der Kostenträger-Auswertung** (grün): **FEK-/MGK-/VVGK-Deckungsdifferenz**.
Der KPTM-Export enthält nur die *verrechnete* Seite. Die Buchhaltung verteilt im
Monatsabschluss die Über-/Unterdeckung jeder Kostenstelle auf die Kostenträger
(„Umlage DDIFF", Kostenart `U` + Kostenstellennummer); die Kostenträger-Auswertung
zeigt das Ergebnis bereits je Produktgruppe. Zugeordnet wird nur, welche
Kostenstelle in welchen Topf gehört – Einkauf/Lager und QS zur MGK, die übrige
Fertigung zur FEK, Verwaltung und Vertrieb zur VVGK (`deckungsdifferenz_umlage` in
`config/kostenart_mapping.json`). Geprüft an Q1 2026: FEK und MGK 16 von 16 Sparten
centgenau, VVGK 13 von 16 – die drei übrigen widersprechen in der Referenz ihrer
eigenen Zeile „VVGK nach Deckungsdifferenz". Fehlt die Datei, bleiben die drei
Zeilen rosa.

**Strukturell null** (grau): **Sonderposten/Einmaleffekte** – in der Referenz in
jeder echten Sparte 0,00; die Beträge stehen nur auf Kostenobjekten ohne Sparte.

Die genaue Konten-Zuordnung steht in `config/kostenart_mapping.json` und
`config/produktgruppen.json` und kann dort angepasst werden – z.B. wenn durch die
laufende Produktgruppen-Reorganisation neue Codes hinzukommen. Ob die
Automatisierung nach einer Änderung noch stimmt, prüft (mit `--alte-struktur` ohne
Umschlüsselung, also je Sparte wie die Referenz):

```bash
python scripts/validate.py --rohdaten rohdaten/2026/Q1 --referenz "referenz/01-03 2026-Spartenrechnung Perforator.xlsx" --referenz-sheet "Spartenrechnung Jan-März 2026"
```

## Arbeitsweise beim monatlichen Abschluss

Hochzuladen sind drei Dateien plus – einmal im Jahr – der Anfangsbestand:

| Datei | Woher |
|---|---|
| `KPTM_Wertsummen_*.xls` | SQL-Abfrage ans ERP, kumuliert seit Jahresbeginn |
| Kostenträger-Auswertung | Buchhaltung, nach dem Monatsabschluss gezogen |
| SuSa (Summen- und Saldenliste), **optional** | Buchhaltung, gleicher Zeitraum – nur für die Abstimmung mit der FiBu |
| `PWBS_Werkstattbestand_*.xls` **mit Spalte Produktgruppe** | Buchhaltung, **am Stichtag** gezogen |

### Abstimmung mit der Buchhaltung – Prüfung ohne Referenzrechnung

Eine manuelle Referenzrechnung gibt es nur für Q1 2026. Damit auch jede spätere
Rechnung geprüft ist, hält das Tool sein Ergebnis gegen zwei Quellen, die
**unabhängig vom KPTM-Export** entstehen, und zeigt das Ergebnis im Browser sowie im
Blatt „Abstimmung" der Excel-Datei:

1. **Kostenträger-Auswertung:** Aus ihr wird jede Zeile (Erlöse, FE, Material,
   Fremdleistungen, sbA, FEK, MGK, VVGK) je Sparte ein zweites Mal gerechnet.
   Abweichungen unter 1 % gelten als Buchungsstand. Geeicht an Q1 2026: Sie zeigt
   dort genau die Abweichungen, die auch die Referenzrechnung zeigt.
2. **SuSa:** Erlöse gesamt gegen alle Erlöskonten der FiBu (auch solche, die das
   Mapping noch nicht kennt), Werkstattbestand zum Stichtag gegen die Bilanzkonten
   1040000 + 1041000, und der hinterlegte Anfangsbestand gegen die FiBu zum 01.01.

Material und Fremdleistungen sind gegen die FiBu bewusst nicht abgestimmt – die
Kostenrechnung bewertet anders, die Differenz ist strukturell. Die
Kostenstellen-Auswertung wird nicht gebraucht: Die Kostenstellen sind nach der
Umlage nicht auf null abgerechnet, eine Abstimmung daraus wäre Interpretation.

Zusätzlich lehnt das Tool Dateien aus dem falschen Zeitraum ab: Kostenträger-
Auswertung und SuSa müssen denselben Zeitraum abdecken wie der KPTM-Export, der
Werkstattbestand den passenden Stichtag tragen (Dateiname, z.B. `..._011026` für
Jan–Sep).

### Werkstattbestand: mit Produktgruppe, am Stichtag

Die Zeile **Bestandsveränderung UFE** braucht für jeden offenen Fertigungsauftrag
seine Sparte. Seit Oktober 2026 liefert die Buchhaltung den Werkstattbestand mit
der Spalte „Produktgruppe" – damit ist jeder Auftrag direkt zugeordnet, der frühere
Umweg über den PFAK-Export entfällt.

Wichtig ist der **Zeitpunkt**: Der Werkstattbestand muss am Stichtag gezogen werden
(für September also am 01.10.), nicht Wochen später. Ein nachträglich gezogener
Export zeigt die Aufträge mit ihrem *heutigen* Stand – geprüft an einem rückwirkend
gezogenen Bestand zum 01.01.2026, der deutlich vom Bilanzwert abweicht, während der
am Stichtag gezogene centgenau passt.

Ältere Exporte ohne Produktgruppen-Spalte funktionieren weiter über die
Kommandozeile mit `--pfak` (neuester PFAK-Export). Ohne beides blockiert das Tool
die UFE-Zeile, statt eine ungenaue Zahl auszugeben (Schwelle: 1 % nicht
zuordenbarer Bestand).

### Der Anfangsbestand: einmal im Jahr, aus dem ERP

Die UFE-Zeile ist `Bestand(Stichtag) − Bestand(01.01.)`. Den Stichtagswert rechnet
das Tool aus der PWBS-Datei; den Wert zum 01.01. bildet keiner der Exporte ab,
weil alle am 01.01. *beginnen*. Er ist als Jahreskonstante hinterlegt.

Von Hand pflegen muss man ihn trotzdem nicht. In der Weboberfläche gibt es dafür nur
das Feld **Betriebsparameter** (`betriebsparameter.xlsx`):

1. Beim **Jahresabschluss** (Stichtag 01.01. des Folgejahres) zeigt das Tool nach der
   Berechnung unter **„Werkstattbestand als Betriebsparameter sichern"** den
   errechneten Bestand als fertige **Excel-Tabelle** an. Einmal speichern – die
   übrigen Inhalte der Datei (andere Jahre, Zuschlagssätze, Umschlüsselung) bleiben
   dabei erhalten.
2. Diese Datei im Feld „Betriebsparameter" laden. Sie wird im Browser gespeichert und
   danach automatisch verwendet – das reicht für alle Monate des Jahres.

Die Kommandozeile kennt zusätzlich `--pwbs-anfang` (Werkstattbestand zum 01.01.
direkt angeben).

Geprüft: Der so ausgegebene Bestand ist als Anfangsbestand der Folgeperiode
centgenau verwendbar (`UFE(Jan–Jun) = UFE(Q1) + [Bestand(01.07.) − Bestand(01.04.)]`,
Abweichung 0,00 € über alle 14 Sparten).

## Produktgruppen-Reorganisation 2026

Seit Periode 4/2026 bucht das ERP auf neue Produktgruppen (Pressbohrtechnik:
PB0001 geteilt, PB0003 „Pilot Tubes" in PB0002 aufgegangen, neu PB0006–PB0008).
Ein kumulierter Export des Jahres mischt also beide Strukturen. Das Tool rechnet
**2026 durchgängig in der neuen Struktur**:

- **Buchungen** werden je Artikel umgeschlüsselt – nur, wenn sie auf der *alten*
  Gruppe ihres Artikels stehen. Grundlage ist die Liste „Alte vs Neue
  Produktgruppen", hinterlegt als Blatt `Umschluesselung` in den
  Betriebsparametern (einmalig erzeugt mit `scripts/umstellung_neue_struktur.py`).
  Die Zeilensummen bleiben dabei exakt erhalten. In der Ausgabe ist die
  Produktgruppe dort eine Formel, die gebuchte Gruppe steht daneben.
- **Näherung (gelb)** ist, was nur je Gruppe und nicht je Artikel vorliegt:
  Deckungsdifferenzen, die noch auf einer alten, geteilten Gruppe gebucht sind
  (verteilt nach ihrer Bezugsbasis), und der Anfangsbestand 2026 der geteilten
  Gruppen (verteilt nach dem Werkstattbestand zum 01.01. je Auftrag). Über die
  betroffenen Sparten zusammen sind die Werte exakt.
- Drill Pipes und Injektionstechnik sind nicht betroffen.
- Ab welchem Jahr umgeschlüsselt wird, steht in `config/produktgruppen.json`
  (`umstellung_ab_jahr`); frühere Jahre bleiben in ihrer gebuchten Struktur.

## BWA-Datei: hilft nur als Kontroll-Check, nicht zur Sparten-Automatisierung

Die Datei `BWA Perforator *.xlsx` (Downloads-Ordner) wurde geprüft. Sie ist eine
**firmenweite** Betriebswirtschaftliche Auswertung ganz ohne Sparten-Spalte
(Kategorien wie "Personalkosten", "sbA insgesamt" nur als Gesamtsumme). Damit
lässt sich Personalaufwand/sbA/Afa **nicht** automatisch auf DP/PB/IT/SR
verteilen.

Nützlich ist die BWA aber als **Kontroll-Anker**: Die Zeile "Vorläufiges
Ergebnis" (Spalte "kumuliert") entspricht exakt der Zeile "Ergebnis lt. GuV
vor Besserungsschein" der bisherigen manuellen Spartenrechnung (für Q1 2026
auf den Cent geprüft). Mit `--bwa`/`--bwa-sheet` zieht das
Skript diesen Wert automatisch und zeigt die Differenz zur bottom-up
berechneten DB III an. Diese Differenz muss weiterhin manuell erklärt werden
(Sonderposten, "keine Abrechnung BAB", "fehlende Abbildung SuSa in BAB und
Kostenträger") – das war auch im bisherigen Prozess ein manueller Schritt.

### Was in der Spalte „Summe" fehlt – strukturell, nicht nur beim aktuellen Stand

Die manuelle Referenzrechnung führt neben den echten Sparten fünf weitere Spalten
für Sachkonto-Buchungen ohne Produktgruppe: `712000` Photovoltaik Anlage,
`721000` Grundstück Valluhn, `910000` Abgrenzungen Walkenried, `920000`
Weiterberechnung Valluhn, `999999` periodenfremder Ertrag. Keiner dieser fünf
Codes taucht in KPTM als Produktgruppe auf – das Tool kann sie also aus
keinem Zeitraum ableiten. „Summe" ist deshalb die **Summe über die Sparten**,
nicht das Gesamtunternehmensergebnis, und der BWA-Vergleich oben enthält diese
Posten mit, nicht nur die Deckungsdifferenzen.

Konkrete Beträge je Periode stehen nicht hier (öffentliches Repository), sondern
in `intern/HERLEITUNG.md`.

## Projektstruktur

```
docs/                            Weboberfläche (GitHub Pages) – Kernlogik + UI
  spartenrechnung_core.py          gemeinsame Berechnungslogik (Web + CLI)
  index.html / app.js / style.css  Oberfläche (Pyodide, läuft im Browser)
  config/                          Kopie der Mapping-Tabellen für den Browser
scripts/spartenrechnung.py       CLI-Wrapper um dieselbe Kernlogik
scripts/validate.py              zellgenauer Abgleich gegen eine manuelle Referenz
scripts/umstellung_neue_struktur.py  einmalig: Betriebsparameter auf neue Produktgruppen
config/                          Mapping-Tabellen (Kostenart→Zeile, Produktgruppen)

intern/                          ← eigenes, PRIVATES Repository (nicht hier drin)
  HERLEITUNG.md                    Herleitung jeder Zeile inkl. Belegen
  betriebsparameter.xlsx           Anfangsbestände, Zuschlagssätze, Umschlüsselung
rohdaten/ referenz/ output/      lokal, nicht versioniert
```

## Wo liegt was – und warum getrennt

Das Projekt besteht aus **zwei** Repositories, die im selben Ordner liegen:

| | |
|---|---|
| `spartenrechner` (dieses, **öffentlich**) | Code, Mapping-Tabellen, Weboberfläche. GitHub Pages veröffentlicht daraus `docs/`. |
| `spartenrechner-intern` (**privat**, als Unterordner `intern/`) | Herleitungsdokument und Betriebsparameter – alles mit echten Zahlen. |

Der Grund für die Trennung: Aus Stundensatz und Zuschlagssätzen liesse sich die
Preisbildung zurückrechnen. Diese Werte gehören deshalb nicht in ein öffentliches
Repository – wohl aber in die Versionierung, denn sie ändern sich und man will
nachvollziehen können, wann und warum.

**Zwei Regeln, die man kennen muss:**

1. **Keine Firmenzahlen in dieses Repository** – nicht in Code, nicht in
   Kommentaren, nicht in Konfigurationsdateien. `.gitignore` sperrt `intern/`,
   Tabellen jeder Art und jede `betriebsparameter.json`.
2. **`docs/` ist doppelt öffentlich.** Jede Datei dort ist zusätzlich unter
   `https://nikolausridder1301.github.io/spartenrechner/<pfad>` abrufbar, auch
   eine, die keine Seite verlinkt. Dort also erst recht nichts ablegen.

Die Weboberfläche lädt die Betriebsparameter **nie** aus einem Repository: Sie
werden einmalig im Browser hinterlegt und dort gespeichert (`localStorage`), auf
dem Rechner des Nutzers. Die Kommandozeile liest sie aus `intern/`.

Wer nur die Weboberfläche benutzt, braucht das private Repository nicht – nur die
ausgefüllte `betriebsparameter.xlsx`.

## GitHub Pages aktivieren / aktualisieren

Die Seite wird direkt aus `docs/` auf dem `main`-Branch veröffentlicht
(Repo-Einstellungen → Pages → Source: `main` / `docs`). Nach jedem Push auf
`main` aktualisiert sich `https://nikolausridder1301.github.io/spartenrechner/`
automatisch (kann 1–2 Minuten dauern).
