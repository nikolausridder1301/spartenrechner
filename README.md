# Spartenrechner – Perforator

Erzeugt die monatliche/quartalsweise Spartenrechnung (Kostenträgerrechnung je
Produktgruppe/Sparte) automatisch aus dem ERP-Rohdatenexport (Penta/APplus,
Datei `KPTM_Wertsummen_*.xls`).

## 🌐 Weboberfläche (empfohlen)

**[→ Spartenrechner öffnen](https://nikolausridder1301.github.io/spartenrechner/)**

Dateien per Drag & Drop hochladen (KPTM-Rohdaten, optional BWA),
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
obere GuV bis einschließlich **Rohmarge I**. Von 266 vergleichbaren Zellen
stimmen 145 auf den Cent und 15 liegen im Rundungsbereich; alle übrigen hängen
an den Deckungsdifferenzen. Zusätzlich erzeugt das Tool das Blatt
**Fertigungsstunden** – die bisher separat manuell gepflegte Auswertung der
Fertigungsaufträge je Kostenstelle (70 von 70 Zellen centgenau).

**Strukturell null** (grau): **Personalaufwand, sbE, Afa** und **aktivierte
Eigenleistung** gehören nicht auf die Sparten – die Personalkosten stecken bereits
in den Verrechnungssätzen (Stundensatz, MGK-Zuschlag, Overhead-Pool), ein
zusätzlicher Ausweis wäre Doppelzählung. In der Referenzrechnung stehen Werte
ausschließlich auf Sonder-Kostenobjekten („Abgrenzungen Walkenried",
„Weiterberechnung Valluhn"). Belege in `intern/HERLEITUNG.md`.

**Manuell zu ergänzen** (rosa) – nur noch zwei Positionen:
- **FEK-/MGK-/VVGK-Deckungsdifferenz**: Das ERP liefert nur die *verrechnete*
  Seite; die *Ist*-Seite kommt aus dem BAB (Betriebsabrechnungsbogen). Keiner der
  sechs Exporte ist eine Kostenstellenrechnung – die Kostenstellen Lager und
  Einkauf kommen darin gar nicht vor. Über 100 Herleitungshypothesen wurden
  numerisch ausgeschlossen (siehe `intern/HERLEITUNG.md`).
- **Sonderposten/Einmaleffekte**: bewusste Einzelfall-Entscheidungen je Periode.

Die genaue Konten-Zuordnung steht in `config/kostenart_mapping.json` und
`config/produktgruppen.json` und kann dort angepasst werden – z.B. wenn durch die
laufende Produktgruppen-Reorganisation neue Codes hinzukommen. Ob die
Automatisierung nach einer Änderung noch stimmt, prüft:

```bash
python scripts/validate.py --rohdaten rohdaten/2026/Q1 --referenz "referenz/01-03 2026-Spartenrechnung Perforator.xlsx" --referenz-sheet "Spartenrechnung Jan-März 2026"
```

## Arbeitsweise beim monatlichen Abschluss

Eine Besonderheit, die man einmal verstanden haben muss, betrifft die Zeile
**Bestandsveränderung UFE**. Sie beruht darauf, dass jeder offene
Fertigungsauftrag einer Sparte zugeordnet werden kann – und genau das ist beim
*laufenden* Abschluss nur teilweise möglich.

Der Grund: Ein PFAK-Export besteht praktisch vollständig aus **fertiggemeldeten**
Aufträgen (Jan–März 2026: 1.376 von 1.542 „Erledigt", 13 „Teilfertig"). Ein
Auftrag, der am Stichtag noch läuft, ist per Definition nicht erledigt – genau
deshalb steht er ja im Werkstattbestand. Von den 508 laufenden Aufträgen am
01.04.2026 findet der März-Export nur **14**, der Juli-Export dagegen **476**.

| Datenlage | Lücke | Sparten korrekt |
|---|---|---|
| laufender Abschluss (PFAK Jan–März) | 4,7 % – 239 Aufträge, 58.156 € | 7 von 13, bis 27.957 € daneben |
| nachträglich (PFAK Jan–Juli) | 0,1 % – 10 Aufträge, 1.445 € | **13 von 13 centgenau** |

Das Tool blockiert die Zeile oberhalb von **1 %** bewusst, statt eine ungenaue
Zahl auszugeben. Der spätere Export ändert dabei nichts am *Wert* des Bestands –
der kommt unverändert aus der PWBS. Er sagt nur, *wozu* die Aufträge gehören.

**Empfohlenes Vorgehen:**

1. **PFAK: immer die neueste Datei hochladen**, die vorliegt – eine genügt. Je
   aktueller der Export, desto mehr Aufträge sind darin fertiggemeldet und damit
   zuordenbar. Mit dem Juli-Export ist die UFE-Zeile für Q1 in allen 13 Sparten
   centgenau.
2. **Bleibt eine Lücke:** Die Oberfläche zeigt nach der Berechnung, wie viel des
   Werkstattbestands zugeordnet werden konnte, und listet die offenen Aufträge
   **nach Wert sortiert** mit je einem Sparten-Dropdown. Für Q1 reicht eine Runde
   mit den 60 größten, um von 4,71 % auf 0,76 % zu kommen – die restlichen 179
   sind Kleinvieh. Die Zuordnungen bleiben über mehrere Runden erhalten.
3. **Oder abwarten:** denselben Zeitraum ein bis zwei Monate später erneut
   rechnen. So ist die Referenzrechnung Q1 2026 entstanden.

Dauerhaft löst das eine zusätzliche Spalte im PWBS-Export (siehe unten) – dann
entfällt die Handarbeit vollständig.

### Der Anfangsbestand: einmal im Jahr, aus dem ERP

Die UFE-Zeile ist `Bestand(Stichtag) − Bestand(01.01.)`. Den Stichtagswert rechnet
das Tool aus der PWBS-Datei; den Wert zum 01.01. bildet keiner der Exporte ab,
weil alle am 01.01. *beginnen*. Er ist als Jahreskonstante hinterlegt.

Von Hand pflegen muss man ihn trotzdem nicht:

1. Beim **Abschluss des Januars** brauchen Sie ohnehin den Bestand zum 01.01. –
   laden Sie den PWBS-Export mit diesem Stichtag im Feld „Periodenbeginn" hoch.
   Es ist derselbe Report wie der monatliche, nur mit anderem Datum.
2. Das Tool zeigt danach unter **„Werkstattbestand als Betriebsparameter sichern"**
   den errechneten Bestand als fertige **Excel-Tabelle** an (Spalten: Sparte,
   Geschäftsjahr, Anfangsbestand). Einmal speichern – damit sind die restlichen elf
   Monate versorgt, ohne die Datei jedes Mal mitzuladen.

Geprüft: Der so ausgegebene Bestand ist als Anfangsbestand der Folgeperiode
centgenau verwendbar (`UFE(Jan–Jun) = UFE(Q1) + [Bestand(01.07.) − Bestand(01.04.)]`,
Abweichung 0,00 € über alle 14 Sparten).

## Was beim Controlling angefragt werden müsste

Drei Punkte würden die verbleibenden Lücken schließen:

1. **BAB Q1/2026** mit Ist-Kosten je Fertigungskostenstelle und dem verwendeten
   Umlageschlüssel → automatisiert alle drei Deckungsdifferenzen (die einzige
   inhaltlich noch offene Position).
2. **`KPTM_Wertsummen` ohne den Filter `Datenbestand = Istdaten`.** Alle 15.045
   Zeilen des heutigen Exports tragen diesen Wert; enthält dieselbe Tabelle auch
   Soll-/Plandaten, wären die Deckungsdifferenzen direkt ableitbar.
3. **`PWBS_Werkstattbestand` mit einer Produktgruppen-Spalte.** Der kleinste
   Aufwand der drei und sofort spürbar: eine zusätzliche Spalte in einer
   bestehenden Abfrage, kein neuer Report. Die Produktgruppe hängt im ERP am
   Fertigungsauftrag, sie kommt im Export nur nicht mit. Damit entfällt die
   Hilfskonstruktion aus PFAK, KPTM-Kostenobjekt und Handzuordnung – beim
   Q1-Abschluss 239 Aufträge, die sonst nachgetragen werden müssen.

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
config/                          Mapping-Tabellen (Kostenart→Zeile, Produktgruppen)

intern/                          ← eigenes, PRIVATES Repository (nicht hier drin)
  HERLEITUNG.md                    Herleitung jeder Zeile inkl. Belegen
  betriebsparameter.xlsx           Anfangsbestände je Sparte (echte Werte)
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
ausgefüllte `betriebsparameter.json`.

## GitHub Pages aktivieren / aktualisieren

Die Seite wird direkt aus `docs/` auf dem `main`-Branch veröffentlicht
(Repo-Einstellungen → Pages → Source: `main` / `docs`). Nach jedem Push auf
`main` aktualisiert sich `https://nikolausridder1301.github.io/spartenrechner/`
automatisch (kann 1–2 Minuten dauern).
