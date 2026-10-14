/* Spartenrechner – laeuft komplett im Browser (Pyodide/WebAssembly).
   Es wird nichts hochgeladen: Dateien werden ausschliesslich lokal im
   virtuellen Dateisystem von Pyodide verarbeitet. */

const loadingStatus = document.getElementById("loading-status");
const form = document.getElementById("form");
const submitBtn = document.getElementById("submit-btn");
const resultBox = document.getElementById("result");
const errorBox = document.getElementById("error");

let pyodide = null;
let kptmFile = null;
let kostentraegerFile = null;
let susaFile = null;
let pwbsEndeFile = null;

function setupDropzone(zoneId, textId, inputId, onFile) {
  const zone = document.getElementById(zoneId);
  const text = document.getElementById(textId);
  const input = document.getElementById(inputId);
  // Fehlt ein Feld (Seite und Skript aus verschiedenen Versionen), dieses eine Feld
  // auslassen statt das ganze Skript abzubrechen - sonst startet Python nie.
  if (!zone || !text || !input) {
    console.warn("Feld fehlt auf der Seite:", zoneId);
    return;
  }

  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    if (input.files.length) onFile(input.files[0], zone, text);
  });
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0], zone, text);
  });
}

setupDropzone("dropzone-kptm", "dropzone-kptm-text", "file-kptm", (file, zone, text) => {
  kptmFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

setupDropzone("dropzone-kt", "dropzone-kt-text", "file-kt", (file, zone, text) => {
  kostentraegerFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

setupDropzone("dropzone-susa", "dropzone-susa-text", "file-susa", (file, zone, text) => {
  susaFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

setupDropzone("dropzone-pwbs-e", "dropzone-pwbs-e-text", "file-pwbs-e", (file, zone, text) => {
  pwbsEndeFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

/* Kurze Erklaerung je Feld ("?"-Symbol neben der Ueberschrift): woher die Datei
   kommt, was sie enthaelt und wozu sie gebraucht wird - fuer alle, die die
   ERP-Exportnamen nicht auswendig kennen. Ein einziges, wiederverwendetes
   Tooltip-Element statt eines pro Icon, damit die Positionierung an einer
   Stelle berechnet wird und auf jedem Bildschirm im sichtbaren Bereich bleibt. */
const INFO_TEXTE = {
  "kptm": `<strong>Woher</strong>Export aus dem ERP-System (Penta/APplus) &ndash; die SQL-Abfrage
    &bdquo;KPTM_Wertsummen&ldquo;.
    <strong>Was</strong>Alle Buchungen (Erl&ouml;se, Material, L&ouml;hne, &hellip;) je Kostenart und
    Fertigungsauftrag, kumuliert seit Jahresbeginn bis zum gew&uuml;nschten Stichtag.
    <strong>Warum</strong>Grundlage der gesamten Rechnung &ndash; daraus werden Erl&ouml;se, Kosten
    und Deckungsbeitr&auml;ge je Sparte berechnet.`,
  "kostentraeger": `<strong>Woher</strong>Von der Buchhaltung: Auswertung aus dem ERP (Penta)
    nach <em>Kostentr&auml;gern</em>, gezogen nach dem Monatsabschluss. Sie kommt meist zusammen
    mit der Auswertung nach Kostenstellen &ndash; gebraucht wird nur die nach Kostentr&auml;gern.
    <strong>Was</strong>Die Kosten je Produktgruppe, einschlie&szlig;lich der Umlage der
    Deckungsdifferenzen (&bdquo;Umlage DDIFF&ldquo;) aus jeder Kostenstelle.
    <strong>Warum</strong>Daraus kommen die FEK-, MGK- und VVGK-Deckungsdifferenz je Sparte &ndash;
    der Unterschied zwischen kalkulierten und tats&auml;chlichen Gemeinkosten. Ohne sie fehlen
    DB&nbsp;I bis DB&nbsp;III.`,
  "susa": `<strong>Woher</strong>Von der Buchhaltung: die Summen- und Saldenliste der FiBu
    (z.&nbsp;B. &bdquo;SuSa 01-09.2026.csv&ldquo;) f&uuml;r denselben Zeitraum &ndash; sie kommt
    zusammen mit der Kostentr&auml;ger-Auswertung.
    <strong>Was</strong>Anfangs- und Endsaldo jedes Sachkontos.
    <strong>Warum</strong>Nur zur Abstimmung: Das Tool pr&uuml;ft damit, ob die Erl&ouml;se mit der FiBu
    &uuml;bereinstimmen und ob der Werkstattbestand dem Bilanzwert entspricht (richtiger
    Stichtag, vollst&auml;ndige Datei). In die Rechnung selbst flie&szlig;t sie nicht ein &ndash;
    ohne SuSa entsteht dieselbe Spartenrechnung, nur ohne diese Gegenprobe.`,
  "abstimmung": `<strong>Was ist es</strong>Eine Gegenprobe ohne Referenzrechnung: Das Ergebnis wird
    gegen zwei Quellen der Buchhaltung gehalten, die unabh&auml;ngig vom KPTM-Export entstehen.
    <strong>Berechnung</strong>(1) Aus der Kostentr&auml;ger-Auswertung wird jede Zeile je Sparte ein
    zweites Mal gerechnet &ndash; Abweichungen unter 1&nbsp;% sind Buchungsstand. (2) Gegen die SuSa:
    Erl&ouml;se gesamt, Werkstattbestand zum Stichtag (Konten 1040000 + 1041000) und der hinterlegte
    Anfangsbestand. Material und Fremdleistungen sind gegen die FiBu nicht abstimmbar, weil die
    Kostenrechnung anders bewertet.`,
  "pwbs-ende": `<strong>Woher</strong>Von der Buchhaltung: ERP-Export &bdquo;PWBS_Werkstattbestand&ldquo;
    in der Form <em>mit Spalte Produktgruppe</em>, gezogen <em>am</em> Stichtag (erster Tag nach
    dem Auswertungszeitraum) &ndash; nicht nachtr&auml;glich, sonst stimmen die Werte nicht mehr.
    <strong>Was</strong>Alle Fertigungsauftr&auml;ge, die an diesem Stichtag noch nicht fertiggemeldet
    sind, mit ihrem bisher aufgelaufenen Wert (unfertige Erzeugnisse) und ihrer Sparte.
    <strong>Warum</strong>Ohne diese Datei fehlt die Zeile &bdquo;Bestandsver&auml;nderung UFE&ldquo; &ndash;
    Betriebsleistung und Rohmarge w&auml;ren unvollst&auml;ndig.`,
  "betriebsparameter": `<strong>Woher</strong>Die Datei betriebsparameter.xlsx &ndash; einmal von
    Ihnen gesichert (Button am Ende einer Berechnung) oder vom Controlling bereitgestellt.
    Sie wird im Browser gespeichert und danach automatisch verwendet.
    <strong>Was</strong>Der Werkstattbestand zu Jahresbeginn je Sparte (Anfangsbestand), die
    Zuschlagss&auml;tze als Sollwert und &ndash; im Jahr der Produktgruppen-Reorganisation &ndash; die
    Liste, welcher Artikel in welche neue Produktgruppe geh&ouml;rt.
    <strong>Warum</strong>Ohne Anfangsbestand l&auml;sst sich die Ver&auml;nderung des Bestands im
    Jahresverlauf nicht berechnen &ndash; der Endwert allein reicht nicht. Ohne die
    Umschl&uuml;sselung stehen alte und neue Produktgruppen nebeneinander.`,

  // --- Zeilen im Ergebnisteil: was die Zahl ist und wie sie zustande kommt ---
  "zeilen": `<strong>Was ist es</strong>Die Anzahl der Buchungszeilen, die aus dem
    KPTM-Export eingelesen wurden &ndash; eine Zeile ist eine Buchung
    (Kostenart &times; Kostenobjekt &times; Periode).
    <strong>Berechnung</strong>Schlicht die Zeilenzahl der Datei. Sie dient als Kontrolle,
    dass der richtige und vollst&auml;ndige Export geladen wurde.`,
  "erloese": `<strong>Was ist es</strong>Die Umsatzerl&ouml;se des Zeitraums, aufsummiert
    &uuml;ber alle Sparten.
    <strong>Berechnung</strong>Buchungen auf den Erl&ouml;skonten (4440*, 4460*, 4470*) plus
    Frachterl&ouml;se (4521*, 4531*, 4535*), je Produktgruppe summiert und im Vorzeichen
    gedreht &ndash; im ERP stehen Erl&ouml;se als Haben-Buchung negativ. Buchungen ohne
    Produktgruppe z&auml;hlen nicht mit.`,
  "db3": `<strong>Was ist es</strong>Deckungsbeitrag III: was je Sparte nach allen
    zugerechneten Kosten &uuml;brig bleibt.
    <strong>Berechnung</strong>Betriebsleistung (Erl&ouml;se + Bestandsver&auml;nderungen +
    aktivierte Eigenleistung) abz&uuml;glich Material, Fremdleistungen und sonstiger
    Aufwendungen ergibt die Rohmarge I; davon FEK, MGK und VVGK jeweils nach
    Deckungsdifferenz. &bdquo;bottom-up&ldquo; hei&szlig;t: je Sparte gerechnet, dann summiert &ndash;
    das ist <em>nicht</em> das Unternehmensergebnis, weil Buchungen ohne Sparte fehlen.`,
  "mgk_satz": `<strong>Was ist es</strong>Der Materialgemeinkosten-Satz, mit dem Ihre
    Kalkulation arbeitet &ndash; aus den Daten zur&uuml;ckgerechnet, nicht hinterlegt.
    <strong>Berechnung</strong>MGK geteilt durch (Material + Fremdleistungen). Nur eine
    Probe: Alle Werte kommen aus dem ERP, der Satz wird nie zum Rechnen benutzt. Streut
    er zwischen den Sparten, wurde er unterj&auml;hrig ge&auml;ndert oder die Kontenzuordnung
    stimmt nicht mehr.`,
  "vvgk_satz": `<strong>Was ist es</strong>Der Satz f&uuml;r Verwaltungs- und
    Vertriebsgemeinkosten, ebenfalls zur&uuml;ckgerechnet.
    <strong>Berechnung</strong>VVGK geteilt durch die Herstellkosten, also durch
    (Material + Fremdleistungen + MGK + FEK). Wie beim MGK-Satz nur eine Probe.`,
  "unproduktiv": `<strong>Was ist es</strong>Fertigungsstunden, die geleistet, aber auf
    keinen Fertigungsauftrag gebucht wurden. Sie liegen auf den Sammelauftr&auml;gen der
    Kostenstellen &ndash; je Kostenstelle einer. Wirtschaftlich ist das die FEK-Unterdeckung.
    <strong>Berechnung</strong>Alle ISWF-Buchungen auf Fertigungskostenstellen (K2*) mit
    Auftragsart &bdquo;5&ldquo;; die Stunden sind der Betrag geteilt durch den Stundensatz. Der
    Betrag wird bewusst auf keine Sparte verteilt &ndash; welcher Schl&uuml;ssel richtig w&auml;re,
    sagt erst der BAB.`,
  "wb_zugeordnet": `<strong>Was ist es</strong>Der Teil des Werkstattbestands, f&uuml;r den
    die Sparte feststeht. Nur diese Auftr&auml;ge gehen in die Bestandsver&auml;nderung UFE ein.
    <strong>Berechnung</strong>Die Sparte steht im Werkstattbestand selbst (Spalte
    &bdquo;Produktgruppe&ldquo;). Fehlt sie bei einem Auftrag, wird gesucht: Zuordnung von Hand,
    dann &uuml;ber die Auftragsnummer (aus KPTM), zuletzt &uuml;ber die Artikelnummer.
    Summiert wird die Spalte &bdquo;Offener Wert&ldquo;.`,
  "struktur": `<strong>Was ist es</strong>Seit der Reorganisation 2026 gelten neue
    Produktgruppen (vor allem Pressbohrtechnik). Buchungen aus den ersten Monaten stehen noch
    unter der alten Gruppe &ndash; sie werden je Artikel umgeschl&uuml;sselt, damit das ganze
    Jahr in einer Struktur steht.
    <strong>Berechnung</strong>Nachschlagen in der Liste &bdquo;Alte vs Neue Produktgruppen&ldquo;
    (in den Betriebsparametern hinterlegt). Was nur je Gruppe und nicht je Artikel vorliegt
    &ndash; Deckungsdifferenzen auf alten Gruppen, der Anfangsbestand der geteilten Gruppen &ndash;
    wird anteilig verteilt; diese Zellen sind in der Datei gelb (N&auml;herung).`,
  "wb_service": `<strong>Was ist es</strong>Serviceauftr&auml;ge &ndash; Monteureinsatz,
    Inbetriebnahme, Schulung. Dabei wird nichts gefertigt, also geh&ouml;ren sie nicht in den
    Werkstattbestand.
    <strong>Berechnung</strong>Erkannt daran, dass der Auftrag keine Artikelnummer tr&auml;gt.
    Sie werden auch dann nicht zugeordnet, wenn eine Quelle eine Sparte dafuer kennt &ndash;
    sonst verf&auml;lschen einzelne Eins&auml;tze eine ganze Sparte.`,
  "wb_offen": `<strong>Was ist es</strong>Auftr&auml;ge mit Artikelnummer, f&uuml;r die sich
    keine Sparte finden lie&szlig; &ndash; die echte L&uuml;cke.
    <strong>Berechnung</strong>Weder Auftrags- noch Artikelnummer f&uuml;hrten zu einer Sparte.
    Liegt der Anteil &uuml;ber 1 % des Gesamtbestands, wird die Zeile
    &bdquo;Bestandsver&auml;nderung UFE&ldquo; gar nicht erst gerechnet. Diese Auftr&auml;ge lassen sich
    unten von Hand zuordnen.`,
  "wb_gesamt": `<strong>Was ist es</strong>Der gesamte offene Werkstattbestand zum
    Stichtag, also alle noch nicht fertiggemeldeten Fertigungsauftr&auml;ge.
    <strong>Berechnung</strong>Summe der Spalte &bdquo;Offener Wert&ldquo; des PWBS-Exports
    (Istwert abz&uuml;glich bereits abgeliefertem Wert) &uuml;ber alle Zeilen. Entspricht
    zugeordnet + Serviceauftr&auml;ge + nicht zuordenbar.`,
};

(() => {
  const tooltip = document.createElement("div");
  tooltip.className = "info-tooltip";
  tooltip.setAttribute("role", "tooltip");
  tooltip.hidden = true;
  document.body.appendChild(tooltip);
  let aktivesIcon = null;

  function positionieren(icon) {
    const r = icon.getBoundingClientRect();
    const breite = Math.min(300, window.innerWidth - 24);
    tooltip.style.width = breite + "px";
    let left = Math.min(r.left, window.innerWidth - breite - 12);
    left = Math.max(12, left);
    tooltip.style.left = left + "px";
    // Hoehe haengt vom Zeilenumbruch ab, der wiederum von der eben gesetzten
    // Breite abhaengt - deshalb erst hier messen, nicht mit einer festen Zahl raten.
    const hoehe = tooltip.offsetHeight;
    const platzUnten = window.innerHeight - r.bottom;
    let top;
    if (platzUnten >= hoehe + 12) {
      top = r.bottom + 8;
    } else if (r.top >= hoehe + 12) {
      top = r.top - 8 - hoehe;
    } else {
      // Passt weder oben noch unten vollstaendig (sehr niedriger Bildschirm):
      // so gut wie moeglich platzieren und am Rand nicht abschneiden lassen.
      top = Math.max(12, Math.min(r.bottom + 8, window.innerHeight - hoehe - 12));
    }
    tooltip.style.top = top + "px";
  }

  function zeigen(icon) {
    const text = INFO_TEXTE[icon.dataset.info];
    if (!text) return;
    tooltip.innerHTML = text;
    tooltip.hidden = false;
    aktivesIcon = icon;
    positionieren(icon);
  }

  function verstecken(icon) {
    if (icon && icon.classList.contains("open")) return;   // per Klick offen gehalten
    tooltip.hidden = true;
    aktivesIcon = null;
  }

  /* Ueber das Dokument verteilt statt je Icon: die Icons im Ergebnisteil entstehen
     erst nach dem Rechnen (innerHTML) und bei jeder Neuberechnung erneut. Mit
     Listenern am einzelnen Icon haetten genau die keine Erklaerung. mouseover/
     focusin werden benutzt, weil mouseenter/focus nicht aufsteigen. */
  const zuIcon = (e) => (e.target && e.target.closest ? e.target.closest(".info-icon") : null);

  document.addEventListener("mouseover", (e) => { const i = zuIcon(e); if (i) zeigen(i); });
  document.addEventListener("mouseout", (e) => { const i = zuIcon(e); if (i) verstecken(i); });
  document.addEventListener("focusin", (e) => { const i = zuIcon(e); if (i) zeigen(i); });
  document.addEventListener("focusout", (e) => { const i = zuIcon(e); if (i) verstecken(i); });

  document.addEventListener("click", (e) => {
    const icon = zuIcon(e);
    document.querySelectorAll(".info-icon.open").forEach((o) => {
      if (o !== icon) o.classList.remove("open");
    });
    if (!icon) { verstecken(null); return; }
    const warOffen = icon.classList.contains("open");
    icon.classList.toggle("open", !warOffen);
    if (warOffen) verstecken(null); else zeigen(icon);
  });

  document.addEventListener("keydown", (e) => {
    const icon = zuIcon(e);
    if (!icon) return;
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); icon.click(); }
    if (e.key === "Escape") { icon.classList.remove("open"); verstecken(null); icon.blur(); }
  });
  window.addEventListener("scroll", () => { if (aktivesIcon) positionieren(aktivesIcon); }, true);
  window.addEventListener("resize", () => { if (aktivesIcon) positionieren(aktivesIcon); });
})();

/* Betriebsparameter (Stundensatz, Anfangsbestaende).
   Bewusst nicht im Repository - das ist oeffentlich, und aus dem Stundensatz liesse
   sich die Preisbildung zurueckrechnen. Stattdessen legt der Nutzer die Datei einmal
   hier ab; sie liegt dann in localStorage, also nur in seinem Browser auf seinem
   Rechner. Sie wird weder hochgeladen noch an Claude oder sonst jemanden gesendet. */
const PARAM_KEY = "spartenrechner.betriebsparameter";

function parameterLesen() {
  try {
    return localStorage.getItem(PARAM_KEY);
  } catch (e) {
    return null;   // privates Fenster oder Site-Daten blockiert
  }
}

function parameterInDateisystem() {
  const txt = parameterLesen();
  if (!txt || !pyodide) return false;
  pyodide.FS.writeFile("/config/betriebsparameter.json", txt);
  return true;
}

// Ob etwas hinterlegt ist, zeigt allein der "loeschen"-Button (erscheint nur dann) -
// ein zusaetzliches Status-Badge war neben der gelb hervorgehobenen Box ein
// Hervorhebung zu viel.
function parameterStatusZeigen() {
  document.getElementById("param-loeschen").style.display = parameterLesen() ? "inline-block" : "none";
}

/* Nimmt Excel (bevorzugt, weil im Controlling gepflegt) ebenso wie die frueher
   verwendete JSON-Datei. Excel wird beim Ablegen einmal nach JSON uebersetzt und
   nur so gespeichert - localStorage haelt Text, keine Binaerdateien. */
setupDropzone("dropzone-param", "dropzone-param-text", "file-param", async (file, zone, text) => {
  const name = file.name.toLowerCase();
  let inhalt;
  try {
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const ext = name.endsWith(".xls") ? ".xls" : ".xlsx";
      pyodide.FS.writeFile("/param_upload" + ext, new Uint8Array(await file.arrayBuffer()));
      inhalt = await pyodide.runPythonAsync(`
import json
json.dumps(spartenrechnung_core.betriebsparameter_aus_excel("/param_upload${ext}"))
      `);
    } else {
      inhalt = await file.text();
      JSON.parse(inhalt);
    }
  } catch (e) {
    showError(`„${file.name}" konnte nicht gelesen werden: ${e.message || e}`);
    return;
  }
  const geparst = JSON.parse(inhalt);
  if (!geparst.anfangsbestand || !Object.keys(geparst.anfangsbestand).length) {
    showError(`In „${file.name}" wurde kein Anfangsbestand gefunden. Erwartet werden `
      + `die Spalten Sparte, Geschäftsjahr und Anfangsbestand.`);
    return;
  }
  // Die heruntergeladene Vorlage hat dieselben Spalten, nur mit Nullen. Ungefüllt
  // hochgeladen ergäbe sie eine Bestandsveränderung in Höhe des vollen Endbestands –
  // eine Zahl, die plausibel aussieht und um den ganzen Anfangsbestand danebenliegt.
  const werteAlleJahre = Object.values(geparst.anfangsbestand)
    .flatMap((proJahr) => Object.values(proJahr || {}));
  if (!werteAlleJahre.some((w) => Math.abs(w) > 0.005)) {
    showError(`„${file.name}" enthält nur Nullen – das ist die unausgefüllte Vorlage. `
      + `Bitte je Sparte den Werkstattbestand zum 01.01. eintragen, sonst zeigt die Zeile `
      + `„Bestandsveränderung UFE" den vollen Endbestand statt der Veränderung.`);
    return;
  }
  try {
    localStorage.setItem(PARAM_KEY, inhalt);
  } catch (e) {
    showError("Der Browser konnte die Parameter nicht speichern (privates Fenster?). "
      + "Sie gelten für diesen Durchlauf, müssen aber beim nächsten Mal erneut ausgewählt werden.");
  }
  parameterInDateisystem();
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
  // Eine vorherige Fehlermeldung wegräumen - sonst steht nach dem Nachreichen der
  // richtigen Datei noch die Absage zur falschen da und sieht aus, als sei auch
  // diese abgelehnt worden.
  errorBox.style.display = "none";
  parameterStatusZeigen();
});

// Vorlage: wird im Browser erzeugt, damit keine Excel-Datei im Repository liegen muss.
document.getElementById("param-vorlage").addEventListener("click", async (e) => {
  e.preventDefault();
  if (!pyodide) return;
  const jahr = new Date().getFullYear();
  await pyodide.runPythonAsync(`
spartenrechnung_core.schreibe_betriebsparameter_excel(
    {"${jahr}": {"DP0001": 0.0, "PB0001": 0.0}}, "/vorlage.xlsx")
  `);
  const bytes = pyodide.FS.readFile("/vorlage.xlsx");
  const url = URL.createObjectURL(new Blob([bytes],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = "betriebsparameter_vorlage.xlsx";
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById("param-loeschen").addEventListener("click", () => {
  try {
    localStorage.removeItem(PARAM_KEY);
  } catch (e) { /* dann war ohnehin nichts gespeichert */ }
  if (pyodide) {
    try { pyodide.FS.unlink("/config/betriebsparameter.json"); } catch (e) { /* nicht vorhanden */ }
  }
  const zone = document.getElementById("dropzone-param");
  zone.classList.remove("has-file");
  document.getElementById("dropzone-param-text").textContent = "betriebsparameter.json auswählen";
  parameterStatusZeigen();
});

async function init() {
  try {
    pyodide = await loadPyodide();
    loadingStatus.textContent = "⏳ Lade benötigte Python-Pakete (pandas, openpyxl) …";
    await pyodide.loadPackage(["pandas", "micropip"]);
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");
    try {
      await micropip.install("xlrd");
    } catch (e) {
      console.warn("xlrd konnte nicht geladen werden (nur fuer alte .xls-Dateien noetig):", e);
    }

    // Kernmodul + Konfiguration laden (aus diesem Repository, nicht von einem fremden Server)
    // "no-cache" heisst: beim Server nachfragen, ob sich die Datei geaendert hat. Sonst
    // rechnet der Browser nach einer Aktualisierung noch mit dem alten Code oder der
    // alten Konfiguration weiter - ohne dass es irgendwo sichtbar wuerde.
    for (const modul of ["spartenrechnung_core.py", "herleitung_excel.py"]) {
      pyodide.FS.writeFile("/" + modul, await (await fetch(modul, { cache: "no-cache" })).text());
    }
    pyodide.FS.mkdir("/config");
    for (const name of ["kostenart_mapping.json", "produktgruppen.json"]) {
      const txt = await (await fetch("config/" + name, { cache: "no-cache" })).text();
      pyodide.FS.writeFile("/config/" + name, txt);
    }
    // Betriebsparameter kommen nicht aus dem Repository, sondern aus dem Browser
    // des Nutzers - falls er sie schon einmal hinterlegt hat.
    parameterInDateisystem();
    parameterStatusZeigen();
    await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, "/")
import spartenrechnung_core
    `);

    loadingStatus.style.display = "none";
    form.style.display = "block";
    submitBtn.disabled = false;
  } catch (err) {
    loadingStatus.textContent = "❌ Fehler beim Laden der Python-Umgebung: " + err;
  }
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  errorBox.style.display = "none";
  resultBox.style.display = "none";

  if (!kptmFile) {
    showError("Bitte zuerst die KPTM_Wertsummen-Datei auswählen.");
    return;
  }
  if (!kostentraegerFile) {
    showError("Bitte die Kostenträger-Auswertung der Buchhaltung auswählen.");
    return;
  }
  if (!pwbsEndeFile) {
    showError("Bitte den PWBS Werkstattbestand zum Periodenende auswählen.");
    return;
  }

  submitBtn.disabled = true;
  submitBtn.textContent = "Wird berechnet …";
  await rechenAnimationZeigen();

  try {
    // Unter dem Originalnamen ablegen, je Feld in einem eigenen Ordner: Fehlermeldungen
    // nennen dann genau die Datei, die der Nutzer ausgewaehlt hat (wichtig, wenn jemand
    // die Kostenstellen- mit der Kostentraeger-Auswertung verwechselt), und aus dem
    // KPTM-Namen laesst sich notfalls der Zeitraum lesen.
    const kptmPfad = await ablegen("/kptm", kptmFile);
    const ktPfad = await ablegen("/kostentraeger", kostentraegerFile);
    // Optional: ohne SuSa entfaellt nur die Abstimmung mit der FiBu, gerechnet wird trotzdem.
    const susaPfad = susaFile ? await ablegen("/susa", susaFile) : null;

    // Werkstattbestand fuer die Bestandsveraenderung UFE (Pflicht, siehe Pruefung oben).
    // Der Anfangsbestand kommt aus den hinterlegten Betriebsparametern.
    // Unter dem Originalnamen: der Stichtag steht nur im Dateinamen (z.B. ..._010426)
    // und wird gegen den Zeitraum des KPTM-Exports geprueft.
    const pwbsAArg = "None";
    const pwbsEArg = JSON.stringify(await ablegen("/pwbs_ende", pwbsEndeFile));
    // PFAK wird nicht mehr gebraucht: der Werkstattbestand bringt die Sparte selbst mit.
    const pfakArg = "None";

    letzterLauf = { kptmPfad, ktPfad, susaPfad, pwbsAArg, pwbsEArg, pfakArg };
    manuelleZuordnung = {};   // neuer Lauf mit neuen Dateien: alte Zuordnungen verwerfen
    await berechnen();
  } catch (err) {
    showError(fehlertext(err));
  } finally {
    rechenAnimationVerbergen();
    submitBtn.disabled = false;
    submitBtn.textContent = "Spartenrechnung erstellen";
  }
});

// Argumente des letzten Laufs, damit eine manuelle Zuordnung ohne erneutes
// Hochladen der Dateien angewendet werden kann.
let letzterLauf = null;
// Bereits von Hand vorgenommene Zuordnungen. Muss ueber Durchlaeufe hinweg erhalten
// bleiben: nach dem Neuberechnen zeigt die Tabelle nur noch die VERBLEIBENDEN offenen
// Auftraege, die frueheren Eintraege stehen dort also nicht mehr zur Auswahl.
let manuelleZuordnung = {};

/* Rechen-Animation. Das Zaehlwerk wird einmal aufgebaut: Ziffernwalzen mit
   unterschiedlicher Drehzahl (rechts schnell, links langsam) in Euro-Schreibweise. */
function zaehlwerkAufbauen() {
  const ziel = document.getElementById("zaehlwerk");
  if (!ziel || ziel.childElementCount) return;
  const muster = "1.234.567,89 €";
  const ziffern = [...muster].filter((z) => /\d/.test(z)).length;
  let stelle = 0;
  for (const z of muster) {
    if (!/\d/.test(z)) {
      const s = document.createElement("span");
      s.className = "zeichen";
      s.textContent = z === " " ? "\u00a0" : z;
      ziel.appendChild(s);
      continue;
    }
    const walze = document.createElement("span");
    walze.className = "walze";
    const band = document.createElement("span");
    band.className = "band";
    // Rechte Stellen drehen schnell, linke langsam - wie ein echtes Zaehlwerk.
    const rang = ziffern - 1 - stelle;
    band.style.setProperty("--dauer", (0.5 + rang * 0.55).toFixed(2) + "s");
    band.style.animationDelay = (-Math.random() * 3).toFixed(2) + "s";
    for (let i = 0; i < 20; i++) {
      const d = document.createElement("span");
      d.textContent = String(i % 10);
      band.appendChild(d);
    }
    walze.appendChild(band);
    ziel.appendChild(walze);
    stelle++;
  }
}

async function rechenAnimationZeigen() {
  const o = document.getElementById("rechnet-overlay");
  if (!o || !o.hidden) return;
  zaehlwerkAufbauen();
  o.hidden = false;
  // Kurz warten, bis das Overlay gezeichnet ist, bevor Python den Browser blockiert -
  // sonst erschiene es erst, wenn die Berechnung schon vorbei ist. Bewusst ohne
  // Einblend-Effekt: der koennte halbfertig stehen bleiben.
  // Hoechstens 150 ms: Liegt der Tab im Hintergrund, zeichnet der Browser gar nicht -
  // ohne Obergrenze begaenne die Berechnung dann erst bei der Rueckkehr zum Tab.
  await Promise.race([
    new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))),
    new Promise((r) => setTimeout(r, 150)),
  ]);
  await new Promise((r) => setTimeout(r, 50));
}

function rechenAnimationVerbergen() {
  const o = document.getElementById("rechnet-overlay");
  if (o) o.hidden = true;
}

async function berechnen() {
  await rechenAnimationZeigen();
  try {
    await berechnenOhneAnimation();
  } finally {
    rechenAnimationVerbergen();
  }
}

async function berechnenOhneAnimation() {
  const L = letzterLauf;
  const manuellArg = Object.keys(manuelleZuordnung).length
    ? JSON.stringify(manuelleZuordnung) : "None";
  const resultJson = await pyodide.runPythonAsync(`
import json
summary = spartenrechnung_core.generate(
    ${JSON.stringify(L.kptmPfad)}, "/config", "/output.xlsx",
    pwbs_anfang=${L.pwbsAArg}, pwbs_ende=${L.pwbsEArg}, pfak_pfade=${L.pfakArg},
    manuelle_zuordnung=${manuellArg},
    kostentraeger_pfad=${JSON.stringify(L.ktPfad)},
    susa_pfad=${L.susaPfad ? JSON.stringify(L.susaPfad) : "None"},
)
json.dumps(summary)
  `);
  const summary = JSON.parse(resultJson);
  const outBytes = pyodide.FS.readFile("/output.xlsx");
  const blob = new Blob([outBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  // Der Dateiname kommt aus den Rohdaten (Jahr + Perioden), nicht aus einer Eingabe.
  showResult(summary, url, summary.dateiname);
}

/* Eine Zeile der Ergebnisuebersicht. 'info' ist der Schluessel in INFO_TEXTE und
   haengt das "?" mit der Erklaerung an - dieselbe Mechanik wie im Formular. */
function summaryZeile(label, wert, info, fett) {
  const l = fett ? `<strong>${label}</strong>` : label;
  const w = fett ? `<strong>${wert}</strong>` : wert;
  const fragezeichen = info
    ? ` <span class="info-icon" tabindex="0" role="button" data-info="${info}"
         aria-label="Erklärung: ${label}">?</span>`
    : "";
  return `<div class="summary-line"><span>${l}${fragezeichen}</span><span>${w}</span></div>`;
}

/* Ergebnisansicht. Oben nur, was man sofort wissen muss (Ampel, drei Kennzahlen,
   Download); alles Weitere in aufklappbaren Bereichen. Meldungen stehen dort als
   einzeilige Titel - der ausfuehrliche Text erst auf Klick. Vorher stand alles
   untereinander ausgeschrieben, und die wichtigen Punkte gingen darin unter. */
function esc(t) {
  return String(t ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// Kurzer Titel aus einer langen Meldung: der erste Satz, notfalls gekuerzt.
function titelAus(text) {
  const t = String(text || "").trim();
  const ende = t.search(/[.:](\s|$)/);
  const satz = ende > 0 ? t.slice(0, ende) : t;
  return satz.length > 90 ? satz.slice(0, 88) + "…" : satz;
}

/* Eine Meldung: einzeiliger Titel, Text aufklappbar. stufe: "kritisch" | "pruefen" | "info" */
function meldung(stufe, titel, text, details) {
  const zeilen = [].concat(text ? [text] : [], details || []);
  const koerper = zeilen.length
    ? `<div class="meldung-text">${zeilen.map((z) => `<p>${esc(z)}</p>`).join("")}</div>` : "";
  return `<details class="meldung meldung-${stufe}"><summary><span class="punkt"></span>
    <span class="meldung-titel">${esc(titel)}</span></summary>${koerper}</details>`;
}

/* Ein aufklappbarer Bereich mit Kurzinfo rechts (sichtbar auch eingeklappt). */
function bereich(id, titel, meta, metaStufe, inhalt, offen) {
  return `<details class="abschnitt" id="${id}"${offen ? " open" : ""}>
    <summary><span class="abschnitt-titel">${titel}</span>
      <span class="abschnitt-meta meta-${metaStufe}">${meta}</span></summary>
    <div class="abschnitt-inhalt">${inhalt}</div></details>`;
}

function showResult(summary, url, filename) {
  const eurF = (v) => (v || 0).toLocaleString("de-DE", {minimumFractionDigits: 2, maximumFractionDigits: 2});
  const eur = (v) => eurF(v) + " €";
  const pct = (v) => (v * 100).toLocaleString("de-DE", {minimumFractionDigits: 1, maximumFractionDigits: 1}) + " %";
  const abst = summary.abstimmung || [];
  const stat = summary.ufe_statistik;
  const st = summary.struktur;
  const dd = summary.deckungsdifferenzen || {};
  const ufeGesperrt = !!(summary.ufe_warnungen && summary.ufe_warnungen.length);

  // --- Meldungen einsammeln: was Zahlen falsch macht oder fehlen laesst ("zu pruefen"),
  //     und was man wissen sollte, ohne dass etwas fehlt ("Hinweise").
  const pruefen = [];
  const hinweise = [];
  if (ufeGesperrt) {
    pruefen.push(meldung("kritisch", "Bestandsveränderung UFE wurde nicht berechnet",
      "Die Zeile steht auf 0,00 – Betriebsleistung und alle Deckungsbeiträge darunter sind um diesen Betrag zu niedrig.",
      summary.ufe_warnungen));
  }
  for (const p of abst.filter((x) => x.status === "abweichung")) {
    // Die Werkstattbestands-Pruefung steckt bei gesperrter UFE-Zeile schon oben drin.
    if (ufeGesperrt && p.text.startsWith("Werkstattbestand gegen gebuchte")) continue;
    pruefen.push(meldung("pruefen", "Abstimmung: " + p.text, null, p.details));
  }
  if (st && st.fehlt) {
    pruefen.push(meldung("pruefen", "Alte und neue Produktgruppen stehen nebeneinander",
      `Der Export enthält Buchungen auf aufgelösten Produktgruppen (${st.alte.join(", ")}), es ist aber keine Umschlüsselung hinterlegt. Bitte die aktuelle betriebsparameter.xlsx (mit Blatt „Umschluesselung") laden.`));
  }
  for (const h of dd.hinweise || []) {
    pruefen.push(meldung("pruefen", "Deckungsdifferenzen: " + titelAus(h), h));
  }
  if (summary.unbekannte_kostenarten && summary.unbekannte_kostenarten.length) {
    const b = summary.unmapped_betrag || {};
    const summe = Object.values(b).reduce((a, v) => a + v, 0);
    pruefen.push(meldung("pruefen", `Kostenarten ohne Zuordnung – ${eur(summe)} fehlen im Ergebnis`,
      "Diese Konten sind keiner GuV-Zeile zugeordnet; ihre Beträge fehlen. Bitte in config/kostenart_mapping.json eintragen.",
      summary.unbekannte_kostenarten.map((k) => `${k}: ${eur(b[k])}`)));
  }
  if (summary.unbekannte_produktgruppen && summary.unbekannte_produktgruppen.length) {
    pruefen.push(meldung("pruefen", "Unbekannte Produktgruppen: " + summary.unbekannte_produktgruppen.join(", "),
      "Nicht in config/produktgruppen.json – ggf. Reorganisation. Sie erscheinen als eigene Spalte."));
  }
  if (summary.fracht_ohne_ziel && Object.keys(summary.fracht_ohne_ziel).length) {
    pruefen.push(meldung("pruefen", "Ausgangsfracht blieb auf der operativen Sparte",
      "Es fehlt das Kostenobjekt <Familie>0100. Der Deckungsbeitrag dieser Sparte ist dadurch zu niedrig.",
      Object.entries(summary.fracht_ohne_ziel).map(([k, v]) => `${k}: ${eur(v)}`)));
  }
  for (const p of abst.filter((x) => x.status === "hinweis")) {
    hinweise.push(meldung("info", p.text, null, p.details));
  }
  for (const h of (stat && stat.hinweise) || []) {
    hinweise.push(meldung("info", "Werkstattbestand: " + titelAus(h), h));
  }
  for (const h of summary.mgk_hinweise || []) {
    hinweise.push(meldung("info", "Zuschlagssätze: " + titelAus(h), h));
  }
  if (summary.per_regel_zugeordnet && summary.per_regel_zugeordnet.length) {
    const k = summary.per_regel_zugeordnet;
    hinweise.push(meldung("info", `${k.length} Konten per Kontenregel zugeordnet`,
      "Nicht einzeln hinterlegt – alle 6er-Konten gehen auf „sbA\". Bitte kurz prüfen, ob das für diese Konten stimmt; falls nicht, in config/kostenart_mapping.json einzeln eintragen.",
      [k.join(", ")]));
  }
  if (abst.length && !summary.susa_geladen) {
    hinweise.push(meldung("info", "Ohne SuSa: Abstimmung mit der FiBu entfällt",
      "Geprüft wurde nur gegen die Kostenträger-Auswertung. Mit SuSa kommen Erlöse, Werkstattbestand und Anfangsbestand gegen die FiBu dazu."));
  }

  // --- Kopf: Ampel, drei Kennzahlen, Download
  // Bestanden ist alles ausser einer Abweichung - ein Hinweis ist kein Fehler.
  const nOk = abst.filter((p) => p.status !== "abweichung").length;
  const ampel = pruefen.length
    ? `<span class="ampel ampel-rot">⚠ ${pruefen.length} ${pruefen.length === 1 ? "Punkt" : "Punkte"} prüfen</span>`
    : `<span class="ampel ampel-gruen">✓ Abgestimmt</span>`;
  let html = `<div class="ergebnis-kopf"><h2>Spartenrechnung erstellt</h2>
    <span class="ergebnis-zeitraum">${esc(summary.zeitraum)}</span>${ampel}</div>`;
  html += `<div class="kacheln">
    <div class="kachel"><span class="kachel-label">Erlöse</span><span class="kachel-wert">${eurF(summary.erloese_summe)} €</span></div>
    <div class="kachel"><span class="kachel-label">DB III (Sparten)</span><span class="kachel-wert">${eurF(summary.db3_summe)} €</span></div>
    <div class="kachel${abst.length ? " kachel-link" : ""}" id="kachel-abstimmung"${abst.length ? ' role="button" tabindex="0" title="Prüfungen anzeigen"' : ""}><span class="kachel-label">Prüfungen bestanden</span><span class="kachel-wert">${abst.length ? `${nOk} / ${abst.length}${nOk === abst.length ? " ✓" : ""}` : "–"}</span></div>
  </div>`;
  html += `<a class="download-btn" href="${url}" download="${filename}">⬇ ${esc(filename)} herunterladen</a>`;

  // --- Bereiche
  if (pruefen.length) {
    html += bereich("ab-pruefen", "Zu prüfen", String(pruefen.length), "rot",
      `<div class="meldungen">${pruefen.join("")}</div>`, true);
  }

  if (abst.length) {
    const zeichen = { ok: "✓", hinweis: "ℹ", abweichung: "⚠" };
    let inhalt = "";
    for (const p of abst) {
      // Nicht bestandene Pruefungen zeigen ihre Zahlen direkt darunter.
      const unter = p.status !== "ok" && p.details && p.details.length
        ? `<div class="pruef-details">${p.details.slice(0, 6).map((d) => `<div>${esc(d)}</div>`).join("")}</div>` : "";
      inhalt += `<div class="pruef-zeile pruef-${p.status}"><span class="pruef-zeichen">${zeichen[p.status]}</span>
        <span>${esc(p.text)}${unter}</span></div>`;
    }
    inhalt += `<p class="hinweis">Gegenprobe gegen die Kostenträger-Auswertung${summary.susa_geladen ? " und die SuSa" : ""}
      – Quellen, die unabhängig vom KPTM-Export entstehen.
      <span class="info-icon" tabindex="0" role="button" data-info="abstimmung" aria-label="Erklärung: Abstimmung">?</span></p>`;
    html += bereich("ab-abstimmung", "Abstimmung mit der Buchhaltung", `${nOk} von ${abst.length}`,
      nOk === abst.length ? "gruen" : "gelb", inhalt, false);
  }

  let kennzahlen = summaryZeile("Rohdaten-Zeilen verarbeitet", summary.zeilen.toLocaleString("de-DE"), "zeilen");
  kennzahlen += summaryZeile("Erlöse (Summe der Sparten)", eur(summary.erloese_summe), "erloese");
  kennzahlen += summaryZeile("DB III bottom-up (Summe der Sparten)", eur(summary.db3_summe), "db3");
  // Die aus den Rohdaten zurueckgerechneten Zuschlagssaetze - nur eine Probe.
  const saetze = summary.zuschlagssaetze || {};
  for (const [key, label] of [["mgk", "MGK"], ["vvgk", "VVGK"]]) {
    if (saetze[key] === undefined || saetze[key] === null) continue;
    const satz = (saetze[key] * 100).toLocaleString("de-DE", {minimumFractionDigits: 2, maximumFractionDigits: 2});
    kennzahlen += summaryZeile(`${label}-Zuschlagssatz (Probe)`, satz + " %", key + "_satz");
  }
  if (summary.unproduktive_gemeinkosten) {
    const std = summary.unproduktive_stunden.toLocaleString("de-DE", {maximumFractionDigits: 0});
    kennzahlen += summaryZeile(`Nicht auf Kostenträger gebuchte Stunden (${std} h)`,
      eur(summary.unproduktive_gemeinkosten), "unproduktiv");
  }
  if (st && !st.fehlt) {
    kennzahlen += summaryZeile("Umgeschlüsselte Buchungen (neue Produktgruppen)",
      st.zeilen.toLocaleString("de-DE") + ` · ${summary.naeherung_zellen} Zellen Näherung`, "struktur");
  }
  html += bereich("ab-kennzahlen", "Kennzahlen",
    summary.zeilen.toLocaleString("de-DE") + " Buchungen", "neutral", kennzahlen, false);

  if (stat) {
    const zugeordnetQuote = stat.gesamt ? stat.zugeordnet / stat.gesamt : 0;
    let wb = summaryZeile("einer Sparte zugeordnet", `${eur(stat.zugeordnet)} · ${pct(zugeordnetQuote)}`, "wb_zugeordnet");
    wb += summaryZeile("Serviceaufträge (bewusst ausgeschlossen)", `${eur(stat.service)} · ${pct(stat.service_quote)}`, "wb_service");
    wb += summaryZeile("nicht zuordenbar", `${eur(stat.offen)} · ${pct(stat.luecke_quote)}`, "wb_offen");
    wb += summaryZeile("Werkstattbestand gesamt", eur(stat.gesamt), "wb_gesamt", true);

    const bereits = Object.keys(manuelleZuordnung).length;
    if (bereits) {
      wb += summaryZeile("davon von Hand zugeordnet", `${bereits} Aufträge`);
    }
    if (stat.ueber_artikel > 0 || stat.service_mit_quelle > 0) {
      wb += `<details class="zuordnung"><summary>Wie belastbar ist die Zuordnung?</summary>`;
      wb += summaryZeile("über die Auftragsnummer (sicherster Weg)", eur(stat.zugeordnet - stat.ueber_artikel));
      wb += summaryZeile("nur über die Artikelnummer (schwächer)", eur(stat.ueber_artikel));
      if (stat.mehrdeutige_artikel) {
        wb += summaryZeile("Artikel ohne eindeutige Sparte (nicht nutzbar)", String(stat.mehrdeutige_artikel));
      }
      wb += summaryZeile("Serviceaufträge, für die eine Sparte bekannt wäre", eur(stat.service_mit_quelle));
      wb += `<p class="hinweis">Der Weg über die <strong>Artikelnummer</strong> ist schwächer:
        Wird ein Artikel in einem späteren Export mehrdeutig, entfällt die Zuordnung. Die
        genannten Serviceaufträge werden bewusst abgezogen (es wird nichts gefertigt).</p></details>`;
    }
    // Der Bestand zum Stichtag ist der Anfangsbestand der Folgeperiode.
    if (stat.bestand_stichtag && Object.keys(stat.bestand_stichtag).length) {
      wb += `<details class="zuordnung"><summary>Werkstattbestand als Betriebsparameter sichern</summary>`;
      wb += `<p class="hinweis">Der hier errechnete Bestand ist zugleich der
        <strong>Anfangsbestand der Folgeperiode</strong>. Beim Abschluss zum <strong>01.01.</strong>
        ist das die Jahreskonstante des neuen Jahres – speichern, Jahreszahl prüfen, fertig.</p>`;
      wb += `<button type="button" id="param-sichern">⬇ betriebsparameter.xlsx</button></details>`;
    }
    if (stat.offene_auftraege && stat.offene_auftraege.length) {
      const optionen = (summary.produktgruppen_auswahl || [])
        .map((p) => `<option value="${p}">${p}</option>`).join("");
      wb += `<details class="zuordnung" id="zuordnung-block"${bereits ? " open" : ""}><summary>`;
      wb += `${stat.offene_auftraege.length} Aufträge selbst zuordnen `;
      wb += `<span class="meta-label">(nach Wert sortiert)</span></summary>`;
      wb += `<p class="hinweis">Ordnen Sie die größten Posten zu, bis die Lücke unter 1 % fällt –
        dann rechnet die UFE-Zeile. Zugeordnete Aufträge verschwinden aus dieser Liste und bleiben
        gespeichert, solange die Seite geöffnet ist.</p>`;
      wb += `<table class="zuordnung-tabelle"><thead><tr>
        <th>Auftrag</th><th>Bezeichnung</th><th style="text-align:right">Offener Wert</th><th>Sparte</th>
        </tr></thead><tbody>`;
      for (const a of stat.offene_auftraege.slice(0, 60)) {
        wb += `<tr><td>${esc(a.auftrag)}</td><td>${esc(a.bezeichnung)}</td>`;
        wb += `<td style="text-align:right">${eur(a.wert)}</td>`;
        wb += `<td><select class="zuordnung-select" data-auftrag="${esc(a.auftrag)}">
          <option value="">– offen –</option>${optionen}</select></td></tr>`;
      }
      wb += `</tbody></table>`;
      if (stat.offene_auftraege.length > 60) {
        wb += `<p class="hinweis">Angezeigt werden die 60 größten von ${stat.offene_auftraege.length} Aufträgen.</p>`;
      }
      wb += `<button type="button" id="neu-berechnen">Mit dieser Zuordnung neu berechnen</button></details>`;
    }
    html += bereich("ab-werkstatt", "Werkstattbestand",
      // Serviceauftraege sind bewusst ausgeschlossen - massgeblich ist, was OFFEN bleibt.
      ufeGesperrt ? "UFE nicht berechnet"
        : (stat.offen > 0.005 ? `${pct(stat.luecke_quote)} offen` : "vollständig zugeordnet"),
      ufeGesperrt ? "rot" : (stat.luecke_quote < 0.01 ? "gruen" : "gelb"), wb, bereits > 0);
  }

  if (hinweise.length) {
    html += bereich("ab-hinweise", "Hinweise", String(hinweise.length), "gelb",
      `<div class="meldungen">${hinweise.join("")}</div>`, false);
  }

  html += `<div class="legend">In der Excel-Datei:
    <span class="status-badge badge-gruen">automatisch</span>
    <span class="status-badge badge-gelb">Näherung</span>
    <span class="status-badge badge-rot">manuell zu ergänzen</span>
  </div>`;

  resultBox.innerHTML = html;
  resultBox.style.display = "block";

  // Die Kachel "Pruefungen bestanden" oeffnet die Liste der Pruefungen.
  const kachel = document.getElementById("kachel-abstimmung");
  const abschnitt = document.getElementById("ab-abstimmung");
  if (kachel && abschnitt) {
    const oeffnen = () => {
      abschnitt.open = true;
      abschnitt.scrollIntoView({ behavior: "smooth", block: "start" });
    };
    kachel.addEventListener("click", oeffnen);
    kachel.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); oeffnen(); }
    });
  }

  const sichernBtn = document.getElementById("param-sichern");
  if (sichernBtn) {
    sichernBtn.addEventListener("click", async () => {
      // Geschaeftsjahr aus den Rohdaten (Spalte 'Geschaeftsjahr'), nicht aus Text geraten.
      const jahr = summary.jahr || new Date().getFullYear();
      // Fortschreiben statt neu anlegen: die uebrigen Jahre, die Zuschlagssaetze und
      // die Umschluesselung bleiben erhalten - sonst fehlten sie beim naechsten Laden.
      const bestand = JSON.stringify(summary.ufe_statistik.bestand_stichtag);
      const bisher = parameterLesen() || "{}";
      await pyodide.runPythonAsync(`
import json
spartenrechnung_core.betriebsparameter_fortschreiben(
    json.loads(${JSON.stringify(bisher)}), ${JSON.stringify(String(jahr))},
    json.loads(${JSON.stringify(bestand)}), "/betriebsparameter_neu.xlsx")
      `);
      const bytes = pyodide.FS.readFile("/betriebsparameter_neu.xlsx");
      const url = URL.createObjectURL(new Blob([bytes],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url;
      a.download = "betriebsparameter.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    });
  }

  const neuBtn = document.getElementById("neu-berechnen");
  if (neuBtn) {
    neuBtn.addEventListener("click", async () => {
      let neue = 0;
      for (const sel of document.querySelectorAll(".zuordnung-select")) {
        if (sel.value) { manuelleZuordnung[sel.dataset.auftrag] = sel.value; neue++; }
      }
      if (!neue) {
        alert("Bitte mindestens einem Auftrag eine Sparte zuordnen.");
        return;
      }
      neuBtn.disabled = true;
      neuBtn.textContent = "Wird neu berechnet …";
      try {
        await berechnen();
      } catch (err) {
        showError(fehlertext(err));
      }
    });
  }
  resultBox.scrollIntoView({ behavior: "smooth", block: "start" });
}

/* Legt eine hochgeladene Datei im virtuellen Dateisystem ab - in einem eigenen Ordner
   je Feld und unter ihrem echten Namen (nur Pfadtrenner werden ersetzt). */
async function ablegen(ordner, datei) {
  try { pyodide.FS.mkdir(ordner); } catch (e) { /* existiert schon */ }
  const pfad = ordner + "/" + datei.name.replace(/[/\\]/g, "_");
  pyodide.FS.writeFile(pfad, new Uint8Array(await datei.arrayBuffer()));
  return pfad;
}

/* Aus einem Python-Fehler nur die Meldung holen. Die Pruefungen im Kern (falsche Datei,
   fehlende Spalten, falscher Zeitraum) werfen ValueError mit einem Satz, der sagt, was zu
   tun ist - der stand bisher am Ende eines Tracebacks, den niemand liest. Bei allen
   anderen Fehlern bleibt der volle Text stehen, der wird zur Fehlersuche gebraucht. */
function fehlertext(err) {
  const text = String((err && err.message) || err).trim();
  const letzte = text.split("\n").pop();
  const m = letzte.match(/^ValueError: (.+)$/);
  return m ? m[1] : text;
}

function showError(msg) {
  errorBox.textContent = "❌ " + msg;
  errorBox.style.display = "block";
}

init();
