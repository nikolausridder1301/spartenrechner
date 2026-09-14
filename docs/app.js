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
let bwaFile = null;
let pwbsAnfangFile = null;
let pwbsEndeFile = null;
let pfakFiles = [];

function setupDropzone(zoneId, textId, inputId, onFile) {
  const zone = document.getElementById(zoneId);
  const text = document.getElementById(textId);
  const input = document.getElementById(inputId);

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

setupDropzone("dropzone-bwa", "dropzone-bwa-text", "file-bwa", async (file, zone, text) => {
  bwaFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
  await populateBwaSheets(file);
});

setupDropzone("dropzone-pwbs-a", "dropzone-pwbs-a-text", "file-pwbs-a", (file, zone, text) => {
  pwbsAnfangFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

setupDropzone("dropzone-pwbs-e", "dropzone-pwbs-e-text", "file-pwbs-e", (file, zone, text) => {
  pwbsEndeFile = file;
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
});

// PFAK erlaubt mehrere Dateien - deshalb eigener Handler statt setupDropzone
(() => {
  const zone = document.getElementById("dropzone-pfak");
  const text = document.getElementById("dropzone-pfak-text");
  const input = document.getElementById("file-pfak");
  const uebernehmen = (files) => {
    pfakFiles = Array.from(files);
    zone.classList.add("has-file");
    text.textContent = `✓ ${pfakFiles.length} Datei(en): ` + pfakFiles.map((f) => f.name).join(", ");
  };
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => { if (input.files.length) uebernehmen(input.files); });
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (e.dataTransfer.files.length) uebernehmen(e.dataTransfer.files);
  });
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

function parameterStatusZeigen() {
  const badge = document.getElementById("parameter-status");
  const loeschen = document.getElementById("param-loeschen");
  const txt = parameterLesen();
  if (!txt) {
    badge.textContent = "nicht hinterlegt";
    badge.className = "param-badge";
    loeschen.style.display = "none";
    return;
  }
  let info = "hinterlegt";
  try {
    const p = JSON.parse(txt);
    const jahre = Object.keys(p.anfangsbestand || {});
    const teile = [];
    if (p.standard_stundensatz) teile.push("Stundensatz");
    if (jahre.length) teile.push("Anfangsbestand " + jahre.join(", "));
    if (teile.length) info = "hinterlegt: " + teile.join(" · ");
  } catch (e) { /* Anzeige ist nachrangig */ }
  badge.textContent = info;
  badge.className = "param-badge ok";
  loeschen.style.display = "inline-block";
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
  try {
    localStorage.setItem(PARAM_KEY, inhalt);
  } catch (e) {
    showError("Der Browser konnte die Parameter nicht speichern (privates Fenster?). "
      + "Sie gelten für diesen Durchlauf, müssen aber beim nächsten Mal erneut ausgewählt werden.");
  }
  parameterInDateisystem();
  zone.classList.add("has-file");
  text.textContent = "✓ " + file.name;
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

async function populateBwaSheets(file) {
  const field = document.getElementById("bwa-sheet-field");
  const select = document.getElementById("bwa-sheet");
  select.innerHTML = "";
  field.style.display = "none";
  if (!pyodide) return;

  const bytes = new Uint8Array(await file.arrayBuffer());
  pyodide.FS.writeFile("/bwa_scan.xlsx", bytes);
  const sheetsJson = await pyodide.runPythonAsync(`
import openpyxl, json, warnings
warnings.filterwarnings("ignore")
wb = openpyxl.load_workbook("/bwa_scan.xlsx", read_only=True)
sheets = [s for s in wb.sheetnames if s.strip().upper().startswith("BWA")]
json.dumps(sheets)
  `);
  const sheets = JSON.parse(sheetsJson);
  if (sheets.length === 0) return;
  for (const s of sheets) {
    const opt = document.createElement("option");
    opt.value = s; opt.textContent = s;
    select.appendChild(opt);
  }
  field.style.display = "block";
}

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
    for (const modul of ["spartenrechnung_core.py", "herleitung_excel.py"]) {
      pyodide.FS.writeFile("/" + modul, await (await fetch(modul)).text());
    }
    pyodide.FS.mkdir("/config");
    for (const name of ["kostenart_mapping.json", "produktgruppen.json"]) {
      const txt = await (await fetch("config/" + name)).text();
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

  submitBtn.disabled = true;
  submitBtn.textContent = "Wird berechnet …";

  try {
    // Unter dem Originalnamen ablegen: Fehlermeldungen nennen dann die Datei, die
    // der Nutzer ausgewaehlt hat, und aus dem Namen laesst sich notfalls der
    // Zeitraum lesen, falls dem Export die Periodenspalten fehlen.
    const kptmPfad = "/" + kptmFile.name.replace(/[^\w.-]+/g, "_");
    pyodide.FS.writeFile(kptmPfad, new Uint8Array(await kptmFile.arrayBuffer()));

    let bwaPathArg = "None";
    let bwaSheetArg = "None";
    const bwaSheetSelect = document.getElementById("bwa-sheet");
    if (bwaFile && bwaSheetSelect.value) {
      const bwaBytes = new Uint8Array(await bwaFile.arrayBuffer());
      pyodide.FS.writeFile("/bwa.xlsx", bwaBytes);
      bwaPathArg = `"/bwa.xlsx"`;
      bwaSheetArg = JSON.stringify(bwaSheetSelect.value);
    }

    // Optional: Werkstattbestand fuer die Bestandsveraenderung UFE.
    // Der Anfangsbestand kommt aus der hinterlegten Jahreskonstante, eine PWBS-Datei
    // zum Periodenbeginn ist daher nur ein Zusatz.
    let pwbsAArg = "None";
    let pwbsEArg = "None";
    let pfakArg = "None";
    if (pwbsEndeFile) {
      const extE = pwbsEndeFile.name.toLowerCase().endsWith(".xls") ? ".xls" : ".xlsx";
      pyodide.FS.writeFile("/pwbs_ende" + extE, new Uint8Array(await pwbsEndeFile.arrayBuffer()));
      pwbsEArg = JSON.stringify("/pwbs_ende" + extE);
      const pfakPfade = [];
      for (let i = 0; i < pfakFiles.length; i++) {
        const p = `/pfak_${i}.xlsx`;
        pyodide.FS.writeFile(p, new Uint8Array(await pfakFiles[i].arrayBuffer()));
        pfakPfade.push(p);
      }
      pfakArg = pfakPfade.length ? JSON.stringify(pfakPfade) : "None";
      if (pwbsAnfangFile) {
        const extA = pwbsAnfangFile.name.toLowerCase().endsWith(".xls") ? ".xls" : ".xlsx";
        pyodide.FS.writeFile("/pwbs_anfang" + extA, new Uint8Array(await pwbsAnfangFile.arrayBuffer()));
        pwbsAArg = JSON.stringify("/pwbs_anfang" + extA);
      }
    }

    letzterLauf = { kptmPfad, bwaPathArg, bwaSheetArg, pwbsAArg, pwbsEArg, pfakArg };
    manuelleZuordnung = {};   // neuer Lauf mit neuen Dateien: alte Zuordnungen verwerfen
    await berechnen();
  } catch (err) {
    showError("Fehler bei der Verarbeitung: " + err);
  } finally {
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

async function berechnen() {
  const L = letzterLauf;
  const manuellArg = Object.keys(manuelleZuordnung).length
    ? JSON.stringify(manuelleZuordnung) : "None";
  const resultJson = await pyodide.runPythonAsync(`
import json
summary = spartenrechnung_core.generate(
    ${JSON.stringify(L.kptmPfad)}, "/config", "/output.xlsx",
    bwa_path=${L.bwaPathArg}, bwa_sheet=${L.bwaSheetArg},
    pwbs_anfang=${L.pwbsAArg}, pwbs_ende=${L.pwbsEArg}, pfak_pfade=${L.pfakArg},
    manuelle_zuordnung=${manuellArg},
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

function showResult(summary, url, filename) {
  let html = "<h2>✅ Spartenrechnung erstellt</h2>";
  html += `<div class="summary-line"><span>Rohdaten-Zeilen verarbeitet</span><span>${summary.zeilen.toLocaleString("de-DE")}</span></div>`;
  html += `<div class="summary-line"><span>Erlöse (Summe der Sparten)</span><span>${summary.erloese_summe.toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2})} €</span></div>`;
  html += `<div class="summary-line"><span>DB III bottom-up (Summe der Sparten)</span><span>${summary.db3_summe.toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2})} €</span></div>`;
  if (summary.bwa_ergebnis !== null && summary.bwa_ergebnis !== undefined) {
    const delta = summary.db3_summe - summary.bwa_ergebnis;
    html += `<div class="summary-line"><span>Ergebnis lt. GuV (BWA, kumuliert)</span><span>${summary.bwa_ergebnis.toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2})} €</span></div>`;
    html += `<div class="summary-line"><span>Δ zu erklären</span><span>${delta.toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2})} €</span></div>`;
    html += `<p class="hinweis" style="margin-top:0.4rem">„Summe der Sparten" ist nicht das
      Gesamtunternehmensergebnis: Sachkonto-Buchungen ohne Produktgruppe (z.&nbsp;B.
      Grundstücksverkäufe, Abgrenzungen, periodenfremde Erträge) tragen keine Sparte und
      erscheinen in keinem KPTM-Export – die Δ-Zeile enthält deshalb auch diese Posten,
      nicht nur die Deckungsdifferenzen. Details im Hinweise-Blatt der Ausgabedatei.</p>`;
  }

  // Die aus den Rohdaten zurueckgerechneten Zuschlagssaetze. Frueher stand hier der
  // erwartete Satz als Vergleichswert - der gehoert nicht in oeffentlichen Quellcode.
  const saetze = summary.zuschlagssaetze || {};
  for (const [key, label] of [["mgk", "MGK"], ["vvgk", "VVGK"]]) {
    if (saetze[key] === undefined || saetze[key] === null) continue;
    const satz = (saetze[key] * 100).toLocaleString("de-DE", {minimumFractionDigits: 2, maximumFractionDigits: 2});
    html += `<div class="summary-line"><span>${label}-Zuschlagssatz (Probe – gegen die eigene Kalkulation prüfen)</span><span>${satz} %</span></div>`;
  }

  if (summary.unproduktive_gemeinkosten) {
    const betrag = summary.unproduktive_gemeinkosten.toLocaleString("de-DE", {minimumFractionDigits: 2, maximumFractionDigits: 2});
    const std = summary.unproduktive_stunden.toLocaleString("de-DE", {maximumFractionDigits: 0});
    html += `<div class="summary-line"><span>Nicht auf Kostenträger gebuchte Stunden (${std} h)</span><span>${betrag} €</span></div>`;
  }

  if (summary.mgk_hinweise && summary.mgk_hinweise.length) {
    html += `<div class="warning-box">⚠️ ${summary.mgk_hinweise.join(" ")}</div>`;
  }

  if (summary.ufe_warnungen && summary.ufe_warnungen.length) {
    html += `<div class="warning-box">⚠️ Bestandsveränderung UFE wurde <strong>nicht</strong> berechnet: `;
    html += summary.ufe_warnungen.join(" ");
    html += `</div>`;
  }

  const stat = summary.ufe_statistik;
  if (stat) {
    const eur = (v) => v.toLocaleString("de-DE", {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " €";
    const pct = (v) => (v * 100).toLocaleString("de-DE", {minimumFractionDigits: 1, maximumFractionDigits: 1}) + " %";
    const zugeordnetQuote = stat.gesamt ? stat.zugeordnet / stat.gesamt : 0;
    html += `<h2 style="margin-top:1.5rem">Werkstattbestand – Zuordnung</h2>`;
    html += `<div class="summary-line"><span>einer Sparte zugeordnet</span><span>${eur(stat.zugeordnet)} · ${pct(zugeordnetQuote)}</span></div>`;
    html += `<div class="summary-line"><span>Serviceaufträge (bewusst ausgeschlossen)</span><span>${eur(stat.service)} · ${pct(stat.service_quote)}</span></div>`;
    html += `<div class="summary-line"><span>nicht zuordenbar</span><span>${eur(stat.offen)} · ${pct(stat.luecke_quote)}</span></div>`;
    html += `<div class="summary-line"><span><strong>Werkstattbestand gesamt</strong></span><span><strong>${eur(stat.gesamt)}</strong></span></div>`;

    // Warnungen, die die Zeile NICHT blockieren, aber die Belastbarkeit einzelner
    // Sparten betreffen. Die Prozentschwelle misst am Gesamtbestand und sagt darum
    // nichts ueber eine kleine Sparte aus.
    if (stat.hinweise && stat.hinweise.length) {
      html += `<div class="warning-box">⚠️ ${stat.hinweise.join(" ")}</div>`;
    }
    if (stat.ueber_artikel > 0 || stat.service_mit_quelle > 0) {
      html += `<details class="zuordnung"><summary>Wie belastbar ist die Zuordnung?</summary>`;
      html += `<div class="summary-line"><span>über die Auftragsnummer (sicherster Weg)</span>
        <span>${eur(stat.zugeordnet - stat.ueber_artikel)}</span></div>`;
      html += `<div class="summary-line"><span>nur über die Artikelnummer (schwächer)</span>
        <span>${eur(stat.ueber_artikel)}</span></div>`;
      if (stat.mehrdeutige_artikel) {
        html += `<div class="summary-line"><span>Artikel ohne eindeutige Sparte (nicht nutzbar)</span>
          <span>${stat.mehrdeutige_artikel}</span></div>`;
      }
      html += `<div class="summary-line"><span>Serviceaufträge, für die eine Sparte bekannt wäre</span>
        <span>${eur(stat.service_mit_quelle)}</span></div>`;
      html += `<p class="hinweis">Der Weg über die <strong>Artikelnummer</strong> ist schwächer:
        Wird ein Artikel in einem späteren Export mehrdeutig, entfällt die Zuordnung – mehr
        Daten können das Ergebnis dort also verschlechtern. Die genannten Serviceaufträge
        werden bewusst abgezogen (es wird nichts gefertigt); das ist gegen die
        Referenzrechnung Q1 2026 geprüft, der Betrag wächst aber.</p></details>`;
    }

    const bereits = Object.keys(manuelleZuordnung).length;
    if (bereits) {
      html += `<div class="summary-line"><span>davon von Hand zugeordnet</span><span>${bereits} Aufträge</span></div>`;
    }

    // Der Bestand zum Stichtag ist der Anfangsbestand der Folgeperiode. Faellt der
    // Stichtag auf den 01.01., ist es die Jahreskonstante des naechsten Jahres -
    // dann muss sie nicht von Hand gepflegt werden, sondern faellt hier heraus.
    if (stat.bestand_stichtag && Object.keys(stat.bestand_stichtag).length) {
      html += `<details class="zuordnung"><summary>Werkstattbestand als Betriebsparameter sichern</summary>`;
      html += `<p class="hinweis">Der hier errechnete Bestand ist zugleich der
        <strong>Anfangsbestand der Folgeperiode</strong>. Wenn Sie den Abschluss zum
        <strong>01.01.</strong> rechnen, ist das die Jahreskonstante des neuen Jahres –
        speichern, Jahreszahl in der Tabelle prüfen, fertig. Dann muss nichts von Hand
        gepflegt werden.</p>`;
      html += `<button type="button" id="param-sichern">⬇ betriebsparameter.xlsx</button></details>`;
    }

    if (stat.offene_auftraege && stat.offene_auftraege.length) {
      const optionen = (summary.produktgruppen_auswahl || [])
        .map((p) => `<option value="${p}">${p}</option>`).join("");
      html += `<details class="zuordnung" id="zuordnung-block"${bereits ? " open" : ""}><summary>`;
      html += `${stat.offene_auftraege.length} Aufträge selbst zuordnen `;
      html += `<span class="meta-label">(nach Wert sortiert)</span></summary>`;
      html += `<p class="hinweis">Ordnen Sie die größten Posten zu, bis die Lücke unter 1 % fällt –
        dann rechnet die UFE-Zeile. Zugeordnete Aufträge verschwinden aus dieser Liste und bleiben
        gespeichert, solange die Seite geöffnet ist.</p>`;
      html += `<table class="zuordnung-tabelle"><thead><tr>
        <th>Auftrag</th><th>Bezeichnung</th><th style="text-align:right">Offener Wert</th><th>Sparte</th>
        </tr></thead><tbody>`;
      for (const a of stat.offene_auftraege.slice(0, 60)) {
        html += `<tr><td>${a.auftrag ?? ""}</td><td>${a.bezeichnung}</td>`;
        html += `<td style="text-align:right">${eur(a.wert)}</td>`;
        html += `<td><select class="zuordnung-select" data-auftrag="${a.auftrag ?? ""}">
          <option value="">– offen –</option>${optionen}</select></td></tr>`;
      }
      html += `</tbody></table>`;
      if (stat.offene_auftraege.length > 60) {
        html += `<p class="hinweis">Angezeigt werden die 60 größten von ${stat.offene_auftraege.length} Aufträgen.</p>`;
      }
      html += `<button type="button" id="neu-berechnen">Mit dieser Zuordnung neu berechnen</button>`;
      html += `</details>`;
    }
  }

  // Per Auffangregel zugeordnete Konten ausweisen - sie sind korrekt verbucht,
  // aber niemand hat sie einzeln geprueft. Das gehoert vor Augen, nicht ins Log.
  if (summary.fracht_ohne_ziel && Object.keys(summary.fracht_ohne_ziel).length) {
    const eurF = (v) => (v || 0).toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2});
    const teile = Object.entries(summary.fracht_ohne_ziel).map(([k, v]) => `${k} (${eurF(v)} €)`);
    html += `<div class="warning-box">⚠️ Ausgangsfracht konnte nicht auf ein Frachtkostenobjekt
      umgebucht werden und blieb auf der operativen Sparte: <strong>${teile.join(", ")}</strong>.
      Es fehlt das Kostenobjekt &lt;Familie&gt;0100. Der Deckungsbeitrag dieser Sparte ist
      dadurch zu niedrig.</div>`;
  }

  if (summary.per_regel_zugeordnet && summary.per_regel_zugeordnet.length) {
    html += `<div class="warning-box">ℹ️ Nach Kontenregel zugeordnet (nicht einzeln
      hinterlegt): <strong>${summary.per_regel_zugeordnet.join(", ")}</strong>.
      Alle 6er-Konten gehen auf „sbA". Bitte kurz prüfen, ob das für diese Konten
      stimmt – falls nicht, in <code>config/kostenart_mapping.json</code> einzeln
      eintragen.</div>`;
  }

  if (summary.unbekannte_produktgruppen.length || summary.unbekannte_kostenarten.length) {
    html += `<div class="warning-box">⚠️ `;
    if (summary.unbekannte_produktgruppen.length) {
      html += `Unbekannte Produktgruppen gefunden (nicht in config/produktgruppen.json, ggf. Reorganisation): <strong>${summary.unbekannte_produktgruppen.join(", ")}</strong>. `;
    }
    if (summary.unbekannte_kostenarten.length) {
      // Mit Betrag: erst der sagt, wie dringend es ist.
      const b = summary.unmapped_betrag || {};
      const eurF = (v) => (v || 0).toLocaleString("de-DE", {minimumFractionDigits:2, maximumFractionDigits:2});
      const teile = summary.unbekannte_kostenarten.map((k) => `${k} (${eurF(b[k])} €)`);
      const summe = Object.values(b).reduce((a, v) => a + v, 0);
      html += `Kostenarten ohne Mapping – diese Beträge <strong>fehlen im Ergebnis</strong>:
        <strong>${teile.join(", ")}</strong>, zusammen ${eurF(summe)} €. `;
    }
    html += `Siehe README für Details.</div>`;
  }

  html += `<a class="download-btn" href="${url}" download="${filename}">⬇ ${filename} herunterladen</a>`;
  html += `<div class="legend">
    <span class="status-badge badge-gruen">automatisch</span>
    <span class="status-badge badge-gelb">Näherung</span>
    <span class="status-badge badge-rot">manuell zu ergänzen</span>
    – Details siehe Hinweise-Sheet in der Datei.
  </div>`;

  resultBox.innerHTML = html;
  resultBox.style.display = "block";

  const sichernBtn = document.getElementById("param-sichern");
  if (sichernBtn) {
    sichernBtn.addEventListener("click", async () => {
      // Geschaeftsjahr aus den Rohdaten (Spalte 'Geschaeftsjahr'), nicht aus Text geraten.
      const jahr = summary.jahr || new Date().getFullYear();
      const werte = JSON.stringify({ [String(jahr)]: summary.ufe_statistik.bestand_stichtag });
      await pyodide.runPythonAsync(`
import json
spartenrechnung_core.schreibe_betriebsparameter_excel(
    json.loads(${JSON.stringify(werte)}), "/betriebsparameter_neu.xlsx")
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
        showError("Fehler bei der Neuberechnung: " + err);
      }
    });
  }
  resultBox.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showError(msg) {
  errorBox.textContent = "❌ " + msg;
  errorBox.style.display = "block";
}

init();
