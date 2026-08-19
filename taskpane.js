'use strict';

/* ═══════════════════════════════════════════════════════════════════════════
   ZUGFeRD Word Add-in – Task Pane Logik
   
   Ablauf:
   1. Office.onReady → Content Controls lesen → Vorschau befüllen
   2. Nutzer wählt Profil + klickt "Export"
   3. Add-in exportiert das Dokument als PDF (via Word API)
   4. PDF + Rechnungsdaten → POST an localhost:3737/generate
   5. Server antwortet mit Dateipfad → Erfolgsmeldung
═══════════════════════════════════════════════════════════════════════════ */

const SERVER_URL = 'http://127.0.0.1:3737';
const SERVER_CHECK_INTERVAL = 10_000; // ms

// ── State ─────────────────────────────────────────────────────────────────────
let selectedProfile  = 'EN16931';
let invoiceData      = null;
let serverReachable  = false;

// ── DOM-Referenzen ────────────────────────────────────────────────────────────
const serverBanner  = document.getElementById('serverBanner');
const serverMsg     = document.getElementById('serverMsg');
const exportBtn     = document.getElementById('exportBtn');
const progressDiv   = document.getElementById('progress');
const progressMsg   = document.getElementById('progressMsg');
const resultDiv     = document.getElementById('result');
const previewLoad   = document.getElementById('previewLoading');
const previewCont   = document.getElementById('previewContent');

// ── Office.js initialisieren ──────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    showError('Dieses Add-in funktioniert nur in Microsoft Word.');
    return;
  }

  // Profil-Karten
  document.querySelectorAll('.profile-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('.profile-card').forEach(c => c.classList.remove('selected'));
      card.classList.add('selected');
      selectedProfile = card.dataset.profile;
    });
  });

  // Export-Button
  exportBtn.addEventListener('click', runExport);

  // Neu berechnen Button
  document.getElementById('recalcBtn').addEventListener('click', async () => {
    const btn = document.getElementById('recalcBtn');
    btn.disabled = true;
    btn.textContent = '⏳ Berechne...';
    try {
      await loadInvoiceData();
    } finally {
      btn.disabled = false;
      btn.textContent = '↻ Neu berechnen';
    }
  });

  // Content Controls lesen
  loadInvoiceData();

  // Server-Status prüfen
  checkServer();
  setInterval(checkServer, SERVER_CHECK_INTERVAL);
});

// ─────────────────────────────────────────────────────────────────────────────
// Content Controls auslesen
// ─────────────────────────────────────────────────────────────────────────────
async function loadInvoiceData() {
  try {
    await Word.run(async context => {
      const controls = context.document.contentControls;
      controls.load('tag, text, type');
      await context.sync();

      // Alle Controls in ein Lookup-Objekt umwandeln
      const cc = {};
      for (const ctrl of controls.items) {
        if (ctrl.tag) {
          cc[ctrl.tag] = ctrl.text || '';
        }
      }

      // ── Positionen: direkt aus der Word-Tabelle lesen ──────────────
      // Tabelle mit 7 Spalten suchen (Pos/Beschr/Menge/Einheit/Preis/MwSt/Betrag)
      // Spaltenanzahl via erste Zeile / Zellanzahl ermitteln (columnCount nicht
      // direkt verfuegbar in Word JS API).
      const positions = [];
      const POS_TABLE_COLS = 7;

      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();

      // Alle Tabellen-Zeilen vorladen um Spaltenanzahl zu ermitteln
      for (const tbl of tables.items) {
        tbl.rows.load('items');
      }
      await context.sync();

      // Erste Zeile jeder Tabelle laden
      for (const tbl of tables.items) {
        if (tbl.rows.items.length > 0) {
          tbl.rows.items[0].cells.load('items');
        }
      }
      await context.sync();

      // Tabelle mit 7 Spalten finden
      let posTable = null;
      for (const tbl of tables.items) {
        const firstRow = tbl.rows.items[0];
        if (firstRow && firstRow.cells.items.length === POS_TABLE_COLS) {
          posTable = tbl;
          break;
        }
      }

      // Betrag-Zellen merken fuer späteres Zurückschreiben
      const betragZellen = []; // { cell, rowIdx }

      if (posTable) {
        const rows = posTable.rows.items;

        // Alle Zeilen-Zellen laden
        for (let ri = 1; ri < rows.length; ri++) {
          rows[ri].cells.load('items');
        }
        await context.sync();

        for (let ri = 1; ri < rows.length; ri++) {
          const cells = rows[ri].cells.items;
          if (cells.length < POS_TABLE_COLS) continue;
          for (const cell of cells) cell.load('value');
        }
        await context.sync();

        for (let ri = 1; ri < rows.length; ri++) {
          const cells = rows[ri].cells.items;
          if (cells.length < POS_TABLE_COLS) continue;

          const colText = cells.map(c => (c.value || '').trim());

          // Leere Zeilen ueberspringen
          const beschr = colText[1] || '';
          if (!beschr && !parseGermanFloat(colText[6])) continue;

          // Beschreibung: erste Zeile = Titel, weitere = Detail
          const lines = beschr.split('\n').map(l => l.trim()).filter(Boolean);

          positions.push({
            pos_nr:             colText[0] || String(ri),
            beschreibung:       beschr,
            beschreibung_titel: lines[0] || beschr,
            beschreibung_detail:lines.slice(1).join(' ') || '',
            menge:              colText[2] || '0',
            einheit:            colText[3] || 'Stk.',
            einzelpreis:        colText[4] || '0,00',
            mwst_satz:          colText[5] || '19 %',
            betrag:             colText[6] || '0,00',
          });

          // Betrag-Zelle (Spalte 6) fuer späteres Zurückschreiben merken
          betragZellen.push({ cell: cells[6], posIdx: positions.length - 1 });
        }
        console.log('[ZUGFeRD] Positionen gelesen:', positions.length);
      } else {
        console.warn('[ZUGFeRD] Positionstabelle (7 Spalten) nicht gefunden. Tabellen:', tables.items.length);
      }

      // ── Summen aus Positionen berechnen ─────────────────────────────
      // Betraege und MwSt aus den gelesenen Positionen ermitteln.
      // Dabei werden Menge x Einzelpreis pro Position berechnet,
      // sowie Netto, MwSt-Aufschluesselung und Brutto summiert.

      // Hilfsfunktion: deutschen Betrag zu Float
      function parseDE(s) {
        if (!s) return 0;
        return parseFloat(
          String(s).replace(/[^0-9,.-]/g, '').replace(/\./g, '').replace(',', '.')
        ) || 0;
      }

      // Hilfsfunktion: Float zu deutschem Betrag-String "1.234,56 €"
      function fmtDE(n) {
        return n.toLocaleString('de-DE', {
          minimumFractionDigits: 2, maximumFractionDigits: 2
        }) + ' €';
      }

      // Zeilenbetrag = Menge x Einzelpreis (Einzelpreis ohne €-Zeichen)
      let nettoGesamt = 0;
      const mwstGruppen = {}; // { '19': { basis: 0, steuer: 0 }, ... }

      for (let idx = 0; idx < positions.length; idx++) {
        const pos = positions[idx];
        const menge      = parseDE(pos.menge);
        const einzelpreis= parseDE(pos.einzelpreis);
        const zeilenbetrag = Math.round(menge * einzelpreis * 100) / 100;

        // Berechneten Betrag in Position eintragen
        pos.betrag = fmtDE(zeilenbetrag);

        // Betrag direkt in Tabellenzelle schreiben
        const zelle = betragZellen.find(z => z.posIdx === idx);
        if (zelle) {
          try {
            zelle.cell.body.insertText(fmtDE(zeilenbetrag), 'Replace');
          } catch(e) {
            console.warn('[ZUGFeRD] Betrag-Zelle nicht beschreibbar:', e.message);
          }
        }

        nettoGesamt += zeilenbetrag;

        // MwSt-Satz ermitteln: "19 %" -> 19
        const mwstMatch = String(pos.mwst_satz).match(/(\d+)/);
        const mwstSatz  = mwstMatch ? parseInt(mwstMatch[1], 10) : 0;
        const key = String(mwstSatz);
        if (!mwstGruppen[key]) mwstGruppen[key] = { satz: mwstSatz, basis: 0, steuer: 0 };
        mwstGruppen[key].basis  += zeilenbetrag;
        mwstGruppen[key].steuer += Math.round(zeilenbetrag * mwstSatz / 100 * 100) / 100;
      }

      nettoGesamt = Math.round(nettoGesamt * 100) / 100;

      const mwst19  = mwstGruppen['19'] || { basis: 0, steuer: 0 };
      const mwst7   = mwstGruppen['7']  || { basis: 0, steuer: 0 };
      const mwst0   = mwstGruppen['0']  || { basis: 0, steuer: 0 };

      const steuerGesamt = Math.round((mwst19.steuer + mwst7.steuer) * 100) / 100;
      const bruttoGesamt = Math.round((nettoGesamt + steuerGesamt) * 100) / 100;

      // Berechnete Summen in Content Controls zurueckschreiben
      const summenMap = {
        'summe_netto':   fmtDE(nettoGesamt),
        'summe_mwst_19': fmtDE(mwst19.steuer),
        'summe_mwst_7':  fmtDE(mwst7.steuer),
        'summe_mwst_0':  fmtDE(mwst0.basis),   // Steuerfreie = Nettobasis
        'summe_brutto':  fmtDE(bruttoGesamt),
      };

      // Betrag-Zellen in Tabelle synchronisieren
      await context.sync();

      // Summen-Controls aktualisieren
      for (const ctrl of controls.items) {
        if (ctrl.tag in summenMap) {
          try {
            ctrl.insertText(summenMap[ctrl.tag], 'Replace');
          } catch(e) {
            console.warn('[ZUGFeRD] Summe konnte nicht geschrieben werden:', ctrl.tag, e.message);
          }
        }
      }
      await context.sync();
      console.log('[ZUGFeRD] Summen berechnet:',
        'Netto', fmtDE(nettoGesamt),
        '| MwSt19', fmtDE(mwst19.steuer),
        '| MwSt7', fmtDE(mwst7.steuer),
        '| Brutto', fmtDE(bruttoGesamt)
      );

      // Rechnungsdaten zusammenstellen
      invoiceData = {
        // Rechnungsdetails
        rechnung_nummer:   extractValue(cc['rechnung_nummer']   || ''),
        rechnung_datum:    extractValue(cc['rechnung_datum']    || ''),
        lieferdatum:       extractValue(cc['lieferdatum']       || ''),
        faelligkeitsdatum: extractValue(cc['faelligkeitsdatum'] || ''),
        zahlungsziel:      extractValue(cc['zahlungsziel']      || ''),
        leistungsort:      extractValue(cc['leistungsort']      || ''),
        leistungsmonat:    extractValue(cc['leistungsmonat']    || ''),
        waehrung:          extractValue(cc['waehrung']          || 'EUR'),
        sprache:           extractValue(cc['sprache']      || 'Deutsch'),

        // Empfänger
        empfaenger_firma:           cc['empfaenger_firma']           || '',
        empfaenger_ansprechpartner: cc['empfaenger_ansprechpartner'] || '',
        empfaenger_strasse:         cc['empfaenger_strasse']         || '',
        empfaenger_plz_ort:         cc['empfaenger_plz_ort']         || '',
        empfaenger_land:            cc['empfaenger_land']            || 'Deutschland',
        empfaenger_land_code:       countryCode(cc['empfaenger_land'] || 'Deutschland'),
        empfaenger_ust_id:          extractEmpfaengerValue(cc['empfaenger_ust_id'] || '', ['USt-ID Empfänger:', 'USt-ID Empfaenger:']),
        empfaenger_kundennr:        extractEmpfaengerValue(cc['empfaenger_kundennr'] || '', ['Kunden-Nr.:']),
        empfaenger_bestellnr:       extractEmpfaengerValue(cc['empfaenger_bestellnr'] || '', ['Ihre Bestellnummer:']),

        // Summen - aus Positions-Berechnung (oben berechnet)
        summe_netto:    fmtDE(nettoGesamt),
        summe_mwst_19:  fmtDE(mwst19.steuer),
        summe_mwst_7:   fmtDE(mwst7.steuer),
        summe_mwst_0:   fmtDE(mwst0.basis),
        summe_brutto:   fmtDE(bruttoGesamt),

        // Zahlung
        zahlung_betrag:            cc['zahlung_betrag']            || '',
        zahlung_faellig:           cc['zahlung_faellig']           || '',
        zahlung_verwendungszweck:  cc['zahlung_verwendungszweck']  || '',

        // Positionen
        positions,
      };

      updatePreview(invoiceData);
      updateExportButton();
    });
  } catch (err) {
    console.error('Content Controls lesen fehlgeschlagen:', err);
    previewLoad.textContent = 'Fehler beim Lesen des Dokuments: ' + err.message;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Vorschau aktualisieren
// ─────────────────────────────────────────────────────────────────────────────
function updatePreview(data) {
  previewLoad.style.display = 'none';
  previewCont.style.display = 'block';

  document.getElementById('pv-nr').textContent     = data.rechnung_nummer || '(leer)';
  document.getElementById('pv-datum').textContent  = data.rechnung_datum  || '(leer)';
  document.getElementById('pv-emp').textContent    = data.empfaenger_firma || '(leer)';
  document.getElementById('pv-pos').textContent    = `${data.positions.length} Position(en)`;
  document.getElementById('pv-netto').textContent  = data.summe_netto  || '–';
  document.getElementById('pv-brutto').textContent = data.summe_brutto || '–';
}

function updateExportButton() {
  const ready = serverReachable && invoiceData && invoiceData.rechnung_nummer;
  exportBtn.disabled = !ready;
  exportBtn.textContent = ready
    ? '⬇ ZUGFeRD PDF erstellen'
    : (serverReachable ? 'Dokument nicht bereit' : 'Server nicht erreichbar');

  // Neu berechnen: aktiv sobald Dokument geladen (unabhaengig vom Server)
  const recalcBtn = document.getElementById('recalcBtn');
  if (recalcBtn) recalcBtn.disabled = !(invoiceData && invoiceData.rechnung_nummer);
}

// ─────────────────────────────────────────────────────────────────────────────
// Export-Ablauf
// ─────────────────────────────────────────────────────────────────────────────
async function runExport() {
  if (!invoiceData) return;

  showProgress('Dokument wird als PDF exportiert…');
  hideResult();

  try {
    // ── Schritt 1: Rechnungsdaten nochmals aktuell lesen ──────────────────
    await loadInvoiceData();

    // ── Schritt 2: Dokument als PDF exportieren ───────────────────────────
    showProgress('PDF wird erzeugt…');
    const pdfBase64 = await exportDocumentAsPdf();

    // ── Schritt 3: An Server senden ───────────────────────────────────────
    showProgress('ZUGFeRD XML wird generiert und eingebettet…');

    // Dateiname: Leistungsmonat - Rechnungsnummer - Bestellnummer
    const nameParts = [
      invoiceData.leistungsmonat,
      invoiceData.rechnung_nummer,
      invoiceData.empfaenger_bestellnr,
    ].filter(p => p && p.trim());  // leere Teile weglassen
    const safeName = nameParts
      .join('-')
      .replace(/[/\\?%*:|"<>]/g, '_')  // ungueltige Zeichen ersetzen
      .replace(/\s+/g, '_')             // Leerzeichen durch Unterstrich
      .replace(/-+/g, '-')              // mehrfache Bindestriche bereinigen
      .replace(/^-|-$/g, '');           // fuehrende/abschliessende Bindestriche

    const response = await fetch(`${SERVER_URL}/generate`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        profile:   selectedProfile,
        pdfBase64: pdfBase64,
        filename:  safeName,
        invoice:   invoiceData,
      }),
    });

    const json = await response.json();

    if (json.success) {
      showSuccess(
        `✓ ZUGFeRD PDF erfolgreich erstellt!\n` +
        `Profil: ${selectedProfile}\n` +
        `Positionen: ${invoiceData.positions.length}`,
        json.filePath,
      );
    } else {
      showError('Server-Fehler: ' + json.error);
    }

  } catch (err) {
    console.error('Export fehlgeschlagen:', err);
    showError('Export fehlgeschlagen: ' + err.message);
  } finally {
    hideProgress();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Dokument als PDF exportieren (Office.js)
// ─────────────────────────────────────────────────────────────────────────────
async function exportDocumentAsPdf() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Pdf,
      { sliceSize: 65536 },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        const file   = result.value;
        const slices = [];
        let   sliceIdx = 0;

        function getNextSlice() {
          file.getSliceAsync(sliceIdx, sliceResult => {
            if (sliceResult.status === Office.AsyncResultStatus.Failed) {
              file.closeAsync();
              reject(new Error(sliceResult.error.message));
              return;
            }
            slices.push(sliceResult.value.data);
            sliceIdx++;
            if (sliceIdx < file.sliceCount) {
              getNextSlice();
            } else {
              file.closeAsync();
              // Alle Slices zusammensetzen → Base64
              const allBytes = mergeUint8Arrays(slices);
              const base64   = uint8ArrayToBase64(allBytes);
              resolve(base64);
            }
          });
        }

        getNextSlice();
      }
    );
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Server-Erreichbarkeit prüfen
// ─────────────────────────────────────────────────────────────────────────────
async function checkServer() {
  try {
    const res = await fetch(`${SERVER_URL}/ping`, {
      signal: AbortSignal.timeout(3000),
    });
    const json = await res.json();
    if (json.status === 'ok') {
      serverReachable = true;
      serverBanner.className = 'ok';
      serverMsg.textContent  = `Server erreichbar (v${json.version})`;
    } else {
      throw new Error('Unerwartete Antwort');
    }
  } catch {
    serverReachable = false;
    serverBanner.className = 'err';
    serverMsg.textContent  = 'Server nicht erreichbar – bitte start-server.bat starten';
  }
  updateExportButton();
}

// ─────────────────────────────────────────────────────────────────────────────
// Hilfsfunktionen
// ─────────────────────────────────────────────────────────────────────────────

function mergeUint8Arrays(arrays) {
  const totalLength = arrays.reduce((sum, a) => sum + a.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const arr of arrays) {
    result.set(arr, offset);
    offset += arr.length;
  }
  return result;
}

function uint8ArrayToBase64(bytes) {
  let binary = '';
  const len = bytes.length;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function parseGermanFloat(str) {
  if (!str) return 0;
  return parseFloat(str.replace(/[^\d,.-]/g, '').replace('.', '').replace(',', '.')) || 0;
}

// Land → ISO 3166-1 Alpha-2
function countryCode(name) {
  const map = {
    'Deutschland': 'DE', 'Österreich': 'AT', 'Schweiz': 'CH',
    'Frankreich': 'FR', 'Niederlande': 'NL', 'Belgien': 'BE',
    'Polen': 'PL', 'Italien': 'IT', 'Spanien': 'ES',
  };
  return map[name] || (name && name.length === 2 ? name.toUpperCase() : 'DE');
}

// ── UI-Hilfsfunktionen ────────────────────────────────────────────────────────
function showProgress(msg) {
  progressMsg.textContent   = msg;
  progressDiv.style.display = 'flex';
  exportBtn.disabled        = true;
}
function hideProgress() {
  progressDiv.style.display = 'none';
  updateExportButton();
}
function showSuccess(msg, filePath) {
  resultDiv.className   = 'success';
  resultDiv.innerHTML   = msg.replace(/\n/g, '<br>') +
    (filePath ? `<div class="path">${filePath}</div>` : '');
  resultDiv.style.display = 'block';
}
function showError(msg) {
  resultDiv.className     = 'error';
  resultDiv.textContent   = msg;
  resultDiv.style.display = 'block';
}
function hideResult() {
  resultDiv.style.display = 'none';
}

// Extrahiert Wert aus "Label:Wert" String - z.B. "Rechnungsnummer:RE-2024-0001" -> "RE-2024-0001"
function extractValue(text) {
  if (!text) return '';
  const colonIdx = text.indexOf(':');
  if (colonIdx === -1) return text.trim();
  return text.slice(colonIdx + 1).trim();
}

// Bereinigt Empfaenger-Felder mit Label-Prefix
function extractEmpfaengerValue(text, prefixes) {
  if (!text) return text;
  for (const prefix of prefixes) {
    if (text.startsWith(prefix)) return text.slice(prefix.length).trim();
  }
  return text.trim();
}
