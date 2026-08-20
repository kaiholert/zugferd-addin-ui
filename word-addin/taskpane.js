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
let kositAvailable   = false;

// ── DOM-Referenzen ────────────────────────────────────────────────────────────
const serverBanner  = document.getElementById('serverBanner');
const serverMsg     = document.getElementById('serverMsg');
const exportBtn     = document.getElementById('exportBtn');
const progressDiv   = document.getElementById('progress');
const progressMsg   = document.getElementById('progressMsg');
const resultDiv     = document.getElementById('result');
const previewLoad   = document.getElementById('previewLoading');
const previewCont   = document.getElementById('previewContent');
const kositBtn      = document.getElementById('kositBtn');
const kositResultDiv       = document.getElementById('kositResult');
const kositUnavailableDiv  = document.getElementById('kositUnavailable');

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

  // KoSIT-Validierung Button
  kositBtn.addEventListener('click', runKositValidation);
  checkKositStatus();

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

      // Berechnete Summen + Labels in Content Controls zurueckschreiben
      // Rechnungsnummer und Empfaenger aus Detail-Controls extrahieren
      const rechnungsNr  = extractValue(cc['rechnung_nummer']   || '');
      const empfFirma    = cc['empfaenger_firma'] || '';

      // Rechnungsdatum/Lieferdatum/Faelligkeitsdatum automatisch aus
      // Leistungsmonat (manuell) + Zahlungsziel (manuell) berechnen:
      //   Rechnungsdatum = Lieferdatum = Monatsletzter des Leistungsmonats
      //   Faelligkeitsdatum = Rechnungsdatum + Zahlungsziel (Tage)
      const { rechnungsDatumBerechnet, lieferdatumBerechnet, faelligkeitsdatumBerechnet } =
        berechneDatumsfelder(cc['leistungsmonat'], cc['zahlungsziel']);

      // Fuer den Zahlungssatz: berechnetes Faelligkeitsdatum bevorzugen,
      // sonst (falls Leistungsmonat noch nicht auswertbar) alten Wert nutzen
      const faelligkeit = faelligkeitsdatumBerechnet || extractValue(cc['faelligkeitsdatum'] || '');

      const summenMap = {
        // Summen
        'summe_netto':    fmtDE(nettoGesamt),
        'summe_mwst_19':  fmtDE(mwst19.steuer),
        'summe_mwst_7':   fmtDE(mwst7.steuer),
        'summe_mwst_0':   fmtDE(mwst0.steuer),
        'summe_brutto':   fmtDE(bruttoGesamt),
        // MwSt-Labels mit aktuellen Basisbetragen
        'label_mwst_19':  `MwSt. 19 % (auf ${fmtDE(mwst19.basis)}):`,
        'label_mwst_7':   `MwSt. 7 % (auf ${fmtDE(mwst7.basis)}):`,
        'label_mwst_0':   `MwSt. 0 % (auf ${fmtDE(mwst0.basis)}):`,
        // Titel: Rechnungsnummer aus Detail uebernehmen
        'rechnung_nummer_titel': `Rechnung ${rechnungsNr}`,
        // Zahlungsblock: Verwendungszweck
        'zahlung_verwendungszweck': `Verwendungszweck:  ${rechnungsNr} / ${empfFirma}`,
        // Zahlungsblock: Zahlungsaufforderungs-Satz
        'zahlung_satz': `Bitte überweisen Sie den Rechnungsbetrag von ${fmtDE(bruttoGesamt)} bis zum ${faelligkeit} auf folgendes Konto:`,
      };

      // Rechnungsdatum/Lieferdatum/Faelligkeitsdatum nur ueberschreiben,
      // wenn sie sich aus dem Leistungsmonat/Zahlungsziel berechnen liessen
      // (sonst bleibt der bisherige Inhalt der Controls unangetastet)
      if (rechnungsDatumBerechnet)    summenMap['rechnung_datum']      = `Rechnungsdatum:${rechnungsDatumBerechnet}`;
      if (lieferdatumBerechnet)       summenMap['lieferdatum']         = `Lieferdatum:${lieferdatumBerechnet}`;
      if (faelligkeitsdatumBerechnet) summenMap['faelligkeitsdatum']   = `Fälligkeitsdatum:${faelligkeitsdatumBerechnet}`;

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
        rechnung_datum:    rechnungsDatumBerechnet    || extractValue(cc['rechnung_datum']    || ''),
        lieferdatum:       lieferdatumBerechnet       || extractValue(cc['lieferdatum']       || ''),
        faelligkeitsdatum: faelligkeitsdatumBerechnet || extractValue(cc['faelligkeitsdatum'] || ''),
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
        summe_mwst_0:   fmtDE(mwst0.steuer),   // Steuerbetrag 0%
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

  // KoSIT-Pruefung: aktiv sobald Dokument geladen, Server erreichbar und Validator eingerichtet
  kositBtn.disabled = !(kositAvailable && serverReachable && invoiceData && invoiceData.rechnung_nummer);
}

// ─────────────────────────────────────────────────────────────────────────────
// KoSIT-Validierung
// ─────────────────────────────────────────────────────────────────────────────
async function checkKositStatus() {
  try {
    const res    = await fetch(`${SERVER_URL}/validate-status`, { signal: AbortSignal.timeout(5000) });
    const status = await res.json();
    kositAvailable = !!status.available;

    if (!kositAvailable) {
      kositUnavailableDiv.style.display = 'block';
      kositUnavailableDiv.innerHTML =
        'KoSIT-Validierung nicht eingerichtet:<ul>' +
        (status.reasons || []).map(r => `<li>${escapeHtml(r)}</li>`).join('') +
        '</ul>Siehe setup-kosit-validator.bat im Projektordner.';
    } else {
      kositUnavailableDiv.style.display = 'none';
      kositUnavailableDiv.innerHTML = '';
    }
  } catch {
    kositAvailable = false;
    kositUnavailableDiv.style.display = 'block';
    kositUnavailableDiv.textContent = 'KoSIT-Status konnte nicht abgefragt werden (Server nicht erreichbar).';
  }
  updateExportButton();
}

async function runKositValidation() {
  if (!invoiceData) return;

  kositBtn.disabled = true;
  const prevLabel = kositBtn.textContent;
  kositBtn.textContent = '⏳ Prüfe...';

  try {
    const response = await fetch(`${SERVER_URL}/validate`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ profile: selectedProfile, invoice: invoiceData }),
    });
    const json = await response.json();

    if (json.success) {
      renderKositResult(json.validation);
    } else {
      renderKositResult({ ran: false, valid: null, reason: json.error || 'Unbekannter Fehler' });
    }
  } catch (err) {
    renderKositResult({ ran: false, valid: null, reason: 'Anfrage fehlgeschlagen: ' + err.message });
  } finally {
    kositBtn.textContent = prevLabel;
    updateExportButton();
  }
}

/**
 * Zeigt das Ergebnis einer KoSIT-Validierung im #kositResult Panel an.
 * validation: { ran, valid, errorCount, warningCount, messages, reason, ... }
 */
function renderKositResult(validation) {
  kositResultDiv.style.display = 'block';

  if (!validation || validation.ran === false) {
    kositResultDiv.className = 'off';
    kositResultDiv.innerHTML =
      `<div class="kosit-title">ⓘ KoSIT-Prüfung nicht durchgeführt</div>` +
      escapeHtml((validation && validation.reason) || 'Unbekannter Grund');
    return;
  }

  if (validation.valid === null) {
    kositResultDiv.className = 'warn';
    kositResultDiv.innerHTML =
      `<div class="kosit-title">⚠ KoSIT-Prüfung ohne eindeutiges Ergebnis</div>` +
      escapeHtml(validation.reason || 'Report konnte nicht ausgewertet werden.');
    return;
  }

  const msgs = validation.messages || [];
  const MAX_SHOWN = 20;

  if (validation.valid) {
    kositResultDiv.className = 'ok';
    kositResultDiv.innerHTML =
      `<div class="kosit-title">✓ KoSIT: XML ist gültig</div>` +
      (validation.warningCount ? `${validation.warningCount} Warnung(en)` : 'Keine Fehler oder Warnungen.');
  } else {
    kositResultDiv.className = 'fail';
    kositResultDiv.innerHTML =
      `<div class="kosit-title">✗ KoSIT: XML ist NICHT gültig</div>` +
      `${validation.errorCount} Fehler, ${validation.warningCount} Warnung(en)`;
  }

  if (msgs.length) {
    const shown = msgs.slice(0, MAX_SHOWN);
    kositResultDiv.innerHTML += '<ul>' + shown.map(m => `<li>${
      m.severity === 'error' ? '✗' : (m.severity === 'warning' ? '⚠' : 'ⓘ')
    } ${escapeHtml(m.message)}${m.location ? `<span class="kosit-loc">${escapeHtml(m.location)}</span>` : ''}</li>`).join('') + '</ul>';
    if (msgs.length > MAX_SHOWN) {
      kositResultDiv.innerHTML += `<div class="kosit-more">+ ${msgs.length - MAX_SHOWN} weitere Meldung(en)</div>`;
    }
  }
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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
      if (json.validation) renderKositResult(json.validation);
    } else {
      showError('Server-Fehler: ' + json.error);
      if (json.validation) renderKositResult(json.validation);
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

// ─────────────────────────────────────────────────────────────────────────────
// Rechnungsdatum / Lieferdatum / Faelligkeitsdatum aus Leistungsmonat +
// Zahlungsziel berechnen
// ─────────────────────────────────────────────────────────────────────────────

const MONATE_DE = {
  januar: 0, februar: 1, 'märz': 2, maerz: 2, april: 3, mai: 4, juni: 5,
  juli: 6, august: 7, september: 8, oktober: 9, november: 10, dezember: 11,
};

// "Leistungsmonat:April 2026" -> extractValue() -> "April 2026" -> {jahr, monatIndex}
function parseLeistungsmonat(text) {
  const wert = extractValue(text || '');
  const m = wert.trim().match(/^(\p{L}+)\s+(\d{4})$/u);
  if (!m) return null;
  const monatIndex = MONATE_DE[m[1].toLowerCase()];
  if (monatIndex === undefined) return null;
  return { jahr: parseInt(m[2], 10), monatIndex };
}

// Letzter Tag eines Monats: Tag 0 des Folgemonats
function letzterTagDesMonats(jahr, monatIndex) {
  return new Date(jahr, monatIndex + 1, 0);
}

function addTage(datum, tage) {
  const d = new Date(datum);
  d.setDate(d.getDate() + tage);
  return d;
}

function fmtDatumDE(d) {
  const tt = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${tt}.${mm}.${d.getFullYear()}`;
}

/**
 * Berechnet Rechnungsdatum, Lieferdatum und Faelligkeitsdatum:
 *   Rechnungsdatum = Lieferdatum = Monatsletzter des Leistungsmonats
 *   Faelligkeitsdatum = Rechnungsdatum + Zahlungsziel (Tage, aus z.B. "30 Tage netto")
 * Liefert '' fuer Felder, die sich nicht berechnen lassen (z.B. Leistungsmonat
 * leer/nicht auswertbar) - in dem Fall bleibt das jeweilige Control unangetastet.
 */
function berechneDatumsfelder(leistungsmonatText, zahlungszielText) {
  const leistungsmonat = parseLeistungsmonat(leistungsmonatText);
  if (!leistungsmonat) {
    console.warn('[ZUGFeRD] Leistungsmonat konnte nicht geparst werden:', leistungsmonatText);
    return { rechnungsDatumBerechnet: '', lieferdatumBerechnet: '', faelligkeitsdatumBerechnet: '' };
  }

  const monatsletzter = letzterTagDesMonats(leistungsmonat.jahr, leistungsmonat.monatIndex);
  const rechnungsDatumBerechnet = fmtDatumDE(monatsletzter);
  const lieferdatumBerechnet    = rechnungsDatumBerechnet;

  const zahlungszielWert = extractValue(zahlungszielText || '');
  const tageMatch        = zahlungszielWert.match(/(\d+)/);
  let faelligkeitsdatumBerechnet = '';
  if (tageMatch) {
    faelligkeitsdatumBerechnet = fmtDatumDE(addTage(monatsletzter, parseInt(tageMatch[1], 10)));
  } else {
    console.warn('[ZUGFeRD] Zahlungsziel konnte nicht geparst werden:', zahlungszielText);
  }

  return { rechnungsDatumBerechnet, lieferdatumBerechnet, faelligkeitsdatumBerechnet };
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
