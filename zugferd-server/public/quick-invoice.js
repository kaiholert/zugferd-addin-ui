'use strict';

/* ═══════════════════════════════════════════════════════════════════════════
   ZUGFeRD Schnellerfassung – Web-Formular
   Läuft im Browser (kein Office.js), spricht mit dem lokalen ZUGFeRD-Server.
═══════════════════════════════════════════════════════════════════════════ */

const SERVER_URL = ''; // gleicher Origin (Server liefert diese Seite selbst aus)

const MONATSNAMEN_DE = [
  'Januar', 'Februar', 'März', 'April', 'Mai', 'Juni',
  'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember',
];

const serverBanner = document.getElementById('serverBanner');
const serverMsg    = document.getElementById('serverMsg');
const submitBtn    = document.getElementById('submitBtn');
const resultDiv    = document.getElementById('result');
const form         = document.getElementById('quickForm');

let serverReachable = false;

async function checkServer() {
  try {
    const res  = await fetch(`${SERVER_URL}/ping`, { signal: AbortSignal.timeout(3000) });
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
  submitBtn.disabled = !serverReachable;
}

// "2026-04" (aus <input type="month">) -> "April 2026"
function monatEingabeZuDeutsch(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return '';
  const jahr = m[1];
  const monatIndex = parseInt(m[2], 10) - 1;
  const name = MONATSNAMEN_DE[monatIndex];
  return name ? `${name} ${jahr}` : '';
}

form.addEventListener('submit', async ev => {
  ev.preventDefault();
  hideResult();

  const rechnungNummer = document.getElementById('rechnungNummer').value.trim();
  const leistungsmonat = monatEingabeZuDeutsch(document.getElementById('leistungsmonat').value);
  const bestellnummer  = document.getElementById('bestellnummer').value.trim();
  const stunden        = document.getElementById('stunden').value.trim();

  if (!rechnungNummer || !leistungsmonat || !bestellnummer || !stunden) {
    showError('Bitte alle Felder ausfüllen.');
    return;
  }

  submitBtn.disabled  = true;
  const prevLabel     = submitBtn.textContent;
  submitBtn.textContent = '⏳ Word wird geöffnet...';

  try {
    const response = await fetch(`${SERVER_URL}/quick-invoice`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rechnung_nummer: rechnungNummer,
        leistungsmonat:  leistungsmonat,
        bestellnummer:   bestellnummer,
        stunden:         stunden,
      }),
    });
    const json = await response.json();

    if (json.success) {
      showSuccess(`✓ Word geöffnet und befüllt (${rechnungNummer}, ${leistungsmonat}). Bitte in Word prüfen, "Neu berechnen" klicken und exportieren.`);
    } else {
      showError('Fehler: ' + json.error);
    }
  } catch (err) {
    showError('Anfrage fehlgeschlagen: ' + err.message);
  } finally {
    submitBtn.disabled    = !serverReachable;
    submitBtn.textContent = prevLabel;
  }
});

function showSuccess(msg) {
  resultDiv.className     = 'success';
  resultDiv.textContent    = msg;
  resultDiv.style.display  = 'block';
}
function showError(msg) {
  resultDiv.className     = 'error';
  resultDiv.textContent   = msg;
  resultDiv.style.display = 'block';
}
function hideResult() {
  resultDiv.style.display = 'none';
}

checkServer();
setInterval(checkServer, 10_000);
