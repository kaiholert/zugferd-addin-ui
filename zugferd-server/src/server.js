'use strict';

/**
 * ZUGFeRD Lokaler Server
 * Läuft auf http://localhost:3737
 *
 * Endpunkte:
 *   GET  /ping           – Erreichbarkeitscheck vom Add-in
 *   POST /generate       – Empfängt JSON + PDF, liefert ZUGFeRD PDF
 */

const express   = require('express');
const cors      = require('cors');
const path      = require('path');
const fs        = require('fs');
const { execFile } = require('child_process');
const { generateZugferdXml, PROFILES } = require('./zugferd-xml');
const { embedZugferdXml }              = require('./pdf-embedder');
const { checkAvailability: checkKositAvailability, validateXml: validateWithKosit } = require('./kosit-validator');

// ── Konfiguration laden ───────────────────────────────────────────────────────
const CONFIG_PATH = path.join(__dirname, '..', 'config.json');
let config = {};
console.log(`[INFO] Suche config.json unter: ${CONFIG_PATH}`);
if (!fs.existsSync(CONFIG_PATH)) {
  console.error(`[FEHLER] Datei nicht vorhanden: ${CONFIG_PATH}`);
  console.error('[FEHLER] Bitte config.json im Ordner "zugferd-server" anlegen.');
  config = { seller: {}, port: 3737, outputDir: '' };
} else {
  try {
    const raw = fs.readFileSync(CONFIG_PATH, 'utf-8');
    config = JSON.parse(raw);
    console.log('[OK] config.json erfolgreich geladen');
  } catch (err) {
    console.error(`[FEHLER] config.json gefunden, aber ungueltiges JSON: ${err.message}`);
    console.error('[FEHLER] Bitte Syntax pruefen (fehlendes Komma, Anfuehrungszeichen etc.)');
    config = { seller: {}, port: 3737, outputDir: '' };
  }
}

const PORT       = config.port || 3737;
const OUTPUT_DIR = config.outputDir
  ? path.resolve(config.outputDir)
  : path.join(require('os').homedir(), 'Desktop');

// Ausgabeverzeichnis anlegen falls nicht vorhanden
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// ── Express-App ───────────────────────────────────────────────────────────────
const app = express();

// CORS: localhost + konfigurierte GitHub Pages URL erlauben
const ALLOWED_ORIGINS = [
  /^https?:\/\/(localhost|127\.0\.0\.1)/,          // lokaler Dev-Server
  /^https:\/\/[a-z0-9-]+\.github\.io$/,              // GitHub Pages (alle Subdomains)
  /^https:\/\/[a-z0-9-]+\.github\.io\/.*/,          // GitHub Pages mit Repo-Pfad
  ...(config.allowedOrigins || []).map(o => new RegExp('^' + o.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$')),
];

app.use(cors({
  origin: (origin, cb) => {
    // kein Origin = direkter Aufruf (z.B. curl, Postman) -> erlauben
    if (!origin) return cb(null, true);
    const allowed = ALLOWED_ORIGINS.some(pattern => pattern.test(origin));
    if (allowed) {
      cb(null, true);
    } else {
      console.warn(`[CORS] Abgelehnt: ${origin}`);
      console.warn(`[CORS] Erlaubt sind: localhost, *.github.io`);
      console.warn(`[CORS] Weitere Origins in config.json unter "allowedOrigins" eintragen`);
      cb(new Error(`CORS: Origin nicht erlaubt: ${origin}`));
    }
  },
  methods: ['GET', 'POST', 'OPTIONS'],
}));

// Body-Parser: JSON bis 50 MB (PDF als Base64 kann groß sein)
app.use(express.json({ limit: '50mb' }));

// Statische Dateien (z.B. quick-invoice.html – Schnellerfassung)
app.use(express.static(path.join(__dirname, '..', 'public')));

// ── Routen ────────────────────────────────────────────────────────────────────

/** Erreichbarkeitscheck */
app.get('/ping', (_req, res) => {
  res.json({ status: 'ok', version: '1.0.0' });
});

/**
 * GET /validate-status
 * Prüft ob die KoSIT-Validierung einsatzbereit ist (Konfiguration + Java + Dateien).
 * Wird vom Add-in genutzt um den "XML prüfen"-Button ein-/auszublenden.
 */
app.get('/validate-status', async (_req, res) => {
  try {
    const status = await checkKositAvailability(config.kositValidation || {});
    res.json(status);
  } catch (err) {
    res.status(500).json({ available: false, reasons: [err.message] });
  }
});

/**
 * POST /validate
 *
 * Erzeugt das ZUGFeRD-XML aus den übergebenen Rechnungsdaten und validiert es
 * per KoSIT Validator – unabhängig vom eigentlichen PDF-Export ("Nur prüfen").
 *
 * Request-Body: { profile: "EN16931", invoice: { ... } }
 * Response: { success: true, validation: { ran, valid, errorCount, warningCount, messages, ... } }
 */
app.post('/validate', async (req, res) => {
  try {
    const { profile: profileKey, invoice } = req.body;
    if (!invoice || !invoice.rechnung_nummer) {
      return res.status(400).json({ success: false, error: 'invoice.rechnung_nummer fehlt' });
    }

    invoice.seller = config.seller || {};
    const xmlString = generateZugferdXml(invoice, profileKey);

    const validation = await validateWithKosit(xmlString, config.kositValidation || {}, {
      filenameHint: invoice.rechnung_nummer,
    });

    res.json({ success: true, validation });
  } catch (err) {
    console.error('[FEHLER] /validate', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * POST /generate
 *
 * Request-Body (JSON):
 * {
 *   profile:   "EN16931",          // ZUGFeRD-Profil
 *   pdfBase64: "<base64-string>",  // PDF von Word (als Base64)
 *   filename:  "RE-2024-0001",     // Dateiname ohne Extension
 *   invoice: {
 *     rechnung_nummer: "RE-2024-0001",
 *     rechnung_datum:  "01.01.2025",
 *     // ... alle Content-Control-Werte
 *     positions: [
 *       { beschreibung_titel: "...", menge: "40,00", ... },
 *       ...
 *     ]
 *   }
 * }
 *
 * Response:
 *   200 JSON { success: true, filePath: "...", xmlPreview: "..." }
 *   400 JSON { success: false, error: "..." }
 */
app.post('/generate', async (req, res) => {
  try {
    const { profile: profileKey, pdfBase64, filename, invoice } = req.body;

    // ── Validierung ──────────────────────────────────────────────────────────
    if (!pdfBase64) {
      return res.status(400).json({ success: false, error: 'pdfBase64 fehlt' });
    }
    if (!invoice || !invoice.rechnung_nummer) {
      return res.status(400).json({ success: false, error: 'invoice.rechnung_nummer fehlt' });
    }

    const prof = PROFILES[profileKey] || PROFILES.EN16931;
    const safeName = (filename || invoice.rechnung_nummer).replace(/[/\\?%*:|"<>]/g, '-');

    // ── Verkäuferdaten aus config.json injizieren ────────────────────────────
    invoice.seller = config.seller || {};

    // ── ZUGFeRD XML generieren ───────────────────────────────────────────────
    const xmlString = generateZugferdXml(invoice, profileKey);

    // ── KoSIT-Validierung (optional, config.json → kositValidation) ──────────
    let validation = null;
    if (config.kositValidation && config.kositValidation.enabled) {
      try {
        validation = await validateWithKosit(xmlString, config.kositValidation, {
          filenameHint: safeName,
          persistDir:   OUTPUT_DIR,
        });
      } catch (err) {
        console.warn('[WARN] KoSIT-Validierung fehlgeschlagen:', err.message);
        validation = { ran: false, valid: null, reason: err.message };
      }

      if (validation.ran && validation.valid === false && config.kositValidation.blockOnInvalid) {
        return res.status(422).json({
          success: false,
          error:   'ZUGFeRD-XML ist laut KoSIT-Validator nicht gültig (siehe validation.messages).',
          validation,
        });
      }
    }

    // ── PDF dekodieren ───────────────────────────────────────────────────────
    const pdfBytes = Buffer.from(pdfBase64, 'base64');

    // ── XML in PDF/A-3 einbetten ─────────────────────────────────────────────
    const resultBytes = await embedZugferdXml(
      pdfBytes,
      xmlString,
      prof.urn,
      prof.level,
      invoice.rechnung_nummer,
    );

    // ── Datei speichern ──────────────────────────────────────────────────────
    const outPath = path.join(OUTPUT_DIR, `${safeName}_ZUGFeRD.pdf`);
    fs.writeFileSync(outPath, resultBytes);

    // XML-Debug-Datei (optional, zum Prüfen)
    const xmlPath = path.join(OUTPUT_DIR, `${safeName}_factur-x.xml`);
    fs.writeFileSync(xmlPath, xmlString, 'utf-8');

    console.log(`[OK] ${outPath}`);
    res.json({
      success:    true,
      filePath:   outPath,
      xmlPath:    xmlPath,
      xmlPreview: xmlString.slice(0, 500) + '...',
      validation,
    });

  } catch (err) {
    console.error('[FEHLER]', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * POST /quick-invoice
 *
 * Schnellerfassung fuer Standard-Rechnungen: erzeugt per Word-COM-Automation
 * (PowerShell) ein neues Dokument aus der .dotx-Vorlage und befuellt
 * Rechnungsnummer, Leistungsmonat, Bestellnummer sowie die Stunden-Menge der
 * (auf 1 Zeile reduzierten) Positionstabelle. Word wird gestartet falls nicht
 * bereits aktiv, sonst wird die laufende Instanz genutzt.
 *
 * Request-Body: { rechnung_nummer, leistungsmonat, bestellnummer, stunden }
 * Response: { success: true } oder { success: false, error: "..." }
 */
app.post('/quick-invoice', (req, res) => {
  const { rechnung_nummer, leistungsmonat, bestellnummer, stunden } = req.body || {};

  const fehlt = ['rechnung_nummer', 'leistungsmonat', 'bestellnummer', 'stunden']
    .filter(feld => !String(req.body && req.body[feld] || '').trim());
  if (fehlt.length) {
    return res.status(400).json({ success: false, error: `Pflichtfeld(er) fehlen: ${fehlt.join(', ')}` });
  }

  const stundenZahl = parseFloat(String(stunden).replace(',', '.'));
  if (!isFinite(stundenZahl) || stundenZahl <= 0) {
    return res.status(400).json({ success: false, error: 'Stunden muss eine Zahl größer 0 sein' });
  }

  const qc = config.quickInvoice || {};
  if (!qc.templatePath) {
    return res.status(400).json({ success: false, error: 'quickInvoice.templatePath fehlt in config.json' });
  }
  if (!fs.existsSync(qc.templatePath)) {
    return res.status(400).json({ success: false, error: `Vorlage nicht gefunden: ${qc.templatePath}` });
  }

  const scriptPath = path.join(__dirname, 'open-quick-invoice.ps1');
  const args = [
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
    '-RechnungNummer', String(rechnung_nummer),
    '-Leistungsmonat', String(leistungsmonat),
    '-Bestellnummer', String(bestellnummer),
    '-Stunden', String(stundenZahl),
    '-TemplatePfad', qc.templatePath,
  ];

  execFile('powershell', args, { timeout: qc.timeoutMs || 30000 }, (err, stdout, stderr) => {
    if (err) {
      console.error('[FEHLER] /quick-invoice', stderr || err.message);
      return res.status(500).json({ success: false, error: (stderr || err.message || '').trim() || 'PowerShell-Skript fehlgeschlagen' });
    }
    console.log('[OK] Schnellerfassung:', rechnung_nummer, leistungsmonat, bestellnummer, stundenZahl);
    res.json({ success: true });
  });
});

// ── Nicht gefundene Routen ────────────────────────────────────────────────────
app.use((_req, res) => res.status(404).json({ error: 'Nicht gefunden' }));

// ── Server starten ───────────────────────────────────────────────────────────
app.listen(PORT, '127.0.0.1', () => {
  console.log('═══════════════════════════════════════════════');
  console.log(`  ZUGFeRD Server läuft auf http://127.0.0.1:${PORT}`);
  console.log(`  Ausgabeverzeichnis: ${OUTPUT_DIR}`);
  console.log('═══════════════════════════════════════════════');
  console.log('  Bereit – Add-in kann jetzt Exporte starten.');
  console.log('  Beenden: Strg+C');
  console.log('═══════════════════════════════════════════════');
});

module.exports = app;
