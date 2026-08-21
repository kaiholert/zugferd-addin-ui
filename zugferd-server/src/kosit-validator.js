'use strict';

/**
 * KoSIT-Validierung für ZUGFeRD-XML
 * ──────────────────────────────────
 * Ruft den offiziellen KoSIT Validator (Java, https://github.com/itplr-kosit/validator)
 * als externen Prozess auf und prüft das erzeugte CII-XML gegen XML-Schema +
 * Schematron-Regeln (ZUGFeRD-Konfiguration, z. B.
 * https://github.com/LandrixSoftware/validator-configuration-zugferd).
 *
 * Läuft komplett lokal – keine Cloud-Abhängigkeit. Benötigt:
 *   - eine installierte Java-Laufzeit (JRE 11+)
 *   - den Validator (validator-<version>-standalone.jar)
 *   - eine Regel-Konfiguration (scenarios.xml + Schema/Schematron-Dateien)
 *
 * Setup: siehe setup-kosit-validator.bat im Projekt-Root.
 * Konfiguration: config.json → "kositValidation" (siehe PROJEKT_KONTEXT.md).
 */

const { spawn } = require('child_process');
const fs         = require('fs');
const path       = require('path');
const os         = require('os');
const { XMLParser } = require('fast-xml-parser');

/**
 * Prüft ob Java erreichbar ist (spawnt "java -version").
 */
function checkJavaAvailable(javaPath) {
  return new Promise(resolve => {
    let proc;
    try {
      proc = spawn(javaPath || 'java', ['-version'], { windowsHide: true });
    } catch {
      resolve(false);
      return;
    }
    proc.on('error', () => resolve(false));
    proc.on('exit', code => resolve(code === 0));
  });
}

/**
 * Prüft ob die KoSIT-Validierung einsatzbereit ist (Konfiguration + Dateien + Java).
 * Wird vom Add-in via GET /validate-status abgefragt, um den Button ein-/auszublenden.
 *
 * @param {object} cfg – config.json → kositValidation
 * @returns {Promise<{available: boolean, reasons: string[]}>}
 */
async function checkAvailability(cfg) {
  const reasons  = [];
  const javaPath = (cfg && cfg.javaPath) || 'java';
  const jarPath        = cfg && cfg.validatorJar;
  const scenariosPath  = cfg && cfg.scenariosXml;

  if (!cfg || cfg.enabled !== true) {
    reasons.push('KoSIT-Validierung ist deaktiviert (config.json → kositValidation.enabled = false).');
  }
  if (!jarPath || !fs.existsSync(jarPath)) {
    reasons.push(`Validator-JAR nicht gefunden: ${jarPath || '(nicht konfiguriert)'}`);
  }
  if (!scenariosPath || !fs.existsSync(scenariosPath)) {
    reasons.push(`scenarios.xml nicht gefunden: ${scenariosPath || '(nicht konfiguriert)'}`);
  }

  const javaOk = await checkJavaAvailable(javaPath);
  if (!javaOk) {
    reasons.push(`Java nicht gefunden (Pfad/Befehl: "${javaPath}"). Bitte JRE 11+ installieren.`);
  }

  return { available: reasons.length === 0, reasons };
}

/**
 * Startet den Validator-Prozess und wartet auf das Ende.
 */
function runValidator(javaPath, args, { cwd, timeoutMs }) {
  return new Promise((resolve, reject) => {
    let proc;
    try {
      proc = spawn(javaPath, args, { cwd, windowsHide: true });
    } catch (err) {
      reject(err);
      return;
    }

    let stdout = '';
    let stderr = '';
    let settled = false;

    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      proc.kill();
      reject(new Error(`KoSIT Validator hat nicht innerhalb von ${timeoutMs} ms geantwortet (Timeout).`));
    }, timeoutMs);

    proc.stdout.on('data', d => { stdout += d.toString('utf-8'); });
    proc.stderr.on('data', d => { stderr += d.toString('utf-8'); });
    proc.on('error', err => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      reject(err);
    });
    proc.on('exit', code => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, stdout, stderr });
    });
  });
}

/** Extrahiert reinen Text aus einem geparsten fast-xml-parser Knoten */
function extractText(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (Array.isArray(v)) return v.map(extractText).join(' ').trim();
  if (typeof v === 'object') {
    if (typeof v['#text'] === 'string') return v['#text'].trim();
    // Gemischter Inhalt (z.B. svrl:text mit verschachtelten Elementen) – alles einsammeln
    return Object.values(v).map(extractText).join(' ').trim();
  }
  return String(v).trim();
}

/** Normalisiert Schweregrad-Angaben ("fatal", "error", "warning", "information", ...) */
function normalizeSeverity(raw) {
  const s = String(raw || '').toLowerCase();
  if (s.includes('warn'))  return 'warning';
  if (s.includes('info'))  return 'info';
  if (s.includes('fatal') || s.includes('error') || s === '') return 'error';
  return 'error'; // unbekannter Wert -> sicherheitshalber als Fehler werten
}

/**
 * Durchsucht das geparste Report-XML rekursiv (Namespace-Präfixe wurden beim
 * Parsen bereits entfernt) nach dem tatsächlichen VARL-Reportformat des KoSIT
 * Validators 1.6.2 (empirisch anhand echter Reports ermittelt, nicht nur aus
 * generischer Schematron-/XOEV-Doku geraten):
 *
 *   - <rep:report valid="true|false">          – NICHT das Gesamturteil! Kann
 *     "false" sein, obwohl das Dokument als "ACCEPTABLE" durchgeht (z.B. bei
 *     reinen Warnungen). Wird hier bewusst NICHT als Verdikt verwendet.
 *   - <rep:assessment><rep:accept|reject>      – das tatsächliche, maßgebliche
 *     Verdikt (entspricht der "Acceptance"-Spalte der KoSIT-CLI-Ausgabe).
 *   - <rep:noScenarioMatched>                  – Guideline-ID/Dokumenttyp
 *     wurde von keinem Prüfszenario erkannt (führt zu <rep:reject>, aber ohne
 *     <rep:message>-Einträge – daher eigene Behandlung nötig).
 *   - <rep:scenarioMatched><s:scenario><s:name> – Name des erkannten Profils.
 *   - <rep:message level="error|warning|..." xpathLocation="..." code="...">
 *     TEXT</rep:message>  – die eigentlichen Schema-/Schematron-Befunde.
 *   - <failed-assert>/<successful-report> (klassisches SVRL) – zusätzlich als
 *     Fallback unterstützt, falls eine andere Validator-Konfiguration/Version
 *     rohes SVRL statt des rep:message-Formats liefert.
 */
function walkReport(node, ctx) {
  if (node == null) return;
  if (Array.isArray(node)) {
    for (const n of node) walkReport(n, ctx);
    return;
  }
  if (typeof node !== 'object') return;

  for (const key of Object.keys(node)) {
    const value = node[key];

    if (key === 'accept') {
      ctx.accepted = true;
    } else if (key === 'reject') {
      ctx.accepted = false;
    } else if (key === 'noScenarioMatched') {
      ctx.noScenarioMatched = true;
    } else if (key === 'scenario' && value && typeof value === 'object' && value.name && !ctx.scenarioName) {
      const name = extractText(value.name);
      if (name) ctx.scenarioName = name;
    } else if (key === 'message') {
      const items = Array.isArray(value) ? value : [value];
      for (const item of items) {
        const text = extractText(item);
        if (!text) continue;
        const level = item && typeof item === 'object' ? item['@_level'] : undefined;
        ctx.messages.push({
          severity: normalizeSeverity(level),
          message:  text,
          location: item && typeof item === 'object' ? (item['@_xpathLocation'] || item['@_xpath'] || item['@_location']) : undefined,
          code:     item && typeof item === 'object' ? item['@_code'] : undefined,
          source:   'schematron',
        });
      }
    } else if (key === 'failed-assert' || key === 'successful-report') {
      // Fallback fuer klassisches SVRL-Format (andere Validator-Versionen/Konfigurationen)
      const items = Array.isArray(value) ? value : [value];
      for (const item of items) {
        const flag = (item && (item['@_flag'] || item['@_role'])) || (key === 'failed-assert' ? 'error' : 'info');
        ctx.messages.push({
          severity: normalizeSeverity(flag),
          message:  extractText(item && item.text) || '(keine Meldung im Report)',
          location: item && item['@_location'],
          source:   'schematron',
        });
      }
    }

    if (value && typeof value === 'object') {
      walkReport(value, ctx);
    }
  }
}

/**
 * Validiert einen ZUGFeRD-XML-String gegen den KoSIT Validator.
 *
 * @param {string} xmlString  – das erzeugte ZUGFeRD-CII-XML
 * @param {object} cfg        – config.json → kositValidation
 * @param {object} [opts]
 * @param {string} [opts.filenameHint]  – Basisname für temporäre Dateien (z.B. Rechnungsnummer)
 * @param {string} [opts.persistDir]    – falls gesetzt: Report-XML dorthin kopieren
 * @param {boolean}[opts.keepWorkDir]   – Temp-Verzeichnis nicht löschen (Debugging)
 * @returns {Promise<object>}  Ergebnisobjekt, siehe unten
 */
async function validateXml(xmlString, cfg, opts = {}) {
  cfg = cfg || {};

  if (cfg.enabled !== true) {
    return { ran: false, valid: null, reason: 'KoSIT-Validierung ist deaktiviert (config.json → kositValidation.enabled = false).' };
  }

  const javaPath       = cfg.javaPath || 'java';
  const jarPath        = cfg.validatorJar;
  const scenariosPath  = cfg.scenariosXml;
  const timeoutMs      = cfg.timeoutMs || 30000;

  if (!jarPath || !fs.existsSync(jarPath)) {
    return { ran: false, valid: null, reason: `Validator-JAR nicht gefunden: ${jarPath || '(nicht konfiguriert)'}. Siehe setup-kosit-validator.bat.` };
  }
  if (!scenariosPath || !fs.existsSync(scenariosPath)) {
    return { ran: false, valid: null, reason: `scenarios.xml nicht gefunden: ${scenariosPath || '(nicht konfiguriert)'}. Siehe setup-kosit-validator.bat.` };
  }

  const repoDir = cfg.repositoryDir ? path.resolve(cfg.repositoryDir) : path.dirname(scenariosPath);
  const baseName = String(opts.filenameHint || 'invoice').replace(/[/\\?%*:|"<>]/g, '-') || 'invoice';

  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'zugferd-kosit-'));
  const xmlFile = path.join(workDir, `${baseName}.xml`);
  fs.writeFileSync(xmlFile, xmlString, 'utf-8');

  const args = ['-jar', jarPath, '-s', scenariosPath, '-r', repoDir, '-o', workDir, '-p', xmlFile];

  let run;
  try {
    run = await runValidator(javaPath, args, { cwd: workDir, timeoutMs });
  } catch (err) {
    if (!opts.keepWorkDir) safeRmDir(workDir);
    return { ran: false, valid: null, reason: `KoSIT Validator konnte nicht gestartet werden: ${err.message}` };
  }

  // Report-Datei suchen: Standardmuster "<basename>-report.xml", sonst neueste .xml im Arbeitsverzeichnis
  let reportPath = path.join(workDir, `${baseName}-report.xml`);
  if (!fs.existsSync(reportPath)) {
    const candidates = fs.readdirSync(workDir)
      .filter(f => f.toLowerCase().endsWith('.xml') && f !== `${baseName}.xml`)
      .map(f => path.join(workDir, f))
      .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
    reportPath = candidates[0];
  }

  if (!reportPath || !fs.existsSync(reportPath)) {
    const result = {
      ran: true,
      valid: null,
      reason: 'KoSIT Validator hat keinen auswertbaren Report erzeugt.',
      exitCode: run.code,
      stdout: run.stdout,
      stderr: run.stderr,
    };
    if (!opts.keepWorkDir) safeRmDir(workDir);
    return result;
  }

  const reportXml = fs.readFileSync(reportPath, 'utf-8');
  const parser = new XMLParser({
    ignoreAttributes:     false,
    attributeNamePrefix:  '@_',
    removeNSPrefix:        true,
    textNodeName:         '#text',
    trimValues:           true,
  });

  let parsed;
  try {
    parsed = parser.parse(reportXml);
  } catch (err) {
    if (!opts.keepWorkDir) safeRmDir(workDir);
    return {
      ran: true,
      valid: null,
      reason: `Report-XML konnte nicht geparst werden: ${err.message}`,
      exitCode: run.code,
      stdout: run.stdout,
    };
  }

  const ctx = { accepted: null, noScenarioMatched: false, scenarioName: null, messages: [] };
  walkReport(parsed, ctx);

  // Kein Prüfszenario erkannt (Guideline-ID unbekannt) => eigene, verständliche
  // Meldung, da der Report in diesem Fall keine <rep:message>-Einträge enthält.
  if (ctx.noScenarioMatched) {
    ctx.messages.push({
      severity: 'error',
      message:  'Kein passendes Prüfszenario gefunden – die Guideline-ID (Profil) im XML wird von der Validator-Konfiguration nicht erkannt.',
      source:   'scenario',
    });
  }

  const { messages } = ctx;
  const errorCount   = messages.filter(m => m.severity === 'error').length;
  const warningCount = messages.filter(m => m.severity === 'warning').length;
  // Massgeblich ist das <rep:accept>/<rep:reject>-Verdikt (entspricht der
  // "Acceptance"-Spalte der KoSIT-CLI), NICHT das rep:report[@valid]-Attribut -
  // das kann "false" sein obwohl das Dokument insgesamt akzeptiert wird
  // (z.B. bei reinen Warnungen ohne echte Fehler).
  const valid = ctx.accepted !== null ? ctx.accepted : errorCount === 0;

  let persistedReportPath = null;
  if (opts.persistDir) {
    try {
      fs.mkdirSync(opts.persistDir, { recursive: true });
      const destXml = path.join(opts.persistDir, `${baseName}_kosit-report.xml`);
      fs.writeFileSync(destXml, reportXml, 'utf-8');
      persistedReportPath = destXml;
    } catch {
      // Nicht kritisch – Validierungsergebnis bleibt trotzdem gültig
    }
  }

  if (!opts.keepWorkDir) safeRmDir(workDir);

  return {
    ran: true,
    valid,
    errorCount,
    warningCount,
    messages,
    scenarioName: ctx.scenarioName,
    reportXmlPath: persistedReportPath,
    exitCode: run.code,
  };
}

function safeRmDir(dir) {
  try {
    fs.rmSync(dir, { recursive: true, force: true });
  } catch {
    // Aufräumfehler ignorieren (Temp-Verzeichnis, kein kritischer Datenverlust)
  }
}

module.exports = { checkAvailability, validateXml };
