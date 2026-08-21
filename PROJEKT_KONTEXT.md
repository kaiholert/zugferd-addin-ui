# ZUGFeRD Word Add-in – Projektkontext für Claude Code

## Übersicht

Ziel: Word-Dokumente (Rechnungen) direkt als ZUGFeRD/Factur-X-konforme PDF/A-3-Dateien
exportieren – mit eingebettetem XML nach EN 16931. Alles läuft lokal, keine Cloud-Abhängigkeit
für Rechnungsdaten.

---

## Architektur

```
┌─────────────────────────────────────────────────────────────┐
│ Microsoft Word                                              │
│  ├── Rechnungsvorlage_ZUGFeRD_ContentControls.dotx         │
│  │    ├── Content Controls (Kopfdaten, Summen)             │
│  │    └── Positionstabelle (freie Zellen, 7 Spalten)       │
│  └── Office.js Task Pane Add-in                            │
│       └── lädt von GitHub Pages (HTTPS)                    │
└────────────────────┬────────────────────────────────────────┘
                     │ POST /generate (JSON + PDF als Base64)
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ Lokaler Node.js Server (http://127.0.0.1:3737)              │
│  ├── src/server.js      – Express, CORS, Routing            │
│  ├── src/zugferd-xml.js – ZUGFeRD CII XML Generator        │
│  └── src/pdf-embedder.js – PDF/A-3b Einbettung via pdf-lib │
│  config.json            – Firmendaten, outputDir, Port      │
└─────────────────────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ GitHub Pages (https://USERNAME.github.io/REPO/)             │
│  ├── taskpane.html  – Add-in UI                             │
│  ├── taskpane.js    – Kernlogik                             │
│  ├── commands.html  – Pflichtdatei für Manifest             │
│  └── diagnose.html  – Diagnose-Tool                        │
└─────────────────────────────────────────────────────────────┘
```

---

## Projektstruktur

```
zugferd-addin/                         ← Git-Repo-Root (ein Repo fürs gesamte Projekt)
├── word-addin/                        ← wird via GitHub Actions nach GitHub Pages deployt
│   ├── manifest.xml                   ← Add-in Manifest (URLs auf GitHub Pages)
│   ├── taskpane.html                  ← UI: Profil-Auswahl, Vorschau, Buttons, KoSIT-Panel
│   ├── taskpane.js                    ← Kernlogik (siehe unten)
│   ├── commands.html                  ← Leer, Pflicht für Manifest
│   ├── diagnose.html                  ← Diagnose-Tool (läuft in Task Pane)
│   └── assets/                        ← Ribbon-Icons (icon-16/32/80.png)
│
├── zugferd-server/                    ← Node.js Server
│   ├── src/
│   │   ├── server.js                  ← Express Server Port 3737
│   │   ├── zugferd-xml.js             ← XML Generator (CII/EN16931)
│   │   ├── pdf-embedder.js            ← PDF/A-3b via pdf-lib
│   │   ├── kosit-validator.js         ← KoSIT-Validator-Aufruf (Java-Subprozess) + Report-Parsing
│   │   └── open-quick-invoice.ps1     ← Word-COM-Automation für die Schnellerfassung
│   ├── public/                        ← quick-invoice.html/.js (Schnellerfassungs-Formular)
│   ├── tools/kosit/                   ← Validator-JAR + Regel-Konfiguration (via Setup-Skript, NICHT in Git!)
│   ├── config.json                    ← Firmendaten (NICHT in Git!)
│   └── package.json
│
├── .github/workflows/deploy-pages.yml ← Pages-Deploy von word-addin/ bei Push
├── start-server.bat                   ← ZUGFeRD Server starten
├── schnellrechnung.bat                ← Schnellerfassung öffnen (startet Server bei Bedarf)
├── setup-kosit-validator.bat          ← Lädt KoSIT Validator + ZUGFeRD-Regeln herunter
├── start-taskpane-server.bat          ← HTTPS Dev-Server (nicht mehr nötig)
├── install-addin.bat                  ← Shared Folder + Registry
├── clear-office-cache.bat             ← WEF + WebView2 Cache leeren
├── setup-github.bat                   ← überholt, siehe Datei-Inhalt
└── install-certificate.bat            ← HTTPS Dev-Zertifikat (nicht mehr nötig)
```

---

## Content Controls (Tags) in der Vorlage

### Rechnungsdetails (rechte Spalte, Block-SDTs in Tabellenzellen)

| Tag | Inhalt im Dokument | Beispiel |
|-----|-------------------|---------|
| `rechnung_nummer` | `"Rechnungsnummer:RE-2024-0001"` | extractValue() → `"RE-2024-0001"` |
| `zahlungsziel` | `"Zahlungsziel:30 Tage netto"` (manuell) | extractValue() → `"30 Tage netto"` |
| `leistungsort` | `"Leistungsort:München"` | extractValue() → `"München"` |
| `leistungsmonat` | `"Leistungsmonat:April 2026"` (manuell) | extractValue() → `"April 2026"` |
| `waehrung` | `"Währung:EUR"` | extractValue() → `"EUR"` |
| `sprache` | `"Sprache:Deutsch"` | extractValue() → `"Deutsch"` |

**Wichtig:** Alle Detail-Controls enthalten `"Label:Wert"` in einem einzigen Run.
`extractValue()` in taskpane.js schneidet alles bis zum ersten `:` ab.

`rechnung_datum`, `lieferdatum` und `faelligkeitsdatum` werden **nicht mehr manuell
gepflegt** – siehe "Automatisch berechnete Controls" unten.

### Empfänger (linke Spalte, Block-SDTs, je ein Paragraph pro Control)

| Tag | Inhalt |
|-----|--------|
| `empfaenger_firma` | Firmenname (reiner Text) |
| `empfaenger_ansprechpartner` | z. Hd. Name |
| `empfaenger_strasse` | Straße + Hausnummer |
| `empfaenger_plz_ort` | PLZ Ort |
| `empfaenger_land` | Land (ausgeschrieben) |
| `empfaenger_ust_id` | `"USt-ID Empfänger: DE..."` → extractEmpfaengerValue() |
| `empfaenger_kundennr` | `"Kunden-Nr.: KD-001"` → extractEmpfaengerValue() |
| `empfaenger_bestellnr` | `"Ihre Bestellnummer: PO-..."` → extractEmpfaengerValue() |

### Automatisch berechnete Controls (vom Add-in befüllt, nie manuell)

| Tag | Beschreibung |
|-----|-------------|
| `rechnung_nummer_titel` | `"Rechnung RE-2024-0001"` – aus rechnung_nummer |
| `rechnung_datum` | = Monatsletzter des `leistungsmonat` (z. B. Leistungsmonat "April 2026" → `"Rechnungsdatum:30.04.2026"`) |
| `lieferdatum` | = `rechnung_datum` (identischer Monatsletzter) |
| `faelligkeitsdatum` | = `rechnung_datum` + Tage aus `zahlungsziel` (z. B. "30 Tage netto" → +30 Tage) |
| `summe_netto` | Summe aller Zeilenbeträge |
| `summe_mwst_19` | Steuerbetrag 19% |
| `summe_mwst_7` | Steuerbetrag 7% |
| `summe_mwst_0` | Steuerbetrag 0% (immer 0,00 €) |
| `summe_brutto` | Netto + alle Steuerbeträge |
| `label_mwst_19` | `"MwSt. 19 % (auf X.XXX,XX €):"` |
| `label_mwst_7` | `"MwSt. 7 % (auf X.XXX,XX €):"` |
| `label_mwst_0` | `"MwSt. 0 % (auf X.XXX,XX €):"` |
| `zahlung_satz` | Zahlungsaufforderungs-Satz mit Brutto + Fälligkeit |
| `zahlung_verwendungszweck` | `"Verwendungszweck: RE-... / Firma GmbH"` |

**Datumsberechnung (seit 2026-08-20):** `rechnung_datum`/`lieferdatum`/`faelligkeitsdatum`
werden bei "Neu berechnen" aus `leistungsmonat` (z. B. `"April 2026"`, volle deutsche
Monatsnamen) und `zahlungsziel` (führende Zahl wird als Tage interpretiert, z. B.
`"30 Tage netto"` → 30) berechnet und ins jeweilige Control zurückgeschrieben
(`berechneDatumsfelder()` in taskpane.js). Lässt sich der Leistungsmonat nicht parsen
(leer/unbekanntes Format), bleiben die drei Controls unverändert – kein Überschreiben
mit leeren Werten.

### Positionstabelle (KEINE Content Controls – freie Tabellenzellen)

Die Positionstabelle hat **7 Spalten** und wird direkt als Word-Tabelle gelesen.
Zeile 0 = Header, ab Zeile 1 = Daten. Beliebig viele Zeilen möglich.

| Spalte | Index | Inhalt |
|--------|-------|--------|
| Pos. | 0 | Positionsnummer |
| Beschreibung | 1 | Mehrzeilig möglich (Zeile 1 = Titel, weitere = Detail) |
| Menge | 2 | Dezimal, deutsches Format `"40,00"` |
| Einheit | 3 | `"Std."`, `"Psch."`, `"Stk."` etc. |
| Einzelpreis | 4 | `"120,00 €"` (€-Zeichen wird beim Parsen entfernt) |
| MwSt. | 5 | `"19 %"`, `"7 %"`, `"0 %"` |
| Betrag (€) | 6 | Wird vom Add-in berechnet und zurückgeschrieben |

---

## taskpane.js – Ablauflogik

### loadInvoiceData() – wird beim Start und bei "Neu berechnen" aufgerufen

```
1. Word.run() öffnen
2. contentControls laden (tag, text, type)
3. document.body.tables laden
4. Erste Zeile jeder Tabelle laden → Zellanzahl prüfen
5. Tabelle mit 7 Spalten = Positionstabelle
6. Alle Datenzeilen lesen, Zell-Referenzen der Betrag-Spalte merken
7. Für jede Position berechnen:
   - Zeilenbetrag = Menge × Einzelpreis
   - Betrag-Zelle direkt beschreiben (cell.body.insertText)
   - MwSt-Gruppe akkumulieren
8. Summen berechnen: Netto, MwSt19/7/0, Brutto
9. summenMap aufbauen: alle automatischen Controls
10. Controls zurückschreiben (ctrl.insertText)
11. context.sync()
12. invoiceData zusammenstellen (für Vorschau + Export)
13. updatePreview() aufrufen
14. updateExportButton() aufrufen
```

### Wichtige Hilfsfunktionen

```javascript
extractValue(text)
// "Rechnungsnummer:RE-2024-0001" → "RE-2024-0001"
// Alles nach dem ersten ":" zurückgeben

extractEmpfaengerValue(text, prefixes)
// "Kunden-Nr.: KD-0001" → "KD-0001"
// Bekannte Präfixe abschneiden

parseGermanFloat(str)
// "1.234,56 €" → 1234.56
// Punkte entfernen, Komma → Punkt, € entfernen

fmtDE(number)
// 1234.56 → "1.234,56 €"
// toLocaleString('de-DE') + " €"
```

### Dateiname beim Export

```javascript
[leistungsmonat, rechnung_nummer, empfaenger_bestellnr]
  .filter(Boolean)
  .join('-')
  .replace(/[Leerzeichen]/g, '_')
// → "April_2026-RE-2024-0001-PO-2024-001_ZUGFeRD.pdf"
```

### POST /generate – Request-Body

```json
{
  "profile": "EN16931",
  "pdfBase64": "<Word-PDF als Base64>",
  "filename": "April_2026-RE-2024-0001-PO-2024-001",
  "invoice": {
    "rechnung_nummer": "RE-2024-0001",
    "rechnung_datum": "01.01.2025",
    "empfaenger_firma": "Firma GmbH",
    "positions": [
      {
        "beschreibung_titel": "Softwareentwicklung",
        "beschreibung_detail": "Sprint 3",
        "menge": "40,00",
        "einheit": "Std.",
        "einzelpreis": "120,00 €",
        "mwst_satz": "19 %",
        "betrag": "4.800,00 €"
      }
    ],
    "summe_netto": "4.800,00 €",
    "summe_brutto": "5.712,00 €",
    "seller": { ... }
  }
}
```

---

## KoSIT-Validierung

Das erzeugte ZUGFeRD-XML kann optional lokal gegen den offiziellen
[KoSIT Validator](https://github.com/itplr-kosit/validator) (Java) geprüft werden – Schema +
Schematron, mit der ZUGFeRD-Regel-Konfiguration von
[LandrixSoftware/validator-configuration-zugferd](https://github.com/LandrixSoftware/validator-configuration-zugferd).
Läuft komplett lokal, keine Cloud-Abhängigkeit.

### Setup (einmalig)

1. Java-Laufzeit (JRE 11+) installieren, falls nicht vorhanden (z. B. https://adoptium.net).
2. `setup-kosit-validator.bat` ausführen → lädt Validator-JAR + Regel-Konfiguration nach
   `zugferd-server/tools/kosit/`.
3. In `zugferd-server/config.json` den Block `kositValidation` ausfüllen (Pfade aus Schritt 2,
   siehe Skript-Ausgabe) und `enabled: true` setzen.
4. Server neu starten.

### Server (`zugferd-server/src/kosit-validator.js`)

- `checkAvailability(cfg)` – prüft Konfiguration + Java + Dateien (für `GET /validate-status`).
- `validateXml(xmlString, cfg, opts)` – ruft `java -jar <jar> -s <scenarios.xml> -r <repoDir>
  -o <workDir> -p <xmlFile>` als Kindprozess auf, parst den erzeugten Report (Namespace-Präfixe
  werden beim Parsen entfernt, `fast-xml-parser`), sammelt `<valid>`-Flags sowie
  Schematron-Meldungen (`failed-assert`/`successful-report`, SVRL) und XML-Schema-Meldungen
  (`<message level="...">`, XOEV-Report-Format) zu einer einheitlichen `messages[]`-Liste
  ({ severity, message, location, source }) zusammen.

### Neue Endpunkte

| Endpunkt | Beschreibung |
|----------|-------------|
| `GET /validate-status` | `{ available, reasons[] }` – ob KoSIT-Validierung einsatzbereit ist |
| `POST /validate` | Body `{ profile, invoice }` → erzeugt XML + validiert, ohne PDF-Export ("Nur prüfen") |
| `POST /generate` | liefert zusätzlich `validation: {...}` im Response, wenn `kositValidation.enabled = true` |

### config.json → `kositValidation`

```json
"kositValidation": {
  "enabled": false,
  "javaPath": "java",
  "validatorJar": "C:\\...\\tools\\kosit\\validator-1.6.2-standalone.jar",
  "scenariosXml": "C:\\...\\tools\\kosit\\validator-configuration-zugferd-main\\scenarios.xml",
  "timeoutMs": 30000,
  "blockOnInvalid": false
}
```

`blockOnInvalid: true` lässt `/generate` mit `422` fehlschlagen, wenn die Rechnung laut
KoSIT-Validator ungültig ist – Standard ist `false` (nur Anzeige, Export läuft trotzdem durch).

### Task Pane

Neue Sektion "KoSIT-Validierung" (`taskpane.html`/`taskpane.js`):
- Button **"XML prüfen (KoSIT)"** – ruft `POST /validate` unabhängig vom Export auf.
- Wird beim Öffnen automatisch per `GET /validate-status` geprüft; ist der Validator nicht
  eingerichtet, wird der Button deaktiviert und die fehlenden Voraussetzungen werden angezeigt.
- Nach **"ZUGFeRD PDF erstellen"** wird das `validation`-Ergebnis aus `/generate` (falls vorhanden)
  im selben Panel angezeigt.

---

## Bekannte Stolpersteine

### Office.js / Word Add-in

- **Block-SDTs vs. Inline-SDTs**: `contentControls` API gibt nur Block-Level SDTs zurück.
  Inline-SDTs (innerhalb von `<w:r>`) werden ignoriert. Alle Controls müssen als Block-SDTs
  direkt in Tabellenzellen (`parent=tc`) oder Body-Paragraphen eingebettet sein.

- **columnCount existiert nicht**: `table.columnCount` gibt `undefined` zurück.
  Stattdessen: `table.rows.items[0].cells.items.length` für die Spaltenanzahl.

- **Kein verschachtelter Word.run**: Innerhalb eines laufenden `Word.run()` darf kein
  weiterer `Word.run()` aufgerufen werden → führt zu stillem Fehler.

- **WebView2-Cache**: Word cached Task Pane aggressiv. Nach Updates immer
  `clear-office-cache.bat` ausführen (löscht WEF-Cache + WebView2-Cache).

- **Trust Center**: Katalog-URL muss UNC-Pfad sein (`\\localhost\C$\...`).
  Lokale Pfade (`C:\...`) werden nicht als "Freigegebener Ordner" angezeigt.

- **HTTPS-Pflicht**: Task Pane URL muss HTTPS sein → GitHub Pages als Lösung.

### CORS

Server erlaubt: `localhost`, `127.0.0.1`, `*.github.io`.
Weitere Origins können in `config.json` unter `"allowedOrigins"` eingetragen werden.

### Windows / Batch-Skripte

- `.bat`-Dateien müssen **reines ASCII** sein – keine Umlaute, keine Box-Drawing-Zeichen.
  Sonst interpretiert cmd.exe Sonderzeichen als Befehle.

### ZUGFeRD XML

- Alle 4 Profile implementiert: `MINIMUM`, `BASIC_WL`, `EN16931`, `EXTENDED`
- Positionen nur in `EN16931` und `EXTENDED` im XML enthalten
- Einheiten-Mapping: `Std.`→`HUR`, `Psch.`→`LS`, `Stk.`→`C62`, `km`→`KMT` etc.
- Datumsformat: `DD.MM.YYYY` → `YYYYMMDD` (ZUGFeRD-intern)
- Beträge: deutsches Format `"1.234,56 €"` → parseAmount() → Float → fmt() → `"1234.56"`

---

## Schnellerfassung für Standard-Rechnungen

Für den Regelfall (nur Rechnungsnummer, Leistungsmonat, Bestellnummer und
Stunden ändern sich, eine einzige Position, immer derselbe Kunde) gibt es eine
Web-Eingabemaske, die Word per COM-Automation startet/befüllt – ohne Zusatz-Tools,
nur Windows-/Office-Bordmittel (PowerShell + Word-COM).

**Ablauf:** `schnellrechnung.bat` doppelklicken (startet bei Bedarf
`start-server.bat` mit und öffnet den Browser) → 4 Felder ausfüllen → **Word
öffnen & befüllen** → in Word prüfen → **Neu berechnen** → Profil wählen →
**ZUGFeRD PDF erstellen**.

| Datei | Beschreibung |
|-------|-------------|
| `schnellrechnung.bat` | Launcher (Projekt-Root): startet Server falls nötig, öffnet Browser |
| `zugferd-server/public/quick-invoice.html` + `.js` | Eingabemaske (kein Office.js, läuft im Browser) |
| `zugferd-server/src/server.js` → `POST /quick-invoice` | Nimmt die 4 Werte entgegen, ruft PowerShell-Skript auf |
| `zugferd-server/src/open-quick-invoice.ps1` | Word-COM-Automation: neues Dokument aus Vorlage, Content Controls befüllen, Positionstabelle auf 1 Zeile reduzieren |
| `config.json` → `quickInvoice.templatePath` | Pfad zur `.dotx`-Vorlage (lokal, nicht in Git) |

**Was das Skript setzt:** `rechnung_nummer`, `leistungsmonat`, `empfaenger_bestellnr`
(Content Controls, mit exakt der gleichen Label:Wert/Tab-Formatierung wie die
manuell erfassten Felder) sowie die Menge (Spalte 3) in Zeile 1 der 7-spaltigen
Positionstabelle; Zeilen 2+3 (Beispieldaten der Vorlage) werden gelöscht.

**Was unverändert bleibt:** Beschreibung/Einzelpreis/Einheit/MwSt-Satz der
verbleibenden Position (Zeile 1 der Vorlage – dort einmalig auf die echten
Standardwerte pflegen, falls noch nicht geschehen) sowie `zahlungsziel`.
`rechnung_datum`/`lieferdatum`/`faelligkeitsdatum` werden wie gewohnt erst beim
Klick auf **Neu berechnen** im Add-in aus dem Leistungsmonat berechnet (siehe
`berechneDatumsfelder()` weiter oben).

Word-Instanz wird wiederverwendet falls bereits aktiv (`Marshal::GetActiveObject`),
sonst neu gestartet. Läuft **nicht** über `pwsh`/PowerShell 7 (dort existiert
`GetActiveObject` nicht) – `execFile('powershell', ...)` ruft bewusst die
klassische Windows-PowerShell (5.1, .NET Framework) auf.

---

## Täglicher Workflow

1. `start-server.bat` starten (Port 3737, bleibt offen)
2. Word: Neue Datei aus Vorlage `.dotx` erstellen (oder `schnellrechnung.bat`
   für den Regelfall – siehe oben)
3. Rechnungsdaten befüllen (Empfänger, Details, Positionen – Menge + Einzelpreis)
4. Add-in öffnen → **Neu berechnen** (befüllt Betrag, Summen, Verwendungszweck etc.)
5. Profil wählen (Standard: EN 16931)
6. **ZUGFeRD PDF erstellen** → PDF + XML werden in `outputDir` gespeichert

---

## Offene Punkte / mögliche nächste Schritte

- ~~Validierung des erzeugten ZUGFeRD XML (KoSIT Validator)~~ → erledigt, siehe Abschnitt
  "KoSIT-Validierung" (Setup über `setup-kosit-validator.bat` noch manuell auszuführen)
- Mehr Steuer-Szenarien: Reverse Charge, innergemeinschaftliche Lieferung
- Mehrsprachige Rechnungen
- Automatisches Speichern des Word-Dokuments vor Export
- Fehlerbehandlung wenn Pflichtfelder leer sind (Validierung vor Export)
- Verteilung an Kollegen (Microsoft 365 Admin Center statt Shared Folder)
