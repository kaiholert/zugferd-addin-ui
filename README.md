# ZUGFeRD Word Add-in

E-Rechnungen direkt aus Microsoft Word erstellen – als PDF/A-3b mit eingebettetem ZUGFeRD/Factur-X XML (EN 16931-konform).

---

## Projektstruktur

```
zugferd-addin/
├── word-addin/               ← Office.js Task Pane Add-in
│   ├── manifest.xml          ← Add-in-Manifest (für Word-Registrierung)
│   ├── taskpane.html         ← Benutzeroberfläche
│   ├── commands.html         ← Pflichtdatei für Manifest
│   └── src/
│       └── taskpane.js       ← Kernlogik: Content Controls lesen, Export steuern
│
├── zugferd-server/           ← Lokaler Node.js-Server
│   ├── config.json           ← ⚠ HIER IHRE FIRMENDATEN EINTRAGEN
│   ├── src/
│   │   ├── server.js         ← Express-Server (Port 3737)
│   │   ├── zugferd-xml.js    ← ZUGFeRD XML-Generator (CII / EN 16931)
│   │   └── pdf-embedder.js   ← PDF/A-3b Einbettung via pdf-lib
│   └── package.json
│
├── start-server.bat          ← Server starten (Windows)
├── install-addin.bat         ← Add-in einmalig registrieren
└── README.md
```

---

## Voraussetzungen

- **Windows** 10/11
- **Microsoft Word** (Microsoft 365 oder Office 2019+)
- **Node.js** ≥ 18 ([nodejs.org](https://nodejs.org))
- Rechnungsvorlage: `Rechnungsvorlage_ZUGFeRD_ContentControls.docx`

---

## Einrichtung (einmalig)

### 1. Firmendaten konfigurieren

Öffne `zugferd-server/config.json` und trage deine Daten ein:

```json
{
  "port": 3737,
  "outputDir": "C:\\Users\\IhrName\\Desktop\\E-Rechnungen",
  "seller": {
    "name":         "Ihre Firma GmbH",
    "strasse":      "Musterstraße 1",
    "plz":          "12345",
    "ort":          "Musterstadt",
    "land":         "DE",
    "ust_id":       "DE123456789",
    "steuernummer": "",
    "iban":         "DE12 3456 7890 1234 5678 90",
    "bic":          "MUBADE12"
  }
}
```

> **Hinweis:** `outputDir` ist der Ordner, in dem die fertigen PDFs gespeichert werden.

### 2. Server-Abhängigkeiten installieren

```
start-server.bat
```
Beim ersten Start werden die npm-Pakete automatisch installiert.

### 3. Add-in-Webserver einrichten (einmalig)

Das Add-in benötigt einen lokalen HTTPS-Server für die Task Pane.

```bat
:: Einmalig installieren:
npm install -g office-addin-dev-certs http-server
npx office-addin-dev-certs install

:: Danach zum Starten der Task Pane (im word-addin-Verzeichnis):
npx http-server word-addin -p 3000 -S ^
  -C "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" ^
  -K "%USERPROFILE%\.office-addin-dev-certs\localhost.key"
```

### 4. Add-in in Word registrieren

1. `install-addin.bat` ausführen
2. Word öffnen
3. **Einfügen → Add-ins → Meine Add-ins** → ZUGFeRD Export auswählen

---

## Tägliche Nutzung

1. `start-server.bat` starten (Fenster bleibt offen)
2. Word-Vorlage öffnen, Rechnung befüllen
3. Im Add-in-Panel: Profil wählen → **ZUGFeRD PDF erstellen**
4. PDF + XML werden im konfigurierten `outputDir` gespeichert

---

## ZUGFeRD-Profile

| Profil | Inhalt | Empfehlung |
|--------|--------|-----------|
| **MINIMUM** | Nur Pflichtfelder (Aussteller, Empfänger, Betrag) | Für sehr einfache Fälle |
| **BASIC WL** | + Zahlungsinfos (IBAN, Fälligkeit) | – |
| **EN 16931** | Vollständig mit Positionen | ✅ Standard B2B |
| **EXTENDED** | Wie EN 16931 + optionale Felder | Für spezielle Anforderungen |

> Ab 2025 gilt für B2B-Rechnungen in Deutschland die E-Rechnungspflicht.  
> **EN 16931** erfüllt die Anforderungen gemäß § 14 UStG.

---

## Content Controls (Tag-Referenz)

Das Add-in liest diese Tags aus den Word-Content-Controls:

### Rechnungsdetails
| Tag | Beschreibung |
|-----|-------------|
| `rechnung_nummer` | Rechnungsnummer |
| `rechnung_datum` | Rechnungsdatum (DD.MM.YYYY) |
| `lieferdatum` | Liefer-/Leistungsdatum |
| `faelligkeitsdatum` | Zahlungsfällig am |
| `zahlungsziel` | Zahlungsziel-Text (z.B. "30 Tage netto") |
| `leistungsort` | Leistungsort |
| `leistungsmonat` | Leistungsmonat/-zeitraum |
| `waehrung` | Währung (Standard: EUR) |

### Empfänger
| Tag | Beschreibung |
|-----|-------------|
| `empfaenger_firma` | Firmenname |
| `empfaenger_ansprechpartner` | Ansprechpartner |
| `empfaenger_strasse` | Straße + Hausnummer |
| `empfaenger_plz_ort` | PLZ Ort (durch Leerzeichen getrennt) |
| `empfaenger_land` | Land (ausgeschrieben, z.B. "Deutschland") |
| `empfaenger_ust_id` | USt-ID des Empfängers |
| `empfaenger_kundennr` | Kundennummer |
| `empfaenger_bestellnr` | Bestellreferenz des Käufers |

### Positionen (dynamisch, beliebig viele)
| Tag-Muster | Beschreibung |
|-----------|-------------|
| `pos_01_pos_nr` | Positionsnummer |
| `pos_01_beschreibung` | Beschreibung (erste Zeile = Titel, weitere = Detail) |
| `pos_01_menge` | Menge (deutsches Format: "40,00") |
| `pos_01_einheit` | Einheit (Std., Stk., Psch., …) |
| `pos_01_einzelpreis` | Einzelpreis (z.B. "120,00 €") |
| `pos_01_mwst_satz` | MwSt-Satz (z.B. "19 %", "7 %", "0 %") |
| `pos_01_betrag` | Zeilenbetrag netto (z.B. "4.800,00 €") |

Für weitere Positionen: `pos_02_*`, `pos_03_*`, … (automatisch erkannt)

### Summen
| Tag | Beschreibung |
|-----|-------------|
| `summe_netto` | Nettobetrag gesamt |
| `summe_mwst_19` | MwSt. 19% |
| `summe_mwst_7` | MwSt. 7% |
| `summe_mwst_0` | Steuerfreie Positionen |
| `summe_brutto` | Bruttobetrag gesamt |

---

## Validierung der erzeugten Rechnungen

Empfohlene kostenlose Validatoren:

- **Factur-X Validator:** https://fnfe-mpe.org/factur-x/factur-x_en/
- **KoSIT Validator:** https://ecosio.com/de/peppol-und-xml-online-validator/
- **ZUGFeRD Online-Portal:** https://www.e-rechnung.tools

---

## Hinweise

**Antivirus:** Node.js-Prozesse können von Antivirensoftware als verdächtig eingestuft werden. Das Batch-Skript und der Server sind unbedenklich – ggf. eine Ausnahme für den Projektordner einrichten.

**Datenschutz:** Alle Daten bleiben lokal. Es werden keine Rechnungsdaten an externe Server übertragen.

**Mehrseitige Rechnungen:** Word exportiert automatisch alle Seiten ins PDF. Das Add-in überträgt das vollständige Dokument.

**Einheiten-Mapping** (ISO 6523 / UN/ECE Rec 20):
| Word-Einheit | XML-Code | Bedeutung |
|-------------|----------|-----------|
| Std., h | HUR | Stunde |
| Stk., Stück | C62 | Stück |
| Psch., pauschal | LS | Pauschale |
| km | KMT | Kilometer |
| Tag, Tage | DAY | Tag |
| Mon., Monat | MON | Monat |
| (sonstige) | ZZ | Nicht standardisiert |
