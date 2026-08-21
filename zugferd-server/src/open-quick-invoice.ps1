<#
.SYNOPSIS
  Erstellt ein neues ZUGFeRD-Rechnungsdokument aus der Word-Vorlage und
  befuellt die Schnellerfassungs-Felder (Rechnungsnummer, Leistungsmonat,
  Bestellnummer, Stunden).

.DESCRIPTION
  Wird vom lokalen ZUGFeRD-Server (POST /quick-invoice) aufgerufen. Nutzt
  Word-COM-Automation (reines Office-/Windows-Bordmittel, keine Zusatz-Tools):
    - Nutzt eine bereits laufende Word-Instanz weiter, falls vorhanden,
      sonst wird Word neu gestartet.
    - Erzeugt ein neues Dokument aus der .dotx-Vorlage (wie Datei > Neu).
    - Schreibt die 4 Werte in die passenden Content Controls (per Tag),
      mit exakt der gleichen Label:Wert-Formatierung wie die manuell
      erfassten Felder (siehe Kommentare unten), damit die
      Spaltenausrichtung im Dokument erhalten bleibt.
    - Reduziert die 7-spaltige Positionstabelle auf eine Datenzeile und
      setzt dort die Menge (Stunden). Beschreibung/Einzelpreis/Einheit/
      MwSt der verbleibenden Zeile werden NICHT angefasst - sie bleiben so,
      wie sie aktuell in Zeile 1 der Vorlage stehen (dort einmalig pflegen).
    - Rechnungsdatum/Lieferdatum/Faelligkeitsdatum/Zahlungsziel werden
      bewusst nicht gesetzt - die werden im Add-in beim Klick auf
      "Neu berechnen" aus dem Leistungsmonat berechnet.

.PARAMETER RechnungNummer
  z.B. "RE-2024-00042"

.PARAMETER Leistungsmonat
  Deutscher Monatsname + Jahr, z.B. "April 2026" (vom Server bereits
  aus dem <input type="month"> Wert umgewandelt).

.PARAMETER Bestellnummer
  Kunden-Bestellnummer, z.B. "PO-2024-001"

.PARAMETER Stunden
  Menge als Zahl (Punkt oder Komma als Dezimaltrenner), wird intern ins
  deutsche Format "40,00" umgewandelt.

.PARAMETER TemplatePfad
  Vollstaendiger Pfad zur .dotx-Vorlage.
#>
param(
    [Parameter(Mandatory = $true)][string]$RechnungNummer,
    [Parameter(Mandatory = $true)][string]$Leistungsmonat,
    [Parameter(Mandatory = $true)][string]$Bestellnummer,
    [Parameter(Mandatory = $true)][string]$Stunden,
    [Parameter(Mandatory = $true)][string]$TemplatePfad
)

$ErrorActionPreference = 'Stop'

function Set-ContentControlText {
    param($Doc, [string]$Tag, [string]$Text)

    $ccs = $Doc.SelectContentControlsByTag($Tag)
    if ($ccs.Count -lt 1) {
        Write-Warning "Content Control mit Tag '$Tag' nicht gefunden - uebersprungen."
        return
    }
    $ccs.Item(1).Range.Text = $Text
}

try {
    if (-not (Test-Path -LiteralPath $TemplatePfad)) {
        throw "Vorlage nicht gefunden: $TemplatePfad"
    }

    # Stunden ins deutsche Dezimalformat (Komma, 2 Nachkommastellen) bringen.
    # Eingabe kann Punkt oder Komma als Trenner haben (Browser-Locale-abhaengig).
    $stundenZahl = [double]($Stunden -replace ',', '.')
    $stundenDE   = $stundenZahl.ToString('N2', [System.Globalization.CultureInfo]::GetCultureInfo('de-DE'))

    # ── Word-Instanz holen oder starten ──────────────────────────────────────
    $word = $null
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    } catch {
        $word = $null
    }
    if (-not $word) {
        $word = New-Object -ComObject 'Word.Application'
    }
    $word.Visible = $true

    # ── Neues Dokument aus Vorlage erzeugen ──────────────────────────────────
    $doc = $word.Documents.Add($TemplatePfad)

    # ── Content Controls befuellen (Formatierung wie manuelle Felder) ───────
    # "Rechnungsnummer:" + TAB + Wert
    Set-ContentControlText -Doc $doc -Tag 'rechnung_nummer' -Text ("Rechnungsnummer:`t$RechnungNummer")
    # "Leistungsmonat: " (Leerzeichen) + TAB + Wert
    Set-ContentControlText -Doc $doc -Tag 'leistungsmonat' -Text ("Leistungsmonat: `t$Leistungsmonat")
    # "Ihre Bestellnummer:  " (zwei Leerzeichen) + Wert - kein Tab in diesem Bereich
    Set-ContentControlText -Doc $doc -Tag 'empfaenger_bestellnr' -Text ("Ihre Bestellnummer:  $Bestellnummer")

    # ── Positionstabelle finden (7 Spalten) und auf 1 Datenzeile reduzieren ──
    $posTable = $null
    foreach ($tbl in $doc.Tables) {
        if ($tbl.Columns.Count -eq 7) { $posTable = $tbl; break }
    }
    if (-not $posTable) {
        Write-Warning "Positionstabelle (7 Spalten) nicht gefunden - Stunden konnten nicht gesetzt werden."
    } else {
        # Von unten nach oben loeschen, bis nur Header + 1 Datenzeile uebrig sind
        while ($posTable.Rows.Count -gt 2) {
            $posTable.Rows.Item($posTable.Rows.Count).Delete()
        }
        if ($posTable.Rows.Count -ge 2) {
            # Spalte 3 = Menge (Pos=1, Beschreibung=2, Menge=3, Einheit=4, ...)
            $posTable.Cell(2, 3).Range.Text = $stundenDE
        }
    }

    $word.Activate()
    $doc.Activate()

    Write-Output "OK"
    exit 0
} catch {
    Write-Error "Schnellerfassung fehlgeschlagen: $($_.Exception.Message)"
    exit 1
}
