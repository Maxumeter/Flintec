#Requires -Version 5.1
<#
.SYNOPSIS
    Traegt Arbeitsbeginn, Arbeitsende, Pause und Projektstunden automatisch
    in den Excel-Stundennachweis ein.

.DESCRIPTION
    Die Windows-Aufgabenplanung ruft dieses Skript auf:
      - beim Anmelden und Entsperren   -> -Aktion Kommen
        Ist fuer heute noch kein Arbeitsbeginn erfasst, erscheint oben rechts
        ein kleines Fenster, in das man den Arbeitsbeginn eintippt.
      - beim Herunterfahren            -> -Aktion Herunterfahren
        Die Uhrzeit wird als Arbeitsende eingetragen.
    Sperren (Windows-Taste + L) loest nichts aus.
    Alle Angaben landen zuerst in einer lokalen Datei und werden dann in die
    Excel-Datei uebertragen (klappt das nicht, beim naechsten Mal).

    Aendert man Zeiten in der Excel-Datei von Hand, merkt das Skript das und
    ueberschreibt diesen Tag nicht mehr.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Beginn
    Oeffnet das Fenster fuer den Arbeitsbeginn (z. B. zum Korrigieren).

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Anzeigen
    Zeigt die Stempel der letzten Tage an.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Eintragen -Datum 2026-11-09 -Beginn 7:30 -Ende 16:15
    Traegt einen Tag von Hand ein, z. B. wenn der PC nicht an war.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Test
    Prueft, ob die Excel-Datei gefunden und die Spalte fuer heute erkannt wird.
    Dabei wird nichts geaendert.
#>
[CmdletBinding()]
param(
    [ValidateSet('Kommen', 'Beginn', 'Gehen', 'Herunterfahren', 'Sync', 'Eintragen', 'Anzeigen', 'Test')]
    [string]$Aktion = 'Sync',
    [switch]$OhneExcel,
    [datetime]$Datum = (Get-Date).Date,
    [string]$Beginn,
    [string]$Ende,
    [string]$Pause,
    [string]$Projekt
)

$ErrorActionPreference = 'Stop'
$Inv = [Globalization.CultureInfo]::InvariantCulture
$DeKultur = [Globalization.CultureInfo]::GetCultureInfo('de-DE')

$DatenOrdner = Join-Path $env:LOCALAPPDATA 'Stundennachweis'
New-Item -ItemType Directory -Force -Path $DatenOrdner | Out-Null
$StempelDatei = Join-Path $DatenOrdner 'stempel.csv'
$StatusDatei  = Join-Path $DatenOrdner 'uebertragen.json'
$LogDatei     = Join-Path $DatenOrdner 'protokoll.log'

# Zeilen im Tabellenblatt (siehe 2026_Stundennachweis.xlsx)
$Z_DATUM = 3; $Z_BEGINN = 4; $Z_ENDE = 5; $Z_PAUSE = 6
$Z_PROJEKT_VON = 8; $Z_PROJEKT_BIS = 24
$Z_ABWESEND_VON = 27; $Z_ABWESEND_BIS = 31   # Elternzeit, Krankheit, Urlaub, Feiertage, Kurzarbeit
$S_ERSTER_TAG = 3; $S_LETZTER_TAG = 33        # Spalten C..AG

# ---------------------------------------------------------------- Hilfsfunktionen

function Write-Protokoll([string]$Text) {
    $zeile = '{0}  [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Aktion, $Text
    try { Add-Content -Path $LogDatei -Value $zeile -Encoding UTF8 } catch { }
    Write-Host $Text
}

function ConvertTo-Zeitspanne([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim().Replace('.', ':').Replace(',', ':')
    if ($t -match '^(\d{1,2})(\d{2})$') { $t = $Matches[1] + ':' + $Matches[2] }   # 745 -> 7:45
    if ($t -notmatch '^(\d{1,2}):(\d{2})$' -or [int]$Matches[1] -gt 23 -or [int]$Matches[2] -gt 59) {
        throw "Ungueltige Uhrzeit '$Text' (erwartet z. B. 7:45)."
    }
    return New-Object TimeSpan ([int]$Matches[1]), ([int]$Matches[2]), 0
}

function Format-Zeit($Zeitspanne) {
    if ($null -eq $Zeitspanne) { return '' }
    return '{0}:{1:00}' -f [int][math]::Floor($Zeitspanne.TotalHours), $Zeitspanne.Minutes
}

function Get-Minutengenau([datetime]$Zeit) {
    return New-Object TimeSpan $Zeit.Hour, $Zeit.Minute, 0
}

function Get-Einstellungen {
    $pfad = Join-Path $PSScriptRoot 'einstellungen.json'
    if (-not (Test-Path -LiteralPath $pfad)) { throw "Einstellungsdatei fehlt: $pfad" }
    $e = Get-Content -LiteralPath $pfad -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $e.TageRueckwirkend) { $e | Add-Member -Force NoteProperty TageRueckwirkend 14 }
    return $e
}

function Get-DateiFuerJahr([int]$Jahr) {
    return $Einst.Datei.Replace('{Jahr}', [string]$Jahr)
}

# ---------------------------------------------------------------- Stempel

function Add-Stempel([string]$Art, [datetime]$Zeit = (Get-Date)) {
    $zeile = '{0};{1}' -f $Zeit.ToString('yyyy-MM-ddTHH:mm:ss', $Inv), $Art
    Add-Content -Path $StempelDatei -Value $zeile -Encoding UTF8
}

function Get-Stempel {
    if (-not (Test-Path -LiteralPath $StempelDatei)) { return @() }
    $liste = foreach ($zeile in (Get-Content -LiteralPath $StempelDatei -Encoding UTF8)) {
        $teile = $zeile -split ';'
        if ($teile.Count -lt 2) { continue }
        try {
            [pscustomobject]@{
                Zeit = [datetime]::ParseExact($teile[0].Trim(), 'yyyy-MM-ddTHH:mm:ss', $Inv)
                Art  = $teile[1].Trim()
            }
        } catch { }
    }
    return @($liste)
}

function Remove-AlteStempel {
    # Haelt die Stempeldatei klein: alles aelter als 90 Tage entfernen.
    if (-not (Test-Path -LiteralPath $StempelDatei)) { return }
    $grenze = (Get-Date).Date.AddDays(-90)
    $alle = Get-Stempel
    $behalten = @($alle | Where-Object { $_.Zeit -ge $grenze })
    if ($behalten.Count -lt $alle.Count) {
        $behalten | ForEach-Object { '{0};{1}' -f $_.Zeit.ToString('yyyy-MM-ddTHH:mm:ss', $Inv), $_.Art } |
            Set-Content -Path $StempelDatei -Encoding UTF8
    }
}

function Test-AutomatischerNeustart {
    # Ein Neustart durch Windows Update o. Ae. ist kein Feierabend.
    # Ereignis 1074: Eigenschaft 7 (Index 6) = Benutzer, der heruntergefahren hat.
    try {
        $ev = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'User32'; Id = 1074; StartTime = (Get-Date).AddMinutes(-5) } -MaxEvents 1 -ErrorAction Stop
        $wer = [string]$ev.Properties[6].Value
        return ($wer -match '(^|\\)SYSTEM$' -or $wer -match 'TrustedInstaller')
    } catch {
        return $false
    }
}

function Get-TagesZeiten([datetime]$Tag, $Stempel) {
    $tagesStempel = @($Stempel | Where-Object { $_.Zeit.Date -eq $Tag.Date })
    if ($tagesStempel.Count -eq 0) { return $null }

    # Arbeitsbeginn = zuletzt im Fenster eingetippte Uhrzeit (spaetere Eingabe korrigiert eine fruehere)
    $beginnStempel = $tagesStempel | Where-Object { $_.Art -eq 'Beginn' } | Select-Object -Last 1
    $gehen = $tagesStempel | Where-Object { $_.Art -eq 'Gehen' } | Sort-Object Zeit | Select-Object -Last 1
    if (-not $beginnStempel -and -not $gehen) { return $null }

    $b = $null; $e = $null
    if ($beginnStempel) { $b = Get-Minutengenau $beginnStempel.Zeit }
    if ($gehen -and (-not $beginnStempel -or $gehen.Zeit -gt $beginnStempel.Zeit)) { $e = Get-Minutengenau $gehen.Zeit }
    return [pscustomobject]@{ Tag = $Tag.Date; Beginn = $b; Ende = $e }
}

# ---------------------------------------------------------------- Status (was wurde schon eingetragen?)

function Get-Status {
    $h = @{}
    if (Test-Path -LiteralPath $StatusDatei) {
        try {
            $o = Get-Content -LiteralPath $StatusDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = [string]$p.Value }
        } catch { }
    }
    return $h
}

function Save-Status($Status) {
    # Nur die letzten 120 Tage behalten
    $grenze = (Get-Date).Date.AddDays(-120).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($Status.Keys | Sort-Object)) { if ($k -ge $grenze) { $o[$k] = $Status[$k] } }
    ($o | ConvertTo-Json) | Set-Content -LiteralPath $StatusDatei -Encoding UTF8
}

# ---------------------------------------------------------------- Excel

function Open-Mappe([string]$Pfad) {
    if (-not (Test-Path -LiteralPath $Pfad)) {
        throw "Datei nicht gefunden: $Pfad  (Ist Laufwerk S: verbunden?)"
    }
    $name = Split-Path $Pfad -Leaf

    # Ist die Datei schon in Excel geoeffnet? Dann direkt dort eintragen.
    $xl = $null
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
    if ($xl) {
        foreach ($wb in $xl.Workbooks) {
            if ($wb.Name -ieq $name) {
                if ($wb.ReadOnly) { throw 'Die Datei ist in Excel schreibgeschuetzt geoeffnet.' }
                return [pscustomobject]@{ Excel = $xl; Mappe = $wb; Eigen = $false }
            }
        }
    }

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    try {
        $wb = $xl.Workbooks.Open($Pfad, 0, $false)
    } catch {
        $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        throw
    }
    if ($wb.ReadOnly) {
        $wb.Close($false); $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        throw 'Die Datei ist gerade woanders geoeffnet (nur schreibgeschuetzt verfuegbar). Neuer Versuch beim naechsten Mal.'
    }
    return [pscustomobject]@{ Excel = $xl; Mappe = $wb; Eigen = $true }
}

function Close-Mappe($Sitzung, [bool]$Speichern) {
    if (-not $Sitzung) { return }
    try {
        if ($Speichern) { $Sitzung.Mappe.Save() }
    } finally {
        if ($Sitzung.Eigen) {
            try { $Sitzung.Mappe.Close($false) } catch { }
            try { $Sitzung.Excel.Quit() } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Sitzung.Mappe) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Sitzung.Excel) } catch { }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-Blatt($Mappe, [datetime]$Tag) {
    $name = $DeKultur.DateTimeFormat.GetMonthName($Tag.Month)   # Januar, Februar, Maerz ...
    try { return $Mappe.Worksheets.Item($name) } catch { throw "Tabellenblatt '$name' nicht gefunden." }
}

function Find-Spalte($Blatt, [datetime]$Tag) {
    for ($c = $S_ERSTER_TAG; $c -le $S_LETZTER_TAG; $c++) {
        $v = $Blatt.Cells.Item($Z_DATUM, $c).Value2
        if ($v -is [double] -and $v -gt 30000 -and [datetime]::FromOADate($v).Date -eq $Tag.Date) { return $c }
    }
    throw ("Spalte fuer {0:dd.MM.yyyy} im Blatt '{1}' nicht gefunden." -f $Tag, $Blatt.Name)
}

function Get-Zellzeit($Blatt, [int]$Zeile, [int]$Spalte) {
    $v = $Blatt.Cells.Item($Zeile, $Spalte).Value2
    if ($v -is [double]) {
        $ts = [TimeSpan]::FromDays($v)
        return New-Object TimeSpan $ts.Days, $ts.Hours, ([int][math]::Round($ts.Minutes + $ts.Seconds / 60.0)), 0
    }
    return $null
}

function Test-ZelleLeer($Blatt, [int]$Zeile, [int]$Spalte) {
    $v = $Blatt.Cells.Item($Zeile, $Spalte).Value2
    return ($null -eq $v -or ([string]$v).Trim() -eq '')
}

function Get-Feiertage($Mappe) {
    $liste = @()
    try { $blatt = $Mappe.Worksheets.Item('Feiertage') } catch { return $liste }
    for ($r = 1; $r -le 60; $r++) {
        foreach ($c in 1, 3) {   # Spalte A: Feiertage, Spalte C: Brueckentage
            $v = $blatt.Cells.Item($r, $c).Value2
            if ($v -is [double] -and $v -gt 30000) { $liste += [datetime]::FromOADate($v).Date }
        }
    }
    return $liste
}

function Find-Projektzeile($Blatt, [string]$Nummer) {
    $frei = $null
    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        $a = ([string]$Blatt.Cells.Item($r, 1).Value2).Trim()
        $b = ([string]$Blatt.Cells.Item($r, 2).Value2).Trim()
        if ($a -ieq $Nummer -or $b -ieq $Nummer) { return $r }
        if (-not $frei -and $a -eq '' -and $b -eq '') {
            $inhalt = $Blatt.Application.WorksheetFunction.CountA($Blatt.Range("C${r}:AG${r}"))
            if ($inhalt -eq 0) { $frei = $r }
        }
    }
    if ($frei) {
        $Blatt.Cells.Item($frei, 1).NumberFormat = '@'
        $Blatt.Cells.Item($frei, 1).Value2 = $Nummer
        return $frei
    }
    return $null
}

function Get-StatusWert($Beginn, $Ende, $PauseZeit) {
    return '{0}|{1}|{2}' -f (Format-Zeit $Beginn), (Format-Zeit $Ende), (Format-Zeit $PauseZeit)
}

function Get-PauseFuer([TimeSpan]$Anwesenheit) {
    if ($Anwesenheit.TotalHours -ge 6) { return ConvertTo-Zeitspanne $Einst.Pause }
    return ConvertTo-Zeitspanne $Einst.PauseBisSechsStunden
}

function Set-Tag($Blatt, [int]$Spalte, [datetime]$Tag, $B, $E, $P, [string]$ProjektNr) {
    # Schreibt die Werte eines Tages. $B/$E/$P duerfen $null sein (= nicht aendern).
    if ($null -ne $B) { $Blatt.Cells.Item($Z_BEGINN, $Spalte).Value2 = $B.TotalDays }
    if ($null -ne $E) { $Blatt.Cells.Item($Z_ENDE, $Spalte).Value2 = $E.TotalDays }
    if ($null -ne $P) { $Blatt.Cells.Item($Z_PAUSE, $Spalte).Value2 = $P.TotalDays }

    $b2 = Get-Zellzeit $Blatt $Z_BEGINN $Spalte
    $e2 = Get-Zellzeit $Blatt $Z_ENDE $Spalte
    $p2 = Get-Zellzeit $Blatt $Z_PAUSE $Spalte
    if (-not $ProjektNr -or $null -eq $b2 -or $null -eq $e2) { return }

    $stunden = [math]::Round(($e2 - $b2 - $(if ($p2) { $p2 } else { [TimeSpan]::Zero })).TotalHours, 2)
    if ($stunden -le 0) { return }

    $zeile = Find-Projektzeile $Blatt $ProjektNr
    if (-not $zeile) { Write-Protokoll "Keine freie Projektzeile fuer $ProjektNr gefunden."; return }

    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        if ($r -ne $zeile -and -not (Test-ZelleLeer $Blatt $r $Spalte)) {
            Write-Protokoll ("{0:dd.MM.}: Projektstunden nicht eingetragen, weil an diesem Tag schon andere Projekte gebucht sind." -f $Tag)
            return
        }
    }
    $Blatt.Cells.Item($zeile, $Spalte).Value2 = [double]$stunden
}

# ---------------------------------------------------------------- Ablaeufe

function Invoke-Sync {
    $stempel = Get-Stempel
    $status = Get-Status
    $heute = (Get-Date).Date

    $offen = @()
    for ($i = [int]$Einst.TageRueckwirkend; $i -ge 0; $i--) {
        $tag = $heute.AddDays(-$i)
        $key = $tag.ToString('yyyy-MM-dd')
        if ($status[$key] -eq 'manuell' -or $status[$key] -eq 'uebersprungen') { continue }
        $soll = Get-TagesZeiten $tag $stempel
        if (-not $soll) { continue }
        $offen += $soll
    }
    if ($offen.Count -eq 0) { return }

    foreach ($gruppe in ($offen | Group-Object { $_.Tag.Year })) {
        $pfad = Get-DateiFuerJahr ([int]$gruppe.Name)
        $sitzung = $null
        $geaendert = $false
        try {
            # Erst pruefen, ob sich ueberhaupt etwas geaendert hat - sonst Excel gar nicht oeffnen.
            $kandidaten = @()
            foreach ($soll in $gruppe.Group) {
                $key = $soll.Tag.ToString('yyyy-MM-dd')
                $p = $null
                if ($soll.Ende -and $soll.Beginn) { $p = Get-PauseFuer ($soll.Ende - $soll.Beginn) }
                $wert = Get-StatusWert $soll.Beginn $soll.Ende $p
                if ($status[$key] -ne $wert) { $kandidaten += [pscustomobject]@{ Soll = $soll; Pause = $p; Wert = $wert; Key = $key } }
            }
            if ($kandidaten.Count -eq 0) { continue }

            $sitzung = Open-Mappe $pfad
            $feiertage = @()
            if ($Einst.FeiertageUeberspringen) { $feiertage = Get-Feiertage $sitzung.Mappe }

            foreach ($k in $kandidaten) {
                $tag = $k.Soll.Tag
                $text = $tag.ToString('ddd dd.MM.', $DeKultur)

                if ($Einst.WochenendeUeberspringen -and ($tag.DayOfWeek -eq 'Saturday' -or $tag.DayOfWeek -eq 'Sunday')) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Wochenende - nichts eingetragen."; continue
                }
                if ($feiertage -contains $tag) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Feiertag/Brueckentag - nichts eingetragen."; continue
                }

                $blatt = Get-Blatt $sitzung.Mappe $tag
                $spalte = Find-Spalte $blatt $tag

                $abwesend = $false
                for ($r = $Z_ABWESEND_VON; $r -le $Z_ABWESEND_BIS; $r++) {
                    if (-not (Test-ZelleLeer $blatt $r $spalte)) { $abwesend = $true }
                }
                if ($abwesend) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Urlaub/Krankheit o. Ae. eingetragen - nichts geaendert."; continue
                }

                # Von Hand geaendert? Dann Finger weg.
                $ist = Get-StatusWert (Get-Zellzeit $blatt $Z_BEGINN $spalte) (Get-Zellzeit $blatt $Z_ENDE $spalte) (Get-Zellzeit $blatt $Z_PAUSE $spalte)
                $zuletzt = $status[$k.Key]
                $leer = '||'
                if (($zuletzt -and $ist -ne $zuletzt) -or (-not $zuletzt -and $ist -ne $leer)) {
                    $status[$k.Key] = 'manuell'
                    Write-Protokoll "${text}: In der Datei steht schon etwas ($ist), das nicht vom Skript stammt - bleibt unveraendert."
                    continue
                }

                $projektNr = $null
                if ($Einst.ProjektstundenEintragen) { $projektNr = $Einst.Projekt }
                Set-Tag $blatt $spalte $tag $k.Soll.Beginn $k.Soll.Ende $k.Pause $projektNr
                $status[$k.Key] = $k.Wert
                $geaendert = $true
                Write-Protokoll ("{0}: eingetragen  Beginn {1}  Ende {2}  Pause {3}" -f $text, (Format-Zeit $k.Soll.Beginn), (Format-Zeit $k.Soll.Ende), (Format-Zeit $k.Pause))
            }
        } catch {
            Write-Protokoll "FEHLER: $($_.Exception.Message)"
        } finally {
            try { Close-Mappe $sitzung $geaendert } catch { Write-Protokoll "FEHLER beim Speichern: $($_.Exception.Message)"; $status = Get-Status; $geaendert = $false }
        }
    }
    Save-Status $status
}

function Invoke-Eintragen {
    $b = ConvertTo-Zeitspanne $Beginn
    $e = ConvertTo-Zeitspanne $Ende
    $p = ConvertTo-Zeitspanne $Pause
    if ($null -eq $b -and $null -eq $e -and $null -eq $p) { throw 'Bitte mindestens -Beginn, -Ende oder -Pause angeben.' }

    $sitzung = Open-Mappe (Get-DateiFuerJahr $Datum.Year)
    $ok = $false
    try {
        $blatt = Get-Blatt $sitzung.Mappe $Datum
        $spalte = Find-Spalte $blatt $Datum
        if ($null -eq $p) {
            $b2 = $b; if ($null -eq $b2) { $b2 = Get-Zellzeit $blatt $Z_BEGINN $spalte }
            $e2 = $e; if ($null -eq $e2) { $e2 = Get-Zellzeit $blatt $Z_ENDE $spalte }
            if ($null -ne $b2 -and $null -ne $e2 -and (Test-ZelleLeer $blatt $Z_PAUSE $spalte)) { $p = Get-PauseFuer ($e2 - $b2) }
        }
        $nr = $Projekt
        if (-not $nr -and $Einst.ProjektstundenEintragen) { $nr = $Einst.Projekt }
        Set-Tag $blatt $spalte $Datum $b $e $p $nr
        $ok = $true
    } finally {
        Close-Mappe $sitzung $ok
    }
    $status = Get-Status
    $status[$Datum.ToString('yyyy-MM-dd')] = 'manuell'
    Save-Status $status
    Write-Protokoll ("{0}: von Hand eingetragen (Beginn {1}, Ende {2}, Pause {3})." -f $Datum.ToString('ddd dd.MM.yyyy', $DeKultur), $Beginn, $Ende, $Pause)
}

function Test-BeginnErfasst([datetime]$Tag) {
    return [bool](Get-Stempel | Where-Object { $_.Art -eq 'Beginn' -and $_.Zeit.Date -eq $Tag.Date } | Select-Object -First 1)
}

function Show-BeginnFenster([datetime]$Tag) {
    # Kleines Fenster oben rechts. Gibt die eingetippte Uhrzeit zurueck oder $null.
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $vorher = Get-TagesZeiten $Tag (Get-Stempel)
    $vorschlag = if ($vorher -and $vorher.Beginn) { $vorher.Beginn } else { Get-Minutengenau (Get-Date) }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Stundennachweis'
    $f.FormBorderStyle = 'FixedToolWindow'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.TopMost = $true
    $f.ShowInTaskbar = $true
    $f.StartPosition = 'Manual'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $f.ClientSize = New-Object System.Drawing.Size(300, 92)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Arbeitsbeginn am ' + $Tag.ToString('ddd dd.MM.', $DeKultur) + ':'
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.AutoSize = $true

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = Format-Zeit $vorschlag
    $tb.Location = New-Object System.Drawing.Point(12, 46)
    $tb.Width = 70
    $tb.TextAlign = 'Center'

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Eintragen'
    $ok.Location = New-Object System.Drawing.Point(92, 44)
    $ok.Size = New-Object System.Drawing.Size(100, 30)

    $spaeter = New-Object System.Windows.Forms.Button
    $spaeter.Text = 'Sp' + [char]0x00E4 + 'ter'
    $spaeter.Location = New-Object System.Drawing.Point(198, 44)
    $spaeter.Size = New-Object System.Drawing.Size(90, 30)
    $spaeter.DialogResult = 'Cancel'

    $script:FensterZeit = $null
    $ok.Add_Click({
        try {
            $script:FensterZeit = ConvertTo-Zeitspanne $tb.Text
            $f.DialogResult = 'OK'
            $f.Close()
        } catch {
            $tb.BackColor = [System.Drawing.Color]::MistyRose
            $tb.SelectAll(); $tb.Focus()
        }
    })

    $f.Controls.AddRange(@($lbl, $tb, $ok, $spaeter))
    $f.AcceptButton = $ok
    $f.CancelButton = $spaeter

    $bereich = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $f.Location = New-Object System.Drawing.Point(($bereich.Right - $f.Width - 12), ($bereich.Top + 12))
    $f.Add_Shown({ $f.Activate(); $tb.Focus(); $tb.SelectAll() })

    $ergebnis = $f.ShowDialog()
    $f.Dispose()
    if ($ergebnis -eq [System.Windows.Forms.DialogResult]::OK -and $null -ne $script:FensterZeit) {
        return $Tag.Date.Add($script:FensterZeit)
    }
    return $null
}

function Invoke-BeginnAbfrage([bool]$Immer) {
    $heute = (Get-Date).Date
    if (-not $Immer) {
        if (Test-BeginnErfasst $heute) { return }
        if ($Einst.WochenendeUeberspringen -and ($heute.DayOfWeek -eq 'Saturday' -or $heute.DayOfWeek -eq 'Sunday')) { return }
    }
    # Nur ein Fenster gleichzeitig
    $fm = New-Object System.Threading.Mutex($false, 'Local\StundennachweisFenster')
    $hat = $false
    try {
        try { $hat = $fm.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $hat = $true }
        if (-not $hat) { return }
        if (-not $Immer -and (Test-BeginnErfasst $heute)) { return }
        $zeit = Show-BeginnFenster $heute
        if ($zeit) {
            Add-Stempel 'Beginn' $zeit
            Write-Protokoll ('Arbeitsbeginn {0} erfasst.' -f $zeit.ToString('HH:mm'))
        }
    } finally {
        if ($hat) { $fm.ReleaseMutex() }
        $fm.Dispose()
    }
}

function Invoke-Anzeigen {
    $stempel = Get-Stempel
    $status = Get-Status
    $heute = (Get-Date).Date
    Write-Host ''
    Write-Host 'Tag              Beginn  Ende   Status'
    Write-Host '---------------  ------  -----  ----------------'
    for ($i = 13; $i -ge 0; $i--) {
        $tag = $heute.AddDays(-$i)
        $z = Get-TagesZeiten $tag $stempel
        if (-not $z) { continue }
        $s = $status[$tag.ToString('yyyy-MM-dd')]
        if (-not $s) { $s = 'noch nicht eingetragen' } elseif ($s -notin 'manuell', 'uebersprungen') { $s = 'eingetragen' }
        Write-Host ('{0,-15}  {1,-6}  {2,-5}  {3}' -f $tag.ToString('ddd dd.MM.yyyy', $DeKultur), (Format-Zeit $z.Beginn), (Format-Zeit $z.Ende), $s)
    }
    Write-Host ''
    Write-Host "Protokoll: $LogDatei"
}

function Invoke-Test {
    $pfad = Get-DateiFuerJahr $Datum.Year
    Write-Host "Datei:    $pfad"
    $sitzung = Open-Mappe $pfad
    try {
        $blatt = Get-Blatt $sitzung.Mappe $Datum
        $spalte = Find-Spalte $blatt $Datum
        $buchstabe = ''; $n = $spalte
        while ($n -gt 0) { $m = ($n - 1) % 26; $buchstabe = [char](65 + $m) + $buchstabe; $n = [int][math]::Floor(($n - 1) / 26) }
        Write-Host "Blatt:    $($blatt.Name)"
        Write-Host ("Spalte:   {0} ({1:dd.MM.yyyy})" -f $buchstabe, $Datum)
        $zeile = $null
        for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
            $a = ([string]$blatt.Cells.Item($r, 1).Value2).Trim(); $b = ([string]$blatt.Cells.Item($r, 2).Value2).Trim()
            if ($a -ieq $Einst.Projekt -or $b -ieq $Einst.Projekt) { $zeile = $r; break }
        }
        if ($zeile) { Write-Host "Projekt:  $($Einst.Projekt) in Zeile $zeile" }
        else { Write-Host "Projekt:  $($Einst.Projekt) steht noch nicht im Blatt - wird in die erste freie Zeile geschrieben." }
        Write-Host ("Feiertage gefunden: {0}" -f @(Get-Feiertage $sitzung.Mappe).Count)
        Write-Host 'Test OK - es wurde nichts geaendert.' -ForegroundColor Green
    } finally {
        Close-Mappe $sitzung $false
    }
}

# ---------------------------------------------------------------- Start

$mutex = New-Object System.Threading.Mutex($false, 'Local\Stundennachweis')
$hatMutex = $false
try {
    # Stempel immer sofort setzen (schnell, auch beim Herunterfahren).
    if ($Aktion -eq 'Gehen') { Add-Stempel $Aktion }
    if ($Aktion -eq 'Herunterfahren') {
        if (Test-AutomatischerNeustart) { Write-Protokoll 'Automatischer Neustart durch Windows - kein Arbeitsende eingetragen.'; return }
        Add-Stempel 'Gehen'
    }
    if ($OhneExcel) { return }

    $Einst = Get-Einstellungen
    # Fenster fuer den Arbeitsbeginn (ausserhalb der Sperre, es kann ja eine Weile offen sein)
    if ($Aktion -eq 'Kommen') { Invoke-BeginnAbfrage $false }
    if ($Aktion -eq 'Beginn') { Invoke-BeginnAbfrage $true }

    try { $hatMutex = $mutex.WaitOne([TimeSpan]::FromMinutes(5)) } catch [System.Threading.AbandonedMutexException] { $hatMutex = $true }
    if (-not $hatMutex) { Write-Protokoll 'Ein anderer Lauf ist noch aktiv - abgebrochen.'; return }

    switch ($Aktion) {
        'Kommen'    { Remove-AlteStempel; Invoke-Sync }
        'Beginn'    { Invoke-Sync }
        'Gehen'     { Invoke-Sync }
        'Herunterfahren' { Invoke-Sync }
        'Sync'      { Invoke-Sync }
        'Eintragen' { Invoke-Eintragen }
        'Anzeigen'  { Invoke-Anzeigen }
        'Test'      { Invoke-Test }
    }
} catch {
    Write-Protokoll "FEHLER: $($_.Exception.Message)"
    exit 1
} finally {
    if ($hatMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
