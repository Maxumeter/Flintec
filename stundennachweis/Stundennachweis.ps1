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
$FensterDatei = Join-Path $DatenOrdner 'fenster.json'

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
    foreach ($k in ($Status.Keys | Sort-Object)) {
        if ($k -match '(\d{4}-\d{2}-\d{2})$' -and $Matches[1] -lt $grenze) { continue }
        $o[$k] = $Status[$k]
    }
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

function Find-Projektzeile($Blatt, [string]$Nummer, [string]$Bezeichnung, [switch]$NichtAnlegen) {
    $frei = $null
    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        $a = ([string]$Blatt.Cells.Item($r, 1).Value2).Trim()
        $b = ([string]$Blatt.Cells.Item($r, 2).Value2).Trim()
        if ($a -ieq $Nummer) {
            # Bezeichnung in Spalte B aktuell halten
            if ($Bezeichnung -and $b -ne $Bezeichnung) { $Blatt.Cells.Item($r, 2).Value2 = $Bezeichnung }
            return $r
        }
        if ($b -ieq $Nummer) { return $r }   # aeltere Schreibweise: Nummer steht in Spalte B
        if (-not $frei -and $a -eq '' -and $b -eq '') {
            $inhalt = $Blatt.Application.WorksheetFunction.CountA($Blatt.Range("C${r}:AG${r}"))
            if ($inhalt -eq 0) { $frei = $r }
        }
    }
    if ($frei -and -not $NichtAnlegen) {
        $Blatt.Cells.Item($frei, 1).NumberFormat = '@'
        $Blatt.Cells.Item($frei, 1).Value2 = $Nummer
        if ($Bezeichnung) { $Blatt.Cells.Item($frei, 2).Value2 = $Bezeichnung }
        return $frei
    }
    return $null
}

function Get-StatusWert($Beginn, $Ende, $PauseZeit) {
    return '{0}|{1}|{2}' -f (Format-Zeit $Beginn), (Format-Zeit $Ende), (Format-Zeit $PauseZeit)
}

function Get-PauseFuer([TimeSpan]$Anwesenheit) {
    # Erst wenn 6 Stunden ueberschritten sind, gilt die lange Pause.
    if ($Anwesenheit.TotalHours -gt 6) { return ConvertTo-Zeitspanne $Einst.Pause }
    return ConvertTo-Zeitspanne $Einst.PauseBisSechsStunden
}

function ConvertTo-Anteil([string]$Text) {
    # Leer = Rest, "2:30" / "2,5" = Stunden, "40%" = Prozent der Arbeitszeit
    $t = ([string]$Text).Trim()
    if ($t -eq '') { return @{ Art = 'Rest'; Wert = 0.0 } }
    if ($t -match '^(\d+(?:[.,]\d+)?)\s*%$') { return @{ Art = 'Prozent'; Wert = [double]::Parse($Matches[1].Replace(',', '.'), $Inv) } }
    if ($t -match '^(\d{1,2}):(\d{2})$') { return @{ Art = 'Stunden'; Wert = [int]$Matches[1] + [int]$Matches[2] / 60.0 } }
    if ($t -match '^(\d+(?:[.,]\d+)?)\s*h?$') { return @{ Art = 'Stunden'; Wert = [double]::Parse($Matches[1].Replace(',', '.'), $Inv) } }
    throw "Ungueltiger Anteil '$Text' (z. B. 2:30, 2,5 oder 40%)."
}

function Get-Verteilung([double]$Gesamt, $Projekte, [switch]$Still) {
    # Verteilt die Arbeitszeit auf die Projekte.
    #  - feste Stunden und Prozente zuerst
    #  - der Rest gleichmaessig auf die Projekte ohne Angabe
    #  - gibt es keins ohne Angabe, bekommt das letzte Projekt den Rest
    #  - sind die festen Anteile groesser als die Arbeitszeit, werden sie anteilig gekuerzt
    $liste = @()
    foreach ($pj in $Projekte) {
        try { $a = ConvertTo-Anteil $pj.Anteil } catch { if (-not $Still) { Write-Protokoll $_.Exception.Message }; $a = @{ Art = 'Rest'; Wert = 0.0 } }
        $soll = 0.0
        if ($a.Art -eq 'Stunden') { $soll = $a.Wert }
        if ($a.Art -eq 'Prozent') { $soll = $Gesamt * $a.Wert / 100.0 }
        $liste += [pscustomobject]@{ Nummer = $pj.Nummer; Bezeichnung = $pj.Bezeichnung; Rest = ($a.Art -eq 'Rest'); Stunden = $soll }
    }
    if ($liste.Count -eq 0) { return @() }

    $fest = 0.0; foreach ($x in $liste) { if (-not $x.Rest) { $fest += $x.Stunden } }
    if ($fest -gt $Gesamt -and $fest -gt 0) {
        $faktor = $Gesamt / $fest
        foreach ($x in $liste) { if (-not $x.Rest) { $x.Stunden = $x.Stunden * $faktor } }
        $fest = $Gesamt
    }
    $rest = @($liste | Where-Object { $_.Rest })
    if ($rest.Count -gt 0) { foreach ($x in $rest) { $x.Stunden = ($Gesamt - $fest) / $rest.Count } }
    else { $liste[-1].Stunden += ($Gesamt - $fest) }

    # Auf 2 Nachkommastellen runden, Rundungsdifferenz ausgleichen
    $summe = 0.0
    foreach ($x in $liste) { $x.Stunden = [math]::Round($x.Stunden, 2); $summe += $x.Stunden }
    $diff = [math]::Round([math]::Round($Gesamt, 2) - $summe, 2)
    if ($diff -ne 0) {
        $ziel = if ($rest.Count -gt 0) { $rest[-1] } else { $liste[-1] }
        $ziel.Stunden = [math]::Round($ziel.Stunden + $diff, 2)
    }
    return $liste
}

function Set-Tag($Blatt, [int]$Spalte, [datetime]$Tag, $B, $E, $P, $Projekte, $Status) {
    # Schreibt die Werte eines Tages. $B/$E/$P duerfen $null sein (= nicht aendern).
    if ($null -ne $B) { $Blatt.Cells.Item($Z_BEGINN, $Spalte).Value2 = $B.TotalDays }
    if ($null -ne $E) { $Blatt.Cells.Item($Z_ENDE, $Spalte).Value2 = $E.TotalDays }
    if ($null -ne $P) { $Blatt.Cells.Item($Z_PAUSE, $Spalte).Value2 = $P.TotalDays }

    $b2 = Get-Zellzeit $Blatt $Z_BEGINN $Spalte
    $e2 = Get-Zellzeit $Blatt $Z_ENDE $Spalte
    $p2 = Get-Zellzeit $Blatt $Z_PAUSE $Spalte
    if (@($Projekte).Count -eq 0 -or $null -eq $b2 -or $null -eq $e2) { return }

    $gesamt = [math]::Round(($e2 - $b2 - $(if ($p2) { $p2 } else { [TimeSpan]::Zero })).TotalHours, 2)
    if ($gesamt -le 0) { return }

    $key = 'proj-' + $Tag.ToString('yyyy-MM-dd')
    $text = $Tag.ToString('dd.MM.')

    # Zeilen, die das Skript an diesem Tag frueher schon beschrieben hat (z. B. vor einer Aenderung der Verteilung)
    $frueher = @{}
    if ($Status -and $Status[$key]) {
        foreach ($nr in ($Status[$key] -split ';')) {
            if (-not $nr) { continue }
            $z = Find-Projektzeile $Blatt $nr '' -NichtAnlegen
            if ($z) { $frueher[$z] = $true }
        }
    }

    # Von Hand auf andere Projekte gebucht? Dann die Projektstunden nicht anfassen.
    $nummern = @($Projekte | ForEach-Object { $_.Nummer })
    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        if ($frueher[$r] -or (Test-ZelleLeer $Blatt $r $Spalte)) { continue }
        $a = ([string]$Blatt.Cells.Item($r, 1).Value2).Trim(); $bb = ([string]$Blatt.Cells.Item($r, 2).Value2).Trim()
        if (@($nummern | Where-Object { $_ -ieq $a -or $_ -ieq $bb }).Count -gt 0) { continue }   # eines der eigenen Projekte
        Write-Protokoll "${text}: Projektstunden nicht eingetragen, weil an diesem Tag schon von Hand auf andere Projekte gebucht wurde."
        return
    }

    $verteilung = Get-Verteilung $gesamt $Projekte
    $neu = @{}
    foreach ($v in $verteilung) {
        $z = Find-Projektzeile $Blatt $v.Nummer $v.Bezeichnung
        if (-not $z) { Write-Protokoll "Keine freie Projektzeile fuer $($v.Nummer) gefunden."; continue }
        if ($neu.ContainsKey($z)) { $neu[$z] += $v.Stunden } else { $neu[$z] = $v.Stunden }
    }
    # Alte Buchungen des Skripts entfernen, die nicht mehr zur Verteilung gehoeren
    foreach ($z in $frueher.Keys) { if (-not $neu.ContainsKey($z)) { $Blatt.Cells.Item($z, $Spalte).Value2 = $null } }
    foreach ($z in $neu.Keys) {
        if ($neu[$z] -gt 0) { $Blatt.Cells.Item($z, $Spalte).Value2 = [double][math]::Round($neu[$z], 2) }
        else { $Blatt.Cells.Item($z, $Spalte).Value2 = $null }
    }
    if ($Status) { $Status[$key] = ($verteilung | ForEach-Object { $_.Nummer }) -join ';' }
    Write-Protokoll ("${text}: Projektstunden " + (($verteilung | ForEach-Object { '{0} {1:0.00} h' -f $_.Nummer, $_.Stunden }) -join ', '))
}

function Get-ProjektSignatur($Projekte) {
    return (@($Projekte) | ForEach-Object { '{0}={1}' -f $_.Nummer, $_.Anteil }) -join ';'
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
    # Aktuelles Jahr immer pruefen, damit geaenderte Personalnummer/Name auch ohne neue Zeiten ankommen.
    $jahre = @($offen | ForEach-Object { $_.Tag.Year }) + $heute.Year | Sort-Object -Unique
    foreach ($jahr in $jahre) {
        $gruppe = [pscustomobject]@{ Name = [string]$jahr; Group = @($offen | Where-Object { $_.Tag.Year -eq $jahr }) }
        $pfad = Get-DateiFuerJahr $jahr
        $sitzung = $null
        $geaendert = $false
        try {
            # Erst pruefen, ob sich ueberhaupt etwas geaendert hat - sonst Excel gar nicht oeffnen.
            $kandidaten = @()
            $projekte = @()
            if ($Einst.ProjektstundenEintragen) { $projekte = @(Get-Projekte) }
            $projSig = Get-ProjektSignatur $projekte
            foreach ($soll in $gruppe.Group) {
                $key = $soll.Tag.ToString('yyyy-MM-dd')
                $p = $null
                if ($soll.Ende -and $soll.Beginn) { $p = Get-PauseFuer ($soll.Ende - $soll.Beginn) }
                $wert = Get-StatusWert $soll.Beginn $soll.Ende $p
                # Heute zusaetzlich neu verteilen, wenn die Projekte im Fenster geaendert wurden
                $sigNeu = ($soll.Tag -eq $heute -and $soll.Beginn -and $soll.Ende -and $status['sig-' + $key] -ne $projSig)
                if ($status[$key] -ne $wert -or $sigNeu) { $kandidaten += [pscustomobject]@{ Soll = $soll; Pause = $p; Wert = $wert; Key = $key } }
            }
            $fe = Get-FensterEinst
            $kopfWert = '{0}|{1}' -f $fe.Personalnummer, $fe.Name
            $kopfKey = 'kopf-' + $gruppe.Name
            $kopfNeu = ($kopfWert -ne '|' -and $status[$kopfKey] -ne $kopfWert)
            if ($kandidaten.Count -eq 0 -and -not $kopfNeu) { continue }

            $sitzung = Open-Mappe $pfad
            if ($kopfNeu) {
                if (Set-Kopfdaten $sitzung.Mappe $fe.Personalnummer $fe.Name) { $geaendert = $true }
                $status[$kopfKey] = $kopfWert
            }
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

                Set-Tag $blatt $spalte $tag $k.Soll.Beginn $k.Soll.Ende $k.Pause $projekte $status
                $status[$k.Key] = $k.Wert
                if ($k.Soll.Beginn -and $k.Soll.Ende) { $status['sig-' + $k.Key] = $projSig }
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
        $status = Get-Status
        $projekte = @()
        if ($Projekt) { $projekte = @([pscustomobject]@{ Nummer = $Projekt; Bezeichnung = ''; Anteil = '' }) }
        elseif ($Einst.ProjektstundenEintragen) { $projekte = @(Get-Projekte) }
        Set-Tag $blatt $spalte $Datum $b $e $p $projekte $status
        $ok = $true
    } finally {
        Close-Mappe $sitzung $ok
    }
    $status[$Datum.ToString('yyyy-MM-dd')] = 'manuell'
    Save-Status $status
    Write-Protokoll ("{0}: von Hand eingetragen (Beginn {1}, Ende {2}, Pause {3})." -f $Datum.ToString('ddd dd.MM.yyyy', $DeKultur), $Beginn, $Ende, $Pause)
}

function Test-BeginnErfasst([datetime]$Tag) {
    return [bool](Get-Stempel | Where-Object { $_.Art -eq 'Beginn' -and $_.Zeit.Date -eq $Tag.Date } | Select-Object -First 1)
}

function Get-FensterEinst {
    # Gemerkte Angaben aus dem Fenster: Haken, Personalnummer, Name
    $e = @{ AutomatischSchliessen = $false; Personalnummer = ''; Name = ''; Projekte = @() }
    try {
        if (Test-Path -LiteralPath $FensterDatei) {
            $o = Get-Content -LiteralPath $FensterDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $o.AutomatischSchliessen) { $e.AutomatischSchliessen = [bool]$o.AutomatischSchliessen }
            if ($o.Personalnummer) { $e.Personalnummer = ([string]$o.Personalnummer).Trim() }
            if ($o.Name) { $e.Name = ([string]$o.Name).Trim() }
            if ($o.PSObject.Properties['Projekte']) {
                $e.Projekte = @($o.Projekte | Where-Object { $_ -and ([string]$_.Nummer).Trim() } | ForEach-Object {
                    [ordered]@{ Nummer = ([string]$_.Nummer).Trim(); Bezeichnung = ([string]$_.Bezeichnung).Trim(); Anteil = ([string]$_.Anteil).Trim() } })
            } elseif ($o.Projekt) {
                # Aeltere Version: nur ein Projekt
                $e.Projekte = @([ordered]@{ Nummer = ([string]$o.Projekt).Trim(); Bezeichnung = ([string]$o.Bezeichnung).Trim(); Anteil = '' })
            }
        }
    } catch { }
    return $e
}

function Get-Projekte {
    # Projekte aus dem Fenster; solange dort keins steht, das aus einstellungen.json
    $liste = @((Get-FensterEinst).Projekte | ForEach-Object { [pscustomobject]$_ })
    if ($liste.Count -eq 0 -and $Einst.Projekt) {
        $liste = @([pscustomobject]@{ Nummer = [string]$Einst.Projekt; Bezeichnung = ''; Anteil = '' })
    }
    return $liste
}

function Save-FensterEinst($E) {
    try { ($E | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $FensterDatei -Encoding UTF8 } catch { }
}

function Set-Kopfdaten($Mappe, [string]$Personalnummer, [string]$Name) {
    # Personalnummer steht in E1, der Name in J1 jedes Monatsblatts.
    # Zellen mit Formel (z. B. =Oktober!J1) bleiben unangetastet - die holen sich den Wert selbst.
    $geaendert = $false
    for ($m = 1; $m -le 12; $m++) {
        try { $blatt = $Mappe.Worksheets.Item($DeKultur.DateTimeFormat.GetMonthName($m)) } catch { continue }
        foreach ($paar in @(@('E1', $Personalnummer), @('J1', $Name))) {
            $wert = $paar[1]
            if ([string]::IsNullOrWhiteSpace($wert)) { continue }
            $zelle = $blatt.Range($paar[0])
            if ($zelle.HasFormula) { continue }
            if (([string]$zelle.Value2).Trim() -eq $wert) { continue }
            if ($wert -match '^\d{1,15}$') { $zelle.Value2 = [double]$wert } else { $zelle.Value2 = $wert }
            $geaendert = $true
        }
    }
    if ($geaendert) { Write-Protokoll "Personalnummer/Name in die Datei eingetragen." }
    return $geaendert
}

function Show-BeginnFenster([datetime]$Tag) {
    # Kleines Fenster oben rechts. Gibt die eingetippte Uhrzeit zurueck oder $null.
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $vorher = Get-TagesZeiten $Tag (Get-Stempel)
    $vorschlag = if ($vorher -and $vorher.Beginn) { $vorher.Beginn } else { Get-Minutengenau (Get-Date) }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Stundennachweis'
    $f.FormBorderStyle = 'FixedSingle'
    $f.MaximizeBox = $false; $f.MinimizeBox = $true   # minimieren erlaubt
    $f.TopMost = $true
    $f.ShowInTaskbar = $true
    $f.StartPosition = 'Manual'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $f.ClientSize = New-Object System.Drawing.Size(420, 352)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Arbeitsbeginn ' + $Tag.ToString('ddd dd.MM.', $DeKultur) + ':'
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.AutoSize = $true

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = Format-Zeit $vorschlag
    $tb.Location = New-Object System.Drawing.Point(12, 46)
    $tb.Width = 70
    $tb.TextAlign = 'Center'

    # Gesamtstunden des Tages bis jetzt (Beginn bis aktuelle Uhrzeit, Pause abgezogen)
    $lblGesamt = New-Object System.Windows.Forms.Label
    $lblGesamt.Location = New-Object System.Drawing.Point(200, 12)
    $lblGesamt.Size = New-Object System.Drawing.Size(208, 22)
    $lblGesamt.TextAlign = 'TopRight'
    $lblGesamt.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($lblGesamt, 'Arbeitszeit heute von Beginn bis jetzt, Pause bereits abgezogen')
    $zeigeGesamt = {
        try { $bz = ConvertTo-Zeitspanne $tb.Text } catch { $bz = $null }
        if ($null -eq $bz) { $lblGesamt.Text = 'Heute: -'; return }
        $anw = (Get-Minutengenau (Get-Date)) - $bz
        $std = 0.0
        if ($anw.TotalMinutes -gt 0) {
            $arbeit = $anw - (Get-PauseFuer $anw)
            if ($arbeit.TotalMinutes -gt 0) { $std = $arbeit.TotalHours }
        }
        $lblGesamt.Text = 'Heute: {0} h ({1})' -f (Format-Zeit ([TimeSpan]::FromHours($std))), $std.ToString('0.00', $DeKultur)

        # Stunden pro Projekt in der Tabelle (Spalte "Stunden")
        if (-not $grid) { return }
        $verteilung = @(Get-Verteilung ([math]::Round($std, 2)) (& $leseGrid) -Still)
        $i = 0
        foreach ($zeile in $grid.Rows) {
            if ($zeile.IsNewRow) { continue }
            if (-not ([string]$zeile.Cells[0].Value).Trim()) { $zeile.Cells[3].Value = ''; continue }
            if ($i -lt $verteilung.Count) { $zeile.Cells[3].Value = $verteilung[$i].Stunden.ToString('0.00', $DeKultur) }
            $i++
        }
    }
    $uhr = New-Object System.Windows.Forms.Timer
    $uhr.Interval = 15000
    $uhr.Add_Tick({ & $zeigeGesamt })

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Eintragen'
    $ok.Location = New-Object System.Drawing.Point(92, 44)
    $ok.Size = New-Object System.Drawing.Size(100, 30)

    $spaeter = New-Object System.Windows.Forms.Button
    $spaeter.Text = 'Sp' + [char]0x00E4 + 'ter'
    $spaeter.Location = New-Object System.Drawing.Point(198, 44)
    $spaeter.Size = New-Object System.Drawing.Size(90, 30)
    $spaeter.DialogResult = 'Cancel'

    $sekunden = 30
    if ($Einst.FensterSekunden) { $sekunden = [int]$Einst.FensterSekunden }

    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = ('Automatisch schlie' + [char]0x00DF + 'en nach {0} s') -f $sekunden
    $cb.Location = New-Object System.Drawing.Point(12, 84)
    $cb.AutoSize = $true
    $fe = Get-FensterEinst
    $cb.Checked = $fe.AutomatischSchliessen

    # Eingabefelder, die sich das Fenster merkt
    $felder = @()
    $beschriftungen = @()
    function New-Feld([string]$Text, [int]$Y, [string]$Wert) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Text
        $l.Location = New-Object System.Drawing.Point(12, ($Y + 3))
        $l.AutoSize = $true
        $t = New-Object System.Windows.Forms.TextBox
        $t.Text = $Wert
        $t.Location = New-Object System.Drawing.Point(104, $Y)
        $t.Width = 304
        $script:FeldLabel = $l
        return $t
    }
    $tbPnr  = New-Feld 'Personalnr.:' 118 $fe.Personalnummer; $beschriftungen += $script:FeldLabel
    $tbName = New-Feld 'Name:'        150 $fe.Name;           $beschriftungen += $script:FeldLabel
    $felder = @($tbPnr, $tbName)

    # Projekte: Tabelle mit Auftragsnr., Bezeichnung und Anteil
    $lblPj = New-Object System.Windows.Forms.Label
    $lblPj.Text = 'Projekte  (Anteil: 2:30 / 2,5 / 40%  -  leer = Rest)'
    $lblPj.Location = New-Object System.Drawing.Point(12, 186)
    $lblPj.AutoSize = $true
    $beschriftungen += $lblPj

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 210)
    $grid.Size = New-Object System.Drawing.Size(396, 130)
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersWidth = 28
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
    $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
    [void]$grid.Columns.Add('Nummer', 'Auftragsnr.')
    [void]$grid.Columns.Add('Bezeichnung', 'Bezeichnung')
    [void]$grid.Columns.Add('Anteil', 'Anteil')
    [void]$grid.Columns.Add('Stunden', 'Stunden')
    $grid.Columns[0].FillWeight = 28
    $grid.Columns[1].FillWeight = 40
    $grid.Columns[2].FillWeight = 16
    $grid.Columns[3].FillWeight = 16
    $grid.Columns[3].ReadOnly = $true
    $grid.Columns[3].DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
    $grid.Columns[3].DefaultCellStyle.Alignment = 'MiddleRight'
    foreach ($pj in (Get-Projekte)) { [void]$grid.Rows.Add($pj.Nummer, $pj.Bezeichnung, $pj.Anteil) }

    $leseGrid = {
        $liste = @()
        foreach ($zeile in $grid.Rows) {
            if ($zeile.IsNewRow) { continue }
            $nr = ([string]$zeile.Cells[0].Value).Trim()
            if (-not $nr) { continue }
            $liste += [ordered]@{ Nummer = $nr; Bezeichnung = ([string]$zeile.Cells[1].Value).Trim(); Anteil = ([string]$zeile.Cells[2].Value).Trim() }
        }
        , $liste
    }

    # Sofort merken, sobald sich etwas aendert (bleibt auch nach Neustart erhalten)
    $merken = {
        $e2 = Get-FensterEinst
        $e2.Personalnummer = $tbPnr.Text.Trim()
        $e2.Name = $tbName.Text.Trim()
        $e2.Projekte = & $leseGrid
        $e2.AutomatischSchliessen = $cb.Checked
        Save-FensterEinst $e2
    }

    # Countdown: laeuft nur, solange der Haken gesetzt ist und nichts eingetippt wurde.
    # Bei 0 wird die angezeigte Uhrzeit uebernommen.
    $script:Rest = $sekunden
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $script:Rest--
        if ($script:Rest -le 0) { $timer.Stop(); $ok.PerformClick() }
        else { $ok.Text = 'Eintragen ({0})' -f $script:Rest }
    })
    $stoppen = {
        $timer.Stop()
        $ok.Text = 'Eintragen'
    }
    $cb.Add_CheckedChanged({
        $e2 = Get-FensterEinst; $e2.AutomatischSchliessen = $cb.Checked; Save-FensterEinst $e2
        if ($cb.Checked) { $script:Rest = $sekunden; $ok.Text = 'Eintragen ({0})' -f $script:Rest; $timer.Start() }
        else { & $stoppen }
    })
    $tb.Add_KeyDown({ & $stoppen })
    $tb.Add_TextChanged({ & $zeigeGesamt })
    $tb.Add_MouseDown({ & $stoppen })
    foreach ($feld in $felder) {
        $feld.Add_KeyDown({ & $stoppen })
        $feld.Add_MouseDown({ & $stoppen })
        $feld.Add_TextChanged({ & $merken })
    }
    $grid.Add_CellBeginEdit({ & $stoppen })
    $grid.Add_MouseDown({ & $stoppen })
    $grid.Add_CellEndEdit({
        & $merken
        $zelle = $grid.Rows[$_.RowIndex].Cells[2]
        try { [void](ConvertTo-Anteil ([string]$zelle.Value)); $zelle.Style.BackColor = [System.Drawing.Color]::Empty }
        catch { $zelle.Style.BackColor = [System.Drawing.Color]::MistyRose }
        & $zeigeGesamt
    })
    $grid.Add_UserDeletedRow({ & $merken; & $zeigeGesamt })

    $script:FensterZeit = $null
    $ok.Add_Click({
        foreach ($zeile in $grid.Rows) {
            if ($zeile.IsNewRow) { continue }
            try { [void](ConvertTo-Anteil ([string]$zeile.Cells[2].Value)) }
            catch { $zeile.Cells[2].Style.BackColor = [System.Drawing.Color]::MistyRose; $grid.CurrentCell = $zeile.Cells[2]; return }
        }
        try {
            $script:FensterZeit = ConvertTo-Zeitspanne $tb.Text
            $f.DialogResult = 'OK'
            $f.Close()
        } catch {
            $tb.BackColor = [System.Drawing.Color]::MistyRose
            $tb.SelectAll(); $tb.Focus()
        }
    })

    $f.Controls.AddRange(@($lbl, $lblGesamt, $tb, $ok, $spaeter, $cb) + $beschriftungen + $felder + @($grid))
    # Beim Schliessen ebenfalls merken (auch bei "Spaeter")
    $f.Add_FormClosing({ & $merken })
    $f.AcceptButton = $ok
    $f.CancelButton = $spaeter

    $bereich = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $f.Location = New-Object System.Drawing.Point(($bereich.Right - $f.Width - 12), ($bereich.Top + 12))
    $f.Add_Shown({
        $f.Activate(); $tb.Focus(); $tb.SelectAll()
        & $zeigeGesamt; $uhr.Start()
        if ($cb.Checked) { $ok.Text = 'Eintragen ({0})' -f $script:Rest; $timer.Start() }
    })

    $ergebnis = $f.ShowDialog()
    $timer.Stop(); $timer.Dispose()
    $uhr.Stop(); $uhr.Dispose()
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
        foreach ($pj in (Get-Projekte)) {
            $zeile = Find-Projektzeile $blatt $pj.Nummer '' -NichtAnlegen
            $anteil = if ($pj.Anteil) { $pj.Anteil } else { 'Rest' }
            if ($zeile) { Write-Host "Projekt:  $($pj.Nummer) ($anteil) in Zeile $zeile" }
            else { Write-Host "Projekt:  $($pj.Nummer) ($anteil) steht noch nicht im Blatt - wird in die erste freie Zeile geschrieben." }
        }
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
