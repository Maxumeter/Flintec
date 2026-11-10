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
$SkriptPfad = $PSCommandPath
$Inv = [Globalization.CultureInfo]::InvariantCulture
$DeKultur = [Globalization.CultureInfo]::GetCultureInfo('de-DE')

$DatenOrdner = Join-Path $env:LOCALAPPDATA 'Stundennachweis'
New-Item -ItemType Directory -Force -Path $DatenOrdner | Out-Null
$StempelDatei = Join-Path $DatenOrdner 'stempel.csv'
$StatusDatei  = Join-Path $DatenOrdner 'uebertragen.json'
$LogDatei     = Join-Path $DatenOrdner 'protokoll.log'
$FensterDatei = Join-Path $DatenOrdner 'fenster.json'
$WocheDatei   = Join-Path $DatenOrdner 'woche.json'
$TageProjekteDatei = Join-Path $DatenOrdner 'projekte-je-tag.json'   # Projektverteilung je Tag (Zwischenspeicher)
$Z_GESAMT = 32   # Zeile "Gesamtstunden" (Projekte + Urlaub/Feiertag/Krankheit ...)

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

function Get-Wochenanfang([datetime]$Tag) {
    return $Tag.Date.AddDays(-(([int]$Tag.DayOfWeek + 6) % 7))   # Montag
}

function Update-WochenCache($Mappe, [int]$Jahr) {
    # Merkt sich die Gesamtstunden (Zeile 32) der Tage dieser Woche, damit das Fenster
    # die Wochenstunden anzeigen kann, ohne Excel oeffnen zu muessen.
    $cache = @{}
    try {
        if (Test-Path -LiteralPath $WocheDatei) {
            $o = Get-Content -LiteralPath $WocheDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($pp in $o.PSObject.Properties) { $cache[$pp.Name] = [double]$pp.Value }
        }
    } catch { }
    $heute = (Get-Date).Date
    $montag = Get-Wochenanfang $heute
    for ($tag = $montag; $tag -le $heute; $tag = $tag.AddDays(1)) {
        if ($tag.Year -ne $Jahr) { continue }
        try {
            $blatt = Get-Blatt $Mappe $tag
            $spalte = Find-Spalte $blatt $tag
            $v = $blatt.Cells.Item($Z_GESAMT, $spalte).Value2
            $cache[$tag.ToString('yyyy-MM-dd')] = $(if ($v -is [double]) { [math]::Round($v, 2) } else { 0.0 })
        } catch { }
    }
    # Nur diese und letzte Woche behalten
    $grenze = $montag.AddDays(-7).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($cache.Keys | Sort-Object)) { if ($k -ge $grenze) { $o[$k] = $cache[$k] } }
    try { ($o | ConvertTo-Json) | Set-Content -LiteralPath $WocheDatei -Encoding UTF8 } catch { }
}

function Get-WochenstundenBisGestern {
    # Summe der gemerkten Gesamtstunden von Montag bis gestern
    $summe = 0.0
    try {
        if (-not (Test-Path -LiteralPath $WocheDatei)) { return 0.0 }
        $o = Get-Content -LiteralPath $WocheDatei -Raw -Encoding UTF8 | ConvertFrom-Json
        $heute = (Get-Date).Date
        $von = (Get-Wochenanfang $heute).ToString('yyyy-MM-dd')
        $bis = $heute.ToString('yyyy-MM-dd')
        foreach ($pp in $o.PSObject.Properties) {
            if ($pp.Name -ge $von -and $pp.Name -lt $bis) { $summe += [double]$pp.Value }
        }
    } catch { }
    return $summe
}

# ---------------------------------------------------------------- Ablaeufe

function Get-ProjektSchnappschuesse {
    $h = @{}
    try {
        if (Test-Path -LiteralPath $TageProjekteDatei) {
            $o = Get-Content -LiteralPath $TageProjekteDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($pp in $o.PSObject.Properties) {
                $h[$pp.Name] = @($pp.Value | ForEach-Object { [pscustomobject]@{ Nummer = [string]$_.Nummer; Bezeichnung = [string]$_.Bezeichnung; Anteil = [string]$_.Anteil } })
            }
        }
    } catch { }
    return $h
}

function Save-ProjektSchnappschuss([datetime]$Tag) {
    # Merkt sich die Projektverteilung des Tages lokal. Ist die Excel-Datei nicht erreichbar,
    # wird der Tag spaeter mit genau dieser Verteilung nachgetragen - auch wenn die Projekte
    # im Fenster inzwischen geaendert wurden.
    $h = Get-ProjektSchnappschuesse
    $h[$Tag.ToString('yyyy-MM-dd')] = @(Get-Projekte)
    $grenze = (Get-Date).Date.AddDays(-120).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($h.Keys | Sort-Object)) { if ($k -ge $grenze) { $o[$k] = @($h[$k]) } }
    try { ($o | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $TageProjekteDatei -Encoding UTF8 } catch { }
}

function Get-SyncKandidaten($Stempel, $Status) {
    # Alle Tage, deren Zeiten noch nicht (oder nicht aktuell) in der Excel-Datei stehen.
    # Es wird weit zurueckgeschaut, damit nichts verloren geht, wenn die Datei laenger nicht erreichbar war.
    $heute = (Get-Date).Date
    $tageZurueck = [math]::Max([int]$Einst.TageRueckwirkend, 60)
    $schnapp = Get-ProjektSchnappschuesse
    $aktuell = @()
    if ($Einst.ProjektstundenEintragen) { $aktuell = @(Get-Projekte) }
    $liste = @()
    for ($i = $tageZurueck; $i -ge 0; $i--) {
        $tag = $heute.AddDays(-$i)
        $key = $tag.ToString('yyyy-MM-dd')
        if ($Status[$key] -eq 'manuell' -or $Status[$key] -eq 'uebersprungen') { continue }
        $soll = Get-TagesZeiten $tag $Stempel
        if (-not $soll) { continue }

        $projekte = $aktuell
        $hatSchnapp = $schnapp.ContainsKey($key)
        if ($Einst.ProjektstundenEintragen -and $hatSchnapp) { $projekte = @($schnapp[$key]) }
        $sig = Get-ProjektSignatur $projekte

        $p = $null
        if ($soll.Ende -and $soll.Beginn) { $p = Get-PauseFuer ($soll.Ende - $soll.Beginn) }
        $wert = Get-StatusWert $soll.Beginn $soll.Ende $p
        # Neu verteilen, wenn sich die Projekte dieses Tages geaendert haben
        $sigNeu = (($tag -eq $heute -or $hatSchnapp) -and $soll.Beginn -and $soll.Ende -and $Status['sig-' + $key] -ne $sig)
        if ($Status[$key] -ne $wert -or $sigNeu) {
            $liste += [pscustomobject]@{ Soll = $soll; Pause = $p; Wert = $wert; Key = $key; Projekte = $projekte; Sig = $sig }
        }
    }
    return $liste
}

function Get-OffeneTage {
    # Tage vor heute, die wegen nicht erreichbarer Datei noch auf die Uebertragung warten
    try {
        $heute = (Get-Date).Date
        return @(Get-SyncKandidaten (Get-Stempel) (Get-Status) | Where-Object {
            $_.Soll.Tag -lt $heute -and -not ($Einst.WochenendeUeberspringen -and ($_.Soll.Tag.DayOfWeek -eq 'Saturday' -or $_.Soll.Tag.DayOfWeek -eq 'Sunday'))
        } | ForEach-Object { $_.Soll.Tag })
    } catch { return @() }
}

function Invoke-Sync {
    $stempel = Get-Stempel
    $status = Get-Status
    $heute = (Get-Date).Date

    $alleKandidaten = @(Get-SyncKandidaten $stempel $status)
    $nachgetragen = @()
    # Aktuelles Jahr immer pruefen, damit geaenderte Personalnummer/Name auch ohne neue Zeiten ankommen.
    $jahre = @($alleKandidaten | ForEach-Object { $_.Soll.Tag.Year }) + $heute.Year | Sort-Object -Unique
    foreach ($jahr in $jahre) {
        $gruppe = [pscustomobject]@{ Name = [string]$jahr }
        $pfad = Get-DateiFuerJahr $jahr
        $sitzung = $null
        $geaendert = $false
        try {
            # Nur Excel oeffnen, wenn es etwas zu tun gibt.
            $kandidaten = @($alleKandidaten | Where-Object { $_.Soll.Tag.Year -eq $jahr })
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

                Set-Tag $blatt $spalte $tag $k.Soll.Beginn $k.Soll.Ende $k.Pause $k.Projekte $status
                $status[$k.Key] = $k.Wert
                if ($k.Soll.Beginn -and $k.Soll.Ende) { $status['sig-' + $k.Key] = $k.Sig }
                if ($tag -lt $heute) { $nachgetragen += $text }
                $geaendert = $true
                Write-Protokoll ("{0}: eingetragen  Beginn {1}  Ende {2}  Pause {3}" -f $text, (Format-Zeit $k.Soll.Beginn), (Format-Zeit $k.Soll.Ende), (Format-Zeit $k.Pause))
            }
            Update-WochenCache $sitzung.Mappe $jahr
        } catch {
            $anzahl = @($kandidaten).Count
            if ($anzahl -gt 0) {
                Write-Protokoll ("Excel-Datei nicht verfuegbar ({0}). {1} Tag(e) bleiben lokal gespeichert und werden beim naechsten Mal uebertragen." -f $_.Exception.Message, $anzahl)
            } else {
                Write-Protokoll "FEHLER: $($_.Exception.Message)"
            }
            $nachgetragen = @()
        } finally {
            try { Close-Mappe $sitzung $geaendert } catch { Write-Protokoll "FEHLER beim Speichern: $($_.Exception.Message)"; $status = Get-Status; $geaendert = $false; $nachgetragen = @() }
        }
    }
    Save-Status $status
    if ($nachgetragen.Count -gt 0) { Write-Protokoll ('Nachgetragen: ' + ($nachgetragen -join ', ')) }
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
        Update-WochenCache $sitzung.Mappe $Datum.Year
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

function Start-Hintergrundsync {
    # Uebertraegt in einem eigenen, unsichtbaren Prozess in die Excel-Datei,
    # damit das Fenster dabei nicht haengt.
    try {
        $ps = Join-Path $PSHOME 'powershell.exe'
        $argumente = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Aktion Sync' -f $SkriptPfad
        Start-Process -FilePath $ps -ArgumentList $argumente -WindowStyle Hidden | Out-Null
    } catch {
        Write-Protokoll "FEHLER beim Starten der Uebertragung: $($_.Exception.Message)"
    }
}

function Start-App([string]$Modus) {
    # Laeuft den ganzen Tag im Hintergrund (Symbol im Infobereich neben der Uhr).
    # $Modus 'Pruefen': Fenster nur zeigen, wenn heute noch kein Arbeitsbeginn erfasst ist.
    # $Modus 'Zeigen' : Fenster immer zeigen.
    $evZeigen  = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisZeigen')
    $evPruefen = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisPruefen')
    $appMutex = New-Object System.Threading.Mutex($false, 'Local\StundennachweisApp')
    $hat = $false
    try { $hat = $appMutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $hat = $true }
    if (-not $hat) {
        # Laeuft schon - der laufenden App Bescheid geben und beenden.
        if ($Modus -eq 'Zeigen') { [void]$evZeigen.Set() } else { [void]$evPruefen.Set() }
        return
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Remove-AlteStempel
        Start-Hintergrundsync   # liegengebliebene Tage nachtragen

        $exe = Join-Path $PSScriptRoot 'Stundennachweis.exe'
        $symbol = [System.Drawing.SystemIcons]::Information
        if (Test-Path -LiteralPath $exe) { try { $symbol = [System.Drawing.Icon]::ExtractAssociatedIcon($exe) } catch { } }

        $tray = New-Object System.Windows.Forms.NotifyIcon
        $tray.Icon = $symbol
        $tray.Text = 'Stundennachweis'

        $script:Tag = (Get-Date).Date
        $script:Beenden = $false
        $script:HinweisGezeigt = $false

        $f = New-Object System.Windows.Forms.Form
        $f.Text = 'Stundennachweis'
        $f.FormBorderStyle = 'FixedSingle'
        $f.MaximizeBox = $false; $f.MinimizeBox = $true   # minimieren erlaubt
        $f.TopMost = $true
        $f.ShowInTaskbar = $true
        $f.StartPosition = 'Manual'
        $f.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $fenstergroesse = New-Object System.Drawing.Size(460, 372)
        $f.ClientSize = $fenstergroesse

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(12, 12)
        $lbl.AutoSize = $true

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(12, 46)
        $tb.Width = 70
        $tb.TextAlign = 'Center'

        # Gesamtstunden des Tages bis jetzt (Beginn bis aktuelle Uhrzeit, Pause abgezogen)
        $lblGesamt = New-Object System.Windows.Forms.Label
        $lblGesamt.Location = New-Object System.Drawing.Point(240, 12)
        $lblGesamt.Size = New-Object System.Drawing.Size(208, 22)
        $lblGesamt.TextAlign = 'TopRight'
        $lblGesamt.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($lblGesamt, 'Arbeitszeit heute von Beginn bis jetzt, Pause bereits abgezogen')
        $lblWoche = New-Object System.Windows.Forms.Label
        $lblWoche.Location = New-Object System.Drawing.Point(240, 84)
        $lblWoche.Size = New-Object System.Drawing.Size(208, 22)
        $lblWoche.TextAlign = 'TopRight'
        $lblWoche.Font = $lblGesamt.Font
        $tip.SetToolTip($lblWoche, 'Gesamtstunden dieser Woche (Mo bis heute) laut Stundennachweis plus die Stunden von heute bis jetzt')

        $zeigeGesamt = {
            try { $bz = ConvertTo-Zeitspanne $tb.Text } catch { $bz = $null }
            $std = 0.0
            if ($null -ne $bz) {
                $anw = (Get-Minutengenau (Get-Date)) - $bz
                if ($anw.TotalMinutes -gt 0) {
                    $arbeit = $anw - (Get-PauseFuer $anw)
                    if ($arbeit.TotalMinutes -gt 0) { $std = $arbeit.TotalHours }
                }
                $lblGesamt.Text = 'Heute: {0} h ({1})' -f (Format-Zeit ([TimeSpan]::FromHours($std))), $std.ToString('0.00', $DeKultur)
            } else {
                $lblGesamt.Text = 'Heute: -'
            }
            $woche = (Get-WochenstundenBisGestern) + [math]::Round($std, 2)
            $lblWoche.Text = 'Woche: {0} h ({1})' -f (Format-Zeit ([TimeSpan]::FromHours($woche))), $woche.ToString('0.00', $DeKultur)
            $hinweis = 'Stundennachweis - Heute {0} h, Woche {1} h' -f (Format-Zeit ([TimeSpan]::FromHours($std))), (Format-Zeit ([TimeSpan]::FromHours($woche)))
            if ($hinweis.Length -gt 63) { $hinweis = $hinweis.Substring(0, 63) }   # Windows erlaubt hoechstens 63 Zeichen
            $tray.Text = $hinweis

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
        $uhr.Add_Tick({
            # Neuer Tag (PC lief ueber Nacht)? Dann Fenster auf heute umstellen.
            if ((Get-Date).Date -ne $script:Tag -and -not $f.Visible) { & $ladeTag }
            & $zeigeGesamt
            & $pruefeOffen
        })

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'Eintragen'
        $ok.Location = New-Object System.Drawing.Point(92, 44)
        $ok.Size = New-Object System.Drawing.Size(100, 30)

        $spaeter = New-Object System.Windows.Forms.Button
        $spaeter.Text = 'Sp' + [char]0x00E4 + 'ter'
        $spaeter.Location = New-Object System.Drawing.Point(198, 44)
        $spaeter.Size = New-Object System.Drawing.Size(90, 30)

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
            $t.Width = 344
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
        $grid.Size = New-Object System.Drawing.Size(436, 130)
        $grid.AllowUserToAddRows = $true
        $grid.AllowUserToDeleteRows = $true
        $grid.AllowUserToResizeRows = $false
        $grid.RowHeadersWidth = 28
        $grid.AutoSizeColumnsMode = 'None'
        $grid.AllowUserToResizeColumns = $false
        $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
        $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
        [void]$grid.Columns.Add('Nummer', 'Auftragsnr.')
        [void]$grid.Columns.Add('Bezeichnung', 'Bezeichnung')
        [void]$grid.Columns.Add('Anteil', 'Anteil')
        [void]$grid.Columns.Add('Stunden', 'Stunden')
        # Feste Spaltenbreiten (Pixel). Zusammen mit Zeilenkopf (28) und Bildlaufleiste passt alles ohne waagerechtes Scrollen.
        $grid.Columns[0].Width = 100   # Auftragsnr.  z. B. AB002500473
        $grid.Columns[1].Width = 170   # Bezeichnung
        $grid.Columns[2].Width = 55    # Anteil       z. B. 2:30 / 40%
        $grid.Columns[3].Width = 62    # Stunden      z. B. 12,50
        foreach ($spalte in $grid.Columns) { $spalte.SortMode = 'NotSortable' }
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
            Save-ProjektSchnappschuss (Get-Date).Date
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

        $ok.Add_Click({
            foreach ($zeile in $grid.Rows) {
                if ($zeile.IsNewRow) { continue }
                try { [void](ConvertTo-Anteil ([string]$zeile.Cells[2].Value)) }
                catch { $zeile.Cells[2].Style.BackColor = [System.Drawing.Color]::MistyRose; $grid.CurrentCell = $zeile.Cells[2]; return }
            }
            try { $zeit = ConvertTo-Zeitspanne $tb.Text } catch {
                $tb.BackColor = [System.Drawing.Color]::MistyRose
                $tb.SelectAll(); $tb.Focus()
                return
            }
            $tb.BackColor = [System.Drawing.SystemColors]::Window
            $beginnZeit = $script:Tag.Add($zeit)
            Add-Stempel 'Beginn' $beginnZeit
            Write-Protokoll ('Arbeitsbeginn {0} erfasst.' -f $beginnZeit.ToString('HH:mm'))
            & $merken
            Start-Hintergrundsync   # traegt auch liegengebliebene Tage nach
            $script:LetzterVersuch = Get-Date
            & $zeigeGesamt
            & $verstecken
        })
        $spaeter.Add_Click({ & $verstecken })


        $f.Icon = $symbol
        # Hinweis, wenn Tage noch auf die Uebertragung in die Excel-Datei warten
        $lblOffen = New-Object System.Windows.Forms.Label
        $lblOffen.Location = New-Object System.Drawing.Point(12, 346)
        $lblOffen.Size = New-Object System.Drawing.Size(436, 20)
        $lblOffen.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblOffen.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $tip.SetToolTip($lblOffen, 'Die Excel-Datei war nicht erreichbar. Die Zeiten sind lokal gespeichert und werden automatisch nachgetragen, sobald die Datei wieder verfuegbar ist.')

        $script:OffenVorher = @()
        $script:LetzterVersuch = Get-Date
        $pruefeOffen = {
            $offen = @(Get-OffeneTage)
            if ($offen.Count -gt 0) {
                $lblOffen.Text = ('Noch nicht in Excel: ' + (($offen | ForEach-Object { $_.ToString('ddd dd.MM.', $DeKultur) }) -join ', ') + ' - wird nachgetragen')
                # Alle 30 Minuten einen neuen Versuch starten
                if (((Get-Date) - $script:LetzterVersuch).TotalMinutes -ge 30) { $script:LetzterVersuch = Get-Date; Start-Hintergrundsync }
            } else {
                $lblOffen.Text = ''
                if ($script:OffenVorher.Count -gt 0) {
                    $tray.ShowBalloonTip(5000, 'Stundennachweis',
                        ('In die Excel-Datei nachgetragen: ' + (($script:OffenVorher | ForEach-Object { $_.ToString('ddd dd.MM.', $DeKultur) }) -join ', ')), 'Info')
                }
            }
            $script:OffenVorher = $offen
        }

        $f.Controls.AddRange(@($lbl, $lblGesamt, $lblWoche, $tb, $ok, $spaeter, $cb, $lblOffen) + $beschriftungen + $felder + @($grid))
        $f.AcceptButton = $ok
        $f.CancelButton = $spaeter

        $ladeTag = {
            $script:Tag = (Get-Date).Date
            $lbl.Text = 'Arbeitsbeginn ' + $script:Tag.ToString('ddd dd.MM.', $DeKultur) + ':'
            $vorher = Get-TagesZeiten $script:Tag (Get-Stempel)
            if ($vorher -and $vorher.Beginn) { $tb.Text = Format-Zeit $vorher.Beginn }
            else { $tb.Text = Format-Zeit (Get-Minutengenau (Get-Date)) }
            $tb.BackColor = [System.Drawing.SystemColors]::Window
        }
        $brauchtBeginn = {
            $h = (Get-Date).Date
            $wochenende = ($h.DayOfWeek -eq 'Saturday' -or $h.DayOfWeek -eq 'Sunday')
            (-not (Test-BeginnErfasst $h)) -and -not ($Einst.WochenendeUeberspringen -and $wochenende)
        }
        $zeigen = {
            if (-not $f.Visible) {
                # Neuer Tag oder noch kein Beginn: Vorschlag auf die aktuelle Uhrzeit setzen
                if ((Get-Date).Date -ne $script:Tag -or (& $brauchtBeginn)) { & $ladeTag }
            }
            # Reihenfolge wichtig: erst anzeigen, dann aus dem minimierten Zustand holen,
            # dann Groesse und Position setzen - sonst bleibt nur die kleine Titelleiste uebrig.
            $f.Show()
            if ($f.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal) { $f.WindowState = [System.Windows.Forms.FormWindowState]::Normal }
            $f.ClientSize = $fenstergroesse
            $bereich = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $f.Location = New-Object System.Drawing.Point(($bereich.Right - $f.Width - 12), ($bereich.Top + 12))
            $f.BringToFront()
            $f.Activate()
            $tb.Focus(); $tb.SelectAll()
            & $zeigeGesamt
            if ($cb.Checked -and (& $brauchtBeginn)) {
                $script:Rest = $sekunden; $ok.Text = 'Eintragen ({0})' -f $script:Rest; $timer.Start()
            }
        }
        $verstecken = {
            & $stoppen
            & $merken
            $f.Hide()
            if (-not $script:HinweisGezeigt) {
                $script:HinweisGezeigt = $true
                $tray.ShowBalloonTip(4000, 'Stundennachweis',
                    ('L' + [char]0x00E4 + 'uft im Hintergrund weiter. Klick auf das Symbol ' + [char]0x00F6 + 'ffnet das Fenster wieder.'), 'Info')
            }
        }

        # Minimieren und X schicken das Fenster in den Hintergrund (Infobereich)
        $f.Add_Resize({
            if ($f.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
                & $verstecken
                # Fuer das naechste Oeffnen gleich wieder auf normale Groesse stellen
                $f.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            }
        })
        $f.Add_FormClosing({
            if ($_.CloseReason -eq 'UserClosing' -and -not $script:Beenden) { $_.Cancel = $true; & $verstecken }
            else { & $merken }
        })

        # Menue am Symbol
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        [void]$menu.Items.Add(([char]0x00D6 + 'ffnen'), $null, { & $zeigen })
        [void]$menu.Items.Add(('Excel-Datei ' + [char]0x00F6 + 'ffnen'), $null, {
            try { Start-Process -FilePath (Get-DateiFuerJahr (Get-Date).Year) } catch { Write-Protokoll "FEHLER: $($_.Exception.Message)" }
        })
        [void]$menu.Items.Add('Protokoll anzeigen', $null, { try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$LogDatei`"" } catch { } })
        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        [void]$menu.Items.Add('Beenden', $null, {
            $script:Beenden = $true
            & $merken
            $tray.Visible = $false
            [System.Windows.Forms.Application]::Exit()
        })
        $tray.ContextMenuStrip = $menu
        # Einfach- oder Doppelklick (linke Maustaste) auf das Symbol holt das Fenster zurueck
        $tray.Add_MouseClick({ if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $zeigen } })
        $tray.Add_MouseDoubleClick({ if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $zeigen } })
        $tray.Add_BalloonTipClicked({ & $zeigen })
        $tray.Visible = $true

        # Nachrichten von weiteren Starts (Entsperren, Doppelklick auf die EXE)
        $post = New-Object System.Windows.Forms.Timer
        $post.Interval = 500
        $post.Add_Tick({
            if ($evZeigen.WaitOne(0)) { & $zeigen }
            if ($evPruefen.WaitOne(0) -and (& $brauchtBeginn)) { & $zeigen }
        })
        $post.Start()

        & $ladeTag
        & $zeigeGesamt
        & $pruefeOffen
        $uhr.Start()
        if ($Modus -eq 'Zeigen' -or (& $brauchtBeginn)) { & $zeigen }

        [System.Windows.Forms.Application]::Run()

        $post.Stop(); $uhr.Stop(); $timer.Stop()
        $tray.Visible = $false
        $tray.Dispose()
        $f.Dispose()
    } finally {
        $appMutex.ReleaseMutex()
        $appMutex.Dispose()
        $evZeigen.Dispose(); $evPruefen.Dispose()
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
        try { $Einst = Get-Einstellungen; Save-ProjektSchnappschuss (Get-Date).Date } catch { }
    }
    if ($OhneExcel) { return }

    $Einst = Get-Einstellungen
    # Die App laeuft den ganzen Tag - ausserhalb der Sperre fuer die Excel-Uebertragung.
    if ($Aktion -eq 'Kommen') { Start-App 'Pruefen'; return }
    if ($Aktion -eq 'Beginn') { Start-App 'Zeigen'; return }

    try { $hatMutex = $mutex.WaitOne([TimeSpan]::FromMinutes(5)) } catch [System.Threading.AbandonedMutexException] { $hatMutex = $true }
    if (-not $hatMutex) { Write-Protokoll 'Ein anderer Lauf ist noch aktiv - abgebrochen.'; return }

    switch ($Aktion) {
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
