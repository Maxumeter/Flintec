#Requires -Version 5.1
# Richtet die automatische Zeiterfassung ein (zwei Aufgaben in der Windows-Aufgabenplanung).
# Aufruf: Doppelklick auf "Installieren.cmd"

$ErrorActionPreference = 'Stop'
$skript = Join-Path $PSScriptRoot 'Stundennachweis.ps1'
if (-not (Test-Path -LiteralPath $skript)) { throw "Stundennachweis.ps1 nicht gefunden in $PSScriptRoot" }

# Aus dem Internet geladene Dateien freigeben, sonst blockiert Windows sie.
Get-ChildItem -LiteralPath $PSScriptRoot -File | Unblock-File -ErrorAction SilentlyContinue

$benutzer = "$env:USERDOMAIN\$env:USERNAME"
$ns = 'Root/Microsoft/Windows/TaskScheduler'

function New-Aktion([string]$Argumente) {
    New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $PSScriptRoot `
        -Argument ("-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$skript`" $Argumente")
}

function New-SitzungsTrigger([int]$Art) {
    # 8 = Arbeitsstation entsperrt
    $klasse = Get-CimClass -Namespace $ns -ClassName MSFT_TaskSessionStateChangeTrigger
    $t = New-CimInstance -CimClass $klasse -ClientOnly
    $t.StateChange = $Art
    $t.UserId = $benutzer
    $t.Enabled = $true
    return $t
}

function New-HerunterfahrenTrigger {
    # Ereignis 1074 (User32) = Benutzer faehrt herunter / startet neu
    $klasse = Get-CimClass -Namespace $ns -ClassName MSFT_TaskEventTrigger
    $t = New-CimInstance -CimClass $klasse -ClientOnly
    $t.Subscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''User32''] and EventID=1074]]</Select></Query></QueryList>'
    $t.Enabled = $true
    return $t
}

$principal = New-ScheduledTaskPrincipal -UserId $benutzer -LogonType Interactive -RunLevel Limited
$einst = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances Queue
$einstKurz = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 3) -MultipleInstances Parallel

$aufgaben = @(
    @{ Name = 'Stundennachweis - Kommen'
       Beschreibung = 'Merkt sich beim Anmelden/Entsperren die Uhrzeit und traegt den Arbeitsbeginn ein.'
       Trigger = @((New-ScheduledTaskTrigger -AtLogOn -User $benutzer), (New-SitzungsTrigger 8))
       Aktion = (New-Aktion '-Aktion Kommen'); Einst = $einst },
    @{ Name = 'Stundennachweis - Herunterfahren'
       Beschreibung = 'Traegt beim Herunterfahren das Arbeitsende ein (klappt das nicht mehr, beim naechsten Anmelden).'
       Trigger = @(New-HerunterfahrenTrigger)
       Aktion = (New-Aktion '-Aktion Herunterfahren'); Einst = $einstKurz }
)

# Fruehere Version hat beim Sperren reagiert - diese Aufgabe entfernen.
try { Unregister-ScheduledTask -TaskName 'Stundennachweis - Gehen' -Confirm:$false -ErrorAction Stop; Write-Host 'Entfernt: Stundennachweis - Gehen (Sperren)' } catch { }

$fehler = 0
foreach ($a in $aufgaben) {
    try {
        Register-ScheduledTask -TaskName $a.Name -Description $a.Beschreibung -Trigger $a.Trigger `
            -Action $a.Aktion -Principal $principal -Settings $a.Einst -Force | Out-Null
        Write-Host "OK      $($a.Name)" -ForegroundColor Green
    } catch {
        $fehler++
        Write-Host "FEHLER  $($a.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ''
Write-Host 'Pruefe den Zugriff auf die Excel-Datei ...'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $skript -Aktion Test
Write-Host ''
if ($fehler -eq 0) {
    Write-Host 'Fertig. Ab jetzt werden Beginn und Ende automatisch eingetragen.' -ForegroundColor Green
    Write-Host 'Das Arbeitsende wird beim Herunterfahren eingetragen.'
} else {
    Write-Host 'Nicht alle Aufgaben konnten angelegt werden. Siehe Fehlermeldungen oben.' -ForegroundColor Yellow
}
