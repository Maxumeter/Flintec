#Requires -Version 5.1
# Entfernt die Aufgaben wieder. Die Excel-Datei bleibt unveraendert.
foreach ($name in 'Stundennachweis - Kommen', 'Stundennachweis - Gehen', 'Stundennachweis - Herunterfahren') {
    try {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction Stop
        Write-Host "Entfernt: $name"
    } catch {
        Write-Host "Nicht vorhanden: $name"
    }
}
Write-Host ''
Write-Host "Stempel und Protokoll liegen weiterhin unter $env:LOCALAPPDATA\Stundennachweis (kann geloescht werden)."
