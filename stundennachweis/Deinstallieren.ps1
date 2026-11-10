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
$startmenue = Join-Path ([Environment]::GetFolderPath('Programs')) 'Stundennachweis.lnk'
if (Test-Path -LiteralPath $startmenue) { Remove-Item -LiteralPath $startmenue -Force; Write-Host 'Entfernt: Verknuepfung im Startmenue' }
Write-Host ''
Write-Host 'Laeuft die App noch (Symbol neben der Uhr), dort per Rechtsklick auf Beenden klicken.'
Write-Host "Stempel und Protokoll liegen weiterhin unter $env:LOCALAPPDATA\Stundennachweis (kann geloescht werden)."
