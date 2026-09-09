# Flintec Control Center v1.8.9

## Änderungen

### SDO-Schreibzugriffe
- Fehlerbehandlung im SDO-Write-Pfad korrigiert.
- In v1.8.8 konnte der Fehler des ersten SDO-Schreibversuchs den tatsächlichen Rückgabewert des nachfolgenden Schreibversuchs überdecken.
- Dadurch wurde beim Schreiben von DAD143-Parametern häufig nur die irreführende Meldung
  `HTTP 400: Field type not found`
  angezeigt.
- v1.8.9 zeigt jetzt die tatsächliche Antwort des EtherCAT-SDO-Zugriffs an.

### DAD143 / Kalibrierparameter
- `INTEGER32`-Schreibzugriffe erreichen jetzt den EtherCAT-/Flintec-Teilnehmer.
- Getestet mit `0x2300:11 Zero Range`.
- Bei gesperrten Kalibrierparametern wird jetzt die echte Geräteantwort sichtbar, z. B.
  `Ecat SDO: Data cannot be transferred (local control)`.
- Dadurch ist klar erkennbar, wenn vor dem Schreiben eines geschützten Kalibrierparameters zunächst `0x2300:03 Calibrate Enable` mit dem aktuellen TAC freigegeben werden muss.

### Diagnose
- Verbesserte Fehlermeldungen für EtherCAT CoE / SDO.
- Gerätefehler werden nicht mehr durch einen vorherigen HTTP-Fehler verdeckt.
- Schnellere Fehlersuche bei DAD143-Inbetriebnahme, Kalibrierung und Parameteränderungen.

### Version
- Neue Version: **1.8.9**
- Vorherige Version: **1.8.8**

## Installer

`Flintec_ControlCenter_Setup_1.8.9.exe`

SHA-256:

`5fb17120c7ac09fca896a78d5e20bc32aec387129cc29378d92da7c5d130a45f`

## Hinweis
Der Build wurde als Windows-x64-PE strukturell geprüft. Ein Live-Test gegen eine reale ctrlX-/EtherCAT-Anlage ist außerhalb der Zielanlage nicht möglich.
