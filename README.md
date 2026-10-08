# Flintec Control Center

Control Center for Flintec DAD143 EtherCAT devices.

## Aktuelle Version

**v1.9.5**

- Kalibrierübertragung auf mehrere DAD143 korrigiert: TAC-Freigabe vor jedem `0x2300`-Wert, Kalibrierbefehle `0x2300:04`/`0x2300:0A` werden nie übertragen, Zurücklesen optionale Nullpunkt-Kalibrierung leerer Zielwaagen und optionales EEPROM-Speichern.

Details: [RELEASE_NOTES_v1.9.5.md](RELEASE_NOTES_v1.9.5.md)

## Build

Siehe [source/BUILD_RECOVERY.md](source/BUILD_RECOVERY.md). Die Basis-EXE `diagnostics/installed-1.9.3.exe` liegt im Release-Asset `Flintec_ControlCenter_1.9.4_RecoverySources.zip` und ist nicht im Repository.

## Ältere Version v1.8.9


### Änderungen in v1.8.9
- Fehlerbehandlung beim SDO-Schreiben korrigiert.
- `INTEGER32`-Parameter werden beim Schreiben nicht mehr durch den irreführenden Fehler `HTTP 400: Field type not found` blockiert.
- Der tatsächliche EtherCAT-/DAD143-Fehler wird jetzt angezeigt, z. B. `Ecat SDO: Data cannot be transferred (local control)`.
- Dadurch lassen sich TAC-/Calibrate-Enable-geschützte Kalibrierparameter wie `0x2300:11 Zero Range` korrekt diagnostizieren.
- Versionsnummer von **1.8.8** auf **1.8.9** erhöht.

Details: [RELEASE_NOTES_v1.8.9.md](RELEASE_NOTES_v1.8.9.md)

Geplanter Release-Installer:
`Flintec_ControlCenter_Setup_1.8.9.exe`
