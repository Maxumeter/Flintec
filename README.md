# Flintec Control Center

Control Center for Flintec DAD143 EtherCAT devices.

## Aktuelle Version

**v1.8.9**

### Änderungen in v1.8.9
- Fehlerbehandlung beim SDO-Schreiben korrigiert.
- `INTEGER32`-Parameter werden beim Schreiben nicht mehr durch den irreführenden Fehler `HTTP 400: Field type not found` blockiert.
- Der tatsächliche EtherCAT-/DAD143-Fehler wird jetzt angezeigt, z. B. `Ecat SDO: Data cannot be transferred (local control)`.
- Dadurch lassen sich TAC-/Calibrate-Enable-geschützte Kalibrierparameter wie `0x2300:11 Zero Range` korrekt diagnostizieren.
- Versionsnummer von **1.8.8** auf **1.8.9** erhöht.

Details: [RELEASE_NOTES_v1.8.9.md](RELEASE_NOTES_v1.8.9.md)

Geplanter Release-Installer:
`Flintec_ControlCenter_Setup_1.8.9.exe`
