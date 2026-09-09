# Flintec Control Center v1.8.9

## Fixed

### SDO parameter writes
- Fixed incorrect error handling in the SDO write path.
- In v1.8.8, a failed first SDO write attempt could mask the result of the fallback request and always show:
  `HTTP 400: Field type not found`
- v1.8.9 now reports the actual result of the SDO write request.
- Verified with DAD143 calibration parameter `0x2300:11 Zero Range`: the request now reaches EtherCAT and returns the device-side response when calibration access has not been enabled.

### DAD143 calibration diagnostics
- Improved visibility of device-side SDO aborts such as:
  `Ecat SDO: Data cannot be transferred (local control)`
- This makes TAC/Calibrate Enable protected parameters easier to diagnose.

## Version
- Application version: **1.8.9**
- Previous version: 1.8.8

## Build verification
- Windows x64 installer build tested structurally as a valid PE executable.
- Installer SHA-256:
  `5fb17120c7ac09fca896a78d5e20bc32aec387129cc29378d92da7c5d130a45f`

## Note
The repository currently contains release metadata/documentation only; the original application source code was not present in the repository when v1.8.9 was prepared.
