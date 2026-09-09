# Flintec Control Center

Control Center for Flintec DAD143 EtherCAT devices.

## Latest version

**v1.8.9**

### Highlights
- Fixed SDO write error reporting for DAD143 parameters.
- `INTEGER32` SDO writes now reach the EtherCAT/Flintec device instead of stopping at the previous `HTTP 400: Field type not found` error.
- The application now exposes the actual EtherCAT SDO abort/error returned by the device, e.g. `Ecat SDO: Data cannot be transferred (local control)`.
- Version number updated from v1.8.8 to v1.8.9.

See [RELEASE_NOTES_v1.8.9.md](RELEASE_NOTES_v1.8.9.md) for details.
