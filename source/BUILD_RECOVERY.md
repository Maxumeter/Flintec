# Build v1.9.4 / v1.9.5

Dieses Paket enthält den vollständigen Quellcode des neuen C#-Launchers und die wiederhergestellte Weboberfläche. Die Anwendung ist eine dokumentierte Wiederherstellung aus der mitgelieferten ursprünglichen v1.9.3-Binärdatei, kein Build aus ursprünglichen Go-Quellen.

Im entpackten Hauptverzeichnis:

1. Python 3 mit `pefile==2024.8.26` und `capstone==5.0.9` bereitstellen: `python -m pip install --target diagnostics/python-libs pefile==2024.8.26 capstone==5.0.9`.
2. Anwendung bauen: `python diagnostics/build_1_9_4.py`.
3. Oberfläche prüfen (Node.js erforderlich): `node --check source/app.js` und `node diagnostics/test_frontend.cjs`.
4. Unter Windows den Launcher bauen: `powershell -ExecutionPolicy Bypass -File source/updater/build.ps1`. Das Skript verwendet den .NET-Framework-C#-Compiler und führt die lokalen Updater-Tests aus. Resultat: `release-1.9.4-autoupdate/Flintec_ControlCenter_1.9.4_AutoUpdate.exe`.

Die Basis-EXE ist über SHA-256 festgelegt; das Build-Skript bricht bei Abweichungen ab. Die neue Anwendung und der Launcher enthalten keine Benutzeranmeldedaten. Der ursprüngliche Kalibrier-Timeout ist nicht behoben.

## v1.9.5

`diagnostics/installed-1.9.3.exe` aus dem v1.9.4-Recovery-ZIP übernehmen, dann:

1. `python diagnostics/build_1_9_5.py` (wendet zusätzlich `diagnostics/transfer-fixes.js` an; Ergebnis in `release-1.9.5/`).
2. `node --check source/app.js`, `node diagnostics/test_frontend.cjs` und `node diagnostics/test_transfer.cjs`.

`source/app.js` und `source/index.html` werden vom Build erzeugt und entsprechen dem zuletzt gebauten Stand (v1.9.5).
