# Flintec Control Center v1.9.5

## Behoben: Profil-/Kalibrierübertragung auf mehrere DAD143

In v1.9.4 schlug das Übertragen von Kalibrierparametern auf andere Waagen bei jedem `0x2300`-Wert mit
`Ecat SDO: Data cannot be transferred (local control)` fehl. Ursache: Die Übertragung hat die TAC-Freigabe
(`0x2300:03` lesen und zurückschreiben) nicht direkt vor dem jeweiligen Schreibzugriff gesetzt, wie es der
Kalibrierdialog tut.

### Änderungen
- **Kalibrierbefehle werden nie übertragen.** `0x2300:04` (Verstärkung kalibrieren) und `0x2300:0A`
  (Nullpunkt kalibrieren) lösen auf dem Gerät eine Kalibrierung mit der aktuell aufliegenden Last aus.
  Ein Übertragen hätte jede Zielwaage falsch kalibriert. Zusätzlich werden `0x2300`-Einträge, deren Name
  auf calib/gain/span/adjust/command/execute lautet, ausgeschlossen. Das gilt auch für Profile aus älteren Versionen.
- **Der Nullpunkt der Waage wird nicht übertragen:** `0x2300:02 Absolute zero` ist waagenspezifisch und bleibt auf der Zielwaage
  erhalten (bzw. wird durch die optionale Nullpunkt-Kalibrierung neu gesetzt).
- Vor jedem übertragenen `0x2300`-Einstellwert wird die TAC-Freigabe frisch gesetzt (gleicher Ablauf wie im Kalibrierdialog).
- Die Bestätigung listet die zu schreibenden und die ausgeschlossenen Kalibrierparameter mit Namen auf.
- Nach dem ersten Fehler im Kalibrierblock werden die übrigen Kalibrierwerte dieses Geräts übersprungen statt blind weiterzuschreiben.
- Alle geschriebenen Werte werden zurückgelesen und verglichen; Abweichungen erscheinen im Ergebnis.
- **Dauerhaftes Speichern je Parametergruppe (Handbuch 9.12):** Bisher wurde – auch in allen Vorversionen – nur `0x2004:02` (CS, Kalibrierung)
  ausgelöst; übertragene Konfigurationswerte (Filter, Stillstand, Trigger …) waren nach einem Neustart weg. Jetzt wird je geschriebener Gruppe
  gespeichert: Setup `0x2004:03` (WP), Füllparameter `0x2004:04`, Sollwerte `0x2004:05` (SS), Analogausgang `0x2004:01` (AS, nur DAD143.1,
  Fehler nur als Hinweis) und zuletzt Kalibrierung `0x2004:02` (CS, erhöht TAC). Nur auf Geräten ohne Fehler, Abweichungen oder übersprungene Werte.
- **Optional Nullpunkt kalibrieren:** Nach einer eigenen Rückfrage („Alle Zielwaagen müssen vollständig entlastet sein“) wird auf jeder
  Zielwaage nach den Einstellwerten der Nullpunkt kalibriert (`0x2300:0A` = 0 mit TAC-Freigabe, wie im Kalibrierdialog),
  danach erst im EEPROM gespeichert. Der Vergleich der Werte läuft vor dem Nullpunkt, weil dieser `0x2300:02 Absolute zero`
  mit dem eigenen Wert der Zielwaage überschreibt; das Ergebnis zeigt die Änderung (z. B. `Absolute zero 1219→1313`). Die Verstärkung wird nie verändert. Schlägt der Nullpunkt fehl, wird nicht gespeichert.
- **Nullpunkt nur bei ruhiger Waage:** Das DAD143 nimmt „Calibrate Zero“ nur an, wenn das Signal während NT ms (`0x2100:0B`) um höchstens
  NR Teilungen (`0x2100:0A`) schwankte (Handbuch 9.3). Vor `0x2300:0A` wird das Brutto deshalb mindestens NT + 0,5 s (min. 2 s) gelesen; die
  zulässige Spanne ist min(NR, 2) Teilungen (Teilung = `0x2300:0C` × 10^-`0x2300:0B`). Sonst wird der Befehl nicht gesendet.
- **Diagnose bei abgelehntem Nullpunkt:** Vor dem Nullpunkt und nach einer Ablehnung werden Qualifier (`0x2900:0D`, Bits laut Handbuch 14.1.6),
  Gerätestatus (`0x2900:0A`), interner mV/V-Wert (`0x2900:12`) und A/D-Wert (`0x2900:07`) gelesen, ins Log geschrieben und an die Fehlermeldung angehängt.
- **Fehlerhafte Waagen werden übersprungen:** Antwortet ein Gerät nicht (`Ecat: Timeout`), bekommt es keine weiteren Anfragen; lehnt es den
  Nullpunkt ab oder ist es nicht ruhig, wird es nicht gespeichert. Die Übertragung läuft jeweils mit den übrigen Waagen weiter. Nur wenn in einem
  Lauf drei Waagen nicht auf die Nullpunkt-Kalibrierung antworten, wird der Rest nicht mehr bearbeitet (allgemeines Problem).
- **Fehlgeschlagene erneut übertragen:** Neuer Knopf neben „Auf ausgewählte Geräte übertragen“. Er wiederholt die Übertragung nur für die
  Waagen, die im letzten Lauf fehlgeschlagen oder übersprungen wurden (mit denselben Rückfragen).
- Kein zweiter Bus-Scan und keine neue Dashboard-Abfrage, solange die vorherige noch läuft (bei hängendem Master stauten sich sonst Anfragen über Minuten).
- Die Startzeile im Log enthält eine Build-Kennung (`Version 1.9.5 (Build xxxxxxxx)`), damit Testläufe eindeutig einer EXE zugeordnet werden können.
- Das Log zeigt während der Übertragung `Kalibrierstatus=Profilübertragung <Adresse>` statt des Status der letzten Kalibrierung.

#### Filter-Optimierung (neu, Reiter Kalibrierung)
- Misst bei leerer Waage alle Filterstufen FL 1–8 (`0x2100:04`) im gewählten Modus FM (`0x2100:09`, IIR/FIR) in zwei Durchgängen (auf-/absteigend),
  jeweils Spanne und Standardabweichung in Teilungen, und schlägt den schwächsten (schnellsten) Filter vor, der das Ziel einhält.
- Bewertung: Unruhe = Spanne P5–P95 nach Abzug der linearen Drift; maßgeblich ist der bessere der beiden Durchgänge (Störung im anderen wird
  mit „!“ markiert), Drift wird getrennt angezeigt. Erreicht keine Stufe das Ziel, wird nichts vorgeschlagen und nichts geändert, mit Hinweis
  auf die wahrscheinliche Ursache (langsame Drift bzw. mechanisch/elektrisch). Hardware 1022: Einzeldurchgänge mit 48–51 d gegenüber 3–7 d im
  anderen Durchgang, selbst FL 8 bei 5 d – die erste Fassung hatte daraus fälschlich FL 8 vorgeschlagen.
- Hardware 07.10.2026 (Build 5f11e164): Filter-Optimierung auf 1014 und anschließend auf 1022 sauber durchgelaufen.
- Hardware 07.10.2026 (Build 5f11e164): Übertragung mit Nullpunkt auf 1016–1049: Nullpunkt bei allen ruhigen Waagen angenommen und gespeichert; 1022 nicht ruhig (danach per Kalibrierdialog kalibriert); 1023, 1030, 1037 lehnten den Nullpunkt trotz ruhigem Brutto mit `General error` ab.
- Werte werden nur flüchtig geschrieben; Übernahme mit Speichern (`0x2004:03`) erst nach Bestätigung, sonst wird die vorherige Einstellung
  wiederhergestellt. Einschwingzeiten und Grenzfrequenzen laut Handbuch-Tabellen 9.4.2.

## Wichtig
Der Nullpunkt kann bei leeren Zielwaagen direkt bei der Übertragung kalibriert werden. Das Kalibriergewicht (Verstärkung)
muss bei Bedarf auf **jeder Waage einzeln** mit dem Kalibrierdialog kalibriert werden. Eine Kalibrierung lässt sich nicht von einer Wägezelle auf eine andere kopieren.

## Build und Prüfung
- Gleiche Recovery-Methode wie v1.9.4 (Patch der geprüften v1.9.3-EXE, Oberfläche als PE-Abschnitt); der Build von v1.9.4 wurde vorab bitgenau reproduziert.
- Maschinen-Code unverändert gegenüber v1.9.4. Geändert sind nur die Weboberfläche und die Versionsangaben.
- `node diagnostics/test_frontend.cjs` (bestehende Regressionstests) bestanden.
- Neu: `node diagnostics/test_transfer.cjs` mit simuliertem DAD143, das ohne vorherige TAC-Freigabe mit „local control“ ablehnt. Der Test schlägt mit der v1.9.4-Oberfläche fehl und besteht mit v1.9.5.
- Hardware 07.10.2026: Übertragung auf 1014 ohne Nullpunkt-Schritt erfolgreich (15 OK, 0 Fehler, EEPROM gespeichert).
- Hardware 07.10.2026: Nullpunkt-Schritt auf 1014 erfolgreich; die Vorversion meldete danach fälschlich `0x2300:0x02 Absolute zero: 1313≠1219` und speicherte nicht (behoben).
- Hardware 07.10.2026: Übertragung ohne Nullpunkt auf 1015, 1016, 1020 erfolgreich und gespeichert (mit Build beb65b79, damals noch inkl. Absolute zero der Quellwaage).
- Hardware 07.10.2026: Auf 1015 lief die Nullpunkt-Kalibrierung (`0x2300:0A`) in einen Timeout; danach antwortete das Gerät nicht mehr und stand auf SAFE-OP. Um 11:10 hing auch 1014, die den Befehl zuvor dreimal angenommen hatte; vermutete Ursache: Messwert nicht in Ruhe. Daher die Stillstandsprüfung.
- Hardware 07.10.2026 (Build 2b35233e): Nullpunkt mit Stillstandsprüfung auf 1014, 1015, 1016, 1021, 1022 erfolgreich und gespeichert (Absolute zero 1312, 1317, 1517, 1553, 1874). 1023 lehnte den Nullpunkt nach 33 ms mit `General error` ab. Erst mit einer einzelnen Waage prüfen, bevor auf alle übertragen wird.

SHA-256 `Flintec_ControlCenter_1.9.5_Portable.exe` / `Flintec_ControlCenter_App.exe`:
`d7801f66dcddf461f0b90b8792b55e54986210223895d70ee4e67bb903a112d1`
