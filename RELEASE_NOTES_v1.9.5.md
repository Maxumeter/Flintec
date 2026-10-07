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
- Optional dauerhaftes Speichern im EEPROM (`0x2004:02`), aber nur auf Geräten ohne Schreibfehler, übersprungene Werte oder Abweichungen.
- **Optional Nullpunkt kalibrieren:** Nach einer eigenen Rückfrage („Alle Zielwaagen müssen vollständig entlastet sein“) wird auf jeder
  Zielwaage nach den Einstellwerten der Nullpunkt kalibriert (`0x2300:0A` = 0 mit TAC-Freigabe, wie im Kalibrierdialog),
  danach erst im EEPROM gespeichert. Der Vergleich der Werte läuft vor dem Nullpunkt, weil dieser `0x2300:02 Absolute zero`
  mit dem eigenen Wert der Zielwaage überschreibt; das Ergebnis zeigt die Änderung (z. B. `Absolute zero 1219→1313`). Die Verstärkung wird nie verändert. Schlägt der Nullpunkt fehl, wird nicht gespeichert.
- **Nullpunkt nur bei ruhiger Waage:** Vor `0x2300:0A` wird das Brutto ca. 2 s lang sechsmal gelesen. Schwankt es um mehr als zwei Stellen
  der letzten Ziffer oder ist es nicht lesbar, wird der Befehl nicht gesendet (Übertragung: Hinweis im Ergebnis, Einstellwerte werden
  trotzdem gespeichert; Kalibrierdialog: Abbruch vor dem Nullpunkt). Die Messwerte stehen im Log (`NULLPUNKT …`).
- **Fehlerhafte Waagen werden übersprungen:** Antwortet ein Gerät nicht (`Ecat: Timeout`), bekommt es keine weiteren Anfragen; lehnt es den
  Nullpunkt ab oder ist es nicht ruhig, wird es nicht gespeichert. Die Übertragung läuft jeweils mit den übrigen Waagen weiter. Nur wenn in einem
  Lauf drei Waagen nicht auf die Nullpunkt-Kalibrierung antworten, wird der Rest nicht mehr bearbeitet (allgemeines Problem).
- **Fehlgeschlagene erneut übertragen:** Neuer Knopf neben „Auf ausgewählte Geräte übertragen“. Er wiederholt die Übertragung nur für die
  Waagen, die im letzten Lauf fehlgeschlagen oder übersprungen wurden (mit denselben Rückfragen).
- Kein zweiter Bus-Scan und keine neue Dashboard-Abfrage, solange die vorherige noch läuft (bei hängendem Master stauten sich sonst Anfragen über Minuten).
- Die Startzeile im Log enthält eine Build-Kennung (`Version 1.9.5 (Build xxxxxxxx)`), damit Testläufe eindeutig einer EXE zugeordnet werden können.
- Das Log zeigt während der Übertragung `Kalibrierstatus=Profilübertragung <Adresse>` statt des Status der letzten Kalibrierung.

### Wichtig
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
`44dc38d811224ec3bfdbb449f415e0dba4bc2ecc9240218fe303804d8e103249`
