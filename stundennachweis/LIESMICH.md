# Stundennachweis automatisch ausfüllen

Trägt jeden Arbeitstag **Arbeitsbeginn, Arbeitsende, Pause und Projektstunden** in
`2026_Stundennachweis - Ossowski.xlsx` ein. Dafür musst du nichts tun.

## So funktioniert es

| Wann                         | Was passiert                                                         |
|------------------------------|----------------------------------------------------------------------|
| Anmelden / PC entsperren     | Uhrzeit wird gemerkt. Die **erste** des Tages ist der Arbeitsbeginn. |
| Herunterfahren               | Uhrzeit wird als **Arbeitsende** eingetragen.                        |
| PC sperren (Windows-Taste+L) | Nichts.                                                              |

Danach trägt das Skript in die richtige Spalte im Blatt des Monats ein:

- **Zeile 4/5** – Beginn und Ende, minutengenau
- **Zeile 6** – Pause 0:45 (bei 6 Stunden oder weniger: 0:00)
- **Projektzeile `AB2500473`** – Arbeitszeit in Stunden, z. B. 7,53

Folgende Tage werden **nicht** angefasst:

- Wochenenden sowie Feier- und Brückentage aus dem Blatt „Feiertage“
- Tage, an denen schon Urlaub, Krankheit, Elternzeit usw. eingetragen ist
- Tage, die du **von Hand geändert** hast. Das Skript erkennt das und überschreibt sie nicht mehr.
- Projektstunden an Tagen, an denen schon andere Projekte gebucht sind. Beginn, Ende und Pause trägt es trotzdem ein.

Ist die Datei bei dir in Excel offen, schreibt das Skript direkt hinein und speichert.
Klappt das Eintragen beim Herunterfahren nicht, z. B. weil Windows zu schnell beendet oder jemand anderes
die Datei offen hat, holt das Skript es beim nächsten Anmelden nach. Die Uhrzeit des Herunterfahrens geht dabei nicht verloren.
Neustarts, die Windows selbst auslöst (z. B. durch Updates in der Nacht), zählen nicht als Arbeitsende.

## Einrichten (einmalig, ca. 2 Minuten)

1. Diesen Ordner `stundennachweis` an einen festen Ort auf deinem Arbeits-PC kopieren,
   z. B. nach `Dokumente\Stundennachweis`. Nicht auf das Netzlaufwerk, nicht in „Downloads“.
2. `einstellungen.json` mit dem Editor öffnen und prüfen:
   - `Datei` – Pfad zur Excel-Datei. `{Jahr}` wird automatisch ersetzt, 2027 läuft also von allein weiter.
   - `Projekt` – Auftragsnummer, auf die gebucht wird
   - `Pause` / `PauseBisSechsStunden` – Pausenzeiten
   - `ProjektstundenEintragen` – auf `false` setzen, wenn nur Beginn, Ende und Pause eingetragen werden sollen
3. Doppelklick auf **`Installieren.cmd`**.
   Am Ende muss „Test OK“ in Grün erscheinen.

Ab dann läuft alles automatisch. **Wichtig:** Zum Feierabend den PC herunterfahren.
Hattest du schon die erste Version installiert, `Installieren.cmd` einfach noch einmal ausführen.
Die alte Aufgabe für das Sperren wird dabei entfernt.

## Nützliches

**Was wurde erfasst?** Doppelklick auf `Anzeigen.cmd`.

**Tag von Hand eintragen oder korrigieren**, z. B. wenn du auswärts warst. Den Befehl in PowerShell im Ordner ausführen:

```powershell
.\Stundennachweis.ps1 -Aktion Eintragen -Datum 2026-11-09 -Beginn 7:30 -Ende 16:15
```

Optional geht auch `-Pause 0:30` oder `-Projekt AB2500257`.
Dieser Tag wird danach nicht mehr automatisch überschrieben.

**Protokoll:** `%LOCALAPPDATA%\Stundennachweis\protokoll.log`

**Ausschalten:** Doppelklick auf `Deinstallieren.cmd`.

## Grenzen

- Wird der PC nicht heruntergefahren (nur gesperrt, Energiesparmodus, Akku leer), bleibt das Arbeitsende leer.
  Dann den Tag mit `-Aktion Eintragen -Ende 16:30` nachtragen.
- Mehrmals Herunterfahren am selben Tag: Es zählt das letzte.
- Dienstreisen, Homeoffice ohne Firmen-PC usw. bitte von Hand eintragen.
- Voraussetzungen: Windows mit installiertem Excel und Zugriff auf Laufwerk `S:`.
  Falls die IT PowerShell-Skripte oder geplante Aufgaben sperrt, bricht `Installieren.cmd` mit einer Fehlermeldung ab.
