# Stundennachweis automatisch ausfüllen

Trägt jeden Arbeitstag **Arbeitsbeginn, Arbeitsende, Pause und Projektstunden** in
`2026_Stundennachweis - Ossowski.xlsx` ein. Du tippst nur morgens deinen Arbeitsbeginn ein.

## So funktioniert es

| Wann                         | Was passiert                                                         |
|---|---|
| Erstes Anmelden / Entsperren am Tag | Kleines Fenster **oben rechts**: Arbeitsbeginn eintippen (z. B. `7:45` oder `745`), Enter. |
| Herunterfahren               | Uhrzeit wird als **Arbeitsende** eingetragen. |
| PC sperren (Windows-Taste+L) | Nichts. |

Danach trägt das Skript in die richtige Spalte im Blatt des Monats ein:

- **Zeile 4/5** – Beginn und Ende, minutengenau
- **Zeile 6** – Pause 0:15. Erst wenn die Anwesenheit (Ende minus Beginn) **über 6 Stunden** liegt: 0:45
- **Projektzeilen** – Arbeitszeit in Stunden, z. B. 7,53. Sie kann auf mehrere Projekte verteilt werden, siehe unten.

Im Fenster steht schon die aktuelle Uhrzeit, du kannst sie überschreiben.
Oben rechts im Fenster stehen die **Gesamtstunden des Tages** bis jetzt, z. B. `Heute: 6:12 h (6,20)`.
Gerechnet wird vom Arbeitsbeginn bis zur aktuellen Uhrzeit, die Pause ist schon abgezogen.
Die Anzeige aktualisiert sich alle 15 Sekunden und sofort, wenn du den Beginn änderst.
Klickst du auf „Später“, kommt das Fenster beim nächsten Entsperren wieder.

Unten im Fenster gibt es den Haken **„Automatisch schließen nach 30 s“**. Das Fenster merkt sich die Einstellung.
Ist er gesetzt, läuft auf dem Knopf ein Countdown mit. Bei 0 wird die angezeigte Uhrzeit übernommen und das Fenster schließt sich.
Sobald du ins Uhrzeitfeld tippst oder klickst, stoppt der Countdown.
Die Sekunden stellst du in `einstellungen.json` unter `FensterSekunden` ein.

Darunter stehen **Personalnr.** und **Name**. Die Personalnummer kommt in Zelle E1, der Name in J1, und zwar in jedem Monatsblatt.
Zellen, die schon eine Formel haben (z. B. `=Oktober!J1`), bleiben unverändert.

### Projekte und Verteilung

Ganz unten im Fenster steht eine kleine Tabelle mit den Spalten **Auftragsnr.**, **Bezeichnung**, **Anteil** und **Stunden**.
Die Spalte **Stunden** (grau) zeigt, wie viele der bisherigen Tagesstunden auf jedes Projekt entfallen.
Sie aktualisiert sich zusammen mit der Gesamtanzeige und sofort, wenn du einen Anteil änderst.
Für ein weiteres Projekt tippst du einfach in die leere letzte Zeile.
Zum Löschen klickst du links auf das Kästchen vor der Zeile und drückst `Entf`.

So schreibst du den Anteil:

| Anteil | Bedeutung |
|---|---|
| `2:30` oder `2,5` | feste Stunden |
| `40%` | Anteil an der Arbeitszeit des Tages |
| *leer* | der Rest. Mehrere leere Zeilen teilen sich den Rest gleichmäßig. |

Beispiel bei 7,53 Stunden Arbeitszeit:

| Auftragsnr. | Anteil | gebucht |
|---|---|---|
| AB2500473 | *leer* | 5,03 |
| AB2500257 | `2:30` | 2,50 |

- **Ohne leere Zeile** bekommt das letzte Projekt den Rest, damit immer die ganze Arbeitszeit gebucht ist.
- **Sind die festen Stunden größer als die Arbeitszeit**, werden sie anteilig gekürzt.
- **Ein ungültiger Anteil** wird rot markiert. Das Fenster lässt sich dann erst nach einer Korrektur schließen.

In Spalte A der Projektzeile steht die Auftragsnummer, in Spalte B die Bezeichnung.
Gibt es zu einer Nummer im Monatsblatt noch keine Zeile, legt das Skript sie in der ersten freien Projektzeile an.
Änderst du die Verteilung im Laufe des Tages (mit `Beginn.cmd`), wird der heutige Tag neu verteilt.
Zurückliegende Tage bleiben, wie sie sind.

Alle Angaben im Fenster werden sofort gespeichert und stehen beim nächsten Mal wieder da, auch nach einem Neustart.
Steht in der Tabelle kein Projekt, gilt die Nummer aus `einstellungen.json`.

Das Fenster lässt sich mit dem Strich oben rechts **minimieren** und über die Taskleiste wieder öffnen.

Ist der Beginn für heute erfasst, erscheint das Fenster nicht mehr. Am Wochenende erscheint es gar nicht.
Zum Korrigieren öffnest du es mit **`Beginn.cmd`** jederzeit wieder.

Folgende Tage werden **nicht** angefasst:

- Wochenenden sowie Feier- und Brückentage aus dem Blatt „Feiertage“
- Tage, an denen schon Urlaub, Krankheit, Elternzeit usw. eingetragen ist
- Tage, die du **von Hand geändert** hast. Das Skript erkennt das und überschreibt sie nicht mehr.
- Projektstunden an Tagen, an denen du von Hand auf Projekte gebucht hast, die nicht in der Tabelle stehen. Beginn, Ende und Pause trägt das Skript trotzdem ein.

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
   - `Pause` / `PauseBisSechsStunden` – Pausenzeiten (0:45 erst über 6 Stunden, sonst 0:15)
   - `ProjektstundenEintragen` – auf `false` setzen, wenn nur Beginn, Ende und Pause eingetragen werden sollen
3. Doppelklick auf **`Installieren.cmd`**.
   Am Ende muss „Test OK“ in Grün erscheinen.

Ab dann läuft alles automatisch. **Wichtig:** Zum Feierabend den PC herunterfahren.
Hattest du schon die erste Version installiert, `Installieren.cmd` einfach noch einmal ausführen.
Die Aufgaben werden dabei aktualisiert, die alte Aufgabe für das Sperren wird entfernt.

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
