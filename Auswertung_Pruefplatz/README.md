# Prüfplatz-Auswertung (Excel/VBA)

VBA-Modul `Pruefplatz_Auswertung.bas` für die grafische Auswertung und Analyse der
Prüfplatz-CSV-Exporte (z. B. `Pruefplatz_2_1.csv`).

## Installation

### Variante A: fertige Arbeitsmappe (empfohlen)

1. `Pruefplatz_Auswertung.xlsm` öffnen.
2. In der gelben Leiste **Inhalt aktivieren** klicken. Kommt stattdessen eine rote Leiste
   „Sicherheitsrisiko“: Datei schließen, im Explorer *Rechtsklick → Eigenschaften →
   „Zulassen“* anhaken, dann erneut öffnen.
3. Auf dem Blatt **Start** die Schaltfläche **„1. CSV importieren und auswerten“** klicken.

### Variante B: Modul selbst importieren

1. Excel öffnen → neue leere Arbeitsmappe.
2. `Alt + F11` (VBA-Editor) → **Datei → Datei importieren…** → `Pruefplatz_Auswertung.bas` wählen.
3. Arbeitsmappe als **Excel-Arbeitsmappe mit Makros (`*.xlsm`)** speichern.
4. `Alt + F8` → **`Startseite_Anlegen`** ausführen (legt das Blatt *Start* mit den Schaltflächen an)
   oder direkt **`Start_Auswertung`** ausführen.

Voraussetzung: Excel 2010 oder neuer unter Windows.

Die `.xlsm` wird mit `tools/build_xlsm.py` aus der `.bas`-Datei erzeugt
(VBA-Projekt nach MS-OVBA, nur Quelltext – Excel kompiliert beim ersten Öffnen).

## Makros

| Makro | Funktion |
|---|---|
| `Start_Auswertung` | CSV auswählen, einlesen, komplett auswerten |
| `Neu_Berechnen` | Auswertung nach Änderungen im Blatt *Einstellungen* wiederholen (ohne neuen Import) |
| `Einstellungen_Zuruecksetzen` | Grenzwerte/Einheiten auf Standard zurücksetzen |
| `Startseite_Anlegen` | Blatt *Start* mit Schaltflächen für die Makros anlegen |

## Erzeugte Blätter

| Blatt | Inhalt |
|---|---|
| **Uebersicht** | Kennzahlen (Anzahl, IO/NIO, Ausbeute, Zeitraum), automatisch erkannte Auffälligkeiten, Top-Fehlergründe, Kurzstatistik |
| **Diagramme** | 13 Diagramme (siehe unten) |
| **Statistik** | je Kanal: Min/P5/Q1/Median/Q3/P95/Max, Mittelwert, Std.-Abw., geometr. Mittel, log-Streuung, Grenzwertverletzungen, Ppk (log); Vergleich CSV-Ergebnis ↔ eigene Bewertung; Korrelationsmatrix; Häufigkeitsklassen |
| **Wochen** | je Kalenderwoche: Anzahl, IO, NIO, NIO-Quote, Mediane der Isolationswerte |
| **NIO_Liste** | alle auffälligen Prüfungen mit Fehlergrund (filterbar) |
| **Daten** | aufbereitete Messwerte in Anzeigeeinheit, Prüfzeit als Datum, KW, eigene Bewertung, Fehlergrund; Grenzwertverletzungen rot, Messbereichsende grau |
| **Einstellungen** | Kanaltyp, Umrechnungsfaktor, Einheit, Unter-/Obergrenze, Overrange-Schwelle, Zeitzone usw. |
| **Datenqualitaet** | Importprotokoll (korrigierte Formatfehler im CSV) |
| Rohdaten / Hilfsdaten | interne Tabellen (Hilfsdaten ausgeblendet) |

### Diagramme

1. Isolationswiderstände 5-7 / 5-8 / 6-7 / 6-8 über LfdN (log) mit Grenzlinie
2. Trend: gleitender Median (Fenster einstellbar)
3. Verteilung der Isolationswiderstände (log-Klassen)
4. Pareto der Fehlergründe
5. Anteil IO / NIO
6. Prüfungen und NIO-Quote je Kalenderwoche
7. Median der Isolationswiderstände je Kalenderwoche
8. 5-6 Widerstand – Verlauf mit gleitendem Median und Grenze
9. 5-6 Kapazität – Verlauf mit gleitendem Median
10. 5-6 Kapazität – Verteilung (zeigt die Zweigipfligkeit)
11. Korrelation 5-7 ↔ 5-8 (log-log)
12. Korrelation 6-7 ↔ 6-8 (log-log)
13. Durchgangsprüfung 6-12 / 5-11 (nur Werte > 0 Ω)

## Bewertungslogik

| Typ | Kanäle | NIO wenn … |
|---|---|---|
| `ISO` | 5-7, 5-8, 6-7, 6-8, 5-6 R | Wert < Untergrenze (Messbereichsende zählt als IO, umschaltbar) |
| `KONT` | 6-12, 5-11 | Wert > Obergrenze oder Messbereichsende (= offen) |
| `INFO` | 5-6 C | nur Statistik, keine Bewertung (außer Grenzen werden eingetragen) |

**Die Standard-Grenzwerte sind Annahmen** und müssen an die Prüfspezifikation angepasst werden:
Isolation ≥ 5000 MΩ, 5-6 R ≥ 100 MΩ, Durchgang ≤ 10 Ω.

## Robustheit beim Import

* Dezimalkomma (`2,64E+06`), Dezimalpunkt (`5645613.5`) und Tausenderpunkte (`1.404.108.875`)
  werden erkannt; Abweichungen werden im Blatt *Datenqualitaet* protokolliert.
* UTF-8 (Ω-Zeichen in den Spaltenköpfen) wird korrekt gelesen.
* Prüfzeit = Unix-Zeit in ms → Excel-Datum (Zeitzone einstellbar).
* Einheit der Kapazität wird automatisch erkannt: Rohwerte < 10⁻³ gelten als Farad und werden
  in nF angezeigt (nur solange im Blatt *Einstellungen* noch Faktor 1 / „Rohwert …“ steht).
* Zahlenformate richten sich nach der Größenordnung der Werte je Kanal (auch sehr kleine
  Werte werden nicht als 0 angezeigt).
* Bedingte Formatierungen verweisen auf Zellen in *Einstellungen* und funktionieren
  deshalb unabhängig von der Excel-Sprache (Dezimalkomma/-punkt).

## Ergebnisse für `Pruefplatz_2_1.csv` (mit den Standard-Grenzwerten)

* 1600 Prüfungen, 31.03.2026 – 22.09.2026 (KW 14 – 39)
* **Gesamtergebnis im CSV ist bei allen 1600 Prüfungen `false`**, auch bei unauffälligen
  Messwerten → das Feld wird vom Prüfplatz offenbar nicht korrekt gesetzt.
* Eigene Bewertung: 1352 IO / 248 NIO (84,5 % Ausbeute). Häufigste Gründe:
  5-6 R < 100 MΩ (135), 6-7 und 6-8 < 5000 MΩ (je 100), 5-7 (66), 5-8 (53),
  Durchgang 5-11 (25) und 6-12 (17) offen/hochohmig.
* Die Spaltenköpfe nennen „MΩ“, die Werte sind aber offensichtlich in Ω
  (Median 5-7 ≈ 4,9·10¹⁰ Ω = 49 GΩ). Werte ≈ 9,999·10¹² sind Messbereichsende.
* **5-6 Kapazität ist zweigipflig**: 919 Werte um 2,6·10⁷ und 597 Werte um 2,6·10⁶ –
  gleiche Mantisse, Faktor 10 → Verdacht auf Messbereichs-/Einheitenfehler.
* Die Prüfzeit hat im CSV nur 6 signifikante Stellen (nur 86 verschiedene Zeitstempel,
  Auflösung ≈ 2,8 h) → Zeitauswertung nur je Kalenderwoche sinnvoll.
* LfdN 1600 (und auffällig auch 1599): fast alle Kanäle im Messbereichsende → vermutlich
  Messung ohne kontaktierten Prüfling.
* Isolation: hohe Korrelation 5-7 ↔ 5-8 (r = 0,72) und 6-7 ↔ 6-8 (r = 0,78);
  Ppk(log) gegen 5000 MΩ nur 0,48 – 0,59.
