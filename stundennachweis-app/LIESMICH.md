# Stundennachweis – Android-App

Ergänzung zum PC-Programm im Ordner `stundennachweis`.
Mit der App erfasst du Kommen und Gehen per Tipp auf dem Handy, auch ohne Netz.
Die Zeiten werden später einmal an den PC übertragen, der sie in die Excel-Datei einträgt.
**Die Zeiten vom Handy haben immer Vorrang** vor denen am PC.

## Installieren

1. `Stundennachweis.apk` aufs Handy bringen, z. B. per E-Mail an dich selbst, Google Drive oder USB-Kabel.
2. Die Datei antippen. Android fragt, ob die App (z. B. „Dateien“ oder „Gmail“) Apps installieren darf. Das einmal erlauben.
3. Eventuell warnt Google Play Protect, weil die App nicht aus dem Play Store kommt. Dann auf „Trotzdem installieren“ tippen.

Voraussetzung: Android 6 oder neuer.

## Bedienung

| Was | Wie |
|---|---|
| Arbeitsbeginn | **KOMMEN** antippen |
| Arbeitsende | **GEHEN** antippen. Nochmal tippen überschreibt das Ende. |
| Wieder da nach „Gehen“ | **KOMMEN** antippen. Das Ende wird gelöscht, der Beginn bleibt. |
| Zeit korrigieren | Unten unter „Erfasste Tage“ auf den Tag tippen |
| Projekte | Auftragsnr., Bezeichnung und Anteil wie am PC: `2:30`, `2,5`, `40%`, leer = Rest |

Die Projektverteilung wird beim Tippen auf KOMMEN oder GEHEN und bei jeder Änderung für den heutigen Tag gemerkt.

## An den PC übertragen

Handy und PC müssen sich im Netz „sehen“. Dafür gibt es drei Wege:

| Weg | So geht's |
|---|---|
| **Hotspot am PC** | Am PC: Rechtsklick auf das Uhr-Symbol, „Vom Handy übernehmen …“, dann „Einschalten“. WLAN-Name und Passwort erscheinen. Mit dem Handy in dieses WLAN gehen. Vorteil: Der PC bleibt im Firmennetz. |
| Hotspot am Handy | Hotspot am Handy einschalten und den PC damit verbinden. Solange ist der PC meist nicht im Firmennetz, das Eintragen in Excel passiert dann später automatisch. |
| Gleiches WLAN | z. B. zu Hause. Ein Gäste-WLAN trennt die Geräte oft voneinander, dann geht es dort nicht. |

Dann:
1. In der App auf **„Übertragen …“** tippen. Dort stehen der Code und die Adresse des Handys. Die Anzeige offen lassen.
2. Am PC: Rechtsklick auf das Uhr-Symbol, **„Vom Handy übernehmen …“**. Beim ersten Mal den 6-stelligen Code eintragen, dann auf **„Übernehmen“** klicken.
3. In der App erscheint „Fertig“, und die Tage stehen auf „✓ am PC“.

Findet der PC das Handy nicht, trägst du die Adresse aus der App im PC-Fenster ein.

## Technik

- Die Übertragung läuft nur, solange in der App „Übertragen“ offen ist. Der PC holt die Daten ab. Dafür braucht der PC keine Administratorrechte und keine Firewall-Freigabe.
- Der Code verhindert, dass jemand anderes im selben WLAN die Zeiten abruft.
- Die App braucht nur die Berechtigung „Internet“, um im WLAN erreichbar zu sein. Es werden keine Daten ins Internet geschickt.

## Neu bauen

Ein Linux-Rechner mit den Ubuntu-Paketen `android-sdk-platform-23 aapt apksigner zipalign dalvik-exchange` und einem JDK reicht:

```bash
./build.sh
```

`signatur.keystore` ist der Signaturschlüssel. Updates lassen sich nur mit demselben Schlüssel über die installierte App installieren.
