#!/usr/bin/env bash
# Baut Stundennachweis.apk ohne Gradle - nur mit den Ubuntu-Paketen
#   android-sdk-platform-23 aapt apksigner zipalign dalvik-exchange  (+ ein JDK)
# Aufruf: ./build.sh   -> Ergebnis: Stundennachweis.apk
set -euo pipefail
cd "$(dirname "$0")"

SDK_JAR=/usr/lib/android-sdk/platforms/android-23/android.jar
KS=signatur.keystore
KS_PASS=${KS_PASS:-stundennachweis}

rm -rf build
mkdir -p build/gen build/obj

# Signaturschluessel einmalig anlegen (fuer Updates muss immer derselbe benutzt werden)
if [ ! -f "$KS" ]; then
  keytool -genkeypair -keystore "$KS" -storepass "$KS_PASS" -keypass "$KS_PASS" \
    -alias stundennachweis -keyalg RSA -keysize 2048 -validity 10000 \
    -dname "CN=Stundennachweis" >/dev/null
fi

aapt package -f -m -J build/gen -M AndroidManifest.xml -S res -I "$SDK_JAR" -F build/unsigned.apk
javac --release 8 -encoding UTF-8 -nowarn -classpath "$SDK_JAR" -d build/obj \
  $(find build/gen src -name '*.java')
dalvik-exchange --dex --output=build/classes.dex build/obj
(cd build && aapt add -f unsigned.apk classes.dex >/dev/null)
zipalign -f 4 build/unsigned.apk build/aligned.apk
apksigner sign --ks "$KS" --ks-pass "pass:$KS_PASS" --key-pass "pass:$KS_PASS" \
  --out Stundennachweis.apk build/aligned.apk
apksigner verify Stundennachweis.apk
echo "Fertig: $(pwd)/Stundennachweis.apk"
