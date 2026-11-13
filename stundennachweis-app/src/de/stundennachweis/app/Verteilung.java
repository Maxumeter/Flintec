package de.stundennachweis.app;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Gleiche Regeln wie im PC-Programm (nur fuer die Anzeige auf dem Handy):
 *  - Anteil "2:30" / "2,5" = feste Stunden, "40%" = Prozent, leer = Rest
 *  - Rest gleichmaessig auf Projekte ohne Anteil, sonst auf das letzte Projekt
 *  - Feste Anteile ueber der Arbeitszeit werden anteilig gekuerzt
 * Die verbindliche Berechnung macht das PC-Programm beim Eintragen in Excel.
 */
public final class Verteilung {

    private Verteilung() {
    }

    /** Pause wie am PC: ueber 6 Stunden Anwesenheit 0:45, sonst 0:15 (in Minuten). */
    public static int pause(int anwesenheitMinuten) {
        return anwesenheitMinuten > 6 * 60 ? 45 : 15;
    }

    /** Arbeitszeit in Stunden (Pause abgezogen), 0 wenn nicht berechenbar. */
    public static double arbeitsstunden(String beginn, String ende) {
        int b = Speicher.minuten(beginn);
        int e = Speicher.minuten(ende);
        if (b < 0 || e < 0 || e <= b) return 0;
        int anw = e - b;
        int arbeit = anw - pause(anw);
        return arbeit > 0 ? arbeit / 60.0 : 0;
    }

    /** Liefert null, wenn gueltig, sonst eine Fehlermeldung. */
    public static String pruefeAnteil(String text) {
        return anteil(text) == null ? "Ungültiger Anteil \"" + text + "\" (z. B. 2:30, 2,5 oder 40%)" : null;
    }

    /** {art, wert}: art 0 = Rest, 1 = Stunden, 2 = Prozent; null wenn ungueltig */
    static double[] anteil(String text) {
        String t = text == null ? "" : text.trim();
        if (t.isEmpty()) return new double[]{0, 0};
        try {
            if (t.endsWith("%")) {
                return new double[]{2, Double.parseDouble(t.substring(0, t.length() - 1).trim().replace(',', '.'))};
            }
            if (t.contains(":")) {
                int m = Speicher.minuten(t);
                return m < 0 ? null : new double[]{1, m / 60.0};
            }
            String z = t.toLowerCase(Locale.GERMANY).replace("h", "").trim().replace(',', '.');
            return new double[]{1, Double.parseDouble(z)};
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public static List<Double> verteile(double gesamt, List<Speicher.Projekt> projekte) {
        int n = projekte.size();
        List<Double> stunden = new ArrayList<>();
        if (n == 0) return stunden;
        boolean[] rest = new boolean[n];
        double[] s = new double[n];
        double fest = 0;
        int anzahlRest = 0;
        for (int i = 0; i < n; i++) {
            double[] a = anteil(projekte.get(i).anteil);
            if (a == null || a[0] == 0) {
                rest[i] = true;
                anzahlRest++;
            } else {
                s[i] = a[0] == 1 ? a[1] : gesamt * a[1] / 100.0;
                fest += s[i];
            }
        }
        if (fest > gesamt && fest > 0) {
            double f = gesamt / fest;
            for (int i = 0; i < n; i++) if (!rest[i]) s[i] *= f;
            fest = gesamt;
        }
        if (anzahlRest > 0) {
            for (int i = 0; i < n; i++) if (rest[i]) s[i] = (gesamt - fest) / anzahlRest;
        } else {
            s[n - 1] += gesamt - fest;
        }
        for (int i = 0; i < n; i++) stunden.add(Math.round(s[i] * 100) / 100.0);
        return stunden;
    }

    public static String zahl(double d) {
        return String.format(Locale.GERMANY, "%.2f", d);
    }

    public static String hm(double stunden) {
        int min = (int) Math.round(stunden * 60);
        return Speicher.hhmm(min);
    }
}
