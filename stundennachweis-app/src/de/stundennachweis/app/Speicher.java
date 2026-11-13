package de.stundennachweis.app;

import android.content.Context;
import android.content.SharedPreferences;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.security.SecureRandom;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

/**
 * Lokaler Speicher der App (bleibt auch ohne Netz erhalten).
 *
 * Je Tag: Beginn, Ende (HH:mm), Projektverteilung und ob der Tag schon an den PC uebertragen wurde.
 */
public final class Speicher {

    public static final class Projekt {
        public String nummer = "";
        public String bezeichnung = "";
        public String anteil = "";

        JSONObject toJson() throws JSONException {
            JSONObject o = new JSONObject();
            o.put("nummer", nummer);
            o.put("bezeichnung", bezeichnung);
            o.put("anteil", anteil);
            return o;
        }

        static Projekt fromJson(JSONObject o) {
            Projekt p = new Projekt();
            p.nummer = o.optString("nummer", "").trim();
            p.bezeichnung = o.optString("bezeichnung", "").trim();
            p.anteil = o.optString("anteil", "").trim();
            return p;
        }
    }

    public static final class Tag {
        public String datum;            // yyyy-MM-dd
        public String beginn = "";      // HH:mm oder leer
        public String ende = "";        // HH:mm oder leer
        public List<Projekt> projekte = new ArrayList<>();
        public boolean uebertragen = false;
    }

    private static final String PREFS = "stundennachweis";
    private final SharedPreferences prefs;

    public Speicher(Context context) {
        prefs = context.getSharedPreferences(PREFS, Context.MODE_PRIVATE);
    }

    // ------------------------------------------------------------------ Datum/Zeit

    public static String heute() {
        return datum(new Date());
    }

    public static String datum(Date d) {
        return new SimpleDateFormat("yyyy-MM-dd", Locale.GERMANY).format(d);
    }

    public static String jetzt() {
        return new SimpleDateFormat("HH:mm", Locale.GERMANY).format(new Date());
    }

    /** "Mi 12.11." */
    public static String anzeigeDatum(String datum) {
        try {
            Date d = new SimpleDateFormat("yyyy-MM-dd", Locale.GERMANY).parse(datum);
            return new SimpleDateFormat("EE dd.MM.", Locale.GERMANY).format(d);
        } catch (Exception e) {
            return datum;
        }
    }

    /** Minuten seit Mitternacht oder -1 */
    public static int minuten(String hhmm) {
        if (hhmm == null) return -1;
        String t = hhmm.trim();
        int p = t.indexOf(':');
        if (p < 1) return -1;
        try {
            int h = Integer.parseInt(t.substring(0, p));
            int m = Integer.parseInt(t.substring(p + 1));
            if (h < 0 || h > 23 || m < 0 || m > 59) return -1;
            return h * 60 + m;
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    public static String hhmm(int minuten) {
        return String.format(Locale.GERMANY, "%d:%02d", minuten / 60, minuten % 60);
    }

    // ------------------------------------------------------------------ Tage

    private JSONObject alleTageJson() {
        try {
            return new JSONObject(prefs.getString("tage", "{}"));
        } catch (JSONException e) {
            return new JSONObject();
        }
    }

    public Tag tag(String datum) {
        JSONObject o = alleTageJson().optJSONObject(datum);
        Tag t = new Tag();
        t.datum = datum;
        if (o == null) return t;
        t.beginn = o.optString("beginn", "");
        t.ende = o.optString("ende", "");
        t.uebertragen = o.optBoolean("uebertragen", false);
        JSONArray a = o.optJSONArray("projekte");
        if (a != null) {
            for (int i = 0; i < a.length(); i++) {
                JSONObject p = a.optJSONObject(i);
                if (p != null) t.projekte.add(Projekt.fromJson(p));
            }
        }
        return t;
    }

    public boolean existiert(String datum) {
        return alleTageJson().has(datum);
    }

    public void speichere(Tag t) {
        try {
            JSONObject alle = alleTageJson();
            JSONObject o = new JSONObject();
            o.put("beginn", t.beginn);
            o.put("ende", t.ende);
            o.put("uebertragen", t.uebertragen);
            JSONArray a = new JSONArray();
            for (Projekt p : t.projekte) a.put(p.toJson());
            o.put("projekte", a);
            alle.put(t.datum, o);
            aufraeumen(alle);
            prefs.edit().putString("tage", alle.toString()).apply();
        } catch (JSONException ignored) {
        }
    }

    public void loesche(String datum) {
        JSONObject alle = alleTageJson();
        alle.remove(datum);
        prefs.edit().putString("tage", alle.toString()).apply();
    }

    /** Alle Tage, neueste zuerst */
    public List<Tag> alleTage() {
        List<String> daten = new ArrayList<>();
        Iterator<String> it = alleTageJson().keys();
        while (it.hasNext()) daten.add(it.next());
        Collections.sort(daten, Collections.<String>reverseOrder());
        List<Tag> liste = new ArrayList<>();
        for (String d : daten) liste.add(tag(d));
        return liste;
    }

    public List<Tag> offeneTage() {
        List<Tag> liste = new ArrayList<>();
        for (Tag t : alleTage()) if (!t.uebertragen) liste.add(t);
        Collections.reverse(liste);
        return liste;
    }

    public void alsUebertragenMarkieren(List<String> daten) {
        for (String d : daten) {
            if (!existiert(d)) continue;
            Tag t = tag(d);
            t.uebertragen = true;
            speichere(t);
        }
    }

    /** Uebertragene Tage, die aelter als 90 Tage sind, entfernen. */
    private void aufraeumen(JSONObject alle) {
        Calendar c = Calendar.getInstance();
        c.add(Calendar.DAY_OF_YEAR, -90);
        String grenze = datum(c.getTime());
        List<String> weg = new ArrayList<>();
        Iterator<String> it = alle.keys();
        while (it.hasNext()) {
            String k = it.next();
            JSONObject o = alle.optJSONObject(k);
            if (k.compareTo(grenze) < 0 && o != null && o.optBoolean("uebertragen", false)) weg.add(k);
        }
        for (String k : weg) alle.remove(k);
    }

    // ------------------------------------------------------------------ Projekte (aktuelle Auswahl)

    public List<Projekt> projekte() {
        List<Projekt> liste = new ArrayList<>();
        try {
            JSONArray a = new JSONArray(prefs.getString("projekte", "[]"));
            for (int i = 0; i < a.length(); i++) {
                JSONObject o = a.optJSONObject(i);
                if (o != null) liste.add(Projekt.fromJson(o));
            }
        } catch (JSONException ignored) {
        }
        return liste;
    }

    public void speichereProjekte(List<Projekt> liste) {
        JSONArray a = new JSONArray();
        try {
            for (Projekt p : liste) if (!p.nummer.trim().isEmpty()) a.put(p.toJson());
        } catch (JSONException ignored) {
        }
        prefs.edit().putString("projekte", a.toString()).apply();
    }

    // ------------------------------------------------------------------ Kopplungscode

    /** 6-stelliger Code, den der PC beim ersten Abholen einmal abfragt. */
    public String code() {
        String c = prefs.getString("code", "");
        if (c.length() != 6) {
            SecureRandom r = new SecureRandom();
            c = String.format(Locale.GERMANY, "%06d", r.nextInt(1000000));
            prefs.edit().putString("code", c).apply();
        }
        return c;
    }

    public void neuerCode() {
        prefs.edit().remove("code").apply();
    }

    // ------------------------------------------------------------------ Export

    public static JSONObject tagJson(Tag t) throws JSONException {
        JSONObject o = new JSONObject();
        o.put("datum", t.datum);
        o.put("beginn", t.beginn);
        o.put("ende", t.ende);
        JSONArray a = new JSONArray();
        for (Projekt p : t.projekte) a.put(p.toJson());
        o.put("projekte", a);
        return o;
    }
}
