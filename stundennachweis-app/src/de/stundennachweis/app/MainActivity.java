package de.stundennachweis.app;

import android.app.Activity;
import android.app.AlertDialog;
import android.app.TimePickerDialog;
import android.content.DialogInterface;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.GradientDrawable;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.text.Editable;
import android.text.InputType;
import android.text.TextWatcher;
import android.util.TypedValue;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.view.WindowManager;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.TimePicker;
import android.widget.Toast;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class MainActivity extends Activity {

    private static final int GRUEN = Color.rgb(46, 125, 50);
    private static final int ORANGE = Color.rgb(230, 81, 0);
    private static final int BLAU = Color.rgb(33, 111, 191);
    private static final int GRAU = Color.rgb(110, 110, 110);

    private Speicher speicher;
    private final Handler handler = new Handler(Looper.getMainLooper());

    private TextView tvDatum, tvStatus, tvStunden;
    private LinearLayout projektListe, tageListe;
    private final List<EditText[]> projektFelder = new ArrayList<>();
    private final List<TextView> projektStunden = new ArrayList<>();
    private boolean ladeProjekte = false;

    private final Runnable takt = new Runnable() {
        @Override
        public void run() {
            aktualisiereKopf();
            handler.postDelayed(this, 30000);
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        speicher = new Speicher(this);
        setTitle("Stundennachweis");
        baueOberflaeche();
    }

    @Override
    protected void onResume() {
        super.onResume();
        aktualisiereAlles();
        handler.post(takt);
    }

    @Override
    protected void onPause() {
        super.onPause();
        handler.removeCallbacks(takt);
    }

    // ------------------------------------------------------------------ Oberflaeche

    private int dp(int v) {
        return (int) TypedValue.applyDimension(TypedValue.COMPLEX_UNIT_DIP, v, getResources().getDisplayMetrics());
    }

    private TextView text(String s, int sp, boolean fett, int farbe) {
        TextView t = new TextView(this);
        t.setText(s);
        t.setTextSize(TypedValue.COMPLEX_UNIT_SP, sp);
        t.setTextColor(farbe);
        if (fett) t.setTypeface(Typeface.DEFAULT, Typeface.BOLD);
        return t;
    }

    private Button grosserKnopf(String s, int farbe) {
        Button b = new Button(this);
        b.setText(s);
        b.setTextSize(TypedValue.COMPLEX_UNIT_SP, 24);
        b.setTextColor(Color.WHITE);
        b.setTypeface(Typeface.DEFAULT, Typeface.BOLD);
        GradientDrawable g = new GradientDrawable();
        g.setColor(farbe);
        g.setCornerRadius(dp(14));
        b.setBackground(g);
        return b;
    }

    private TextView ueberschrift(String s) {
        TextView t = text(s, 18, true, Color.rgb(40, 40, 40));
        t.setPadding(0, dp(22), 0, dp(6));
        return t;
    }

    private void baueOberflaeche() {
        ScrollView scroll = new ScrollView(this);
        LinearLayout wurzel = new LinearLayout(this);
        wurzel.setOrientation(LinearLayout.VERTICAL);
        wurzel.setPadding(dp(16), dp(12), dp(16), dp(24));
        scroll.addView(wurzel);

        tvDatum = text("", 20, true, Color.rgb(30, 30, 30));
        tvStatus = text("", 16, false, GRAU);
        tvStatus.setPadding(0, dp(4), 0, 0);
        tvStunden = text("", 22, true, BLAU);
        tvStunden.setPadding(0, dp(6), 0, dp(12));
        wurzel.addView(tvDatum);
        wurzel.addView(tvStatus);
        wurzel.addView(tvStunden);

        LinearLayout knoepfe = new LinearLayout(this);
        knoepfe.setOrientation(LinearLayout.HORIZONTAL);
        Button kommen = grosserKnopf("KOMMEN", GRUEN);
        Button gehen = grosserKnopf("GEHEN", ORANGE);
        LinearLayout.LayoutParams lp = new LinearLayout.LayoutParams(0, dp(110), 1f);
        lp.setMargins(0, 0, dp(8), 0);
        knoepfe.addView(kommen, lp);
        LinearLayout.LayoutParams lp2 = new LinearLayout.LayoutParams(0, dp(110), 1f);
        lp2.setMargins(dp(8), 0, 0, 0);
        knoepfe.addView(gehen, lp2);
        wurzel.addView(knoepfe);

        kommen.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                kommen();
            }
        });
        gehen.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                gehen();
            }
        });

        // Projekte
        wurzel.addView(ueberschrift("Projekte"));
        TextView hilfe = text("Anteil: 2:30 / 2,5 / 40%  –  leer = Rest", 13, false, GRAU);
        wurzel.addView(hilfe);
        LinearLayout kopf = new LinearLayout(this);
        kopf.setOrientation(LinearLayout.HORIZONTAL);
        kopf.setPadding(0, dp(6), 0, 0);
        kopf.addView(text("Auftragsnr.", 13, true, GRAU), new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 3.2f));
        kopf.addView(text("Bezeichnung", 13, true, GRAU), new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 3.4f));
        kopf.addView(text("Anteil", 13, true, GRAU), new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.6f));
        kopf.addView(text("Std.", 13, true, GRAU), new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.4f));
        kopf.addView(new View(this), new LinearLayout.LayoutParams(dp(40), 1));
        wurzel.addView(kopf);
        projektListe = new LinearLayout(this);
        projektListe.setOrientation(LinearLayout.VERTICAL);
        wurzel.addView(projektListe);
        Button neu = new Button(this);
        neu.setText("+ Projekt");
        neu.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                fuegeProjektZeileHinzu(new Speicher.Projekt());
            }
        });
        wurzel.addView(neu, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, ViewGroup.LayoutParams.WRAP_CONTENT));

        // Uebertragen
        wurzel.addView(ueberschrift("An PC übertragen"));
        Button senden = grosserKnopf("Übertragen …", BLAU);
        senden.setTextSize(TypedValue.COMPLEX_UNIT_SP, 18);
        wurzel.addView(senden, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, dp(56)));
        senden.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                uebertragen();
            }
        });

        // Tage
        wurzel.addView(ueberschrift("Erfasste Tage"));
        wurzel.addView(text("Antippen zum Ändern", 13, false, GRAU));
        tageListe = new LinearLayout(this);
        tageListe.setOrientation(LinearLayout.VERTICAL);
        wurzel.addView(tageListe);

        setContentView(scroll);
    }

    // ------------------------------------------------------------------ Kommen / Gehen

    private void kommen() {
        String heute = Speicher.heute();
        Speicher.Tag t = speicher.tag(heute);
        String jetzt = Speicher.jetzt();
        String meldung;
        if (t.beginn.isEmpty()) {
            t.beginn = jetzt;
            meldung = "Kommen " + jetzt + " gespeichert";
        } else if (!t.ende.isEmpty()) {
            // Nach "Gehen" wieder da: Tag geht weiter, Beginn bleibt
            t.ende = "";
            meldung = "Weiter gearbeitet – Beginn bleibt " + t.beginn;
        } else {
            meldung = "Beginn ist schon " + t.beginn + ". Zum Ändern unten auf den Tag tippen.";
        }
        t.projekte = speicher.projekte();
        t.uebertragen = false;
        speicher.speichere(t);
        Toast.makeText(this, meldung, Toast.LENGTH_LONG).show();
        aktualisiereAlles();
    }

    private void gehen() {
        String heute = Speicher.heute();
        Speicher.Tag t = speicher.tag(heute);
        String jetzt = Speicher.jetzt();
        t.ende = jetzt;
        t.projekte = speicher.projekte();
        t.uebertragen = false;
        speicher.speichere(t);
        String extra = t.beginn.isEmpty() ? " (Beginn fehlt noch – unten auf den Tag tippen)" : "";
        Toast.makeText(this, "Gehen " + jetzt + " gespeichert" + extra, Toast.LENGTH_LONG).show();
        aktualisiereAlles();
    }

    // ------------------------------------------------------------------ Anzeige

    private void aktualisiereAlles() {
        ladeProjektZeilen();
        aktualisiereKopf();
        aktualisiereTage();
    }

    private double heuteStunden() {
        Speicher.Tag t = speicher.tag(Speicher.heute());
        if (t.beginn.isEmpty()) return 0;
        String ende = t.ende.isEmpty() ? Speicher.jetzt() : t.ende;
        return Verteilung.arbeitsstunden(t.beginn, ende);
    }

    private void aktualisiereKopf() {
        tvDatum.setText(new SimpleDateFormat("EEEE, dd.MM.yyyy", Locale.GERMANY).format(new Date()));
        Speicher.Tag t = speicher.tag(Speicher.heute());
        String b = t.beginn.isEmpty() ? "–" : t.beginn;
        String e = t.ende.isEmpty() ? "–" : t.ende;
        tvStatus.setText("Kommen " + b + "   ·   Gehen " + e);
        double h = heuteStunden();
        tvStunden.setText("Heute: " + Verteilung.hm(h) + " h (" + Verteilung.zahl(h) + ")");
        aktualisiereProjektStunden(h);
    }

    private void aktualisiereProjektStunden(double gesamt) {
        List<Speicher.Projekt> liste = new ArrayList<>();
        List<Integer> index = new ArrayList<>();
        for (int i = 0; i < projektFelder.size(); i++) {
            Speicher.Projekt p = ausZeile(projektFelder.get(i));
            if (!p.nummer.isEmpty()) {
                liste.add(p);
                index.add(i);
            }
        }
        List<Double> s = Verteilung.verteile(gesamt, liste);
        for (TextView tv : projektStunden) tv.setText("");
        for (int k = 0; k < index.size(); k++) projektStunden.get(index.get(k)).setText(Verteilung.zahl(s.get(k)));
        for (int i = 0; i < projektFelder.size(); i++) {
            EditText anteil = projektFelder.get(i)[2];
            boolean ok = Verteilung.pruefeAnteil(anteil.getText().toString()) == null;
            anteil.setBackgroundColor(ok ? Color.TRANSPARENT : Color.rgb(255, 228, 225));
        }
    }

    private void aktualisiereTage() {
        tageListe.removeAllViews();
        List<Speicher.Tag> tage = speicher.alleTage();
        if (tage.isEmpty()) {
            tageListe.addView(text("Noch nichts erfasst.", 15, false, GRAU));
            return;
        }
        int n = 0;
        for (final Speicher.Tag t : tage) {
            if (++n > 30) break;
            LinearLayout z = new LinearLayout(this);
            z.setOrientation(LinearLayout.HORIZONTAL);
            z.setPadding(0, dp(10), 0, dp(10));
            z.setGravity(Gravity.CENTER_VERTICAL);
            String zeiten = (t.beginn.isEmpty() ? "?" : t.beginn) + " – " + (t.ende.isEmpty() ? "?" : t.ende);
            double h = Verteilung.arbeitsstunden(t.beginn, t.ende);
            z.addView(text(Speicher.anzeigeDatum(t.datum), 16, true, Color.rgb(30, 30, 30)),
                    new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 2.2f));
            z.addView(text(zeiten, 16, false, Color.rgb(30, 30, 30)),
                    new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 3f));
            z.addView(text(h > 0 ? Verteilung.zahl(h) + " h" : "", 15, false, GRAU),
                    new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.6f));
            TextView st = text(t.uebertragen ? "✓ am PC" : "● offen", 14, true, t.uebertragen ? GRUEN : ORANGE);
            z.addView(st, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 2f));
            z.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    bearbeiteTag(t.datum);
                }
            });
            tageListe.addView(z);
            View linie = new View(this);
            linie.setBackgroundColor(Color.rgb(225, 225, 225));
            tageListe.addView(linie, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 1));
        }
    }

    // ------------------------------------------------------------------ Projekte

    private Speicher.Projekt ausZeile(EditText[] f) {
        Speicher.Projekt p = new Speicher.Projekt();
        p.nummer = f[0].getText().toString().trim();
        p.bezeichnung = f[1].getText().toString().trim();
        p.anteil = f[2].getText().toString().trim();
        return p;
    }

    private void ladeProjektZeilen() {
        ladeProjekte = true;
        projektListe.removeAllViews();
        projektFelder.clear();
        projektStunden.clear();
        List<Speicher.Projekt> liste = speicher.projekte();
        if (liste.isEmpty()) liste.add(new Speicher.Projekt());
        for (Speicher.Projekt p : liste) fuegeProjektZeileHinzu(p);
        ladeProjekte = false;
    }

    private EditText feld(String wert, String hinweis) {
        EditText e = new EditText(this);
        e.setText(wert);
        e.setHint(hinweis);
        e.setSingleLine(true);
        e.setTextSize(TypedValue.COMPLEX_UNIT_SP, 15);
        e.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence s, int a, int b, int c) {
            }

            @Override
            public void onTextChanged(CharSequence s, int a, int b, int c) {
            }

            @Override
            public void afterTextChanged(Editable s) {
                projekteGeaendert();
            }
        });
        return e;
    }

    private void fuegeProjektZeileHinzu(Speicher.Projekt p) {
        final LinearLayout z = new LinearLayout(this);
        z.setOrientation(LinearLayout.HORIZONTAL);
        z.setGravity(Gravity.CENTER_VERTICAL);
        final EditText nr = feld(p.nummer, "AB…");
        nr.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_FLAG_CAP_CHARACTERS);
        final EditText bez = feld(p.bezeichnung, "");
        final EditText ant = feld(p.anteil, "Rest");
        TextView std = text("", 14, false, GRAU);
        std.setGravity(Gravity.END);
        Button weg = new Button(this);
        weg.setText("✕");
        weg.setTextSize(TypedValue.COMPLEX_UNIT_SP, 14);
        z.addView(nr, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 3.2f));
        z.addView(bez, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 3.4f));
        z.addView(ant, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.6f));
        z.addView(std, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1.4f));
        z.addView(weg, new LinearLayout.LayoutParams(dp(40), dp(40)));
        final EditText[] felder = new EditText[]{nr, bez, ant};
        projektFelder.add(felder);
        projektStunden.add(std);
        projektListe.addView(z);
        weg.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                int i = projektFelder.indexOf(felder);
                if (i >= 0) {
                    projektFelder.remove(i);
                    projektStunden.remove(i);
                }
                projektListe.removeView(z);
                projekteGeaendert();
            }
        });
    }

    private void projekteGeaendert() {
        if (ladeProjekte) return;
        List<Speicher.Projekt> liste = new ArrayList<>();
        for (EditText[] f : projektFelder) {
            Speicher.Projekt p = ausZeile(f);
            if (!p.nummer.isEmpty()) liste.add(p);
        }
        speicher.speichereProjekte(liste);
        // Heutigen Tag mit der neuen Verteilung merken
        String heute = Speicher.heute();
        if (speicher.existiert(heute)) {
            Speicher.Tag t = speicher.tag(heute);
            t.projekte = liste;
            t.uebertragen = false;
            speicher.speichere(t);
            aktualisiereTage();
        }
        aktualisiereProjektStunden(heuteStunden());
    }

    // ------------------------------------------------------------------ Tag bearbeiten

    private void bearbeiteTag(final String datum) {
        final Speicher.Tag t = speicher.tag(datum);
        LinearLayout box = new LinearLayout(this);
        box.setOrientation(LinearLayout.VERTICAL);
        box.setPadding(dp(20), dp(8), dp(20), 0);

        final Button bBeginn = new Button(this);
        final Button bEnde = new Button(this);
        final Runnable beschriften = new Runnable() {
            @Override
            public void run() {
                bBeginn.setText("Kommen: " + (t.beginn.isEmpty() ? "–" : t.beginn));
                bEnde.setText("Gehen: " + (t.ende.isEmpty() ? "–" : t.ende));
            }
        };
        beschriften.run();
        box.addView(bBeginn);
        box.addView(bEnde);
        StringBuilder pj = new StringBuilder();
        for (Speicher.Projekt p : t.projekte) {
            pj.append(p.nummer);
            if (!p.anteil.isEmpty()) pj.append(" (").append(p.anteil).append(")");
            pj.append("  ");
        }
        if (pj.length() > 0) {
            TextView tp = text("Projekte: " + pj.toString().trim(), 13, false, GRAU);
            tp.setPadding(0, dp(8), 0, 0);
            box.addView(tp);
        }

        bBeginn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                waehleZeit(t.beginn, new Zeitwahl() {
                    @Override
                    public void gewaehlt(String hhmm) {
                        t.beginn = hhmm;
                        beschriften.run();
                    }
                });
            }
        });
        bEnde.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                waehleZeit(t.ende, new Zeitwahl() {
                    @Override
                    public void gewaehlt(String hhmm) {
                        t.ende = hhmm;
                        beschriften.run();
                    }
                });
            }
        });

        new AlertDialog.Builder(this)
                .setTitle(Speicher.anzeigeDatum(datum))
                .setView(box)
                .setPositiveButton("Speichern", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface d, int w) {
                        int b = Speicher.minuten(t.beginn), e = Speicher.minuten(t.ende);
                        if (b >= 0 && e >= 0 && e <= b) {
                            Toast.makeText(MainActivity.this, "Gehen muss nach Kommen liegen – nicht gespeichert.", Toast.LENGTH_LONG).show();
                            return;
                        }
                        t.uebertragen = false;
                        speicher.speichere(t);
                        aktualisiereAlles();
                    }
                })
                .setNeutralButton("Tag löschen", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface d, int w) {
                        new AlertDialog.Builder(MainActivity.this)
                                .setMessage("Tag " + Speicher.anzeigeDatum(datum) + " wirklich löschen?\n(Was schon am PC ist, bleibt dort.)")
                                .setPositiveButton("Löschen", new DialogInterface.OnClickListener() {
                                    @Override
                                    public void onClick(DialogInterface d2, int w2) {
                                        speicher.loesche(datum);
                                        aktualisiereAlles();
                                    }
                                })
                                .setNegativeButton("Abbrechen", null)
                                .show();
                    }
                })
                .setNegativeButton("Abbrechen", null)
                .show();
    }

    private interface Zeitwahl {
        void gewaehlt(String hhmm);
    }

    private void waehleZeit(String vorgabe, final Zeitwahl w) {
        int m = Speicher.minuten(vorgabe);
        if (m < 0) m = Speicher.minuten(Speicher.jetzt());
        new TimePickerDialog(this, new TimePickerDialog.OnTimeSetListener() {
            @Override
            public void onTimeSet(TimePicker view, int h, int min) {
                w.gewaehlt(Speicher.hhmm(h * 60 + min));
            }
        }, m / 60, m % 60, true).show();
    }

    // ------------------------------------------------------------------ Uebertragen

    private void uebertragen() {
        final TextView info = text("", 15, false, Color.rgb(30, 30, 30));
        info.setPadding(dp(20), dp(12), dp(20), 0);
        final TextView status = text("Warte auf den PC …", 15, true, BLAU);
        status.setPadding(dp(20), dp(12), dp(20), dp(8));
        LinearLayout box = new LinearLayout(this);
        box.setOrientation(LinearLayout.VERTICAL);
        box.addView(info);
        box.addView(status);

        int offen = speicher.offeneTage().size();
        List<String> ips = Uebertragung.eigeneAdressen();
        StringBuilder sb = new StringBuilder();
        sb.append("Code: ").append(speicher.code()).append("\n");
        sb.append("Adresse: ").append(ips.isEmpty() ? "– (kein WLAN?)" : android.text.TextUtils.join(", ", ips)).append("\n\n");
        sb.append(offen).append(" Tag(e) noch nicht am PC.\n\n");
        sb.append("1. Handy und PC verbinden: gleiches WLAN, Hotspot am PC oder Hotspot am Handy.\n");
        sb.append("2. Am PC: Rechtsklick auf das Uhr-Symbol → „Vom Handy übernehmen“ → „Übernehmen“.\n");
        sb.append("3. Diese Anzeige offen lassen, bis „Fertig“ erscheint.");
        info.setText(sb.toString());

        final Uebertragung u = new Uebertragung(speicher, new Uebertragung.Meldung() {
            @Override
            public void status(final String text) {
                runOnUiThread(new Runnable() {
                    @Override
                    public void run() {
                        status.setText(text);
                    }
                });
            }

            @Override
            public void bestaetigt(final List<String> tage) {
                runOnUiThread(new Runnable() {
                    @Override
                    public void run() {
                        status.setTextColor(GRUEN);
                        status.setText("Fertig – " + tage.size() + " Tag(e) vom PC übernommen.");
                        aktualisiereTage();
                    }
                });
            }
        });
        try {
            u.start();
        } catch (Exception e) {
            new AlertDialog.Builder(this).setTitle("Übertragen")
                    .setMessage("Die Übertragung konnte nicht gestartet werden:\n" + e.getMessage())
                    .setPositiveButton("OK", null).show();
            return;
        }
        getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
        AlertDialog d = new AlertDialog.Builder(this)
                .setTitle("An PC übertragen")
                .setView(box)
                .setPositiveButton("Schließen", null)
                .create();
        d.setOnDismissListener(new DialogInterface.OnDismissListener() {
            @Override
            public void onDismiss(DialogInterface dialog) {
                u.stop();
                getWindow().clearFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
                aktualisiereTage();
            }
        });
        d.show();
    }
}
