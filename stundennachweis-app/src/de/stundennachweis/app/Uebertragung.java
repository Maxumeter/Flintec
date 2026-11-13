package de.stundennachweis.app;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.Inet4Address;
import java.net.InetAddress;
import java.net.InetSocketAddress;
import java.net.NetworkInterface;
import java.net.ServerSocket;
import java.net.Socket;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Kleiner Server, der nur laeuft, solange der Bildschirm "An PC uebertragen" offen ist.
 * Das PC-Programm holt die Daten ab (der PC baut die Verbindung auf - dafuer braucht
 * der PC weder Administratorrechte noch eine Firewall-Freigabe).
 *
 *   GET /daten?code=123456                  -> noch nicht uebertragene Tage (JSON)
 *   GET /bestaetigen?code=123456&tage=a,b   -> Tage als uebertragen markieren
 */
public final class Uebertragung {

    public static final int PORT = 8765;

    public interface Meldung {
        void status(String text);

        void bestaetigt(List<String> tage);
    }

    private final Speicher speicher;
    private final Meldung meldung;
    private volatile ServerSocket server;
    private Thread thread;

    public Uebertragung(Speicher speicher, Meldung meldung) {
        this.speicher = speicher;
        this.meldung = meldung;
    }

    public void start() throws Exception {
        ServerSocket s = new ServerSocket();
        s.setReuseAddress(true);
        s.bind(new InetSocketAddress(PORT));
        server = s;
        thread = new Thread(new Runnable() {
            @Override
            public void run() {
                schleife();
            }
        }, "Uebertragung");
        thread.setDaemon(true);
        thread.start();
    }

    public void stop() {
        try {
            if (server != null) server.close();
        } catch (Exception ignored) {
        }
        server = null;
    }

    private void schleife() {
        while (server != null && !server.isClosed()) {
            try {
                Socket c = server.accept();
                c.setSoTimeout(5000);
                bearbeite(c);
            } catch (Exception ignored) {
                // Server geschlossen oder Verbindung abgebrochen
            }
        }
    }

    private void bearbeite(Socket c) {
        try {
            BufferedReader in = new BufferedReader(new InputStreamReader(c.getInputStream(), StandardCharsets.UTF_8));
            String zeile = in.readLine();
            if (zeile == null) return;
            // Kopfzeilen ueberspringen
            String h;
            while ((h = in.readLine()) != null && !h.isEmpty()) { /* nichts */ }

            String[] teile = zeile.split(" ");
            String pfad = teile.length > 1 ? teile[1] : "/";
            Map<String, String> q = new HashMap<>();
            int f = pfad.indexOf('?');
            if (f >= 0) {
                for (String paar : pfad.substring(f + 1).split("&")) {
                    int g = paar.indexOf('=');
                    if (g > 0) q.put(paar.substring(0, g), URLDecoder.decode(paar.substring(g + 1), "UTF-8"));
                }
                pfad = pfad.substring(0, f);
            }

            if (!speicher.code().equals(q.get("code"))) {
                antworte(c, 403, "{\"fehler\":\"code\"}");
                meldung.status("Ein PC hat sich mit falschem Code gemeldet.");
                return;
            }

            if (pfad.equals("/daten")) {
                JSONObject o = new JSONObject();
                o.put("app", "Stundennachweis");
                o.put("version", 1);
                JSONArray a = new JSONArray();
                for (Speicher.Tag t : speicher.offeneTage()) a.put(Speicher.tagJson(t));
                o.put("tage", a);
                antworte(c, 200, o.toString());
                meldung.status("PC hat " + a.length() + " Tag(e) abgeholt …");
            } else if (pfad.equals("/bestaetigen")) {
                List<String> tage = new ArrayList<>();
                String liste = q.get("tage");
                if (liste != null) for (String d : liste.split(",")) if (!d.trim().isEmpty()) tage.add(d.trim());
                speicher.alsUebertragenMarkieren(tage);
                antworte(c, 200, "{\"ok\":true}");
                meldung.bestaetigt(tage);
            } else if (pfad.equals("/hallo")) {
                antworte(c, 200, "{\"app\":\"Stundennachweis\"}");
            } else {
                antworte(c, 404, "{}");
            }
        } catch (Exception e) {
            try {
                antworte(c, 500, "{}");
            } catch (Exception ignored) {
            }
        } finally {
            try {
                c.close();
            } catch (Exception ignored) {
            }
        }
    }

    private static void antworte(Socket c, int code, String json) throws Exception {
        byte[] body = json.getBytes(StandardCharsets.UTF_8);
        String text = code == 200 ? "OK" : (code == 403 ? "Forbidden" : (code == 404 ? "Not Found" : "Error"));
        String kopf = "HTTP/1.0 " + code + " " + text + "\r\n"
                + "Content-Type: application/json; charset=utf-8\r\n"
                + "Content-Length: " + body.length + "\r\n"
                + "Connection: close\r\n\r\n";
        OutputStream out = c.getOutputStream();
        out.write(kopf.getBytes(StandardCharsets.US_ASCII));
        out.write(body);
        out.flush();
    }

    /** IPv4-Adressen des Handys (zur Anzeige, falls der PC sie nicht selbst findet). */
    public static List<String> eigeneAdressen() {
        List<String> liste = new ArrayList<>();
        try {
            for (NetworkInterface ni : Collections.list(NetworkInterface.getNetworkInterfaces())) {
                if (!ni.isUp() || ni.isLoopback()) continue;
                for (InetAddress a : Collections.list(ni.getInetAddresses())) {
                    if (a instanceof Inet4Address && !a.isLoopbackAddress()) liste.add(a.getHostAddress());
                }
            }
        } catch (Exception ignored) {
        }
        return liste;
    }
}
