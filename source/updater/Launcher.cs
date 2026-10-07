using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Collections.Generic;
using System.Windows.Forms;

[assembly: AssemblyVersion("1.9.4.0")]
[assembly: AssemblyFileVersion("1.9.4.0")]
[assembly: AssemblyTitle("Flintec Control Center – Auto Update")]
internal static class Launcher
{
    const string Bundled = "1.9.4";
    const string Api = "https://api.github.com/repos/Maxumeter/Flintec/releases/latest";
    const string BaseHash = "fd60169e6239d9977cbcf95d15349cd0512b4f1ab1944ebb0686c205987f4aae";
    static string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Flintec Control Center AutoUpdate");
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    static Func<Dictionary<string, object>> TestRelease;
    static Action<string, string> TestDownload;
    static string LogPath { get { return Path.Combine(Root, "updater.log"); } }
    static void Log(string s) { try { if (File.Exists(LogPath) && new FileInfo(LogPath).Length > 1000000) File.Move(LogPath, LogPath + "." + DateTime.UtcNow.Ticks); File.AppendAllText(LogPath, DateTime.UtcNow.ToString("o") + " " + s + Environment.NewLine); } catch {} }
    static string Str(Dictionary<string, object> d, string k) { return d.ContainsKey(k) ? Convert.ToString(d[k]) : ""; }
    internal static Version Parse(string s) { Version v; return System.Text.RegularExpressions.Regex.IsMatch(s ?? "", @"^v?\d+\.\d+\.\d+$") && Version.TryParse(s.TrimStart('v'), out v) ? v : null; }
    internal static bool AssetName(string name, Version v) { return name == "Flintec_ControlCenter_App.exe" || name == "Flintec_ControlCenter_" + v + "_Portable.exe"; }
    internal static bool GoodUrl(string s) { Uri u; return Uri.TryCreate(s, UriKind.Absolute, out u) && u.Scheme == "https" && u.Host == "github.com" && u.AbsolutePath.StartsWith("/Maxumeter/Flintec/releases/download/", StringComparison.Ordinal); }
    static string Hash(string path) { using (var f = File.OpenRead(path)) using (var h = SHA256.Create()) return BitConverter.ToString(h.ComputeHash(f)).Replace("-", "").ToLowerInvariant(); }
    static void Verify(string path, string expected) { if (Hash(path) != expected) throw new InvalidDataException("SHA-256 stimmt nicht überein."); using (var f = File.OpenRead(path)) { if (f.ReadByte() != 'M' || f.ReadByte() != 'Z') throw new InvalidDataException("Keine Windows-EXE."); } }
    static HttpWebRequest Request(string url) { var r = (HttpWebRequest)WebRequest.Create(url); r.UserAgent = "Flintec-ControlCenter-Updater/1.9.4"; r.Accept = "application/vnd.github+json"; r.Timeout = 6000; r.ReadWriteTimeout = 15000; return r; }
    static Dictionary<string, object> Latest() { using (var response = Request(Api).GetResponse()) using (var reader = new StreamReader(response.GetResponseStream())) return Json.Deserialize<Dictionary<string, object>>(reader.ReadToEnd()); }
    static void Download(string url, string target) { using (var response = Request(url).GetResponse()) using (var input = response.GetResponseStream()) using (var output = File.Create(target)) { byte[] b = new byte[65536]; int n; long total = 0; var watch = Stopwatch.StartNew(); while ((n = input.Read(b, 0, b.Length)) > 0) { total += n; if (total > 150000000 || watch.Elapsed.TotalMinutes > 3) throw new IOException("Download-Limit überschritten."); output.Write(b, 0, n); } } }
    static string AppDir(Version v) { return Path.Combine(Root, "versions", v.ToString()); }
    static string AppPath(Version v) { return Path.Combine(AppDir(v), "Flintec_ControlCenter_App.exe"); }
    static void EnsureBundled() { var v = Parse(Bundled); Directory.CreateDirectory(AppDir(v)); string p = AppPath(v); if (!File.Exists(p) || Hash(p) != BaseHash) { string tmp = p + ".tmp"; using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("FlintecApp")) using (var output = File.Create(tmp)) input.CopyTo(output); Verify(tmp, BaseHash); if (File.Exists(p)) File.Delete(p); File.Move(tmp, p); } }
    static Version Installed() { Version best = Parse(Bundled); foreach (var dir in Directory.GetDirectories(Path.Combine(Root, "versions"))) { var v = Parse(Path.GetFileName(dir)); if (v == null || v <= best) continue; try { Verify(AppPath(v), File.ReadAllText(Path.Combine(dir, "sha256.txt")).Trim()); best = v; } catch { Log("Ungültige lokale Version ignoriert: " + v); } } return best; }
    static Version Update(Version current) {
        var release = TestRelease == null ? Latest() : TestRelease(); var remote = Parse(Str(release, "tag_name"));
        if (remote == null || (release.ContainsKey("prerelease") && Convert.ToBoolean(release["prerelease"])) || (release.ContainsKey("draft") && Convert.ToBoolean(release["draft"]))) throw new InvalidDataException("Kein gültiges stabiles Release.");
        Log("Installiert=" + current + "; GitHub=" + remote);
        if (remote <= current) return current;
        var assets = release["assets"] as System.Collections.IEnumerable;
        foreach (Dictionary<string, object> asset in assets) {
            if (!AssetName(Str(asset, "name"), remote)) continue;
            string url = Str(asset, "browser_download_url"), digest = Str(asset, "digest");
            if (!GoodUrl(url) || !System.Text.RegularExpressions.Regex.IsMatch(digest, "^sha256:[0-9a-fA-F]{64}$")) throw new InvalidDataException("Update-URL oder GitHub-Prüfsumme fehlt/ungültig.");
            string expected = digest.Substring(7).ToLowerInvariant();
            Directory.CreateDirectory(AppDir(remote)); string tmp = Path.Combine(AppDir(remote), Guid.NewGuid() + ".download");
            try { Log("Lade " + Str(asset, "name")); if (TestDownload == null) Download(url, tmp); else TestDownload(url, tmp); Verify(tmp, expected); string dest = AppPath(remote); if (File.Exists(dest)) File.Delete(dest); File.Move(tmp, dest); File.WriteAllText(Path.Combine(AppDir(remote), "sha256.txt"), expected); Log("Update geprüft und bereit: " + remote); return remote; }
            finally { if (File.Exists(tmp)) File.Delete(tmp); }
        }
        Log("Release " + remote + " enthält keine unterstützte App-EXE; vorhandene Version bleibt aktiv.");
        return current;
    }
    static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
    static int SelfTest() {
        Assert(Parse("v1.9.5") > Parse("1.9.4"), "upgrade"); Assert(Parse("v1.9.3") < Parse(Bundled), "downgrade"); Assert(Parse("v1.10.0") > Parse("1.9.9"), "numeric order");
        Assert(Parse("v1.9.5-beta") == null && Parse("../../foo") == null, "version validation");
        Assert(AssetName("Flintec_ControlCenter_App.exe", Parse(Bundled)), "app asset"); Assert(AssetName("Flintec_ControlCenter_1.9.4_Portable.exe", Parse(Bundled)), "portable asset");
        Assert(!AssetName("Flintec_ControlCenter_Setup_1.9.4.exe", Parse(Bundled)), "no installer"); Assert(!AssetName("Flintec_ControlCenter_1.9.4_AutoUpdate.exe", Parse(Bundled)), "no recursive launcher");
        Assert(GoodUrl("https://github.com/Maxumeter/Flintec/releases/download/v1.9.5/Flintec_ControlCenter_App.exe"), "release URL"); Assert(!GoodUrl("https://github.com/Other/Flintec/releases/download/v1/x.exe"), "foreign repo");
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("FlintecApp")) using (var h = SHA256.Create()) Assert(BitConverter.ToString(h.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() == BaseHash, "embedded payload");
        Root = Path.Combine(Path.GetTempPath(), "FlintecUpdaterTest-" + Guid.NewGuid()); Directory.CreateDirectory(Root); EnsureBundled();
        var asset = new Dictionary<string, object> { {"name", "Flintec_ControlCenter_App.exe"}, {"digest", "sha256:" + BaseHash}, {"browser_download_url", "https://github.com/Maxumeter/Flintec/releases/download/v1.9.5/Flintec_ControlCenter_App.exe"} };
        var release = new Dictionary<string, object> { {"tag_name", "v1.9.5"}, {"assets", new object[] {asset}} };
        TestRelease = delegate { return release; }; TestDownload = delegate(string url, string dest) { File.Copy(AppPath(Parse(Bundled)), dest); };
        Assert(Update(Parse(Bundled)) == Parse("1.9.5") && Installed() == Parse("1.9.5"), "verified update promotion");
        release["tag_name"] = "v1.9.3"; TestDownload = delegate { throw new Exception("must not download downgrade"); }; Assert(Update(Parse("1.9.5")) == Parse("1.9.5"), "no downgrade download");
        release["tag_name"] = "v1.9.6"; TestDownload = delegate(string url, string dest) { File.WriteAllText(dest, "corrupt download"); }; bool rejected = false; try { Update(Parse("1.9.5")); } catch (InvalidDataException) { rejected = true; }
        Assert(rejected && Installed() == Parse("1.9.5") && !File.Exists(AppPath(Parse("1.9.6"))), "corrupt update preserves installed version");
        asset["name"] = "Flintec_ControlCenter_Setup_1.9.6.exe"; Assert(Update(Parse("1.9.5")) == Parse("1.9.5"), "setup-only release keeps installed version");
        Console.WriteLine("PASS: version ordering, downgrade prevention, asset selection, URL validation, embedded SHA-256, verified update promotion, corrupt download rejection, setup-only fallback");
        Console.WriteLine("Test artifacts: " + Root); return 0;
    }
    [STAThread] static int Main(string[] args) {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        if (args.Length == 1 && args[0] == "--self-test") return SelfTest();
        if (args.Length == 1 && args[0] == "--check-only") { var r = Latest(); Console.WriteLine("Bundled=" + Bundled + "; GitHub=" + Str(r,"tag_name") + "; newer=" + (Parse(Str(r,"tag_name")) > Parse(Bundled))); return 0; }
        Directory.CreateDirectory(Root);
        using (var mutex = new Mutex(false, "Local\\FlintecControlCenterAutoUpdate")) {
            bool held = false;
            try {
                try { held = mutex.WaitOne(0); } catch (AbandonedMutexException) { held = true; }
                if (!held) { MessageBox.Show("Flintec Control Center läuft bereits über den Auto-Updater."); return 0; }
                EnsureBundled(); Version current = Installed(), chosen = current;
                try { chosen = Update(current); } catch (Exception e) { Log("Update nicht verfügbar: " + e.Message + "; starte " + current); }
                Process child;
                try { child = Process.Start(new ProcessStartInfo(AppPath(chosen)) { WorkingDirectory = AppDir(chosen), UseShellExecute = false }); }
                catch (Exception e) { if (chosen == current) throw; Log("Start fehlgeschlagen: " + e.Message + "; Rückfall auf " + current); File.Delete(Path.Combine(AppDir(chosen), "sha256.txt")); child = Process.Start(new ProcessStartInfo(AppPath(current)) { WorkingDirectory = AppDir(current), UseShellExecute = false }); }
                Log("Programm gestartet; Version=" + chosen); child.WaitForExit(); child.Dispose(); return 0;
            } catch (Exception e) { Log(e.ToString()); MessageBox.Show("Flintec konnte nicht gestartet werden: " + e.Message + "\nLog: " + LogPath, "Flintec Control Center"); return 1; }
            finally { if (held) mutex.ReleaseMutex(); }
        }
    }
}
