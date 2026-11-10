// Kleiner Starter fuer Stundennachweis.ps1.
// Startet PowerShell ohne sichtbares Konsolenfenster und reicht die Aktion durch.
//   Stundennachweis.exe                -> Fenster oeffnen (bzw. in den Vordergrund holen)
//   Stundennachweis.exe Kommen         -> beim Anmelden/Entsperren (Aufgabenplanung)
//   Stundennachweis.exe Herunterfahren -> beim Herunterfahren (Aufgabenplanung)
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("Stundennachweis")]
[assembly: AssemblyProduct("Stundennachweis")]
[assembly: AssemblyDescription("Traegt Arbeitszeiten in den Excel-Stundennachweis ein")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

static class Starter
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

    static readonly string[] Erlaubt = { "Kommen", "Beginn", "Gehen", "Herunterfahren", "Sync", "Test" };

    [STAThread]
    static int Main(string[] args)
    {
        string ordner = AppDomain.CurrentDomain.BaseDirectory;
        string skript = Path.Combine(ordner, "Stundennachweis.ps1");
        if (!File.Exists(skript))
        {
            MessageBoxW(IntPtr.Zero, "Stundennachweis.ps1 wurde nicht gefunden in:\n" + ordner +
                "\n\nDie EXE muss im selben Ordner liegen wie die anderen Dateien.", "Stundennachweis", 0x10);
            return 2;
        }

        string aktion = "Beginn";
        if (args.Length > 0)
        {
            aktion = Array.Find(Erlaubt, a => string.Equals(a, args[0], StringComparison.OrdinalIgnoreCase));
            if (aktion == null)
            {
                MessageBoxW(IntPtr.Zero, "Unbekannte Aktion: " + args[0], "Stundennachweis", 0x10);
                return 2;
            }
        }

        string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
            @"WindowsPowerShell\v1.0\powershell.exe");
        var start = new ProcessStartInfo(powershell,
            "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + skript + "\" -Aktion " + aktion);
        start.UseShellExecute = false;
        start.CreateNoWindow = true;
        start.WorkingDirectory = ordner;

        try
        {
            using (Process p = Process.Start(start))
            {
                p.WaitForExit();
                return p.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBoxW(IntPtr.Zero, "PowerShell konnte nicht gestartet werden:\n" + ex.Message, "Stundennachweis", 0x10);
            return 1;
        }
    }
}
