using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace PopOutlookAddin
{
    /// <summary>
    /// Self-update from the shared Dropbox folder: if a newer add-in DLL is there, install it silently
    /// after Outlook closes. Every PC picks up a new version on its next Outlook restart.
    /// </summary>
    internal static class Updater
    {
        internal const string RelativeFolder = @"ai\outlook-addin";
        internal const string PackageDll = "PopOutlookAddin.dll";
        private static string _pending;

        internal static void CheckForUpdate()
        {
            try
            {
                foreach (var root in DropboxRoots())
                {
                    var setup = Path.Combine(root, RelativeFolder, PackageDll);
                    if (!File.Exists(setup)) continue;
                    var available = new Version(FileVersionInfo.GetVersionInfo(setup).FileVersion ?? "0.0.0.0");
                    var mine = Assembly.GetExecutingAssembly().GetName().Version;
                    if (available > mine) { _pending = setup; Log.Write("Update pending: " + mine + " -> " + available); }
                    return;
                }
            }
            catch (Exception ex) { Log.Write("CheckForUpdate", ex); }
        }

        internal static void RunPendingUpdate()
        {
            if (_pending == null) CheckForUpdate(); // also catch versions published while Outlook was open
            if (_pending == null) return;
            var pending = _pending; _pending = null;
            try
            {
                // Copy the package out of Dropbox first so a sync mid-install cannot disturb it.
                var src = Path.GetDirectoryName(pending);
                var tmp = Path.Combine(Path.GetTempPath(), "PopOutlookAddin-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tmp);
                foreach (var f in new[] { "PopOutlookAddin.dll", "Microsoft.Office.Interop.Outlook.dll", "Install.ps1" })
                    File.Copy(Path.Combine(src, f), Path.Combine(tmp, f));
                Process.Start(new ProcessStartInfo("powershell.exe",
                    "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + Path.Combine(tmp, "Install.ps1") + "\" -Silent")
                { CreateNoWindow = true, UseShellExecute = false });
            }
            catch (Exception ex) { Log.Write("RunPendingUpdate", ex); }
        }

        private static string[] DropboxRoots()
        {
            var list = new System.Collections.Generic.List<string>();
            foreach (var env in new[] { "LOCALAPPDATA", "APPDATA" })
            {
                var info = Path.Combine(Environment.GetEnvironmentVariable(env) ?? "", "Dropbox", "info.json");
                if (!File.Exists(info)) continue;
                foreach (Match m in Regex.Matches(File.ReadAllText(info), "\"path\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\""))
                    list.Add(Regex.Unescape(m.Groups[1].Value));
            }
            list.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Dropbox"));
            return list.ToArray();
        }
    }
}
