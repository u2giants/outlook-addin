using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PopOutlookAddin
{
    /// <summary>Port of ThisOutlookSession "Retention Column Toolkit".</summary>
    internal static class Retention
    {
        internal const bool EnableLiveStamping = true;
        private const int BatchSizeManual = 6000;
        private const int YieldEvery = 50;
        private const bool SkipAlreadyStamped = false;

        private const string PR_POLICY_TAG = "http://schemas.microsoft.com/mapi/proptag/0x30190102";
        private const string PR_RETENTION_PERIOD = "http://schemas.microsoft.com/mapi/proptag/0x301A0003";
        private const string PR_RETENTION_DATE = "http://schemas.microsoft.com/mapi/proptag/0x301C0040";

        private static readonly DateTime MinValid = new DateTime(1900, 1, 1);
        private static readonly DateTime MaxValid = new DateTime(4501, 1, 1);

        private static Outlook.Application App => Connect.App;

        private static void Release(object o) { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }

        // ---------------- Manual stampers ----------------

        internal static void StampCurrentFolder()
        {
            var exp = App.ActiveExplorer();
            if (exp == null) { Warn("Open a mail folder first."); return; }
            var fld = exp.CurrentFolder;
            if (fld.DefaultItemType != Outlook.OlItemType.olMailItem) { Warn("This folder isn't a mail folder."); return; }
            int processed = StampFolderBatch(fld, BatchSizeManual);
            Info("Stamped " + processed + " item(s) in: " + fld.FolderPath + "\r\nRun again for the next batch.");
        }

        internal static void StampCurrentMailbox(long limitTotal)
        {
            var exp = App.ActiveExplorer();
            if (exp == null) { Warn("Open a folder inside the mailbox you want to process."); return; }
            Outlook.Folder root;
            try { root = (Outlook.Folder)exp.CurrentFolder.Store.GetRootFolder(); }
            catch { Warn("Couldn't resolve the mailbox root. Select a folder in the target mailbox and try again."); return; }

            var skipped = new StringBuilder();
            long done = StampTree(root, limitTotal, skipped);
            var msg = "Mailbox pass complete.\r\nStamped: " + done + " item(s) (cap " + limitTotal + ").";
            if (skipped.Length > 0) msg += "\r\n\r\nSkipped:\r\n" + skipped;
            Info(msg);
        }

        private static long StampTree(Outlook.Folder fld, long remaining, StringBuilder skipped)
        {
            long total = 0;
            if (remaining <= 0) return 0;
            try
            {
                if (fld.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    int done = StampFolderBatch(fld, (int)Math.Min(remaining, BatchSizeManual));
                    remaining -= done; total += done;
                }
                foreach (Outlook.Folder sub in fld.Folders)
                {
                    if (remaining <= 0) break;
                    long got = StampTree(sub, remaining, skipped);
                    total += got; remaining -= got;
                    Release(sub);
                }
            }
            catch
            {
                skipped.Append("- ").Append(SafePath(fld)).Append(" (access error)\r\n");
            }
            return total;
        }

        private static int StampFolderBatch(Outlook.MAPIFolder fld, int limit)
        {
            var itms = fld.Items;
            try { itms.Sort("[ReceivedTime]", true); } catch { }
            int done = 0;
            for (int i = itms.Count; i >= 1; i--)
            {
                object it;
                try { it = itms[i]; } catch { continue; }
                try
                {
                    if (it is Outlook.MailItem mi && mi.MessageClass.StartsWith("IPM.Note", StringComparison.Ordinal))
                    {
                        if (!(SkipAlreadyStamped && IsAlreadyStamped(mi)) && StampOneItem(mi))
                        {
                            done++;
                            if (done % YieldEvery == 0) Application.DoEvents();
                            if (done >= limit) break;
                        }
                    }
                }
                finally { Release(it); }
            }
            Release(itms);
            return done;
        }

        // ---------------- Core ----------------

        /// <summary>True if the "Retention" user property changed (saves only on change).</summary>
        internal static bool StampOneItem(Outlook.MailItem mi)
        {
            try
            {
                var pa = mi.PropertyAccessor;
                long days = 0; bool haveDays = false;

                try { days = Convert.ToInt64(pa.GetProperty(PR_RETENTION_PERIOD)); haveDays = days > 0; } catch { }

                if (!haveDays)
                {
                    DateTime exp = default; bool haveExp = false;
                    try { exp = (DateTime)pa.GetProperty(PR_RETENTION_DATE); haveExp = exp > MinValid && exp < MaxValid; } catch { }
                    if (!haveExp)
                    {
                        try { exp = mi.RetentionExpirationDate; haveExp = exp > MinValid && exp < MaxValid; } catch { }
                    }
                    if (haveExp)
                    {
                        var baseDate = PickBestBaseDate(mi);
                        if (baseDate > MinValid && exp > baseDate)
                        {
                            days = (long)Math.Round((exp - baseDate).TotalDays); // VBA CLng = banker's rounding
                            haveDays = days > 0;
                        }
                    }
                }

                string valText = haveDays ? DaysToFriendly(days) : "";
                bool changed = false;

                var upVal = EnsureUserProp(mi, "Retention");
                if (Convert.ToString(upVal.Value) != valText) { upVal.Value = valText; changed = true; }

                try
                {
                    if (pa.GetProperty(PR_POLICY_TAG) is byte[] tag)
                    {
                        var upHex = EnsureUserProp(mi, "Retention (TagHex)");
                        var hex = BitConverter.ToString(tag).Replace("-", "");
                        if (Convert.ToString(upHex.Value) != hex) { upHex.Value = hex; changed = true; }
                    }
                }
                catch { }

                if (changed) mi.Save();
                return changed;
            }
            catch { return false; }
        }

        private static Outlook.UserProperty EnsureUserProp(Outlook.MailItem mi, string name)
        {
            return mi.UserProperties.Find(name) ?? mi.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
        }

        private static bool IsAlreadyStamped(Outlook.MailItem mi)
        {
            try { var up = mi.UserProperties.Find("Retention"); return up != null && Convert.ToString(up.Value).Trim().Length > 0; }
            catch { return false; }
        }

        private static DateTime PickBestBaseDate(Outlook.MailItem mi)
        {
            try { if (mi.ReceivedTime > MinValid) return mi.ReceivedTime; } catch { }
            try { if (mi.SentOn > MinValid) return mi.SentOn; } catch { }
            try { return mi.CreationTime; } catch { return default; }
        }

        internal static string DaysToFriendly(long days)
        {
            if (days <= 0) return "";
            if (days < 60) return days + (days == 1 ? " Day" : " Days");
            long months = (long)Math.Floor(days / 30.0 + 0.5);
            if (days < 730) { if (months < 1) months = 1; return months + (months == 1 ? " Mo" : " Mos"); }
            long years = (long)Math.Floor(days / 365.0 + 0.5);
            if (years < 1) years = 1;
            return years + (years == 1 ? " Yr" : " Yrs");
        }

        private static string SafePath(Outlook.MAPIFolder fld)
        {
            try { var p = fld.FolderPath; return string.IsNullOrEmpty(p) ? "(unknown folder)" : p; } catch { return "(unknown folder)"; }
        }

        // ---------------- Field / view utilities ----------------

        internal static void RegisterFieldOnCurrentMailbox()
        {
            var exp = App.ActiveExplorer();
            if (exp == null) { Warn("Open a folder in the mailbox first."); return; }
            Outlook.Folder root;
            try { root = (Outlook.Folder)exp.CurrentFolder.Store.GetRootFolder(); }
            catch { Warn("Could not get current mailbox root."); return; }

            var skipped = new StringBuilder();
            long count = RegisterTree(root, skipped);
            var msg = "Registered 'Retention' on " + count + " folder(s) in this mailbox.";
            if (skipped.Length > 0) msg += "\r\n\r\nSkipped:\r\n" + skipped;
            Info(msg);
        }

        private static long RegisterTree(Outlook.Folder fld, StringBuilder skipped)
        {
            long touched = 0;
            try
            {
                if (fld.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    if (AddField(fld, "Retention")) touched++;
                    else skipped.Append("- ").Append(SafePath(fld)).Append(" (no write)\r\n");
                    AddField(fld, "Retention (TagHex)");
                }
            }
            catch { skipped.Append("- ").Append(SafePath(fld)).Append(" (access error)\r\n"); return touched; }

            Outlook.Folders children;
            try { children = fld.Folders; }
            catch { skipped.Append("- ").Append(SafePath(fld)).Append(" (cannot enumerate children)\r\n"); return touched; }
            foreach (Outlook.Folder sub in children)
            {
                try { touched += RegisterTree(sub, skipped); }
                catch { skipped.Append("- ").Append(SafePath(sub)).Append(" (child access error)\r\n"); }
            }
            return touched;
        }

        private static bool AddField(Outlook.Folder fld, string name)
        {
            try
            {
                var udfs = fld.UserDefinedProperties;
                if (udfs == null) return false;
                if (udfs.Find(name) == null) udfs.Add(name, Outlook.OlUserPropertyType.olText);
                return true;
            }
            catch { return false; }
        }

        internal static void ApplyActiveViewToCurrentMailbox()
        {
            var exp = App.ActiveExplorer();
            if (exp == null) { Warn("Open a mail folder first."); return; }
            var src = (Outlook.View)exp.CurrentView;
            try { src.Save(); } catch { }
            Outlook.Folder root;
            try { root = (Outlook.Folder)exp.CurrentFolder.Store.GetRootFolder(); }
            catch { Warn("Could not get current mailbox root."); return; }
            long n = CopyViewToTree(root, src, "Retention View");
            Info("Applied your current view to " + n + " mail folder(s) in this mailbox.");
        }

        private static long CopyViewToTree(Outlook.Folder fld, Outlook.View src, string viewName)
        {
            long applied = 0;
            try
            {
                if (fld.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    try { var existing = fld.Views[viewName]; existing?.Delete(); } catch { }
                    try
                    {
                        // Create the view in the target folder and copy the source layout into it.
                        var nv = fld.Views.Add(viewName, src.ViewType, Outlook.OlViewSaveOption.olViewSaveOptionThisFolderEveryone);
                        nv.XML = src.XML;
                        nv.Save();
                        applied++;
                        try { nv.Apply(); } catch { }
                    }
                    catch { }
                }
                foreach (Outlook.Folder sub in fld.Folders) applied += CopyViewToTree(sub, src, viewName);
            }
            catch { }
            return applied;
        }

        internal static void Warn(string m) => MessageBox.Show(m, "POP Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        internal static void Info(string m) => MessageBox.Show(m, "POP Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
