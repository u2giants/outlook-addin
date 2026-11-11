using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PopOutlookAddin
{
    /// <summary>Port of modAttachInlineKit.</summary>
    internal static class AttachmentKit
    {
        private const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
        private const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
        private const RegexOptions Rx = RegexOptions.IgnoreCase | RegexOptions.Multiline;

        private static List<Outlook.MailItem> SelectedMail()
        {
            var exp = Connect.App.ActiveExplorer();
            var list = new List<Outlook.MailItem>();
            if (exp == null) return list;
            foreach (object o in exp.Selection) if (o is Outlook.MailItem mi) list.Add(mi);
            return list;
        }

        private static bool AnySelected()
        {
            var exp = Connect.App.ActiveExplorer();
            if (exp == null || exp.Selection == null || exp.Selection.Count == 0)
            { Retention.Warn("Select one or more emails first."); return false; }
            return true;
        }

        internal static void RemoveInlineImagesSelected()
        {
            if (!AnySelected()) return;
            int done = 0, skipped = 0;
            foreach (var mi in SelectedMail()) { if (RemoveInline(mi)) done++; else skipped++; }
            Retention.Info("Inline images removed from: " + done + "\r\nSkipped/Failed: " + skipped);
        }

        internal static void StripAttachmentsSelected()
        {
            if (!AnySelected()) return;
            if (!AskSavePath("Save copies of attachments before removing?", "Strip Attachments", out var savePath)) return;
            int processed = 0, failed = 0;
            foreach (var mi in SelectedMail()) { if (StripNonInline(mi, savePath, true, true)) processed++; else failed++; }
            var msg = "Processed: " + processed + "\r\nSkipped/Failed: " + failed;
            if (savePath != null) msg += "\r\nSaved to: " + savePath;
            MessageBox.Show(msg, "Strip Attachments", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal static void StripEverythingSelected()
        {
            if (!AnySelected()) return;
            if (!AskSavePath("Save copies of (non-inline) attachments before removing?", "Strip Everything", out var savePath)) return;
            var items = SelectedMail();
            int inDone = 0, inSkip = 0, atDone = 0, atFail = 0;
            foreach (var mi in items) { if (RemoveInline(mi)) inDone++; else inSkip++; }
            foreach (var mi in items) { if (StripNonInline(mi, savePath, true, true)) atDone++; else atFail++; }
            var msg = "Inline images removed from: " + inDone + "  (Skipped/Failed: " + inSkip + ")\r\n" +
                      "Attachments stripped from: " + atDone + "  (Skipped/Failed: " + atFail + ")";
            if (savePath != null) msg += "\r\nSaved copies to: " + savePath;
            MessageBox.Show(msg, "Strip Everything", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>False = user cancelled. savePath null = remove without saving.</summary>
        private static bool AskSavePath(string question, string title, out string savePath)
        {
            savePath = null;
            var r = MessageBox.Show(question + "\r\nYes = pick a folder and save; No = remove without saving.",
                                    title, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.No) return true;
            using (var dlg = new FolderBrowserDialog { Description = "Choose where to save the stripped attachments" })
            {
                if (dlg.ShowDialog() != DialogResult.OK || string.IsNullOrEmpty(dlg.SelectedPath)) return false;
                savePath = dlg.SelectedPath;
            }
            return true;
        }

        // ---------------- Inline image engine ----------------

        private static bool RemoveInline(Outlook.MailItem mi)
        {
            try
            {
                if (mi.BodyFormat != Outlook.OlBodyFormat.olFormatHTML)
                {
                    RemoveInlineAttachments(mi, null);
                    mi.Save();
                    return true;
                }
                string html = mi.HTMLBody;
                var cids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Collect(html, "src\\s*=\\s*['\"]cid:([^'\"]+)['\"]", cids);
                Collect(html, "(?:src|o:href)\\s*=\\s*['\"]cid:([^'\"]+)['\"]", cids);

                html = Regex.Replace(html, "<img\\b[^>]*\\bsrc\\s*=\\s*['\"]cid:[^'\"]+['\"][^>]*>", "", Rx);
                html = Regex.Replace(html, "<v:imagedata\\b[^>]*\\b(?:src|o:href)\\s*=\\s*['\"]cid:[^'\"]+['\"][^>]*/?>", "", Rx);
                html = Regex.Replace(html, "<v:shape\\b[^>]*>\\s*</v:shape>", "", Rx);
                html = Regex.Replace(html, "(?:\\r?\\n\\s*){3,}", "\r\n\r\n", Rx);

                mi.HTMLBody = html;
                RemoveInlineAttachments(mi, cids);
                mi.Save();
                return true;
            }
            catch (Exception ex) { Log.Write("RemoveInline", ex); return false; }
        }

        private static void Collect(string s, string pattern, HashSet<string> set)
        {
            foreach (Match m in Regex.Matches(s, pattern, Rx)) set.Add(m.Groups[1].Value.ToLowerInvariant());
        }

        private static string Cid(Outlook.Attachment a)
        {
            try { return Convert.ToString(a.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)) ?? ""; } catch { return ""; }
        }

        private static bool Hidden(Outlook.Attachment a)
        {
            try { return a.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN) is bool b && b; } catch { return false; }
        }

        private static void RemoveInlineAttachments(Outlook.MailItem mi, HashSet<string> cids)
        {
            for (int i = mi.Attachments.Count; i >= 1; i--)
            {
                try
                {
                    var att = mi.Attachments[i];
                    var cid = Cid(att);
                    bool isInline = cid.Length > 0 || Hidden(att) || att.Type != Outlook.OlAttachmentType.olByValue;
                    if (!isInline) continue;
                    if (cids == null) mi.Attachments.Remove(i);
                    else if (cid.Length > 0 && cids.Contains(cid.ToLowerInvariant())) mi.Attachments.Remove(i);
                }
                catch { }
            }
        }

        // ---------------- Non-inline attachment engine ----------------

        private static bool StripNonInline(Outlook.MailItem mi, string savePath, bool skipInline, bool addNote)
        {
            try
            {
                var removed = new StringBuilder();
                int removedCount = 0;
                for (int i = mi.Attachments.Count; i >= 1; i--)
                {
                    var att = mi.Attachments[i];
                    if (skipInline && IsInline(mi, att)) continue;
                    string name = att.FileName;
                    int size = att.Size;
                    if (savePath != null) att.SaveAsFile(UniquePath(savePath, Sanitize(name)));
                    mi.Attachments.Remove(i);
                    removedCount++;
                    removed.Append("• ").Append(name).Append(" (").Append((size / 1024.0).ToString("0.0")).Append(" KB)\r\n");
                }
                if (removedCount > 0)
                {
                    if (addNote) InsertRemovalNote(mi, removed.ToString());
                    mi.Save();
                }
                return true;
            }
            catch (Exception ex) { Log.Write("StripNonInline", ex); return false; }
        }

        private static bool IsInline(Outlook.MailItem mi, Outlook.Attachment att)
        {
            var cid = Cid(att);
            if (cid.Length > 0)
            {
                string body = "";
                try { body = mi.HTMLBody ?? ""; } catch { }
                if (body.IndexOf("cid:" + cid, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            if (Hidden(att)) return true;
            try { return att.Type != Outlook.OlAttachmentType.olByValue; } catch { return false; }
        }

        private static void InsertRemovalNote(Outlook.MailItem mi, string list)
        {
            try
            {
                var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                if (mi.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                {
                    var note = "<div style=\"border:1px solid #d0d0d0;padding:8px;margin:0 0 8px 0;font-family:Segoe UI,Arial,sans-serif;font-size:12px;background:#fffbe6\">" +
                               "<b>Attachments removed by maintenance macro</b><br/>" +
                               "<div style=\"white-space:pre-wrap\">" + WebUtility.HtmlEncode(list) + "</div>" +
                               "<div style=\"color:#666\">Date: " + WebUtility.HtmlEncode(stamp) + "</div></div>";
                    mi.HTMLBody = note + mi.HTMLBody;
                }
                else
                {
                    mi.Body = "=== Attachments removed by maintenance macro ===\r\n" + list + "Date: " + stamp + "\r\n\r\n" + mi.Body;
                }
            }
            catch { }
        }

        internal static string Sanitize(string name)
        {
            int pos = name.LastIndexOf('.');
            string b = pos > 0 ? name.Substring(0, pos) : name, ext = pos > 0 ? name.Substring(pos) : "";
            b = b.Replace("\\", "_").Replace("/", "_").Replace(":", "-").Replace("*", "_").Replace("?", "_")
                 .Replace("\"", "_").Replace("<", "(").Replace(">", ")").Replace("|", "_").Replace("\t", " ");
            ext = new string(ext.Where(c => Array.IndexOf(Path.GetInvalidFileNameChars(), c) < 0).ToArray());
            return b.Trim() + ext;
        }

        private static string UniquePath(string folder, string file)
        {
            int pos = file.LastIndexOf('.');
            string b = pos > 0 ? file.Substring(0, pos) : file, ext = pos > 0 ? file.Substring(pos) : "";
            string candidate = Path.Combine(folder, b + ext);
            for (int n = 1; File.Exists(candidate); n++) candidate = Path.Combine(folder, b + " (" + n + ")" + ext);
            return candidate;
        }
    }
}
