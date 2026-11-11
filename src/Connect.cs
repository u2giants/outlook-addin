using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PopOutlookAddin
{
    /// <summary>COM entry point Outlook loads. Replaces ThisOutlookSession event wiring.</summary>
    [ComVisible(true)]
    [Guid("AB496E0D-CC70-4F80-854B-0EA3E4C4FBE9")]
    [ProgId("PopOutlookAddin.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        internal static Outlook.Application App;

        // Fields keep COM event sources alive (otherwise GC drops the events).
        private Outlook.Explorers _explorers;
        private Outlook.Explorer _explorer;
        private Outlook.Items _inboxItems;
        private Outlook.Items _sentItems;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            App = (Outlook.Application)Application;
        }

        public void OnStartupComplete(ref Array custom)
        {
            WireEventSinks();
            Updater.CheckForUpdate();
        }

        public void OnBeginShutdown(ref Array custom) { Updater.RunPendingUpdate(); }
        public void OnAddInsUpdate(ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_explorers != null) _explorers.NewExplorer -= OnNewExplorer;
                if (_explorer != null) ((Outlook.ExplorerEvents_10_Event)_explorer).SelectionChange -= OnSelectionChange;
                if (_inboxItems != null) { _inboxItems.ItemAdd -= OnStampItem; _inboxItems.ItemChange -= OnStampItem; }
                if (_sentItems != null) _sentItems.ItemAdd -= OnStampItem;
                if (App != null) App.ItemSend -= OnItemSend;
            }
            catch { }
            _explorers = null; _explorer = null; _inboxItems = null; _sentItems = null; App = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        private void WireEventSinks()
        {
            try
            {
                App.ItemSend += OnItemSend;
                // Outlook skips OnBeginShutdown on fast shutdown; Quit always fires.
                ((Outlook.ApplicationEvents_11_Event)App).Quit += Updater.RunPendingUpdate;
                _explorers = App.Explorers;
                _explorers.NewExplorer += OnNewExplorer;
                var exp = App.ActiveExplorer();
                if (exp == null && _explorers.Count > 0) exp = _explorers[1];
                AttachExplorer(exp);
                var ns = App.Session;
                _inboxItems = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;
                _inboxItems.ItemAdd += OnStampItem;
                _inboxItems.ItemChange += OnStampItem;
                _sentItems = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items;
                _sentItems.ItemAdd += OnStampItem;
            }
            catch (Exception ex) { Log.Write("WireEventSinks", ex); }
        }

        private void AttachExplorer(Outlook.Explorer exp)
        {
            if (exp == null) return;
            if (_explorer != null) ((Outlook.ExplorerEvents_10_Event)_explorer).SelectionChange -= OnSelectionChange;
            _explorer = exp;
            ((Outlook.ExplorerEvents_10_Event)_explorer).SelectionChange += OnSelectionChange;
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            try { AttachExplorer(explorer); } catch { }
        }

        private void OnSelectionChange()
        {
            if (!Retention.EnableLiveStamping) return;
            try
            {
                var sel = _explorer?.Selection;
                if (sel == null) return;
                foreach (object it in sel)
                    if (it is Outlook.MailItem mi) Retention.StampOneItem(mi);
            }
            catch { }
        }

        private void OnStampItem(object item)
        {
            if (!Retention.EnableLiveStamping) return;
            try { if (item is Outlook.MailItem mi) Retention.StampOneItem(mi); } catch { }
        }

        private void OnItemSend(object item, ref bool cancel)
        {
            if (!Retention.EnableLiveStamping) return;
            try { if (item is Outlook.MailItem mi) Retention.StampOneItem(mi); } catch { }
        }

        // ---------------- Ribbon ----------------

        public string GetCustomUI(string RibbonID)
        {
            if (RibbonID != "Microsoft.Outlook.Explorer") return null;
            using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("PopOutlookAddin.Ribbon.xml"))
            using (var r = new StreamReader(s)) return r.ReadToEnd();
        }

        public void OnAction(IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnStampFolder": Retention.StampCurrentFolder(); break;
                    case "btnStampMailbox": Retention.StampCurrentMailbox(6000); break;
                    case "btnStampMailboxCustom":
                        var txt = Microsoft.VisualBasic.Interaction.InputBox("How many emails at most?", "Stamp Mailbox", "3000");
                        if (long.TryParse(txt, out var n) && n > 0) Retention.StampCurrentMailbox(n);
                        break;
                    case "btnRegisterField": Retention.RegisterFieldOnCurrentMailbox(); break;
                    case "btnApplyView": Retention.ApplyActiveViewToCurrentMailbox(); break;
                    case "btnRemoveInline": AttachmentKit.RemoveInlineImagesSelected(); break;
                    case "btnStripAttach": AttachmentKit.StripAttachmentsSelected(); break;
                    case "btnStripAll": AttachmentKit.StripEverythingSelected(); break;
                }
            }
            catch (Exception ex)
            {
                Log.Write(control.Id, ex);
                MessageBox.Show("Something went wrong: " + ex.Message, "POP Tools", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    internal static class Log
    {
        internal static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PopOutlookAddin");
        internal static void Write(string where, Exception ex) => Write(where + ": " + ex);
        internal static void Write(string msg)
        {
            try
            {
                Directory.CreateDirectory(Dir);
                File.AppendAllText(Path.Combine(Dir, "addin.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg + Environment.NewLine);
            }
            catch { }
        }
    }
}
