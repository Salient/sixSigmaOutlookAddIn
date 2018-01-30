using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Diagnostics;
using Timer = System.Timers.Timer;
using System.Text.RegularExpressions;

namespace sixSigmaSecureSend
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;
        private Outlook.Explorers explorers;
        Timer updateSuggestionFlag;


        private Dictionary<Object, editorWrapper> editorWrappersValue = new Dictionary<Object, editorWrapper>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // If somehow loading after windows are already open, find them all and bag 'n tag

            inspectors = this.Application.Inspectors;
            explorers = this.Application.Explorers;
            inspectors.NewInspector +=
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                Inspectors_NewInspector);

            explorers.NewExplorer +=
                new Outlook.ExplorersEvents_NewExplorerEventHandler(
                    Explorers_NewExplorer);

            foreach (Outlook.Inspector inspector in inspectors)
            {
                Inspectors_NewInspector(inspector);
            }

            foreach (Outlook.Explorer explorer in explorers)
            {
                Explorers_NewExplorer(explorer);
            }

            // Setup periodic check of recipients
            updateSuggestionFlag = new Timer();
            updateSuggestionFlag.AutoReset = true;
            updateSuggestionFlag.Elapsed += new System.Timers.ElapsedEventHandler(reviewEditors);
            updateSuggestionFlag.Interval = 5000;
            updateSuggestionFlag.Enabled = true;
            updateSuggestionFlag.Start();


            ((Outlook.ApplicationEvents_11_Event)Application).Quit
+= new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Shutdown);

        }


        private void reviewEditors(Object sender, EventArgs e)
        {
            // stop triggering while we are servicing
            updateSuggestionFlag.Enabled = false;


            try
            {
                foreach (var item in editorWrappersValue.Keys)
                {
                    editorWrapper wrapper = editorWrappersValue[item];
                    bool statusChange = false;

                    Outlook.MailItem emailMsg = null;
                    // Debug.Print(item.ToString);
                    if (item is Outlook.Explorer)
                    {
                        // check if composing
                        Outlook.Explorer exp = item as Outlook.Explorer;
                        emailMsg = exp.ActiveInlineResponse;
                    }
                    else
                    {
                        emailMsg = GetMailItem(item);
                    }
                    if (emailMsg == null)
                    {
                        continue;
                    }


                    // Check if heading outside of Raytheon
                    int numExternal = externalRecipients(emailMsg);

                    if (wrapper.addInVisible != (numExternal>0))
                    {
                        wrapper.addInVisible = !wrapper.addInVisible;
                        statusChange = true;
                    }


                    if (emailMsg.Attachments.Count > 0)
                    {
                        wrapper.addInPaneVisible = wrapper.addInVisible;
                    }
                    
                    if (emailMsg.Subject != null)
                    {

                        emailMsg.Subject = emailMsg.Subject.Replace("[RMSG]", "[RSMG]"); // Fix common typos
                        emailMsg.Subject = emailMsg.Subject.Replace("[PGPWM]", "[RSMG]"); // Let's replace the old keywords while we are at it.
                        
                        bool subjectSet = emailMsg.Subject.Contains("[RSMG]");

                        if (subjectSet)
                        {
                            if (!wrapper.addInVisible)
                            {
                                //wrapper.addInPaneVisible = true; // TODO - Add message in pane that setting secure with no external recipients does nothing
                                wrapper.CustomTaskPane.Visible = true;
                                wrapper.paneNoteNoEffect();
                            }

                            if (!wrapper.addInActive)
                            {
                                setSecure(emailMsg, true);
                                statusChange = true;
                                wrapper.addInActive = true;
                            }
                        }
                    }

                    if (statusChange)
                    {
                        secureSendRibbon.Ribbon.Invalidate();
                    }

                    //  Debug.Print("This message subject: " + emailMsg.Subject + ", have attachements: " + emailMsg.Attachments.Count + ", and sent is " + emailMsg.Sent);
                }
            }
            catch (InvalidOperationException)
            {
                // Do nothing, timer proc'd while window(s) were closing
                // Just being a good digital citizen by catching it here
            }
            // stop possible memory leak
            updateSuggestionFlag.Enabled = true;
        }

        private void ThisAddIn_Shutdown()
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785

            // Manual Application Quit Handler has been created in ThisAddIn_Startup to call this function instead.

            // Prevent from triggering during a shutdown, which would result in exceptions being thrown.
            updateSuggestionFlag.Stop();
            updateSuggestionFlag.Enabled = false;
            updateSuggestionFlag.Dispose();
        }

        private void ThisAddIn_Shutdown(Object sender, EventArgs e)
        {
            // Overload to satisfy Designer assumptions
        }


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new secureSendRibbon();
        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                editorWrappersValue.Add(Inspector, new editorWrapper(Inspector));

            }
        }

        void Explorers_NewExplorer(Outlook.Explorer Explorer)
        {
            // Don't need to check if it's a mail item; can only inline mail items
            editorWrappersValue.Add(Explorer, new editorWrapper(Explorer));
        }

        public Dictionary<Object, editorWrapper> editorWrappers
        {
            get
            {
                return editorWrappersValue;
            }
        }

        public static Outlook.MailItem GetMailItem(Office.IRibbonControl e)
        {
            return GetMailItem(e.Context);
        }

        public static Outlook.MailItem GetMailItem(Object editor)
        {
            // Check to see if a item is select in explorer or we are in inspector.
            if (editor is Outlook.Inspector)
            {
                Outlook.Inspector inspector = (Outlook.Inspector)editor;

                if (inspector.CurrentItem is Outlook.MailItem)
                {
                    return inspector.CurrentItem as Outlook.MailItem;
                }
            }

            if (editor is Outlook.Explorer)
            {
                Outlook.Explorer explorer = (Outlook.Explorer)editor;

                Outlook.Selection selectedItems = explorer.Selection;
                if (selectedItems.Count != 1)
                {
                    return null;
                }

                if (selectedItems[1] is Outlook.MailItem)
                {
                    return selectedItems[1] as Outlook.MailItem;
                }
            }

            return null;
        }

        // Start Add-In features and logic functions...

        internal static void setSecure(Outlook.MailItem mailItem, bool set)
        {

            if (mailItem == null)
            {
                Debug.Print("Error passing handle to container.");
                return;
            }

            string subject = mailItem.Subject;
            string body = "";

            // Tag to encrypt can be placed anywhere in the subject, but common practice is to place it at the beginning.
            if (subject != null && subject != "")
            {
                subject = subject.Replace("[RSMG]", "");
                subject = subject.Replace("[RMSG]", "");
                subject = subject.Replace("[PGPWM]", ""); // Let's replace the old keywords while we are at it.
                subject = subject.Trim();
            }

            mailItem.Subject = (set) ? "[RSMG] " + subject : subject;


            Outlook.Inspector inspector = mailItem.GetInspector;
            object missing = System.Reflection.Missing.Value;

            string bodyTag = "Sent via Raytheon Secure Messaging Gateway";


            if (inspector.IsWordMail() && inspector.EditorType == Outlook.OlEditorType.olEditorWord)
            {
                Word.Document bodyHandle = mailItem.GetInspector.WordEditor;
                // Word.Range msgBody = bodyHandle.Application.Selection.Range;
                Word.Range msgBody = bodyHandle.Range(bodyHandle.Content.Start, bodyHandle.Content.End);
                Word.Find trimLastTag = msgBody.Find;
                trimLastTag.Replacement.ClearFormatting();
                trimLastTag.Replacement.Text = "";

                trimLastTag.Execute(bodyTag, true, true, false, false, false, false, false, false, ref missing, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);

                msgBody.Text.Trim();
                //                msgBody.Find.Execute(ref "Sent via Raytheon Secure Messaging Gateway", true, true, false, false, false, false, true, false, "", Word.WdReplace.wdReplaceOne);
                if (set)
                {
                    msgBody.InsertAfter("\n" + bodyTag);
                }

            }
            else
            {
                body = mailItem.Body;
                body = body.Replace(bodyTag, "");
                body = body.Trim(); // Clean up our act
                mailItem.Body = (set) ? body + "\n" + bodyTag : body;
            }

        }

        private int externalRecipients(Outlook.MailItem mail)
        {
            int numExternal= 0;
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            Outlook.Recipients recips = mail.Recipients;
            foreach (Outlook.Recipient recip in recips)
            {
                Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                string smtpAddress =
                    pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                if (!smtpAddress.EndsWith("@raytheon.com"))
                {
                    numExternal++;
                }
            }
            return numExternal;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        // End ThisAddIn class
    }

    public class editorWrapper
    {
        private Object editor;
        private CustomTaskPane taskPane;

        // Keep track of relevant states of each editor
        private int numExternal= 0;
        private bool showSecureOptions = false;
        private bool suggestSecure = false;
        private bool showPane = false; // default to invisible

        public editorWrapper(Outlook.Inspector Inspector)
        {
            RegisterCallbacks(Inspector);
        }

        public editorWrapper(Outlook.Explorer Explorer)
        {
            RegisterCallbacks(Explorer);
        }

        private void RegisterCallbacks(Object Editor)
        {
            editor = Editor;

            //Register Callbacks
            if (Editor is Outlook.Inspector)
            {
                ((Outlook.InspectorEvents_Event)Editor).Close +=
                new Outlook.InspectorEvents_CloseEventHandler(editor_Close);
            }
            else
            {
                ((Outlook.ExplorerEvents_Event)Editor).Close +=
                new Outlook.ExplorerEvents_CloseEventHandler(editor_Close);
            }

            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                new secureSendPane(), "Secure Email", Editor);
            taskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
            taskPane.Visible = showPane;

        }

        void editor_Close()
        {
            if (taskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }

            taskPane = null;


            if (editor is Outlook.Inspector)
            {
                ((Outlook.InspectorEvents_Event)editor).Close -=
                new Outlook.InspectorEvents_CloseEventHandler(editor_Close);

                Globals.ThisAddIn.editorWrappers.Remove((Outlook.Inspector)editor);
            }
            else
            {
                ((Outlook.ExplorerEvents_Event)editor).Close -=
                new Outlook.ExplorerEvents_CloseEventHandler(editor_Close);
                Globals.ThisAddIn.editorWrappers.Remove((Outlook.Explorer)editor);
            }
            editor = null;
        }

        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (secureSendRibbon.Ribbon != null)
            {
                secureSendRibbon.Ribbon.InvalidateControl("toggleButton1");
            }
        }
        public CustomTaskPane CustomTaskPane
        {
            get
            {
                return taskPane;
            }
        }

        public void refreshPane()
        {
            (taskPane.Control as secureSendPane).setBox_addInActive(showSecureOptions);
        }

        public void paneNoteNoEffect()
        {
            (taskPane.Control as secureSendPane).noteNoEffect();
        }


        internal void refreshRibbon()
        {
            if (suggestSecure)
            {
                //blinkRibbon();
            }
            secureSendRibbon.Ribbon.Invalidate();
        }

        //private void blinkRibbon()
        //{
        //    int count = 3;

        //    Timer blinker = new Timer();
        //    blinker.AutoReset = true;
        //    blinker.Elapsed += new System.Timers.ElapsedEventHandler(blink);
        //    blinker.Interval = 50;
        //    blinker.Enabled = true;
        //    blinker.Start();

        //    void blink(Object sender, EventArgs e)
        //    {
        //        if (count%2 == 1)
        //        {

        //        }
                
        //    }

        //struct ribbonBlinker
        //{
        //    int count;
        //}

        //}

        public bool addInActive
        {
            get => showSecureOptions;
            set => showSecureOptions = value;
        }

        public bool addInVisible
        {
            get => suggestSecure;
            set => suggestSecure = value;
        }

        public bool addInPaneVisible
        {
            get => this.showPane;
            set => showPane = value;
        }

        public int externalRecipients
        {
            get => this.numExternal;
            set => numExternal = value;
        }
    }
}
