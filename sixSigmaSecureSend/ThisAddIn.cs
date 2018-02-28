using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Diagnostics;
using System.Threading;
using Timer = System.Timers.Timer;


// APP FEATURE LIST
//
// Add Just Sent Regret feature?
// Add classifier check for attachments outbound
// Add attachment size counter for RSMG


namespace sixSigmaSecureSend
{
    public partial class ThisAddIn
    {
        // Amazingly there is not a good way to execute a callback when the recipients field changes. Thus, we must periodically check if the field has changed ourselves. Ah, Microsoft...
        private Timer pollTimer;
        internal SynchronizationContext mainThread;
        private Dictionary<Object, editorWrapper> editorWrappersValue = new Dictionary<Object, editorWrapper>();
        private System.Windows.Forms.Form dummyForm = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // In order periodically poll all open editors, we need a timer.
            // In order to have a timer, it needs to spawn its own thread.
            // In order to be useful, the new thread needs to create new UI objects
            // In order to create new UI objects, you must be in the main UI thread
            // In order to do things in the main UI thread from the timer thread, we need a synchonization object.
            // In order to have a synchronization object, you have to create a Form. 

            // Thus, we create this form we don't want and do our best to hide it, so we can do something completely unrelated. 
            // That's how microsoft set up how it works, and somehow thought it was good.

            dummyForm = new System.Windows.Forms.Form();
            dummyForm.Opacity = 0;
            dummyForm.Show();
            dummyForm.Visible = false;

            mainThread = SynchronizationContext.Current;
            

            // Initialize timer object before possibly adding editor instances
            pollTimer = new Timer(2000);
            pollTimer.AutoReset = true;
            //pollTimer.Elapsed += new System.Timers.ElapsedEventHandler(reviewEditors);
            pollTimer.Elapsed += reviewEditors;


            // Plan _was_ to do this with events and callbacks, but the whole VSTO/C# framework is too terrible to really support doing that
            // And since we have a poll timer going anyway, might as well just check everything all the time :-\

            // May the god of efficiency grant me mercy, it is not my fault...


            // If somehow plugin is loading after windows are already open, find them all and bag 'n tag
            //foreach (Outlook.Inspector inspector in this.Application.Inspectors) { editorWrappersValue.Add(inspector, new editorWrapper(inspector)); }
            //foreach (Outlook.Explorer explorer in this.Application.Explorers) { editorWrappersValue.Add(explorer, new editorWrapper(explorer)); }

            //// Register new callbacks to catch new editors opening
            ////this.Application.Inspectors.NewInspector += (s) => { editorWrappersValue.Add(s, new editorWrapper(s)); };
            ////this.Application.Explorers.NewExplorer += (s) => { editorWrappersValue.Add(s, new editorWrapper(s)); };

            //this.Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(test);
            //this.Application.Explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(test);

            //Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(test);
            // Application.Explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(test);


            pollTimer.Enabled = true;

            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Shutdown);
        }

        // Required to create custom ribbon
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() { return new secureSendRibbon(); }
        
        //private void addWrapper(object editor, EventArgs e) {
        //    editorWrappersValue.Add(editor, new editorWrapper(editor));
        //}

        //private void startStopPoll() {
        //    bool run = false;

        //    foreach (Object editor in editorWrappersValue.Keys) {
        //        Outlook.MailItem instance = GetMailItem(editor);
        //        // Test reading or editing message
        //        run = (instance != null && !instance.Sent);

        //        if (run) { break; } // If there is at least one open editor, don't bother checking the rest
        //    }

        //    pollTimer.Enabled = run;
        //}

        internal static Outlook.Application getApp
        {
            get => Globals.ThisAddIn.Application;
        }
        private void reviewEditors(Object sender, EventArgs e)
        {
            //Debug.Print("timer proc")
            // stop triggering while we are servicing, just in case 
            pollTimer.Enabled = false;

            // Maybe we only have to check the top window?
            object topWindow = Application.ActiveWindow();
            //if (item != topWindow)
            //{
            //    continue;
            //}

            //try
            //{

                foreach (Outlook.Inspector inspector in this.Application.Inspectors)
                {
                    if (!editorWrappersValue.ContainsKey(inspector))
                    {
                        editorWrappersValue.Add(inspector, new editorWrapper(inspector));
                    }

                    doEditor(inspector);
                }

                foreach (Outlook.Explorer explorer in this.Application.Explorers)
                {
                    if (!editorWrappersValue.ContainsKey(explorer))
                    {
                        editorWrappersValue.Add(explorer, new editorWrapper(explorer));
                    }

                    doEditor(explorer);
                }
            //}
            //catch (Exception ex)
            //{
            //    //frak if i know
            //}

            pollTimer.Enabled = true;
        }

        private void doEditor(object item)
        {
            try
            {
                bool statusChange = false;

                Outlook.MailItem emailMsg = null;

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
                    return;
                }

                editorWrapper wrapper = editorWrappersValue[item];

                // Check if heading outside of Raytheon
                int numExternal = externalRecipients(emailMsg);
                if (wrapper.externalRecipients != numExternal)
                {
                    wrapper.externalRecipients = numExternal;
                    statusChange = true;
                }

                if (wrapper.addInVisible != (numExternal > 0))
                {
                    wrapper.addInVisible = !wrapper.addInVisible;
                    statusChange = true;
                }


                if (emailMsg.Attachments.Count != wrapper.attachmentsCount)
                {
                    wrapper.attachmentsCount = emailMsg.Attachments.Count;
                    statusChange = true;
                }

                if (emailMsg.Subject != null)
                {

                    emailMsg.Subject = emailMsg.Subject.Replace("[RMSG]", "[RSMG]"); // Fix common typos
                    emailMsg.Subject = emailMsg.Subject.Replace("[PGPWM]", "[RSMG]"); // Let's replace the old keywords while we are at it.

                    bool subjectSet = emailMsg.Subject.Contains("[RSMG]");

                    if (subjectSet)
                    {
                        if (!wrapper.addInVisible && !wrapper.paneShownBefore) { statusChange = true; }

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
                    Debug.Print("Invalidating editor " + item.GetHashCode());
                    secureSendRibbon.Ribbon?.Invalidate();
                    wrapper.refreshPane();
                }
                //  Debug.Print("This message subject: " + emailMsg.Subject + ", have attachements: " + emailMsg.Attachments.Count + ", and sent is " + emailMsg.Sent);
            }

            catch (InvalidOperationException)
            {
                // Do nothing, timer proc'd while window(s) were closing
                // Just being a good digital citizen by catching it here
            }
        }



        private void ThisAddIn_Shutdown()
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785

            // Manual Application Quit Handler has been created in ThisAddIn_Startup to call this function instead.
            // Prevent from polling open editors if exiting Outlook, otherwise might cause exceptions being thrown.
            pollTimer.Enabled = false;
            pollTimer.Dispose();
        }

        // Overload to satisfy Designer assumptions
        private void ThisAddIn_Shutdown(Object sender, EventArgs e) { }


        
        public Dictionary<Object, editorWrapper> editorWrappers { get => editorWrappersValue; }

        public static Outlook.MailItem GetMailItem(Office.IRibbonControl e) { return GetMailItem(e.Context); }

        public static Outlook.MailItem GetMailItem(Object editor)
        {
            if (editor is Outlook.Inspector)
            {
                Outlook.Inspector inspector = (Outlook.Inspector)editor;

                if (inspector.CurrentItem is Outlook.MailItem) { return inspector.CurrentItem as Outlook.MailItem; }
            }

            if (editor is Outlook.Explorer)
            {
                Outlook.Explorer explorer = (Outlook.Explorer)editor;

                Outlook.Selection selectedItems = explorer.Selection;
                if (selectedItems.Count != 1) { return null; }

                if (selectedItems[1] is Outlook.MailItem) { return selectedItems[1] as Outlook.MailItem; }
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

                msgBody.Text = msgBody.Text.Trim();
                //                msgBody.Find.Execute(ref "Sent via Raytheon Secure Messaging Gateway", true, true, false, false, false, false, true, false, "", Word.WdReplace.wdReplaceOne);
                if (set)
                {
                    msgBody.InsertAfter("\n\n" + bodyTag);
                }

            }
            else
            {
                body = mailItem.Body;
                body = body.Replace(bodyTag, "");
                body = body.Trim(); // Clean up our act
                mailItem.Body = (set) ? body + "\n\n" + bodyTag : body;
            }

        }

        private int externalRecipients(Outlook.MailItem mail)
        {
            int numExternal = 0;
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            Outlook.Recipients recips = mail.Recipients;
            foreach (Outlook.Recipient recip in recips)
            {
                if (recip.Address != null) // Check for invalid email addresses
                {
                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;

                    try
                    {
                        string smtpAddress =
                            pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                        if (!smtpAddress.EndsWith("@raytheon.com"))
                        {
                            numExternal++;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        Debug.Print("oh no....it's happenning again....");
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
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


    // Create object to associate and manage ribbon and task pane with email composer
    public class editorWrapper
    {
        // Use email composer as key 
        private Object editor;
        // Custom task pane objects are instanced per email editor; ribbons are single global instance but affect each editor individually. Go figure.
        // Hold reference to task pane object for this instance.
        private CustomTaskPane taskPane;

        // Keep track of relevant states of each editor
        private int numExternal = 0;
        private int numAttached = 0;
        private bool msgSetSecure = false;
        private bool secureOptionsVisible = false; // default to invisible
        private bool showPane = false; // default to invisible
        private bool paneTrigd = false; // Some things we only want to do once after the window opens

        public editorWrapper(Object Editor)
        {
            // Save associated editor object, right now used for cleaning up callbacks
            editor = Editor;

            //Register Callbacks
            if (Editor is Outlook.Inspector && (Editor as Outlook.Inspector).CurrentItem is Outlook.MailItem)
            { ((Outlook.InspectorEvents_Event)Editor).Close += deconstructWrapper; }
            else if (Editor is Outlook.Explorer) { ((Outlook.ExplorerEvents_Event)Editor).Close += deconstructWrapper; }
            else { throw new ArgumentException("Not correct type of editor"); }

            // Setup task pane


            // Since we are probably creating this object from the timer object callback, 
            // and the timer doesn't run from the main UI thread, and thus can't actually call the CustomTaskPanes.Add method,
            // we need to pass it off to the Dispatcher


            //        "screw you that's why" 
            //                    -- Microsoft, probably

            // SynchronizationContext uiContext = SynchronizationContext.Current;
            // this.taskPane = (Globals.ThisAddIn.CustomTaskPanes.Add(new secureSendPane(), "Secure Email", Editor));

            Globals.ThisAddIn.mainThread.Send((s) =>
            {
                this.taskPane = (Globals.ThisAddIn.CustomTaskPanes.Add(new secureSendPane(), "Secure Email", Editor));
                taskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
                taskPane.Visible = showPane;
            }, null);
        }


        // Clean up after ourselves when an editor closes
        void deconstructWrapper()
        {
            if (taskPane != null) { Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane); taskPane = null; }
            if (Globals.ThisAddIn.editorWrappers.ContainsKey(editor)) { Globals.ThisAddIn.editorWrappers.Remove(editor); }

            if (editor is Outlook.Inspector) { ((Outlook.InspectorEvents_Event)editor).Close -= deconstructWrapper; }
            else if (editor is Outlook.Explorer) { ((Outlook.ExplorerEvents_Event)editor).Close -= deconstructWrapper; }

            editor = null;
        }

        internal void refreshPane()
        {
            // Check state of editor and issue appropriate changes to task pane.

            if (!paneTrigd)
            { // Don't want to be super annoying
                if (numExternal == 0 && msgSetSecure) { paneTrigd = true; taskPane.Visible = true; getSecureSendPane.noteNoEffect(); }
                // if (numExternal > 0 && !msgSetSecure) { (taskPane.Control as secureSendPane).suggest(); }
                if (numExternal > 0 && numAttached > 0 && !msgSetSecure) { paneTrigd = true; taskPane.Visible = true; getSecureSendPane.suggest(); }
                //            private int numExternal = 0;
                //private int numAttached = 0;
                //private bool msgSetSecure = false;
                //private bool secureOptionsVisible = false; // default to invisible
                //private bool showPane = false; // default to invisible
            }

        }

        internal secureSendPane getSecureSendPane { get => taskPane.Control as secureSendPane; }
        void TaskPane_VisibleChanged(object sender, EventArgs e) { showPane = taskPane.Visible; secureSendRibbon.Ribbon?.InvalidateControl("toggleButton1"); }

        internal static editorWrapper getWrapper(Office.IRibbonControl control)
        {
            Debug.Print("Getting wrapper for editor " + control.GetHashCode());
            if (Globals.ThisAddIn.editorWrappers.ContainsKey(control.Context))
            {
                return Globals.ThisAddIn.editorWrappers[control.Context];
            }
            return null;
        }

        

        public bool addInActive { get => msgSetSecure; set => msgSetSecure = value; }
        public bool addInVisible { get => secureOptionsVisible; set => secureOptionsVisible = value; }
        public bool addInPaneVisible { get => this.showPane; set => showPane = value; }
        public bool paneShownBefore { get => this.paneTrigd; set => paneTrigd = value; }
        public int externalRecipients { get => this.numExternal; set => numExternal = value; }
        public int attachmentsCount { get => this.numAttached; set => numAttached = value; }
        public CustomTaskPane getTaskPane { get => taskPane; }
    }
}
