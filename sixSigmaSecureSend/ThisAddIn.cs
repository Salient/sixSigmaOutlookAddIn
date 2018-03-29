using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Diagnostics;


using Timer = System.Timers.Timer;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Collections;
using Shell32;


// APP FEATURE LIST
//
// DONE - Add Just Sent Regret feature
// Add classifier check for attachments outbound
// Add attachment size counter for RSMG
// DONE - Only run poll timer when active window
// DONE - Auto fix RMSG typo
// DONE - Warn if using RSMG with no external recipients
// DONE - Warn if sending attachments to external recipients

// This is how I do.
#pragma warning disable IDE1006 // Naming Styles

namespace sixSigmaSecureSend
{
    public partial class ThisAddIn
    {
        // We can't add event callbacks to Application.Inspectors because after the event fires, Application.Inspectors is garbage collected?? 
        // Instead we have to save a handle to it at the class level and then we can do what we want
        Outlook.Inspectors _inspectors;
        Outlook.Explorers _explorers;

        //internal SynchronizationContext mainThread;
        //private System.Windows.Forms.Form dummyForm = null;

        // Need a place to store state information for each editor.
        private Dictionary<int, editorWrapper> editorWrapperCollection = new Dictionary<int, editorWrapper>();

        private string secureTempPath = @"%localappdata%\Temporary Internet Files\Content.Outlook\SSSS";
        private DirectoryInfo secureFolder;

        // Required to create custom ribbon
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() { return new secureSendRibbon(); }

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

            //dummyForm = new System.Windows.Forms.Form();
            //dummyForm.Opacity = 0;
            //dummyForm.Show();
            //dummyForm.Visible = false;

            //mainThread = SynchronizationContext.Current;

            // Store handles to window collections, because otherwise C# stupidty abounds
            _inspectors = Application.Inspectors;
            _explorers = Application.Explorers;

            //while (Application.ActiveWindow() == null)
            //{
            //    await Task.Delay(1000);
            //}

            //editorWrapperCollection.Add(Application.ActiveWindow(), new editorWrapper(Application.ActiveWindow()));

            //// If somehow plugin is loading after windows are already open, find them all and bag 'n tag
            foreach (Outlook.Inspector inspector in _inspectors) { newInspector(inspector); }
            foreach (Outlook.Explorer explorer in _explorers) { newExplorer(explorer); }

            // Register new callbacks to catch new editors opening
            _inspectors.NewInspector += (s) => { newInspector(s); };
            _explorers.NewExplorer += (s) => { newExplorer(s); };

            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Shutdown);

            // Clean/setup secure temporary attachments folder.
            if (Directory.Exists(secureTempPath))
            {
                Directory.Delete(secureTempPath + "/*", true);
                secureFolder = new DirectoryInfo(secureTempPath);
            }
            else { secureFolder = Directory.CreateDirectory(secureTempPath); }
        }

        private void ThisAddIn_Shutdown()
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785

            // Manual Application Quit Handler has been created in ThisAddIn_Startup to call this function instead.
            // Prevent from polling open editors if exiting Outlook, otherwise might cause exceptions being thrown.
        }

        // Overload to satisfy Designer assumptions
        private void ThisAddIn_Shutdown(Object sender, EventArgs e) { }

        private void newInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem)
            {
                // Cast and create wrapper
                Outlook.MailItem mailItem = inspector.CurrentItem;
                editorWrapper newWrapper = new editorWrapper(mailItem);

                // Wrapper created, register Application level callbacks
                inspector.Application.ItemSend += (object item, ref bool cancel) =>
                {
                    inspector.CurrentItem.DeferredDeliveryTime = // Implement 30 second delay if enabled
                        (newWrapper.toggleDelay) ? (DateTime.Now).Add(new TimeSpan(0, 0, 30)) : new DateTime(4501, 1, 1);
                };

                // Register Window level callbacks
                ((Outlook.InspectorEvents_10_Event)inspector).Activate += newWrapper.runTimer;
                inspector.Deactivate += newWrapper.pauseTimer;
                ((Outlook.InspectorEvents_10_Event)inspector).Close += () => { editorWrapperCollection.Remove(inspector.GetHashCode()); };

                // Register mail item callbacks
                mailItem.Open += (ref bool s) => { newWrapper.addInVisible = newWrapper.getTaskPane.Visible = false; }; // Prevent ribbon options from blinking when changing drafts
                mailItem.AttachmentAdd += newWrapper.checkClassification;
                mailItem.Unload += newWrapper.deconstructWrapper;

                // That's a wrap
                editorWrapperCollection.Add(inspector.GetHashCode(), newWrapper);
            }
        }

        private void newExplorer(Outlook.Explorer explorer)
        {

            // Register inline response events
            explorer.InlineResponse += (s) =>
            {
                // Check to see what kind of inline it is. Not sure how many there are.
                if (!(s is Outlook.MailItem)) { Debug.Print("Vat da faaack"); return; }

                Outlook.MailItem mailItem = s as Outlook.MailItem;
                editorWrapper newWrapper = new editorWrapper(mailItem);

                // Wrapper created, register Application level callbacks
                explorer.Application.ItemSend += (object item, ref bool cancel) =>
                {
                    explorer.ActiveInlineResponse.DeferredDeliveryTime = // Implement 30 second delay if enabled
                        (newWrapper.toggleDelay) ? (DateTime.Now).Add(new TimeSpan(0, 0, 30)) : new DateTime(4501, 1, 1);
                };

                // Register Window level callbacks
                ((Outlook.ExplorerEvents_10_Event)explorer).Activate += newWrapper.runTimer;
                explorer.Deactivate += newWrapper.pauseTimer;
                // ((Outlook.ExplorerEvents_10_Event)explorer).Close += () => { editorWrapperCollection.Remove(explorer.GetHashCode()); };

                // Register mail item callbacks
                mailItem.Open += (ref bool b) => { newWrapper.addInVisible = newWrapper.getTaskPane.Visible = false; }; // Prevent ribbon options from blinking when changing drafts
                mailItem.AttachmentAdd += newWrapper.checkClassification;
                // mailItem.Unload += newWrapper.deconstructWrapper;

                // That's a wrap
                editorWrapperCollection.Add(explorer.GetHashCode(), newWrapper);
                // Don't wait the whole second for first poll
                newWrapper.reviewEditor(null);
            };

            explorer.InlineResponseClose += () =>
            {
                editorWrapper thisWrapper = editorWrapperCollection[Application.ActiveExplorer().GetHashCode()];

                ((Outlook.ExplorerEvents_10_Event)explorer).Activate -= thisWrapper.runTimer;
                explorer.Deactivate -= thisWrapper.pauseTimer;

                thisWrapper.deconstructWrapper();
                editorWrapperCollection.Remove(Application.ActiveExplorer().GetHashCode());
            };
        }

        // Some helper functions
        //internal static Outlook.MailItem GetMailItem(Object editor)
        //{
        //    if ((editor is Outlook.Inspector) && (editor as Outlook.Inspector).CurrentItem is Outlook.MailItem) { return (editor as Outlook.Inspector).CurrentItem; }
        //    if (editor is Outlook.Explorer) { return (editor as Outlook.Explorer).ActiveInlineResponse; }
        //    return null;
        //}

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

    internal class helperFunctions
    {
        internal class attachmentInfo
        {
            Dictionary<string, string> metaInfo;
            string securePath;
            classificationData classification;
            attachmentInfo(Dictionary<string, string> newinfo)
            {
                metaInfo = newinfo;
                if (newinfo.ContainsKey("Path")) { securePath = newinfo["Path"]; }
                // TODO 
                // scan all tags for classification data. assign blank or invalid object
            }

            internal class classificationData
            {
                Dictionary<string, string> controlCodes;
                bool validClassification;
                bool emptyClassification;

                classificationData()
                {
                    classificationData temp = new classificationData(true);
                    controlCodes = temp.controlCodes;
                    validClassification = temp.validClassification;
                    emptyClassification = temp.emptyClassification;
                }

                classificationData(bool empty)
                {
                    if (empty)
                    {
                        controlCodes = new Dictionary<string, string>();
                        validClassification = false;
                        emptyClassification = true;
                    }
                    else
                    {
                        controlCodes = new Dictionary<string, string> { // default to unrestricted/undetermined
                        {"rtnipcontrolcode", "unrestricted"},
                        {"rtnipcontrolcodevm", "noipvm"},
                        {"rtnexportcontrolcountry", "usa"},
                        {"rtnexportcontrolcode", "undetermined"},
                        {"rtnexportcontrolcodevm", "piogcgtc5004"}
                    };

                        validClassification = true;
                        emptyClassification = false;
                    }
                }


                classificationData(string data)
                {
                    classificationData temp = parseClassification(data);
                    
                    controlCodes = temp.controlCodes;
                    validClassification = temp.validClassification;
                    emptyClassification = temp.emptyClassification;
                }

                private static classificationData parseClassification(string classification) // String manipulation stuff
                { // Far as I can tell, the classification tool sets up to five keywords, always ordered in the same way, seperated by |

                    if (classification == "") { return new classificationData(); }

                    // Sanitize input
                    classification = classification.Trim();
                    classification = classification.Replace("[", "");
                    classification = classification.Replace("]", "");
                    string[] keywords = classification.Split('|');
                    if (keywords.Length != 5)
                    {
                        Debug.Print("Invalid classification string found?");
                        return new classificationData();
                    } // Replace invalid classification with default

                    classificationData tempObject = new classificationData(false); // Prepopulated, valid classification object
                    for (int x = 0; x < 5; x++)
                    {
                        string code = keywords[x];
                        if (code == "") { continue; }
                        string[] codeparts = code.Split(':');
                        if (codeparts.Length != 2) { continue; }
                        if (tempObject.controlCodes.ContainsKey(codeparts[0])) { tempObject.controlCodes[codeparts[0]] = codeparts[1]; }
                        else { Debug.Print("invalid control code key: " + codeparts[0] + " with value " + codeparts[1] + ", skipping"); }
                    }

                    tempObject.validate();
                    return tempObject;
                }

                private void validate() // Verify each field matches one of the possible values
                {
                    validClassification = false; // Default to invalid
                    emptyClassification = true;

                    if (!(controlCodes.Count == 5
                        && controlCodes.ContainsKey("rtnipcontrolcode")
                        && controlCodes.ContainsKey("rtnipcontrolcodevm")
                        && controlCodes.ContainsKey("rtnexportcontrolcountry")
                        && controlCodes.ContainsKey("rtnexportcontrolcode")
                        && controlCodes.ContainsKey("rtnexportcontrolcodevm"))) { return; } // Must contain exactly these keys, no more, no less

                    emptyClassification = false; 

                    // Validate proper values for each property

                    switch(controlCodes["rtnipcontrolcode"])
                    {
                        case "internaluseonly":
                        case "mostprivate":
                        case "competitionsensitive":
                        case "proprietary":
                        case "thirdpartyproprietary":
                        case "public":
                        case "unrestricted":
                            break;
                        default: return;
                    }

                    switch (controlCodes["rtnipcontrolcodevm"])
                    {
                        case "preexistingipvm":
                        case "rpogc035":
                        case "noipvm":
                            break;
                        default: return;
                    }
                    switch (controlCodes["rtnexportcontrolcountry"])
                    {
                        case "usa":
                            break;
                        default: return;
                    }
                    switch (controlCodes["rtnexportcontrolcode"])
                    {
                        case "otherinfo":
                        case "nonexporteximdetermined":
                        case "itar":
                        case "ear":
                        case "undetermined":
                        case "legacy":
                            break;
                        default: return;
                    }
                    switch (controlCodes["rtnexportcontrolcodevm"])
                    {
                        case "piogcgtc5004":
                        case "nousecvm":
                        case "dodi523024":
                        case "preexistingusecvm":
                            break;
                        default: return;
                    }

                    validClassification = true;
                }

            }
        }


        // Getting all the available Information of a File into a Arraylist
        internal static ArrayList GetDetailedFileInfo(string sFile)
        {
            if (sFile is null)
            {
                return new ArrayList();
            }
            ArrayList aReturn = new ArrayList();
            if (sFile.Length > 0)
            {
                try
                {
                    // Creating a ShellClass Object from the Shell32
                    Shell32.Shell sh = new Shell();
                    // Creating a Folder Object from Folder that inculdes the File
                    Folder dir = sh.NameSpace(Path.GetDirectoryName(sFile));
                    // Creating a new FolderItem from Folder that includes the File
                    FolderItem item = dir.ParseName(Path.GetFileName(sFile));

                    // Debug
                    Debug.Print("Dumping metadata of " + sFile);
                    for (int x = 0; x < int.MaxValue; x++)
                    {
                        string tag = dir.GetDetailsOf(null, x);
                        if (tag is "") { break; }
                        Debug.Print("(" + x + ") [" + tag + "] " + dir.GetDetailsOf(item, x));

                    }
                    // loop throw the Folder Items
                    //for (int i = 0; i < 30; i++)
                    //{
                    //    // read the current detail Info from the FolderItem Object
                    //    //(Retrieves details about an item in a folder. 
                    //    //For example, its size, type, or the time 
                    //    //of its last modification.)

                    //    // some examples:
                    //    // 0 Retrieves the name of the item. 
                    //    // 1 Retrieves the size of the item.
                    //    // 2 Retrieves the type of the item.
                    //    // 3 Retrieves the date and time that the item was last modified.
                    //    // 4 Retrieves the attributes of the item.
                    //    // -1 Retrieves the info tip information for the item. 

                    //    string det = dir.GetDetailsOf(item, i);
                    //    // Create a helper Object for holding the current Information
                    //    // an put it into a ArrayList
                    //    DetailedFileInfo oFileInfo = new DetailedFileInfo(i, det);
                    //    aReturn.Add(oFileInfo);
                    //}

                }
                catch (Exception)
                {

                }
            }
            return aReturn;
        }

        internal static attachmentInfo getAttachmentInfo(string sFile)
        {
            if (sFile is null) { throw new ArgumentNullException("File path"); }

            if (sFile.Length > 0)
            {
                Dictionary<string, string> fileMetaData = new Dictionary<string, string>();

                try
                {
                    // Creating a ShellClass Object from the Shell32
                    Shell32.Shell sh = new Shell();
                    // Creating a Folder Object from Folder that inculdes the File
                    Folder dir = sh.NameSpace(Path.GetDirectoryName(sFile));
                    // Creating a new FolderItem from Folder that includes the File
                    FolderItem item = dir.ParseName(Path.GetFileName(sFile));

                    for (int x = 0; x < int.MaxValue; x++)
                    {
                        if (dir.GetDetailsOf(null, x) is "") { break; }
                        else { fileMetaData.Add(dir.GetDetailsOf(null, x), dir.GetDetailsOf(item, x)); }
                    }
                }
                catch (Exception)
                {

                }

            }
        }


        // Helper Class from holding the detailed File Informations
        // of the System
        public class DetailedFileInfo
        {
            int iID = 0;
            string sValue = "";

            public int ID
            {
                get { return iID; }
                set { iID = value; }
            }
            public string Value
            {
                get { return sValue; }
                set { sValue = value; }
            }

            public DetailedFileInfo(int ID, string Value)
            {
                iID = ID;
                sValue = Value;
            }
        }

    }

    // Create object to associate and manage ribbon and task pane with email composer
    public class editorWrapper
    {
        // Use email composer as key 
        private Outlook.MailItem mailItem;

        // Custom task pane objects are instanced per email editor; ribbons are single global instance but affect each editor individually. Go figure.
        // Hold reference to task pane object for this instance.
        private CustomTaskPane taskPane;

        // Amazingly there is not a good way to execute a callback when the recipients field changes. Thus, we must periodically check if the field has changed ourselves. Ah, Microsoft...
        private Timer pollTimer;

        // Keep track of relevant states of each editor
        private int numExternal = 0; // Number of recipients that are outside of Raytheon
        private int numAttached = 0; // Number of attachments in email
        private int totalSizeAttached = 0;

        private bool msgSetSecure = false; // Bit signifying if this email is to be sent with [RSMG] in the subject line

        private bool secureOptionsVisible = false; // default ribbon buttons to invisible
        private bool showPane = false; // default to invisible
        private bool paneTrigd = false; // Some things we only want to do once after the window opens

        private bool delaySet = true; // default to delay on

        private Dictionary<string, bool> attachmentClassified;

        public editorWrapper(Outlook.MailItem mItem)
        {
            // Store instance of email being composed
            mailItem = mItem;

            // Create a poll timer for this instance
            pollTimer = new Timer(1000); // Check every second (only enabled when window has focus)
            pollTimer.AutoReset = true;
            pollTimer.Elapsed += reviewEditor;

            attachmentClassified = new Dictionary<string, bool>();

            //Globals.ThisAddIn.mainThread.Send((s) =>
            //{
            // Setup task pane
            taskPane = (Globals.ThisAddIn.CustomTaskPanes.Add(new secureSendPane(this), "Secure Email", mItem.Application.ActiveWindow()));
            taskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
            taskPane.Visible = showPane;
            //}, null);

            // Kick-off explicitly
            // pollTimer.Enabled = true;

        }


        // Clean up after ourselves when an editor closes
        internal void deconstructWrapper()
        {
            pollTimer.Enabled = false;
            pollTimer.Dispose();

            if (taskPane != null) { Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane); taskPane = null; }

            mailItem = null;
        }

        internal void checkClassification(object sender)
        {
            //if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            //{
            //    Thread staThread = new Thread(new ParameterizedThreadStart(checkClassification));
            //    staThread.SetApartmentState(ApartmentState.STA);
            //    staThread.Start(sender);
            //    staThread.Join();
            //    return;
            //}


            Outlook.Attachment attach = sender as Outlook.Attachment;

            string attachPath = attach.GetTemporaryFilePath();
            Debug.Print(attach.Type.ToString());

            List<string> arrHeaders = new List<string>();

            ArrayList results = helperFunctions.GetDetailedFileInfo(attachPath);

            //Shell shell = new Shell32.Shell();
            //Folder attachmentFolder = shell.NameSpace(Path.GetDirectoryName(attachPath));
            //FolderItem file = attachmentFolder.ParseName(Path.GetFileName(attachPath));

            //for (int i = 0; i < 30; i++)
            //{
            //    string det = attachmentFolder.GetDetailsOf(file, i);

            //}

        }
        //for (int i = 0; i < short.MaxValue; i++)
        //{

        //    string header = objFolder.GetDetailsOf(null, i);
        //    if (String.IsNullOrEmpty(header))
        //        break;
        //    arrHeaders.Add(header);

        //}

        //foreach (Shell32.FolderItem2 item in objFolder.Items())
        //{
        //    for (int i = 0; i < arrHeaders.Count; i++)
        //    {
        //        Console.WriteLine(
        //          $"{i}\t{arrHeaders[i]}: {objFolder.GetDetailsOf(item, i)}");
        //    }
        //}


        //Debug.Print("Attachment: " + attach.DisplayName + ", " + attach.Position + ", " + attach.Type);


        //Outlook.PropertyAccessor test = attach.PropertyAccessor;

        //dynamic results = test.GetProperties("http://schemas.microsoft.com/mapi/proptag");

        //Debug.Print(attach.PropertyAccessor.GetProperties("http://schemas.microsoft.com/mapi/proptag"));


        //var shellAppType = Type.GetTypeFromProgID("Shell.Application");
        //dynamic shellApp = Activator.CreateInstance(shellAppType);
        //var folder = shellApp.Namespace(attach.GetTemporaryFilePath());
        //foreach (var item in folder.Items())
        //{
        //    var company = item.ExtendedProperty("Company");
        //    var author = item.ExtendedProperty("Author");
        //    // Etc.
        //}
        //attach.PropertyAccessor.GetProperties.
        //folder.
        //var folder = new Shell().Namespace(attach.PathName);

        //attach.Session
        /*
        attach.Size;
        attach.DisplayName;
        attach.Position;
        attach.PropertyAccessor;
        attach.Type;




    I looked into the shellfile class a little more. The answer was staring me right in the face.
    string[] keywords = new string[x];
    var shellFile = ShellFile.FromFilePath(file);
    shellFile.Properties.System.Keywords.Value = keywords;


    to get the keywords already added to the file use:
    var tags = (string[])shellFile.Properties.System.Keywords.ValueAsObject;
    tags = tags ?? new string[0];

    if (tags.Length != 0)
    {
    foreach (string str in tags)
    {
    // code here
    }
    }


    and done!
    */


        private void countExternalRecipients(object emailMsg)
        {

        }

        //    private void reviewAttachments(object msgObj)
        //    {
        //        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
        //        {
        //            Thread staThread = new Thread(new ParameterizedThreadStart(reviewAttachments));
        //            staThread.SetApartmentState(ApartmentState.STA);
        //            staThread.Start(msgObj);
        //            staThread.Join();
        //            return;
        //        }

        //        if (attachmentsCount > 0 /*&& addInVisible */ ) // Only process what's going on with attachments if it can head outside of Raytheon
        //        {
        //            int currentTotalSize = 0;
        //            Outlook.MailItem emailMsg = msgObj as Outlook.MailItem;

        //            Outlook.Attachments fileSet = emailMsg.Attachments;

        //            foreach (Outlook.Attachment attach in emailMsg.Attachments)
        //            {

        //                Outlook.Attachment thisone = attach;

        //                currentTotalSize += attach.Size;
        //                string filename = attach.FileName;
        //                string attachPath = attach.PathName;
        //                Debug.Print(attach.Type.ToString());

        //                if (filename == "" || attachPath is null) { continue; }

        //                List<string> arrHeaders = new List<string>();

        //                ArrayList results = helperFunctions.GetDetailedFileInfo(attachPath);

        //                Shell shell = new Shell32.Shell();
        //                Folder attachmentFolder = shell.NameSpace(Path.GetDirectoryName(attachPath));
        //                FolderItem file = attachmentFolder.ParseName(Path.GetFileName(attachPath));

        //                for (int i = 0; i < 30; i++)
        //                {
        //                    string det = attachmentFolder.GetDetailsOf(file, i);

        //                }


        //                //for (int i = 0; i < short.MaxValue; i++)
        //                //{

        //                //    string header = objFolder.GetDetailsOf(null, i);
        //                //    if (String.IsNullOrEmpty(header))
        //                //        break;
        //                //    arrHeaders.Add(header);

        //                //}

        //                //foreach (Shell32.FolderItem2 item in objFolder.Items())
        //                //{
        //                //    for (int i = 0; i < arrHeaders.Count; i++)
        //                //    {
        //                //        Console.WriteLine(
        //                //          $"{i}\t{arrHeaders[i]}: {objFolder.GetDetailsOf(item, i)}");
        //                //    }
        //                //}


        //                //Debug.Print("Attachment: " + attach.DisplayName + ", " + attach.Position + ", " + attach.Type);


        //                //Outlook.PropertyAccessor test = attach.PropertyAccessor;

        //                //dynamic results = test.GetProperties("http://schemas.microsoft.com/mapi/proptag");

        //                //Debug.Print(attach.PropertyAccessor.GetProperties("http://schemas.microsoft.com/mapi/proptag"));


        //                //var shellAppType = Type.GetTypeFromProgID("Shell.Application");
        //                //dynamic shellApp = Activator.CreateInstance(shellAppType);
        //                //var folder = shellApp.Namespace(attach.GetTemporaryFilePath());
        //                //foreach (var item in folder.Items())
        //                //{
        //                //    var company = item.ExtendedProperty("Company");
        //                //    var author = item.ExtendedProperty("Author");
        //                //    // Etc.
        //                //}
        //                //attach.PropertyAccessor.GetProperties.
        //                //folder.
        //                //var folder = new Shell().Namespace(attach.PathName);

        //                //attach.Session
        //                /*
        //                attach.Size;
        //                attach.DisplayName;
        //                attach.Position;
        //                attach.PropertyAccessor;
        //                attach.Type;


        private void reviewEditor(object unused, EventArgs ev) { reviewEditor(unused); } // Need overload for timer event handler
        internal void reviewEditor(object unused)
        {
            if (Globals.ThisAddIn.Application.ActiveWindow() == null) { return; } // Bail if Outlook is still initializing and there is no active window yet

            // Stop from proc'ing over itself. Turned back on at end of poll.
            pollTimer.Enabled = false;

            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                Thread staThread = new Thread(new ParameterizedThreadStart(reviewEditor));
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start(unused);
                staThread.Join();
                return;
            }


            //Debug.Print("timer proce from object: " + GetHashCode());
            //Debug.Print("Active window object is " + (Globals.ThisAddIn.Application.ActiveWindow() as object).GetHashCode());

            //try
            //{
            bool statusChange = false;

            if (mailItem is null) { /*Debug.Print("reviewing no mail item...?");*/ return; } // Not editing a new email

            // Check if heading outside of Raytheon
            int count = 0;
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            Outlook.Recipients recips = null;
            //Globals.ThisAddIn.mainThread.Send((s) =>
            //{
            recips = mailItem.Recipients;

            //}, null);

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
                            count++;
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

            if (count != numExternal) { statusChange = true; numExternal = count; }

            // If we are heading outside of Raytheon, are we showing some security options?
            if (addInVisible != (numExternal > 0))
            {
                addInVisible = !addInVisible;
                statusChange = true;
            }

            // Are we adding any attachments?
            if (mailItem.Attachments.Count != attachmentsCount)
            {
                attachmentsCount = mailItem.Attachments.Count;
                statusChange = true;

                int currentTotalSize = 0;
                foreach (Outlook.Attachment attach in mailItem.Attachments) { currentTotalSize += attach.Size; }

                totalSizeAttached = (int)Math.Round(currentTotalSize / 1024.0);
            }

            if (mailItem.Subject != null)
            {
                mailItem.Subject = mailItem.Subject.Replace("[RMSG]", "[RSMG]"); // Fix common typos
                mailItem.Subject = mailItem.Subject.Replace("[PGPWM]", "[RSMG]"); // Let's replace the old keywords while we are at it.

                bool subjectSet = mailItem.Subject.Contains("[RSMG]");

                if (subjectSet)
                {
                    if (!addInVisible && !paneShownBefore) { statusChange = true; }

                    if (!addInActive)
                    {
                        setSecure(true);
                        statusChange = true;
                        addInActive = true;
                    }
                }
            }

            if (statusChange) { secureSendRibbon.Ribbon?.Invalidate(); refreshPane(); }

            //  Debug.Print("This message subject: " + emailMsg.Subject + ", have attachements: " + emailMsg.Attachments.Count + ", and sent is " + emailMsg.Sent);
            //}

            //catch (InvalidOperationException)
            //{
            //    // Do nothing, timer proc'd while window(s) were closing
            //    // Just being a good digital citizen by catching it here
            //}
            pollTimer.Enabled = true;
        }


        // Start Add-In features and logic functions...

        internal void setSecure(bool set)
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


            Outlook.Inspector inspector;
            try
            {
                inspector = mailItem.GetInspector;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                // Window not fully initialized yet, bail for now.
                return;
            }
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

        internal void refreshPane()
        {
            // Check state of editor and issue appropriate changes to task pane.
            getSecureSendPane.setBox_addInActive(addInActive);
            taskPane.Visible = showPane;
            getSecureSendPane.updateAttachInfo(numAttached, totalSizeAttached);

            //if (!paneTrigd) { paneTrigd = true;
            //} // Don't want to be super annoying

            if (numAttached > 0) { } else { }

            if (numExternal == 0 && msgSetSecure) { taskPane.Visible = true; getSecureSendPane.noteNoEffect(); }
            // if (numExternal > 0 && !msgSetSecure) { (taskPane.Control as secureSendPane).suggest(); }
            if (numExternal > 0 && numAttached > 0 && !msgSetSecure) { taskPane.Visible = true; getSecureSendPane.suggest(); }
            //            private int numExternal = 0;
            //private int numAttached = 0;
            //private bool msgSetSecure = false;
            //private bool secureOptionsVisible = false; // default to invisible
            //private bool showPane = false; // default to invisible



        }

        internal void updateState(bool set)
        {
            addInActive = set;
            setSecure(set);
            refreshPane();
            getSecureSendPane.setBox_addInActive(addInActive);
            secureSendRibbon.Ribbon?.InvalidateControl("toggleAddInActive");
            secureSendRibbon.Ribbon?.InvalidateControl("toggleAddInActive_inline");
        }

        internal secureSendPane getSecureSendPane { get => taskPane.Control as secureSendPane; }
        void TaskPane_VisibleChanged(object sender, EventArgs e) { showPane = taskPane.Visible; secureSendRibbon.Ribbon?.InvalidateControl("toggleButton1"); }

        internal static editorWrapper getWrapper(Office.IRibbonControl control) { foreach (editorWrapper item in Globals.ThisAddIn.editorWrapperCollection.Values) { if (item.mailItem.Application.ActiveWindow() == control.Context) return item; } return null; }

        internal bool toggleDelay { get => delaySet; set => delaySet = value; }
        public bool addInActive { get => msgSetSecure; set => msgSetSecure = value; }
        public bool addInVisible { get => secureOptionsVisible; set => secureOptionsVisible = value; }
        public bool addInPaneVisible { get => showPane; set => showPane = value; }
        public bool paneShownBefore { get => paneTrigd; set => paneTrigd = value; }
        public int externalRecipients { get => numExternal; set => numExternal = value; }

        public int attachmentsCount { get => numAttached; set => numAttached = value; }

        public CustomTaskPane getTaskPane { get => taskPane; }

        internal void runTimer() { pollTimer.Enabled = true; }
        internal void pauseTimer() { pollTimer.Enabled = false; }
    }
}
#pragma warning restore IDE1006 // Naming Styles