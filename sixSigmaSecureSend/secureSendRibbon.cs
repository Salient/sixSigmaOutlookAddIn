using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Microsoft.Office.Tools;
using stdole;
using System.Windows.Forms;
using System.Drawing;


namespace sixSigmaSecureSend
{
    [ComVisible(true)]
    public class secureSendRibbon : Office.IRibbonExtensibility
    {
        private static Office.IRibbonUI ribbon;

        internal static Office.IRibbonUI Ribbon { get => ribbon; }

        public secureSendRibbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) { return GetResourceText("sixSigmaSecureSend.secureSendRibbon.xml"); }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public bool sendSecureVisible(Office.IRibbonControl control) { return editorWrapper.getWrapper(control)?.addInVisible ?? false; }
        public void toggleShowPane(Office.IRibbonControl control, bool state) { editorWrapper.getWrapper(control).getTaskPane.Visible = state; }
        public IPictureDisp returnRTNSecureLogo(Office.IRibbonControl control) { return ImageConverter.Convert(Properties.Resources.rtnsecuretrans); }
        public IPictureDisp returnRTNLock(Office.IRibbonControl control) { return ImageConverter.Convert(Properties.Resources.rtnlock); }
        //        public string addInStatus(Office.IRibbonControl control) { return (editorWrapper.getWrapper(control)?.addInActive == null || editorWrapper.getWrapper(control).addInActive) ? "AcceptTask" : "Private"; }
        public string addInStatus(Office.IRibbonControl control) { return (editorWrapper.getWrapper(control)?.addInActive ?? false) ? "AcceptTask" : "Private"; }
        public bool isPressed(Office.IRibbonControl control) { return editorWrapper.getWrapper(control)?.getTaskPane.Visible ?? false; }
        public bool addInActive(Office.IRibbonControl control) { return editorWrapper.getWrapper(control)?.addInActive ?? false; }
        public void linkToRTNSecure(Office.IRibbonControl control) { System.Diagnostics.Process.Start("http://web.onertn.ray.com/initiatives/rtnsecurecenter/"); }

        public string numberExternal(Office.IRibbonControl control)
        {
            int? numExternal = editorWrapper.getWrapper(control)?.externalRecipients;

            if (numExternal == 0) { return "There are no external recipients."; }
            if (numExternal == 1) { return "There is 1 external recipient."; }
            return "There are " + numExternal + " external recipients.";
        }

        public void toggleAddInActive(Office.IRibbonControl control, bool set)
        {
            editorWrapper instance = editorWrapper.getWrapper(control);
            instance.addInActive = set;
            instance.getSecureSendPane.setBox_addInActive(set);
            ribbon.InvalidateControl("toggleAddInActive");
            ThisAddIn.setSecure(ThisAddIn.GetMailItem(control), set);
        }
        #endregion

        #region Graphics Helper Functions
        internal class PictureConverter : AxHost
        {
            private PictureConverter() : base(String.Empty) { }

            static public stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }

            static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
            {
                return ImageToPictureDisp(icon.ToBitmap());
            }

            static public Image PictureDispToImage(stdole.IPictureDisp picture)
            {
                return GetPictureFromIPicture(picture);
            }
        }

        internal class ImageConverter : System.Windows.Forms.AxHost
        {
            private ImageConverter() : base(null)
            {
            }

            public static stdole.IPictureDisp Convert(System.Drawing.Image image)
            {
                stdole.IPictureDisp temp = null;
                try
                {
                    temp = (stdole.IPictureDisp)GetIPictureDispFromPicture(image);

                }
                catch (Exception ex)
                {
                    if (ex is System.ArgumentException || ex is System.Runtime.InteropServices.ExternalException)
                    {
                        Debug.Print("doing that thin agian...");
                        throw;
                    }
                }
                return temp;
            }
        }

        // Private and possibly AdpDiagramKeys or DatabaseSetLogonSecurity or ProtectDocument
        // AcceptTask and RelationshipsClearLayout
        // GroupCommunicate

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
