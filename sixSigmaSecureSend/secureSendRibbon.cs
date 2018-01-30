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

        internal static Office.IRibbonUI Ribbon { get => ribbon; set => ribbon = value; }

        public secureSendRibbon() { }

        #region IRibbonExtensibility Members
        
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("sixSigmaSecureSend.secureSendRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public bool sendSecureVisible(Office.IRibbonControl control)
        {
            return (getEditorFromControl(control) as editorWrapper).addInVisible;
        }
        public void toggleShowPane(Office.IRibbonControl control, bool state)
        {
            
            CustomTaskPane secureSendPane = getPaneFromControl(control);
            if (secureSendPane != null)
            {
                secureSendPane.Visible = state;
            }
        }


        public IPictureDisp returnRTNSecureLogo(Office.IRibbonControl control)
        {
            return ImageConverter.Convert(Properties.Resources.rtnsecuretrans);
        }

        public IPictureDisp returnRTNLock(Office.IRibbonControl control)
        {
            return ImageConverter.Convert(Properties.Resources.rtnlock);
        }
            

        public string numberExternal(Office.IRibbonControl control)
        {
            int numExternal = getEditorFromControl(control).externalRecipients;

            if (numExternal == 0)
            {
                return "There are no external recipients.";
            }
            if (numExternal == 1)
            {
                return "There is 1 external recipient.";
            }
            return "There are " + numExternal + " external recipients.";
        }

        public string addInStatus(Office.IRibbonControl control)
        {
            return  (getEditorFromControl(control) as editorWrapper).addInActive ? "AcceptTask" : "Private";
        }

        public bool isPressed(Office.IRibbonControl control)
        {
            CustomTaskPane secureSendPane = getPaneFromControl(control);
            if (secureSendPane != null)
            {
               return secureSendPane.Visible;
            }
            else { return false; }
        }

        public bool addInActive(Office.IRibbonControl control)
        {
            return getEditorFromControl(control).addInActive;
        }

        private editorWrapper getEditorFromControl(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.editorWrappers[control.Context];
        }

        private CustomTaskPane getPaneFromControl(Office.IRibbonControl control)
        {

            if (control.Context is Outlook.Inspector || control.Context is Outlook.Explorer)
            {
                editorWrapper myEditor = Globals.ThisAddIn.editorWrappers[control.Context];
                return myEditor.CustomTaskPane;
            } else
            {
                Debug.Print("oh shnikee");
                return null;
            }
        }

        public void linkToRTNSecure(Office.IRibbonControl control)
        {
            System.Diagnostics.Process.Start("http://web.onertn.ray.com/initiatives/rtnsecurecenter/");
        }
        
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

        
        public void toggleAddInActive(Office.IRibbonControl control, bool set)
        {
            editorWrapper thisEditor = getEditorFromControl(control);
            thisEditor.addInActive = set;
            thisEditor.refreshPane();
            ribbon.InvalidateControl("toggleAddInActive");
            ThisAddIn.setSecure(ThisAddIn.GetMailItem(control), set);
        }

        
        internal class ImageConverter : System.Windows.Forms.AxHost
        {
            private ImageConverter() : base(null)
            {
            }

            public static stdole.IPictureDisp Convert(System.Drawing.Image image)
            {
                try
                {
                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
                } catch(System.ArgumentException)
                {
                    Debug.Print("doing that thin agian...");

                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
                }

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
