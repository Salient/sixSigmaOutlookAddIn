using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections.Generic;


namespace sixSigmaSecureSend
{
    public partial class secureSendPane : UserControl
    {
        editorWrapper myWrapper;

        private Dictionary<string, string> paneText = new Dictionary<string, string>
        {
            {   "suggest",
                "It looks like you are sending an email with an attachment to recipients outside of Raytheon.\n\nRemember to safeguard sensitive or technical information by using PKI encryption or the Raytheon Secure Messaging Gateway."
            }, {
                "standard",
            "Always remember, safeguarding the integrity of sensitive information is a personal responsibility critical to protecting national interests, corporate reputation, and the warfighter. If this message contains technical detail or potentially sensitive information, consider adding [RSMG] to the subject line."
            }, {
                "no effect",
            "You have added [RSMG] to the subject line, indicating you want to safeguard the information in this message when sending to recipients outside of Raytheon. This is an excellent habit, however, I couldn't help noticing there aren't any external recipients. \n\nRaytheon Secure Messaging has no effect with sending to other Raytheon recipients because it never leaves the Raytheon network, and the Raytheon network is always secure :-)"
            },
            {
                "kudos",
                "This message shall be sent to external recipients via the Raytheon Secure Message Gateway. Recipients will receive an email with notification and instructions on how to retrieve this email securely. This is the safest way to transmit any information which may be sensitive."
            }
        };


        public secureSendPane(editorWrapper from)
        {
            myWrapper = from;
            InitializeComponent();
        }

        delegate void StringArgReturningVoidDelegate();

        // Helper functions
        public void setBox_addInActive(bool set)
        {
            // I will never understand all the nonsense created to deal with nonsense problems
            if (this.InvokeRequired)
            {
                this.Invoke(
                    new MethodInvoker(delegate () { setBox_addInActive(set); }));
            }
            else
            {
                checkBox_addInStatus.Checked = set;
                label1.Visible = true;
            }
        }

        // Callbacks

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            Debug.Print("checkbox toggled: " + e);
        }

        public void rtnsecurelogo_Click(object sender, EventArgs e)
        {
            Process.Start("http://web.onertn.ray.com/initiatives/rtnsecurecenter/");
        }
        public void sixsigmalogo_Click(object sender, EventArgs e)
        {
            Process.Start("http://web.onertn.ray.com/functions/etma/ma/r6s/");
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Debug.Print("something creative maybe?");

        }

        private void label1_Click_1(object sender, EventArgs e)
        {
            Debug.Print("something cheeky perhaps?");
        }

        public void suggest() { updateText(paneText["suggest"]); }

        public void stdtext() { updateText(paneText["standard"]); }

        public void noteNoEffect() { updateText(paneText["no effect"]); }

        public void kudos() { updateText(paneText["kudos"]); }

        private void updateText(string text)
        {
            // I will never understand all the nonsense created to deal with nonsense problems
            if (this.InvokeRequired)
            {
                this.Invoke(
                    new MethodInvoker(delegate () { updateText(text); }));
            }
            else
            {
                this.label1.Text = text;
                this.label1.Visible = true;
                this.label1.Show();
            }
        }

        internal void updateAttachInfo(int count, int size) {
            if (this.InvokeRequired) {
                this.Invoke(
                    new MethodInvoker(delegate () { updateAttachInfo(count, size); })); }
            else {

                if (count == 0) {
                    label2.Visible = false;
                    label2.Text = "";
                    return;
                }

                label2.Visible = true;
                String infoString;
                if (count == 1) { infoString = "There is 1 attachment, total size attached: " + size + "kB"; }
                else { infoString = "There are " + count + " attachments, total size attached: " + size + "kB"; }
                label2.Text = infoString;
            }
        }

        private void checkBoxStateChanged(object sender, EventArgs e) { myWrapper.updateState(checkBox_addInStatus.Checked); }

        private void button1_Click(object sender, EventArgs e) { myWrapper.addInPaneVisible = false; myWrapper.refreshPane(); }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
