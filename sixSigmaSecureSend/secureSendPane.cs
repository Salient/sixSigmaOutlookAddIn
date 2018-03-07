using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;

namespace sixSigmaSecureSend
{
    public partial class secureSendPane : UserControl
    {
        editorWrapper myWrapper;
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

        public void suggest()
        {
            string text = "It looks like you are sending an email with an attachment to recipients outside of Raytheon.\n\nRemember to safeguard sensitive or technical information by using PKI encryption or the Raytheon Secure Messaging Gateway.";

            updateText(text);

        }
        public void stdtext()
        {
            string text = "Always remember, safeguarding the integrity of sensitive information is a personal responsibility critical to protecting national interests, corporate reputation, and the warfighter. If this message contains technical detail or potentially sensitive information, consider adding [RSMG] to the subject line.";

            updateText(text);

        }

        public void noteNoEffect()
        {
            string text = "You have added [RSMG] to the subject line, indicating you want to safeguard the information in this message when sending to recipients outside of Raytheon. This is an excellent habit, however, I couldn't help noticing there aren't any external recipients. \n\nRaytheon Secure Messaging has no effect with sending to other Raytheon recipients because it never leaves the Raytheon network, and the Raytheon network is always secure :-)";

            updateText(text);
        }
        public void kudos()
        {
            string text = "This message shall be sent to external recipients via the Raytheon Secure Message Gateway. Recipients will receive an email with notification and instructions on how to retrieve this email securely. This is the safest way to transmit any information which may be sensitive.";

            updateText(text);
        }

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

        private void checkBoxStateChanged(object sender, EventArgs e)
        {
            myWrapper.updateState(checkBox_addInStatus.Checked);
        }
    }
}
