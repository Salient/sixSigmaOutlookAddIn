﻿using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;

namespace sixSigmaSecureSend
{
    public partial class secureSendPane : UserControl
    {
        public secureSendPane()
        {
            InitializeComponent();
        }

        delegate void StringArgReturningVoidDelegate();

        // Helper functions
        public void setBox_addInActive(bool set)
        {
            this.checkBox_addInStatus.Checked = set;
            this.label1.Visible = true;
        }

        // Callbacks
        private void label1_Click(object sender, EventArgs e)
        {
            Debug.Print("something creative maybe?");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            Debug.Print("button changed");
        }

        public void rtnsecurelogo_Click(object sender, EventArgs e)
        {
            Process.Start("http://web.onertn.ray.com/initiatives/rtnsecurecenter/");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        public void suggest()
        {
            string text = "It looks like you are sending an email with an attachment to recipients outside of Raytheon.\n\nRemember to safeguard sensitive or technical information by using PKI encryption or the Raytheon Secure Messaging Gateway.";

            updateText(text);

        }

        public void noteNoEffect()
        {
            string text = "You have added [RSMG] to the subject line, indicating you want to safeguard the information in this message when sending to recipients outside of Raytheon. This is an excellent habit, however, I couldn't help noticing there aren't any external recipients. \n\nRaytheon Secure Messaging has no effect with sending to other Raytheon recipients because it never leaves the Raytheon network, and the Raytheon network is always secure :-)";

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
            }
            this.label1.Visible = true;
            this.label1.Show();
        }
    }
}
