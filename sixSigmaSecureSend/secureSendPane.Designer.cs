namespace sixSigmaSecureSend
{
    partial class secureSendPane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(secureSendPane));
            this.checkBox_addInStatus = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.rtnsecurelogo = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rtnsecurelogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // checkBox_addInStatus
            // 
            this.checkBox_addInStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox_addInStatus.AutoSize = true;
            this.checkBox_addInStatus.Checked = true;
            this.checkBox_addInStatus.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_addInStatus.Cursor = System.Windows.Forms.Cursors.Hand;
            this.checkBox_addInStatus.Location = new System.Drawing.Point(2, 510);
            this.checkBox_addInStatus.Name = "checkBox_addInStatus";
            this.checkBox_addInStatus.Size = new System.Drawing.Size(151, 17);
            this.checkBox_addInStatus.TabIndex = 3;
            this.checkBox_addInStatus.Text = "Send Secure with [RSMG]";
            this.checkBox_addInStatus.UseVisualStyleBackColor = true;
            this.checkBox_addInStatus.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(0, 196);
            this.label1.MaximumSize = new System.Drawing.Size(0, 260);
            this.label1.MinimumSize = new System.Drawing.Size(0, 240);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 260);
            this.label1.TabIndex = 8;
            this.label1.Text = resources.GetString("label1.Text");
            this.label1.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox2.Image = global::sixSigmaSecureSend.Properties.Resources.R6S_defined_RGB;
            this.pictureBox2.Location = new System.Drawing.Point(-3, 859);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(75, 45);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 9;
            this.pictureBox2.TabStop = false;
            // 
            // rtnsecurelogo
            // 
            this.rtnsecurelogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.rtnsecurelogo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.rtnsecurelogo.Image = global::sixSigmaSecureSend.Properties.Resources.rtnsecuretrans;
            this.rtnsecurelogo.Location = new System.Drawing.Point(78, 859);
            this.rtnsecurelogo.Name = "rtnsecurelogo";
            this.rtnsecurelogo.Size = new System.Drawing.Size(82, 45);
            this.rtnsecurelogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.rtnsecurelogo.TabIndex = 4;
            this.rtnsecurelogo.TabStop = false;
            this.rtnsecurelogo.Click += new System.EventHandler(this.rtnsecurelogo_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.AccessibleDescription = "";
            this.pictureBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = global::sixSigmaSecureSend.Properties.Resources.clippy;
            this.pictureBox1.Location = new System.Drawing.Point(2, 30);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(150, 143);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(10, 200);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(141, 20);
            this.textBox1.TabIndex = 10;
            // 
            // secureSendPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CausesValidation = false;
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rtnsecurelogo);
            this.Controls.Add(this.checkBox_addInStatus);
            this.Controls.Add(this.pictureBox1);
            this.DoubleBuffered = true;
            this.Name = "secureSendPane";
            this.Size = new System.Drawing.Size(160, 904);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rtnsecurelogo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox checkBox_addInStatus;
        private System.Windows.Forms.PictureBox rtnsecurelogo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.TextBox textBox1;
    }
}
