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
            this.checkBox_addInStatus = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.sixsigmalogo = new System.Windows.Forms.PictureBox();
            this.rtnsecurelogo = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.closePane = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.sixsigmalogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rtnsecurelogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // checkBox_addInStatus
            // 
            this.checkBox_addInStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox_addInStatus.AutoSize = true;
            this.checkBox_addInStatus.Checked = true;
            this.checkBox_addInStatus.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_addInStatus.Cursor = System.Windows.Forms.Cursors.Hand;
            this.checkBox_addInStatus.Location = new System.Drawing.Point(33, 533);
            this.checkBox_addInStatus.Name = "checkBox_addInStatus";
            this.checkBox_addInStatus.Size = new System.Drawing.Size(151, 17);
            this.checkBox_addInStatus.TabIndex = 3;
            this.checkBox_addInStatus.Text = "Send Secure with [RSMG]";
            this.checkBox_addInStatus.UseVisualStyleBackColor = true;
            this.checkBox_addInStatus.CheckStateChanged += new System.EventHandler(this.checkBoxStateChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(11, 200);
            this.label1.MinimumSize = new System.Drawing.Size(0, 200);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(194, 260);
            this.label1.TabIndex = 8;
            this.label1.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // sixsigmalogo
            // 
            this.sixsigmalogo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sixsigmalogo.Image = global::sixSigmaSecureSend.Properties.Resources.R6S_defined_RGB;
            this.sixsigmalogo.Location = new System.Drawing.Point(124, 630);
            this.sixsigmalogo.Name = "sixsigmalogo";
            this.sixsigmalogo.Size = new System.Drawing.Size(87, 53);
            this.sixsigmalogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.sixsigmalogo.TabIndex = 9;
            this.sixsigmalogo.TabStop = false;
            this.sixsigmalogo.Click += new System.EventHandler(this.sixsigmalogo_Click);
            // 
            // rtnsecurelogo
            // 
            this.rtnsecurelogo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtnsecurelogo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.rtnsecurelogo.Image = global::sixSigmaSecureSend.Properties.Resources.rtnsecuretrans;
            this.rtnsecurelogo.Location = new System.Drawing.Point(8, 630);
            this.rtnsecurelogo.Name = "rtnsecurelogo";
            this.rtnsecurelogo.Size = new System.Drawing.Size(98, 53);
            this.rtnsecurelogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.rtnsecurelogo.TabIndex = 4;
            this.rtnsecurelogo.TabStop = false;
            this.rtnsecurelogo.Click += new System.EventHandler(this.rtnsecurelogo_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.AccessibleDescription = "";
            this.pictureBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = global::sixSigmaSecureSend.Properties.Resources.clippy;
            this.pictureBox1.Location = new System.Drawing.Point(30, 30);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(152, 143);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // closePane
            // 
            this.closePane.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.closePane.Location = new System.Drawing.Point(67, 556);
            this.closePane.Name = "closePane";
            this.closePane.Size = new System.Drawing.Size(77, 23);
            this.closePane.TabIndex = 10;
            this.closePane.Text = "Close";
            this.closePane.UseVisualStyleBackColor = true;
            this.closePane.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(19, 475);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(190, 48);
            this.label2.TabIndex = 11;
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // secureSendPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CausesValidation = false;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.closePane);
            this.Controls.Add(this.checkBox_addInStatus);
            this.Controls.Add(this.sixsigmalogo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rtnsecurelogo);
            this.Controls.Add(this.pictureBox1);
            this.DoubleBuffered = true;
            this.Margin = new System.Windows.Forms.Padding(1);
            this.MinimumSize = new System.Drawing.Size(240, 500);
            this.Name = "secureSendPane";
            this.Size = new System.Drawing.Size(240, 686);
            ((System.ComponentModel.ISupportInitialize)(this.sixsigmalogo)).EndInit();
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
        private System.Windows.Forms.PictureBox sixsigmalogo;
        private System.Windows.Forms.Button closePane;
        private System.Windows.Forms.Label label2;
    }
}
