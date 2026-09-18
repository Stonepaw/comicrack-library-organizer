namespace LibraryOrganizer.Plugin
{
    partial class LibraryOrganizerTestFormSelector
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openConfigureForm = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openConfigureForm
            // 
            this.openConfigureForm.AutoSize = true;
            this.openConfigureForm.Location = new System.Drawing.Point(6, 21);
            this.openConfigureForm.Name = "openConfigureForm";
            this.openConfigureForm.Size = new System.Drawing.Size(108, 26);
            this.openConfigureForm.TabIndex = 0;
            this.openConfigureForm.Text = "Configure Form";
            this.openConfigureForm.UseVisualStyleBackColor = true;
            this.openConfigureForm.Click += new System.EventHandler(this.openConfigureForm_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.openConfigureForm);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(381, 100);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Open Form";
            // 
            // LibraryOrganizerTestFormSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(405, 348);
            this.Controls.Add(this.groupBox1);
            this.Name = "LibraryOrganizerTestFormSelector";
            this.Text = "Library Organizer Test Selector";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button openConfigureForm;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

