namespace LibraryOrganizer.Controls
{
    partial class OverviewConfigControl
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.TableLayoutPanel modeLayout;
            System.Windows.Forms.Label simulateLabel;
            this.moveMode = new System.Windows.Forms.RadioButton();
            this.copyMode = new System.Windows.Forms.RadioButton();
            this.simulateMode = new System.Windows.Forms.RadioButton();
            this.addCopyToLibrary = new System.Windows.Forms.CheckBox();
            this.fileFolderModePanel = new System.Windows.Forms.TableLayoutPanel();
            this.useFileOrganization = new System.Windows.Forms.CheckBox();
            this.useFolderOrganization = new System.Windows.Forms.CheckBox();
            this.baseFolderPanel = new System.Windows.Forms.TableLayoutPanel();
            this.baseFolder = new System.Windows.Forms.TextBox();
            this.baseFolderBrowse = new System.Windows.Forms.Button();
            this.baseFolderLabel = new System.Windows.Forms.Label();
            this.modeGroup = new System.Windows.Forms.GroupBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.copyFileless = new System.Windows.Forms.CheckBox();
            this.filelessImagePanel = new System.Windows.Forms.TableLayoutPanel();
            this.profileViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            modeLayout = new System.Windows.Forms.TableLayoutPanel();
            simulateLabel = new System.Windows.Forms.Label();
            modeLayout.SuspendLayout();
            this.fileFolderModePanel.SuspendLayout();
            this.baseFolderPanel.SuspendLayout();
            this.modeGroup.SuspendLayout();
            this.filelessImagePanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // modeLayout
            // 
            modeLayout.AutoSize = true;
            modeLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            modeLayout.ColumnCount = 3;
            modeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            modeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            modeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            modeLayout.Controls.Add(this.moveMode, 0, 0);
            modeLayout.Controls.Add(this.copyMode, 1, 0);
            modeLayout.Controls.Add(this.simulateMode, 2, 0);
            modeLayout.Controls.Add(this.addCopyToLibrary, 1, 1);
            modeLayout.Controls.Add(simulateLabel, 2, 1);
            modeLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            modeLayout.Location = new System.Drawing.Point(6, 21);
            modeLayout.Name = "modeLayout";
            modeLayout.RowCount = 2;
            modeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            modeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            modeLayout.Size = new System.Drawing.Size(731, 73);
            modeLayout.TabIndex = 3;
            // 
            // moveMode
            // 
            this.moveMode.AutoSize = true;
            this.moveMode.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "IsMoveMode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.moveMode.Location = new System.Drawing.Point(3, 3);
            this.moveMode.Name = "moveMode";
            this.moveMode.Size = new System.Drawing.Size(62, 20);
            this.moveMode.TabIndex = 0;
            this.moveMode.TabStop = true;
            this.moveMode.Text = "Move";
            this.moveMode.UseVisualStyleBackColor = true;
            // 
            // copyMode
            // 
            this.copyMode.AutoSize = true;
            this.copyMode.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "IsCopyMode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.copyMode.Location = new System.Drawing.Point(246, 3);
            this.copyMode.Name = "copyMode";
            this.copyMode.Size = new System.Drawing.Size(60, 20);
            this.copyMode.TabIndex = 1;
            this.copyMode.TabStop = true;
            this.copyMode.Text = "Copy";
            this.copyMode.UseVisualStyleBackColor = true;
            // 
            // simulateMode
            // 
            this.simulateMode.AutoSize = true;
            this.simulateMode.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "IsSimulateMode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.simulateMode.Location = new System.Drawing.Point(489, 3);
            this.simulateMode.Name = "simulateMode";
            this.simulateMode.Size = new System.Drawing.Size(80, 20);
            this.simulateMode.TabIndex = 2;
            this.simulateMode.TabStop = true;
            this.simulateMode.Text = "Simulate";
            this.simulateMode.UseVisualStyleBackColor = true;
            // 
            // addCopyToLibrary
            // 
            this.addCopyToLibrary.AutoSize = true;
            this.addCopyToLibrary.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "AddCopyToLibrary", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.addCopyToLibrary.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileViewModelBindingSource, "IsCopyMode", true, System.Windows.Forms.DataSourceUpdateMode.Never));
            this.addCopyToLibrary.Location = new System.Drawing.Point(246, 29);
            this.addCopyToLibrary.Name = "addCopyToLibrary";
            this.addCopyToLibrary.Size = new System.Drawing.Size(187, 20);
            this.addCopyToLibrary.TabIndex = 3;
            this.addCopyToLibrary.Text = "Add copied book to library";
            this.addCopyToLibrary.UseVisualStyleBackColor = true;
            // 
            // simulateLabel
            // 
            simulateLabel.AutoSize = true;
            simulateLabel.Location = new System.Drawing.Point(489, 26);
            simulateLabel.Name = "simulateLabel";
            simulateLabel.Size = new System.Drawing.Size(156, 32);
            simulateLabel.TabIndex = 4;
            simulateLabel.Text = "No files touched\r\nComplete log file created";
            // 
            // fileFolderModePanel
            // 
            this.fileFolderModePanel.AutoSize = true;
            this.fileFolderModePanel.ColumnCount = 2;
            this.fileFolderModePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.fileFolderModePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.fileFolderModePanel.Controls.Add(this.useFileOrganization, 0, 0);
            this.fileFolderModePanel.Controls.Add(this.useFolderOrganization, 1, 0);
            this.fileFolderModePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.fileFolderModePanel.Location = new System.Drawing.Point(6, 172);
            this.fileFolderModePanel.Name = "fileFolderModePanel";
            this.fileFolderModePanel.Padding = new System.Windows.Forms.Padding(0, 30, 0, 0);
            this.fileFolderModePanel.RowCount = 1;
            this.fileFolderModePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.fileFolderModePanel.Size = new System.Drawing.Size(743, 56);
            this.fileFolderModePanel.TabIndex = 6;
            // 
            // useFileOrganization
            // 
            this.useFileOrganization.AutoSize = true;
            this.useFileOrganization.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "UseFileNaming", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.useFileOrganization.Location = new System.Drawing.Point(3, 33);
            this.useFileOrganization.Name = "useFileOrganization";
            this.useFileOrganization.Size = new System.Drawing.Size(150, 20);
            this.useFileOrganization.TabIndex = 0;
            this.useFileOrganization.Text = "Use file organization";
            this.useFileOrganization.UseVisualStyleBackColor = true;
            // 
            // useFolderOrganization
            // 
            this.useFolderOrganization.AutoSize = true;
            this.useFolderOrganization.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "UseFolderOrganization", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.useFolderOrganization.Location = new System.Drawing.Point(374, 33);
            this.useFolderOrganization.Name = "useFolderOrganization";
            this.useFolderOrganization.Size = new System.Drawing.Size(167, 20);
            this.useFolderOrganization.TabIndex = 1;
            this.useFolderOrganization.Text = "Use folder organization";
            this.useFolderOrganization.UseVisualStyleBackColor = true;
            // 
            // baseFolderPanel
            // 
            this.baseFolderPanel.AutoSize = true;
            this.baseFolderPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.baseFolderPanel.ColumnCount = 3;
            this.baseFolderPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.baseFolderPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.baseFolderPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.baseFolderPanel.Controls.Add(this.baseFolder, 1, 0);
            this.baseFolderPanel.Controls.Add(this.baseFolderBrowse, 2, 0);
            this.baseFolderPanel.Controls.Add(this.baseFolderLabel, 0, 0);
            this.baseFolderPanel.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileViewModelBindingSource, "UseFolderOrganization", true, System.Windows.Forms.DataSourceUpdateMode.Never));
            this.baseFolderPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.baseFolderPanel.Location = new System.Drawing.Point(6, 100);
            this.baseFolderPanel.Margin = new System.Windows.Forms.Padding(0);
            this.baseFolderPanel.Name = "baseFolderPanel";
            this.baseFolderPanel.Padding = new System.Windows.Forms.Padding(0, 40, 0, 0);
            this.baseFolderPanel.RowCount = 1;
            this.baseFolderPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.baseFolderPanel.Size = new System.Drawing.Size(743, 72);
            this.baseFolderPanel.TabIndex = 5;
            // 
            // baseFolder
            // 
            this.baseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.baseFolder.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.profileViewModelBindingSource, "BaseFolder", true));
            this.baseFolder.Location = new System.Drawing.Point(90, 45);
            this.baseFolder.Name = "baseFolder";
            this.baseFolder.Size = new System.Drawing.Size(569, 22);
            this.baseFolder.TabIndex = 0;
            // 
            // baseFolderBrowse
            // 
            this.baseFolderBrowse.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.baseFolderBrowse.AutoSize = true;
            this.baseFolderBrowse.Location = new System.Drawing.Point(665, 43);
            this.baseFolderBrowse.Name = "baseFolderBrowse";
            this.baseFolderBrowse.Size = new System.Drawing.Size(75, 26);
            this.baseFolderBrowse.TabIndex = 1;
            this.baseFolderBrowse.Text = "Browse";
            this.baseFolderBrowse.UseVisualStyleBackColor = true;
            this.baseFolderBrowse.Click += new System.EventHandler(this.baseFolderBrowse_Click);
            // 
            // baseFolderLabel
            // 
            this.baseFolderLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.baseFolderLabel.AutoSize = true;
            this.baseFolderLabel.Location = new System.Drawing.Point(3, 48);
            this.baseFolderLabel.Name = "baseFolderLabel";
            this.baseFolderLabel.Size = new System.Drawing.Size(81, 16);
            this.baseFolderLabel.TabIndex = 2;
            this.baseFolderLabel.Text = "Base Folder";
            // 
            // modeGroup
            // 
            this.modeGroup.Controls.Add(modeLayout);
            this.modeGroup.Dock = System.Windows.Forms.DockStyle.Top;
            this.modeGroup.Location = new System.Drawing.Point(6, 0);
            this.modeGroup.Name = "modeGroup";
            this.modeGroup.Padding = new System.Windows.Forms.Padding(6);
            this.modeGroup.Size = new System.Drawing.Size(743, 100);
            this.modeGroup.TabIndex = 4;
            this.modeGroup.TabStop = false;
            this.modeGroup.Text = "Mode";
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.comboBox1.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.profileViewModelBindingSource, "CopyFilelessBookThumbnailFormat", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.comboBox1.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileViewModelBindingSource, "CopyFilelessBookThumbnail", true, System.Windows.Forms.DataSourceUpdateMode.Never));
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            ".bmp",
            ".jpg",
            ".png"});
            this.comboBox1.Location = new System.Drawing.Point(617, 44);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(123, 24);
            this.comboBox1.TabIndex = 1;
            // 
            // copyFileless
            // 
            this.copyFileless.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.copyFileless.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "CopyFilelessBookThumbnail", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.copyFileless.Location = new System.Drawing.Point(3, 33);
            this.copyFileless.Name = "copyFileless";
            this.copyFileless.Size = new System.Drawing.Size(608, 47);
            this.copyFileless.TabIndex = 0;
            this.copyFileless.Text = "Copy fileless book custom thumbnails to the calculated path.\r\n(Does not affect th" +
    "e original image)";
            this.copyFileless.UseVisualStyleBackColor = true;
            // 
            // filelessImagePanel
            // 
            this.filelessImagePanel.AutoSize = true;
            this.filelessImagePanel.ColumnCount = 2;
            this.filelessImagePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.filelessImagePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.filelessImagePanel.Controls.Add(this.copyFileless, 0, 0);
            this.filelessImagePanel.Controls.Add(this.comboBox1, 1, 0);
            this.filelessImagePanel.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileViewModelBindingSource, "UseFolderOrganization", true, System.Windows.Forms.DataSourceUpdateMode.Never));
            this.filelessImagePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.filelessImagePanel.Location = new System.Drawing.Point(6, 228);
            this.filelessImagePanel.Name = "filelessImagePanel";
            this.filelessImagePanel.Padding = new System.Windows.Forms.Padding(0, 30, 0, 0);
            this.filelessImagePanel.RowCount = 1;
            this.filelessImagePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.filelessImagePanel.Size = new System.Drawing.Size(743, 83);
            this.filelessImagePanel.TabIndex = 8;
            // 
            // profileViewModelBindingSource
            // 
            this.profileViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ProfileViewModel);
            // 
            // OverviewConfigControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.filelessImagePanel);
            this.Controls.Add(this.fileFolderModePanel);
            this.Controls.Add(this.baseFolderPanel);
            this.Controls.Add(this.modeGroup);
            this.Name = "OverviewConfigControl";
            this.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.Size = new System.Drawing.Size(755, 584);
            modeLayout.ResumeLayout(false);
            modeLayout.PerformLayout();
            this.fileFolderModePanel.ResumeLayout(false);
            this.fileFolderModePanel.PerformLayout();
            this.baseFolderPanel.ResumeLayout(false);
            this.baseFolderPanel.PerformLayout();
            this.modeGroup.ResumeLayout(false);
            this.modeGroup.PerformLayout();
            this.filelessImagePanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel fileFolderModePanel;
        private System.Windows.Forms.CheckBox useFileOrganization;
        private System.Windows.Forms.CheckBox useFolderOrganization;
        private System.Windows.Forms.TableLayoutPanel baseFolderPanel;
        private System.Windows.Forms.TextBox baseFolder;
        private System.Windows.Forms.Button baseFolderBrowse;
        private System.Windows.Forms.Label baseFolderLabel;
        private System.Windows.Forms.GroupBox modeGroup;
        private System.Windows.Forms.RadioButton moveMode;
        private System.Windows.Forms.RadioButton copyMode;
        private System.Windows.Forms.RadioButton simulateMode;
        private System.Windows.Forms.CheckBox addCopyToLibrary;
        private System.Windows.Forms.BindingSource profileViewModelBindingSource;
        private System.Windows.Forms.CheckBox copyFileless;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TableLayoutPanel filelessImagePanel;
    }
}
