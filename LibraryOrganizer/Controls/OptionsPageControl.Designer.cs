namespace LibraryOrganizer.Controls
{
    partial class OptionsPageControl
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
            this.removeEmptyFolderExclusionsLayout = new System.Windows.Forms.Panel();
            this.removeEmptyFolderExclusions = new System.Windows.Forms.ListBox();
            this.removeEmptyFoldersExclusionsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.profileViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.removeEmptyFolderExclusionsActionPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.addEmptyFolderExclusion = new System.Windows.Forms.Button();
            this.removeEmptyFolderExclusion = new System.Windows.Forms.Button();
            this.removeEmptyFoldersLabel = new System.Windows.Forms.Label();
            this.removeEmptyFolders = new System.Windows.Forms.CheckBox();
            this.illegalCharacterReplacementsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.illegalCharacterReplacementsLabel = new System.Windows.Forms.Label();
            this.illegalCharacterReplacementsCharacterSelector = new System.Windows.Forms.ComboBox();
            this.illegalCharacterReplacementsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.illegalCharacterReplacementsWith = new System.Windows.Forms.Label();
            this.illegalCharacterReplacementsReplacement = new System.Windows.Forms.TextBox();
            this.addIllegalCharacterReplacement = new System.Windows.Forms.Button();
            this.removeIllegalCharacterReplacement = new System.Windows.Forms.Button();
            this.monthReplacementsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.monthReplacementsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.monthReplacementsMonth = new System.Windows.Forms.Label();
            this.monthReplacementsMonthSelector = new System.Windows.Forms.ComboBox();
            this.monthReplacementsWith = new System.Windows.Forms.Label();
            this.monthReplacementsReplacement = new System.Windows.Forms.TextBox();
            this.autoSelectSingleMultiValueField = new System.Windows.Forms.CheckBox();
            this.copyReadPercentageToReplacement = new System.Windows.Forms.CheckBox();
            this.normalizeMultipleSpaces = new System.Windows.Forms.CheckBox();
            this.removeEmptyFolderExclusionsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.removeEmptyFoldersExclusionsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).BeginInit();
            this.removeEmptyFolderExclusionsActionPanel.SuspendLayout();
            this.illegalCharacterReplacementsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.illegalCharacterReplacementsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.monthReplacementsBindingSource)).BeginInit();
            this.monthReplacementsLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // removeEmptyFolderExclusionsLayout
            // 
            this.removeEmptyFolderExclusionsLayout.Controls.Add(this.removeEmptyFolderExclusions);
            this.removeEmptyFolderExclusionsLayout.Controls.Add(this.removeEmptyFolderExclusionsActionPanel);
            this.removeEmptyFolderExclusionsLayout.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileViewModelBindingSource, "RemoveEmptyFolders", true, System.Windows.Forms.DataSourceUpdateMode.Never));
            this.removeEmptyFolderExclusionsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.removeEmptyFolderExclusionsLayout.Location = new System.Drawing.Point(0, 194);
            this.removeEmptyFolderExclusionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolderExclusionsLayout.Name = "removeEmptyFolderExclusionsLayout";
            this.removeEmptyFolderExclusionsLayout.Size = new System.Drawing.Size(818, 470);
            this.removeEmptyFolderExclusionsLayout.TabIndex = 14;
            // 
            // removeEmptyFolderExclusions
            // 
            this.removeEmptyFolderExclusions.DataSource = this.removeEmptyFoldersExclusionsBindingSource;
            this.removeEmptyFolderExclusions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.removeEmptyFolderExclusions.FormattingEnabled = true;
            this.removeEmptyFolderExclusions.ItemHeight = 16;
            this.removeEmptyFolderExclusions.Location = new System.Drawing.Point(0, 0);
            this.removeEmptyFolderExclusions.Name = "removeEmptyFolderExclusions";
            this.removeEmptyFolderExclusions.Size = new System.Drawing.Size(710, 470);
            this.removeEmptyFolderExclusions.TabIndex = 2;
            this.removeEmptyFolderExclusions.EnabledChanged += new System.EventHandler(this.removeEmptyFolderExclusions_EnabledChanged);
            // 
            // removeEmptyFoldersExclusionsBindingSource
            // 
            this.removeEmptyFoldersExclusionsBindingSource.DataMember = "RemoveEmptyFoldersExclusions";
            this.removeEmptyFoldersExclusionsBindingSource.DataSource = this.profileViewModelBindingSource;
            // 
            // profileViewModelBindingSource
            // 
            this.profileViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ProfileViewModel);
            // 
            // removeEmptyFolderExclusionsActionPanel
            // 
            this.removeEmptyFolderExclusionsActionPanel.AutoSize = true;
            this.removeEmptyFolderExclusionsActionPanel.Controls.Add(this.addEmptyFolderExclusion);
            this.removeEmptyFolderExclusionsActionPanel.Controls.Add(this.removeEmptyFolderExclusion);
            this.removeEmptyFolderExclusionsActionPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.removeEmptyFolderExclusionsActionPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.removeEmptyFolderExclusionsActionPanel.Location = new System.Drawing.Point(710, 0);
            this.removeEmptyFolderExclusionsActionPanel.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolderExclusionsActionPanel.Name = "removeEmptyFolderExclusionsActionPanel";
            this.removeEmptyFolderExclusionsActionPanel.Size = new System.Drawing.Size(108, 470);
            this.removeEmptyFolderExclusionsActionPanel.TabIndex = 1;
            // 
            // addEmptyFolderExclusion
            // 
            this.addEmptyFolderExclusion.Location = new System.Drawing.Point(4, 0);
            this.addEmptyFolderExclusion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 4);
            this.addEmptyFolderExclusion.Name = "addEmptyFolderExclusion";
            this.addEmptyFolderExclusion.Size = new System.Drawing.Size(100, 28);
            this.addEmptyFolderExclusion.TabIndex = 0;
            this.addEmptyFolderExclusion.Text = "Add";
            this.addEmptyFolderExclusion.UseVisualStyleBackColor = true;
            this.addEmptyFolderExclusion.Click += new System.EventHandler(this.addEmptyFolderExclusion_Click);
            // 
            // removeEmptyFolderExclusion
            // 
            this.removeEmptyFolderExclusion.Location = new System.Drawing.Point(4, 36);
            this.removeEmptyFolderExclusion.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolderExclusion.Name = "removeEmptyFolderExclusion";
            this.removeEmptyFolderExclusion.Size = new System.Drawing.Size(100, 28);
            this.removeEmptyFolderExclusion.TabIndex = 1;
            this.removeEmptyFolderExclusion.Text = "Remove";
            this.removeEmptyFolderExclusion.UseVisualStyleBackColor = true;
            this.removeEmptyFolderExclusion.Click += new System.EventHandler(this.removeEmptyFolderExclusion_Click);
            // 
            // removeEmptyFoldersLabel
            // 
            this.removeEmptyFoldersLabel.AutoSize = true;
            this.removeEmptyFoldersLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.removeEmptyFoldersLabel.Location = new System.Drawing.Point(0, 175);
            this.removeEmptyFoldersLabel.Margin = new System.Windows.Forms.Padding(0);
            this.removeEmptyFoldersLabel.Name = "removeEmptyFoldersLabel";
            this.removeEmptyFoldersLabel.Padding = new System.Windows.Forms.Padding(3, 0, 0, 3);
            this.removeEmptyFoldersLabel.Size = new System.Drawing.Size(241, 19);
            this.removeEmptyFoldersLabel.TabIndex = 13;
            this.removeEmptyFoldersLabel.Text = "But do not remove the following folders:";
            // 
            // removeEmptyFolders
            // 
            this.removeEmptyFolders.AutoSize = true;
            this.removeEmptyFolders.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "RemoveEmptyFolders", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.removeEmptyFolders.Dock = System.Windows.Forms.DockStyle.Top;
            this.removeEmptyFolders.Location = new System.Drawing.Point(0, 149);
            this.removeEmptyFolders.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolders.Name = "removeEmptyFolders";
            this.removeEmptyFolders.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.removeEmptyFolders.Size = new System.Drawing.Size(818, 26);
            this.removeEmptyFolders.TabIndex = 12;
            this.removeEmptyFolders.Text = "Remove empty folders";
            this.removeEmptyFolders.UseVisualStyleBackColor = true;
            // 
            // illegalCharacterReplacementsLayout
            // 
            this.illegalCharacterReplacementsLayout.AutoSize = true;
            this.illegalCharacterReplacementsLayout.Controls.Add(this.illegalCharacterReplacementsLabel);
            this.illegalCharacterReplacementsLayout.Controls.Add(this.illegalCharacterReplacementsCharacterSelector);
            this.illegalCharacterReplacementsLayout.Controls.Add(this.illegalCharacterReplacementsWith);
            this.illegalCharacterReplacementsLayout.Controls.Add(this.illegalCharacterReplacementsReplacement);
            this.illegalCharacterReplacementsLayout.Controls.Add(this.addIllegalCharacterReplacement);
            this.illegalCharacterReplacementsLayout.Controls.Add(this.removeIllegalCharacterReplacement);
            this.illegalCharacterReplacementsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.illegalCharacterReplacementsLayout.Location = new System.Drawing.Point(0, 111);
            this.illegalCharacterReplacementsLayout.Name = "illegalCharacterReplacementsLayout";
            this.illegalCharacterReplacementsLayout.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.illegalCharacterReplacementsLayout.Size = new System.Drawing.Size(818, 38);
            this.illegalCharacterReplacementsLayout.TabIndex = 7;
            // 
            // illegalCharacterReplacementsLabel
            // 
            this.illegalCharacterReplacementsLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.illegalCharacterReplacementsLabel.AutoSize = true;
            this.illegalCharacterReplacementsLabel.Location = new System.Drawing.Point(0, 14);
            this.illegalCharacterReplacementsLabel.Margin = new System.Windows.Forms.Padding(0, 0, 4, 0);
            this.illegalCharacterReplacementsLabel.Name = "illegalCharacterReplacementsLabel";
            this.illegalCharacterReplacementsLabel.Size = new System.Drawing.Size(157, 16);
            this.illegalCharacterReplacementsLabel.TabIndex = 0;
            this.illegalCharacterReplacementsLabel.Text = "Replace illegal character";
            // 
            // illegalCharacterReplacementsCharacterSelector
            // 
            this.illegalCharacterReplacementsCharacterSelector.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.illegalCharacterReplacementsCharacterSelector.DataSource = this.illegalCharacterReplacementsBindingSource;
            this.illegalCharacterReplacementsCharacterSelector.DisplayMember = "Character";
            this.illegalCharacterReplacementsCharacterSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.illegalCharacterReplacementsCharacterSelector.FormattingEnabled = true;
            this.illegalCharacterReplacementsCharacterSelector.Location = new System.Drawing.Point(165, 10);
            this.illegalCharacterReplacementsCharacterSelector.Margin = new System.Windows.Forms.Padding(4);
            this.illegalCharacterReplacementsCharacterSelector.Name = "illegalCharacterReplacementsCharacterSelector";
            this.illegalCharacterReplacementsCharacterSelector.Size = new System.Drawing.Size(44, 24);
            this.illegalCharacterReplacementsCharacterSelector.TabIndex = 1;
            this.illegalCharacterReplacementsCharacterSelector.ValueMember = "Character";
            // 
            // illegalCharacterReplacementsBindingSource
            // 
            this.illegalCharacterReplacementsBindingSource.DataMember = "IllegalCharacterReplacements";
            this.illegalCharacterReplacementsBindingSource.DataSource = this.profileViewModelBindingSource;
            // 
            // illegalCharacterReplacementsWith
            // 
            this.illegalCharacterReplacementsWith.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.illegalCharacterReplacementsWith.AutoSize = true;
            this.illegalCharacterReplacementsWith.Location = new System.Drawing.Point(217, 14);
            this.illegalCharacterReplacementsWith.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.illegalCharacterReplacementsWith.Name = "illegalCharacterReplacementsWith";
            this.illegalCharacterReplacementsWith.Size = new System.Drawing.Size(29, 16);
            this.illegalCharacterReplacementsWith.TabIndex = 2;
            this.illegalCharacterReplacementsWith.Text = "with";
            // 
            // illegalCharacterReplacementsReplacement
            // 
            this.illegalCharacterReplacementsReplacement.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.illegalCharacterReplacementsReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.illegalCharacterReplacementsBindingSource, "Replacement", true));
            this.illegalCharacterReplacementsReplacement.Location = new System.Drawing.Point(253, 11);
            this.illegalCharacterReplacementsReplacement.Name = "illegalCharacterReplacementsReplacement";
            this.illegalCharacterReplacementsReplacement.Size = new System.Drawing.Size(57, 22);
            this.illegalCharacterReplacementsReplacement.TabIndex = 6;
            this.illegalCharacterReplacementsReplacement.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.illegalCharacterReplacementsReplacement_KeyPress);
            // 
            // addIllegalCharacterReplacement
            // 
            this.addIllegalCharacterReplacement.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.addIllegalCharacterReplacement.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.addIllegalCharacterReplacement.Location = new System.Drawing.Point(317, 10);
            this.addIllegalCharacterReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.addIllegalCharacterReplacement.Name = "addIllegalCharacterReplacement";
            this.addIllegalCharacterReplacement.Size = new System.Drawing.Size(24, 24);
            this.addIllegalCharacterReplacement.TabIndex = 4;
            this.addIllegalCharacterReplacement.Text = "+";
            this.addIllegalCharacterReplacement.UseVisualStyleBackColor = true;
            this.addIllegalCharacterReplacement.Click += new System.EventHandler(this.addIllegalCharacterReplacement_Click);
            // 
            // removeIllegalCharacterReplacement
            // 
            this.removeIllegalCharacterReplacement.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.removeIllegalCharacterReplacement.Location = new System.Drawing.Point(349, 10);
            this.removeIllegalCharacterReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.removeIllegalCharacterReplacement.Name = "removeIllegalCharacterReplacement";
            this.removeIllegalCharacterReplacement.Size = new System.Drawing.Size(24, 24);
            this.removeIllegalCharacterReplacement.TabIndex = 5;
            this.removeIllegalCharacterReplacement.Text = "-";
            this.removeIllegalCharacterReplacement.UseVisualStyleBackColor = true;
            this.removeIllegalCharacterReplacement.Click += new System.EventHandler(this.removeIllegalCharacterReplacement_Click);
            // 
            // monthReplacementsBindingSource
            // 
            this.monthReplacementsBindingSource.DataMember = "MonthReplacements";
            this.monthReplacementsBindingSource.DataSource = this.profileViewModelBindingSource;
            // 
            // monthReplacementsLayout
            // 
            this.monthReplacementsLayout.AutoSize = true;
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsMonth);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsMonthSelector);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsWith);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsReplacement);
            this.monthReplacementsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.monthReplacementsLayout.Location = new System.Drawing.Point(0, 73);
            this.monthReplacementsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.monthReplacementsLayout.Name = "monthReplacementsLayout";
            this.monthReplacementsLayout.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.monthReplacementsLayout.Size = new System.Drawing.Size(818, 38);
            this.monthReplacementsLayout.TabIndex = 11;
            // 
            // monthReplacementsMonth
            // 
            this.monthReplacementsMonth.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.monthReplacementsMonth.AutoSize = true;
            this.monthReplacementsMonth.Location = new System.Drawing.Point(0, 14);
            this.monthReplacementsMonth.Margin = new System.Windows.Forms.Padding(0, 0, 4, 0);
            this.monthReplacementsMonth.Name = "monthReplacementsMonth";
            this.monthReplacementsMonth.Size = new System.Drawing.Size(43, 16);
            this.monthReplacementsMonth.TabIndex = 0;
            this.monthReplacementsMonth.Text = "Month";
            // 
            // monthReplacementsMonthSelector
            // 
            this.monthReplacementsMonthSelector.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.monthReplacementsMonthSelector.DataSource = this.monthReplacementsBindingSource;
            this.monthReplacementsMonthSelector.DisplayMember = "Month";
            this.monthReplacementsMonthSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.monthReplacementsMonthSelector.FormattingEnabled = true;
            this.monthReplacementsMonthSelector.Location = new System.Drawing.Point(51, 10);
            this.monthReplacementsMonthSelector.Margin = new System.Windows.Forms.Padding(4);
            this.monthReplacementsMonthSelector.Name = "monthReplacementsMonthSelector";
            this.monthReplacementsMonthSelector.Size = new System.Drawing.Size(63, 24);
            this.monthReplacementsMonthSelector.TabIndex = 1;
            this.monthReplacementsMonthSelector.ValueMember = "Month";
            // 
            // monthReplacementsWith
            // 
            this.monthReplacementsWith.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.monthReplacementsWith.AutoSize = true;
            this.monthReplacementsWith.Location = new System.Drawing.Point(122, 14);
            this.monthReplacementsWith.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.monthReplacementsWith.Name = "monthReplacementsWith";
            this.monthReplacementsWith.Size = new System.Drawing.Size(17, 16);
            this.monthReplacementsWith.TabIndex = 2;
            this.monthReplacementsWith.Text = "is";
            // 
            // monthReplacementsReplacement
            // 
            this.monthReplacementsReplacement.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.monthReplacementsReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.monthReplacementsBindingSource, "Replacement", true));
            this.monthReplacementsReplacement.Location = new System.Drawing.Point(147, 11);
            this.monthReplacementsReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.monthReplacementsReplacement.Name = "monthReplacementsReplacement";
            this.monthReplacementsReplacement.Size = new System.Drawing.Size(192, 22);
            this.monthReplacementsReplacement.TabIndex = 3;
            // 
            // autoSelectSingleMultiValueField
            // 
            this.autoSelectSingleMultiValueField.AutoSize = true;
            this.autoSelectSingleMultiValueField.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "AutoSelectSingleMultiValueField", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.autoSelectSingleMultiValueField.Dock = System.Windows.Forms.DockStyle.Top;
            this.autoSelectSingleMultiValueField.Location = new System.Drawing.Point(0, 47);
            this.autoSelectSingleMultiValueField.Margin = new System.Windows.Forms.Padding(4);
            this.autoSelectSingleMultiValueField.Name = "autoSelectSingleMultiValueField";
            this.autoSelectSingleMultiValueField.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.autoSelectSingleMultiValueField.Size = new System.Drawing.Size(818, 26);
            this.autoSelectSingleMultiValueField.TabIndex = 10;
            this.autoSelectSingleMultiValueField.Text = "If there is only one value in a multiple value field then insert it without askin" +
    "g";
            this.autoSelectSingleMultiValueField.UseVisualStyleBackColor = true;
            // 
            // copyReadPercentageToReplacement
            // 
            this.copyReadPercentageToReplacement.AutoSize = true;
            this.copyReadPercentageToReplacement.DataBindings.Add(new System.Windows.Forms.Binding("CheckState", this.profileViewModelBindingSource, "CopyReadPercentageToReplacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.copyReadPercentageToReplacement.Dock = System.Windows.Forms.DockStyle.Top;
            this.copyReadPercentageToReplacement.Location = new System.Drawing.Point(0, 21);
            this.copyReadPercentageToReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.copyReadPercentageToReplacement.Name = "copyReadPercentageToReplacement";
            this.copyReadPercentageToReplacement.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.copyReadPercentageToReplacement.Size = new System.Drawing.Size(818, 26);
            this.copyReadPercentageToReplacement.TabIndex = 9;
            this.copyReadPercentageToReplacement.Text = "When overwriting an existing file, copy the read percentage to the new file";
            this.copyReadPercentageToReplacement.UseVisualStyleBackColor = true;
            // 
            // normalizeMultipleSpaces
            // 
            this.normalizeMultipleSpaces.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileViewModelBindingSource, "NormalizeMultipleSpaces", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.normalizeMultipleSpaces.Dock = System.Windows.Forms.DockStyle.Top;
            this.normalizeMultipleSpaces.Location = new System.Drawing.Point(0, 0);
            this.normalizeMultipleSpaces.Margin = new System.Windows.Forms.Padding(4);
            this.normalizeMultipleSpaces.Name = "normalizeMultipleSpaces";
            this.normalizeMultipleSpaces.Size = new System.Drawing.Size(818, 21);
            this.normalizeMultipleSpaces.TabIndex = 8;
            this.normalizeMultipleSpaces.Text = "Replace multiple spaces with a single space";
            this.normalizeMultipleSpaces.UseVisualStyleBackColor = true;
            // 
            // OptionsPageControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.removeEmptyFolderExclusionsLayout);
            this.Controls.Add(this.removeEmptyFoldersLabel);
            this.Controls.Add(this.removeEmptyFolders);
            this.Controls.Add(this.illegalCharacterReplacementsLayout);
            this.Controls.Add(this.monthReplacementsLayout);
            this.Controls.Add(this.autoSelectSingleMultiValueField);
            this.Controls.Add(this.copyReadPercentageToReplacement);
            this.Controls.Add(this.normalizeMultipleSpaces);
            this.Name = "OptionsPageControl";
            this.Size = new System.Drawing.Size(818, 664);
            this.removeEmptyFolderExclusionsLayout.ResumeLayout(false);
            this.removeEmptyFolderExclusionsLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.removeEmptyFoldersExclusionsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).EndInit();
            this.removeEmptyFolderExclusionsActionPanel.ResumeLayout(false);
            this.illegalCharacterReplacementsLayout.ResumeLayout(false);
            this.illegalCharacterReplacementsLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.illegalCharacterReplacementsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.monthReplacementsBindingSource)).EndInit();
            this.monthReplacementsLayout.ResumeLayout(false);
            this.monthReplacementsLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel removeEmptyFolderExclusionsLayout;
        private System.Windows.Forms.ListBox removeEmptyFolderExclusions;
        private System.Windows.Forms.FlowLayoutPanel removeEmptyFolderExclusionsActionPanel;
        private System.Windows.Forms.Button addEmptyFolderExclusion;
        private System.Windows.Forms.Button removeEmptyFolderExclusion;
        private System.Windows.Forms.Label removeEmptyFoldersLabel;
        private System.Windows.Forms.CheckBox removeEmptyFolders;
        private System.Windows.Forms.FlowLayoutPanel illegalCharacterReplacementsLayout;
        private System.Windows.Forms.Label illegalCharacterReplacementsLabel;
        private System.Windows.Forms.ComboBox illegalCharacterReplacementsCharacterSelector;
        private System.Windows.Forms.Label illegalCharacterReplacementsWith;
        private System.Windows.Forms.TextBox illegalCharacterReplacementsReplacement;
        private System.Windows.Forms.Button addIllegalCharacterReplacement;
        private System.Windows.Forms.Button removeIllegalCharacterReplacement;
        private System.Windows.Forms.FlowLayoutPanel monthReplacementsLayout;
        private System.Windows.Forms.Label monthReplacementsMonth;
        private System.Windows.Forms.ComboBox monthReplacementsMonthSelector;
        private System.Windows.Forms.Label monthReplacementsWith;
        private System.Windows.Forms.TextBox monthReplacementsReplacement;
        private System.Windows.Forms.CheckBox autoSelectSingleMultiValueField;
        private System.Windows.Forms.CheckBox copyReadPercentageToReplacement;
        private System.Windows.Forms.CheckBox normalizeMultipleSpaces;
        private System.Windows.Forms.BindingSource profileViewModelBindingSource;
        private System.Windows.Forms.BindingSource monthReplacementsBindingSource;
        private System.Windows.Forms.BindingSource removeEmptyFoldersExclusionsBindingSource;
        private System.Windows.Forms.BindingSource illegalCharacterReplacementsBindingSource;
    }
}
