namespace LibraryOrganizer.Dialog
{
    partial class ConfigureForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.configurationPanel = new System.Windows.Forms.Panel();
            this.optionsPage = new System.Windows.Forms.TabControl();
            this.optionsTabPage = new System.Windows.Forms.TabPage();
            this.removeEmptyFolderExclusionsLayout = new System.Windows.Forms.Panel();
            this.removeEmptyFolderExclusions = new System.Windows.Forms.ListBox();
            this.profileBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.removeEmptyFoldersExclusionsBindingSource = new System.Windows.Forms.BindingSource(this.components);
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
            this.monthReplacementsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.monthReplacementsMonth = new System.Windows.Forms.Label();
            this.monthReplacementsMonthSelector = new System.Windows.Forms.ComboBox();
            this.monthReplacementsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.monthReplacementsWith = new System.Windows.Forms.Label();
            this.monthReplacementsReplacement = new System.Windows.Forms.TextBox();
            this.autoSelectSingleMultiValueField = new System.Windows.Forms.CheckBox();
            this.copyReadPercentageToReplacement = new System.Windows.Forms.CheckBox();
            this.normalizeMultipleSpaces = new System.Windows.Forms.CheckBox();
            this.emptyValuesTabPage = new System.Windows.Forms.TabPage();
            this.failOperationOnEmptyValueDestinationFolderLayout = new System.Windows.Forms.TableLayoutPanel();
            this.failOperationOnEmptyValueDestinationFolderBrowse = new System.Windows.Forms.Button();
            this.failOperationOnEmptyValueDestinationFolder = new System.Windows.Forms.TextBox();
            this.failOperationOnEmptyValueUseDestinationFolder = new System.Windows.Forms.CheckBox();
            this.failOperationOnEmptyValueFieldsLayout = new System.Windows.Forms.Panel();
            this.failOperationOnEmptyValueFields = new System.Windows.Forms.DataGridView();
            this.emptyValueFieldEnabledColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.emptyValueFieldNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.failOperationOnEmptyValueFieldsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.failOperationOnEmptyValue = new System.Windows.Forms.CheckBox();
            this.emptyFieldReplacementLayout = new System.Windows.Forms.TableLayoutPanel();
            this.emptyFieldReplacementLabel2 = new System.Windows.Forms.Label();
            this.emptyFieldReplacementSelector = new System.Windows.Forms.ComboBox();
            this.emptyFieldReplacementsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.emptyFieldReplacementLabel3 = new System.Windows.Forms.Label();
            this.emptyFieldReplacement = new System.Windows.Forms.TextBox();
            this.emptyFieldReplacementLabel = new System.Windows.Forms.Label();
            this.emptyFolderNameReplacementLabel2 = new System.Windows.Forms.Label();
            this.emptyFolderNameReplacementLayout = new System.Windows.Forms.TableLayoutPanel();
            this.emptyFolderNameReplacement = new System.Windows.Forms.TextBox();
            this.emptyFolderNameReplacementLabel = new System.Windows.Forms.Label();
            this.rulesPage = new System.Windows.Forms.TabControl();
            this.metadataRulesTabPage = new System.Windows.Forms.TabPage();
            this.metadataRulesActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.metadataRulesControlsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.metadataRulesAction = new System.Windows.Forms.ComboBox();
            this.metadataRulesActionLabel1 = new System.Windows.Forms.Label();
            this.metadataRulesMode = new System.Windows.Forms.ComboBox();
            this.metadataRulesActionLabel2 = new System.Windows.Forms.Label();
            this.metadataRulesAddGroup = new System.Windows.Forms.Button();
            this.metadataRulesAddRule = new System.Windows.Forms.Button();
            this.folderRulesTabPage = new System.Windows.Forms.TabPage();
            this.excludedFoldersList = new System.Windows.Forms.ListView();
            this.folderRulesActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.addExcludedFolder = new System.Windows.Forms.Button();
            this.removeExcludedFolder = new System.Windows.Forms.Button();
            this.excludedFolderLabel = new System.Windows.Forms.Label();
            this.folderStructurePage = new System.Windows.Forms.Panel();
            this.folderInsertControlsPanel = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.folderStructureActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.folderSpaceAutomatically = new System.Windows.Forms.CheckBox();
            this.insertFolderSeparator = new System.Windows.Forms.Button();
            this.folderStructurePreviewLayout = new System.Windows.Forms.TableLayoutPanel();
            this.folderPreviewLabel = new System.Windows.Forms.Label();
            this.folderPreview = new System.Windows.Forms.Label();
            this.folderPreviewPrevious = new System.Windows.Forms.Button();
            this.folderPreviewNext = new System.Windows.Forms.Button();
            this.folderStructurePanel = new System.Windows.Forms.Panel();
            this.folderStructureLabel = new System.Windows.Forms.Label();
            this.folderStructure = new System.Windows.Forms.TextBox();
            this.fileStructurePage = new System.Windows.Forms.Panel();
            this.fileStructureInsertControlsPanel = new System.Windows.Forms.Panel();
            this.insertControlsTabPanel = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.fileSpaceAutomatically = new System.Windows.Forms.CheckBox();
            this.fileStructurePreviewLayout = new System.Windows.Forms.TableLayoutPanel();
            this.fileStructurePreviewLabel = new System.Windows.Forms.Label();
            this.fileStructurePreview = new System.Windows.Forms.Label();
            this.fileStructurePreviewPrevious = new System.Windows.Forms.Button();
            this.fileStructurePreviewNext = new System.Windows.Forms.Button();
            this.fileStructurePanel = new System.Windows.Forms.Panel();
            this.fileStructureLabel = new System.Windows.Forms.Label();
            this.fileStructureTextBox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.formActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.okayButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.configurationPanel.SuspendLayout();
            this.optionsPage.SuspendLayout();
            this.optionsTabPage.SuspendLayout();
            this.removeEmptyFolderExclusionsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.removeEmptyFoldersExclusionsBindingSource)).BeginInit();
            this.removeEmptyFolderExclusionsActionPanel.SuspendLayout();
            this.illegalCharacterReplacementsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.illegalCharacterReplacementsBindingSource)).BeginInit();
            this.monthReplacementsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.monthReplacementsBindingSource)).BeginInit();
            this.emptyValuesTabPage.SuspendLayout();
            this.failOperationOnEmptyValueDestinationFolderLayout.SuspendLayout();
            this.failOperationOnEmptyValueFieldsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.failOperationOnEmptyValueFields)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.failOperationOnEmptyValueFieldsBindingSource)).BeginInit();
            this.emptyFieldReplacementLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.emptyFieldReplacementsBindingSource)).BeginInit();
            this.emptyFolderNameReplacementLayout.SuspendLayout();
            this.rulesPage.SuspendLayout();
            this.metadataRulesTabPage.SuspendLayout();
            this.metadataRulesControlsLayout.SuspendLayout();
            this.folderRulesTabPage.SuspendLayout();
            this.folderRulesActionsLayout.SuspendLayout();
            this.folderStructurePage.SuspendLayout();
            this.folderInsertControlsPanel.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.folderStructureActionsLayout.SuspendLayout();
            this.folderStructurePreviewLayout.SuspendLayout();
            this.folderStructurePanel.SuspendLayout();
            this.fileStructurePage.SuspendLayout();
            this.fileStructureInsertControlsPanel.SuspendLayout();
            this.insertControlsTabPanel.SuspendLayout();
            this.fileStructurePreviewLayout.SuspendLayout();
            this.fileStructurePanel.SuspendLayout();
            this.formActionsLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // configurationPanel
            // 
            this.configurationPanel.AutoSize = true;
            this.configurationPanel.Controls.Add(this.optionsPage);
            this.configurationPanel.Controls.Add(this.rulesPage);
            this.configurationPanel.Controls.Add(this.folderStructurePage);
            this.configurationPanel.Controls.Add(this.fileStructurePage);
            this.configurationPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.configurationPanel.Location = new System.Drawing.Point(208, 0);
            this.configurationPanel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.configurationPanel.Name = "configurationPanel";
            this.configurationPanel.Padding = new System.Windows.Forms.Padding(0, 12, 13, 4);
            this.configurationPanel.Size = new System.Drawing.Size(649, 630);
            this.configurationPanel.TabIndex = 0;
            // 
            // optionsPage
            // 
            this.optionsPage.Controls.Add(this.optionsTabPage);
            this.optionsPage.Controls.Add(this.emptyValuesTabPage);
            this.optionsPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.optionsPage.Location = new System.Drawing.Point(0, 12);
            this.optionsPage.Margin = new System.Windows.Forms.Padding(4);
            this.optionsPage.Name = "optionsPage";
            this.optionsPage.SelectedIndex = 0;
            this.optionsPage.Size = new System.Drawing.Size(636, 614);
            this.optionsPage.TabIndex = 10;
            // 
            // optionsTabPage
            // 
            this.optionsTabPage.Controls.Add(this.removeEmptyFolderExclusionsLayout);
            this.optionsTabPage.Controls.Add(this.removeEmptyFoldersLabel);
            this.optionsTabPage.Controls.Add(this.removeEmptyFolders);
            this.optionsTabPage.Controls.Add(this.illegalCharacterReplacementsLayout);
            this.optionsTabPage.Controls.Add(this.monthReplacementsLayout);
            this.optionsTabPage.Controls.Add(this.autoSelectSingleMultiValueField);
            this.optionsTabPage.Controls.Add(this.copyReadPercentageToReplacement);
            this.optionsTabPage.Controls.Add(this.normalizeMultipleSpaces);
            this.optionsTabPage.Location = new System.Drawing.Point(4, 25);
            this.optionsTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.optionsTabPage.Name = "optionsTabPage";
            this.optionsTabPage.Padding = new System.Windows.Forms.Padding(9);
            this.optionsTabPage.Size = new System.Drawing.Size(628, 585);
            this.optionsTabPage.TabIndex = 0;
            this.optionsTabPage.Text = "Options";
            this.optionsTabPage.UseVisualStyleBackColor = true;
            // 
            // removeEmptyFolderExclusionsLayout
            // 
            this.removeEmptyFolderExclusionsLayout.Controls.Add(this.removeEmptyFolderExclusions);
            this.removeEmptyFolderExclusionsLayout.Controls.Add(this.removeEmptyFolderExclusionsActionPanel);
            this.removeEmptyFolderExclusionsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.removeEmptyFolderExclusionsLayout.Location = new System.Drawing.Point(9, 203);
            this.removeEmptyFolderExclusionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolderExclusionsLayout.Name = "removeEmptyFolderExclusionsLayout";
            this.removeEmptyFolderExclusionsLayout.Size = new System.Drawing.Size(610, 373);
            this.removeEmptyFolderExclusionsLayout.TabIndex = 6;
            // 
            // removeEmptyFolderExclusions
            // 
            this.removeEmptyFolderExclusions.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileBindingSource, "RemoveEmptyFolders", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.removeEmptyFolderExclusions.DataSource = this.removeEmptyFoldersExclusionsBindingSource;
            this.removeEmptyFolderExclusions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.removeEmptyFolderExclusions.FormattingEnabled = true;
            this.removeEmptyFolderExclusions.ItemHeight = 16;
            this.removeEmptyFolderExclusions.Location = new System.Drawing.Point(0, 0);
            this.removeEmptyFolderExclusions.Name = "removeEmptyFolderExclusions";
            this.removeEmptyFolderExclusions.Size = new System.Drawing.Size(502, 373);
            this.removeEmptyFolderExclusions.TabIndex = 2;
            this.removeEmptyFolderExclusions.EnabledChanged += new System.EventHandler(this.removeEmptyFolderExclusions_EnabledChanged);
            // 
            // profileBindingSource
            // 
            this.profileBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ProfileViewModel);
            // 
            // removeEmptyFoldersExclusionsBindingSource
            // 
            this.removeEmptyFoldersExclusionsBindingSource.AllowNew = true;
            this.removeEmptyFoldersExclusionsBindingSource.DataMember = "RemoveEmptyFoldersExclusions";
            this.removeEmptyFoldersExclusionsBindingSource.DataSource = this.profileBindingSource;
            // 
            // removeEmptyFolderExclusionsActionPanel
            // 
            this.removeEmptyFolderExclusionsActionPanel.AutoSize = true;
            this.removeEmptyFolderExclusionsActionPanel.Controls.Add(this.addEmptyFolderExclusion);
            this.removeEmptyFolderExclusionsActionPanel.Controls.Add(this.removeEmptyFolderExclusion);
            this.removeEmptyFolderExclusionsActionPanel.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileBindingSource, "RemoveEmptyFolders", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.removeEmptyFolderExclusionsActionPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.removeEmptyFolderExclusionsActionPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.removeEmptyFolderExclusionsActionPanel.Location = new System.Drawing.Point(502, 0);
            this.removeEmptyFolderExclusionsActionPanel.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolderExclusionsActionPanel.Name = "removeEmptyFolderExclusionsActionPanel";
            this.removeEmptyFolderExclusionsActionPanel.Size = new System.Drawing.Size(108, 373);
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
            this.removeEmptyFoldersLabel.Location = new System.Drawing.Point(9, 184);
            this.removeEmptyFoldersLabel.Margin = new System.Windows.Forms.Padding(0);
            this.removeEmptyFoldersLabel.Name = "removeEmptyFoldersLabel";
            this.removeEmptyFoldersLabel.Padding = new System.Windows.Forms.Padding(3, 0, 0, 3);
            this.removeEmptyFoldersLabel.Size = new System.Drawing.Size(241, 19);
            this.removeEmptyFoldersLabel.TabIndex = 5;
            this.removeEmptyFoldersLabel.Text = "But do not remove the following folders:";
            // 
            // removeEmptyFolders
            // 
            this.removeEmptyFolders.AutoSize = true;
            this.removeEmptyFolders.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "RemoveEmptyFolders", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.removeEmptyFolders.Dock = System.Windows.Forms.DockStyle.Top;
            this.removeEmptyFolders.Location = new System.Drawing.Point(9, 158);
            this.removeEmptyFolders.Margin = new System.Windows.Forms.Padding(4);
            this.removeEmptyFolders.Name = "removeEmptyFolders";
            this.removeEmptyFolders.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.removeEmptyFolders.Size = new System.Drawing.Size(610, 26);
            this.removeEmptyFolders.TabIndex = 4;
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
            this.illegalCharacterReplacementsLayout.Location = new System.Drawing.Point(9, 120);
            this.illegalCharacterReplacementsLayout.Name = "illegalCharacterReplacementsLayout";
            this.illegalCharacterReplacementsLayout.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.illegalCharacterReplacementsLayout.Size = new System.Drawing.Size(610, 38);
            this.illegalCharacterReplacementsLayout.TabIndex = 0;
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
            this.illegalCharacterReplacementsBindingSource.DataSource = this.profileBindingSource;
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
            this.illegalCharacterReplacementsReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.illegalCharacterReplacementsBindingSource, "Replacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
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
            // monthReplacementsLayout
            // 
            this.monthReplacementsLayout.AutoSize = true;
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsMonth);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsMonthSelector);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsWith);
            this.monthReplacementsLayout.Controls.Add(this.monthReplacementsReplacement);
            this.monthReplacementsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.monthReplacementsLayout.Location = new System.Drawing.Point(9, 82);
            this.monthReplacementsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.monthReplacementsLayout.Name = "monthReplacementsLayout";
            this.monthReplacementsLayout.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.monthReplacementsLayout.Size = new System.Drawing.Size(610, 38);
            this.monthReplacementsLayout.TabIndex = 3;
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
            // monthReplacementsBindingSource
            // 
            this.monthReplacementsBindingSource.DataMember = "MonthReplacements";
            this.monthReplacementsBindingSource.DataSource = this.profileBindingSource;
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
            this.monthReplacementsReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.monthReplacementsBindingSource, "Replacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.monthReplacementsReplacement.Location = new System.Drawing.Point(147, 11);
            this.monthReplacementsReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.monthReplacementsReplacement.Name = "monthReplacementsReplacement";
            this.monthReplacementsReplacement.Size = new System.Drawing.Size(192, 22);
            this.monthReplacementsReplacement.TabIndex = 3;
            // 
            // autoSelectSingleMultiValueField
            // 
            this.autoSelectSingleMultiValueField.AutoSize = true;
            this.autoSelectSingleMultiValueField.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "AutoSelectSingleMultiValueField", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.autoSelectSingleMultiValueField.Dock = System.Windows.Forms.DockStyle.Top;
            this.autoSelectSingleMultiValueField.Location = new System.Drawing.Point(9, 56);
            this.autoSelectSingleMultiValueField.Margin = new System.Windows.Forms.Padding(4);
            this.autoSelectSingleMultiValueField.Name = "autoSelectSingleMultiValueField";
            this.autoSelectSingleMultiValueField.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.autoSelectSingleMultiValueField.Size = new System.Drawing.Size(610, 26);
            this.autoSelectSingleMultiValueField.TabIndex = 2;
            this.autoSelectSingleMultiValueField.Text = "If there is only one value in a multiple value field then insert it without askin" +
    "g";
            this.autoSelectSingleMultiValueField.UseVisualStyleBackColor = true;
            // 
            // copyReadPercentageToReplacement
            // 
            this.copyReadPercentageToReplacement.AutoSize = true;
            this.copyReadPercentageToReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "CopyReadPercentageToReplacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.copyReadPercentageToReplacement.Dock = System.Windows.Forms.DockStyle.Top;
            this.copyReadPercentageToReplacement.Location = new System.Drawing.Point(9, 30);
            this.copyReadPercentageToReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.copyReadPercentageToReplacement.Name = "copyReadPercentageToReplacement";
            this.copyReadPercentageToReplacement.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.copyReadPercentageToReplacement.Size = new System.Drawing.Size(610, 26);
            this.copyReadPercentageToReplacement.TabIndex = 1;
            this.copyReadPercentageToReplacement.Text = "When overwriting an existing file, copy the read percentage to the new file";
            this.copyReadPercentageToReplacement.UseVisualStyleBackColor = true;
            // 
            // normalizeMultipleSpaces
            // 
            this.normalizeMultipleSpaces.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "NormalizeMultipleSpaces", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.normalizeMultipleSpaces.Dock = System.Windows.Forms.DockStyle.Top;
            this.normalizeMultipleSpaces.Location = new System.Drawing.Point(9, 9);
            this.normalizeMultipleSpaces.Margin = new System.Windows.Forms.Padding(4);
            this.normalizeMultipleSpaces.Name = "normalizeMultipleSpaces";
            this.normalizeMultipleSpaces.Size = new System.Drawing.Size(610, 21);
            this.normalizeMultipleSpaces.TabIndex = 0;
            this.normalizeMultipleSpaces.Text = "Replace multiple spaces with a single space";
            this.normalizeMultipleSpaces.UseVisualStyleBackColor = true;
            // 
            // emptyValuesTabPage
            // 
            this.emptyValuesTabPage.Controls.Add(this.failOperationOnEmptyValueDestinationFolderLayout);
            this.emptyValuesTabPage.Controls.Add(this.failOperationOnEmptyValueUseDestinationFolder);
            this.emptyValuesTabPage.Controls.Add(this.failOperationOnEmptyValueFieldsLayout);
            this.emptyValuesTabPage.Controls.Add(this.failOperationOnEmptyValue);
            this.emptyValuesTabPage.Controls.Add(this.emptyFieldReplacementLayout);
            this.emptyValuesTabPage.Controls.Add(this.emptyFieldReplacementLabel);
            this.emptyValuesTabPage.Controls.Add(this.emptyFolderNameReplacementLabel2);
            this.emptyValuesTabPage.Controls.Add(this.emptyFolderNameReplacementLayout);
            this.emptyValuesTabPage.Location = new System.Drawing.Point(4, 25);
            this.emptyValuesTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.emptyValuesTabPage.Name = "emptyValuesTabPage";
            this.emptyValuesTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.emptyValuesTabPage.Size = new System.Drawing.Size(628, 585);
            this.emptyValuesTabPage.TabIndex = 1;
            this.emptyValuesTabPage.Text = "Empty Values";
            this.emptyValuesTabPage.UseVisualStyleBackColor = true;
            // 
            // failOperationOnEmptyValueDestinationFolderLayout
            // 
            this.failOperationOnEmptyValueDestinationFolderLayout.AutoSize = true;
            this.failOperationOnEmptyValueDestinationFolderLayout.ColumnCount = 2;
            this.failOperationOnEmptyValueDestinationFolderLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.failOperationOnEmptyValueDestinationFolderLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.failOperationOnEmptyValueDestinationFolderLayout.Controls.Add(this.failOperationOnEmptyValueDestinationFolderBrowse, 1, 0);
            this.failOperationOnEmptyValueDestinationFolderLayout.Controls.Add(this.failOperationOnEmptyValueDestinationFolder, 0, 0);
            this.failOperationOnEmptyValueDestinationFolderLayout.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileBindingSource, "FailOperationOnEmptyValueDestinationFolderEnabled", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValueDestinationFolderLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.failOperationOnEmptyValueDestinationFolderLayout.Location = new System.Drawing.Point(4, 308);
            this.failOperationOnEmptyValueDestinationFolderLayout.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValueDestinationFolderLayout.Name = "failOperationOnEmptyValueDestinationFolderLayout";
            this.failOperationOnEmptyValueDestinationFolderLayout.Padding = new System.Windows.Forms.Padding(16, 0, 0, 0);
            this.failOperationOnEmptyValueDestinationFolderLayout.RowCount = 1;
            this.failOperationOnEmptyValueDestinationFolderLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.failOperationOnEmptyValueDestinationFolderLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.failOperationOnEmptyValueDestinationFolderLayout.Size = new System.Drawing.Size(620, 34);
            this.failOperationOnEmptyValueDestinationFolderLayout.TabIndex = 2;
            // 
            // failOperationOnEmptyValueDestinationFolderBrowse
            // 
            this.failOperationOnEmptyValueDestinationFolderBrowse.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.failOperationOnEmptyValueDestinationFolderBrowse.AutoSize = true;
            this.failOperationOnEmptyValueDestinationFolderBrowse.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.failOperationOnEmptyValueDestinationFolderBrowse.Location = new System.Drawing.Point(554, 4);
            this.failOperationOnEmptyValueDestinationFolderBrowse.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValueDestinationFolderBrowse.Name = "failOperationOnEmptyValueDestinationFolderBrowse";
            this.failOperationOnEmptyValueDestinationFolderBrowse.Size = new System.Drawing.Size(62, 26);
            this.failOperationOnEmptyValueDestinationFolderBrowse.TabIndex = 0;
            this.failOperationOnEmptyValueDestinationFolderBrowse.Text = "Browse";
            this.failOperationOnEmptyValueDestinationFolderBrowse.UseVisualStyleBackColor = true;
            this.failOperationOnEmptyValueDestinationFolderBrowse.Click += new System.EventHandler(this.failOperationOnEmptyValueDestinationFolderBrowse_Click);
            // 
            // failOperationOnEmptyValueDestinationFolder
            // 
            this.failOperationOnEmptyValueDestinationFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.failOperationOnEmptyValueDestinationFolder.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.profileBindingSource, "FailOperationOnEmptyValueDestinationFolder", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValueDestinationFolder.Location = new System.Drawing.Point(20, 6);
            this.failOperationOnEmptyValueDestinationFolder.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValueDestinationFolder.Name = "failOperationOnEmptyValueDestinationFolder";
            this.failOperationOnEmptyValueDestinationFolder.ReadOnly = true;
            this.failOperationOnEmptyValueDestinationFolder.Size = new System.Drawing.Size(526, 22);
            this.failOperationOnEmptyValueDestinationFolder.TabIndex = 1;
            // 
            // failOperationOnEmptyValueUseDestinationFolder
            // 
            this.failOperationOnEmptyValueUseDestinationFolder.AutoSize = true;
            this.failOperationOnEmptyValueUseDestinationFolder.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "FailOperationOnEmptyValuesUseDestinationFolder", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValueUseDestinationFolder.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileBindingSource, "FailOperationOnEmptyValue", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValueUseDestinationFolder.Dock = System.Windows.Forms.DockStyle.Top;
            this.failOperationOnEmptyValueUseDestinationFolder.Location = new System.Drawing.Point(4, 284);
            this.failOperationOnEmptyValueUseDestinationFolder.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValueUseDestinationFolder.Name = "failOperationOnEmptyValueUseDestinationFolder";
            this.failOperationOnEmptyValueUseDestinationFolder.Padding = new System.Windows.Forms.Padding(20, 4, 0, 0);
            this.failOperationOnEmptyValueUseDestinationFolder.Size = new System.Drawing.Size(620, 24);
            this.failOperationOnEmptyValueUseDestinationFolder.TabIndex = 0;
            this.failOperationOnEmptyValueUseDestinationFolder.Text = "and move/copy them to this folder:";
            this.failOperationOnEmptyValueUseDestinationFolder.UseVisualStyleBackColor = true;
            // 
            // failOperationOnEmptyValueFieldsLayout
            // 
            this.failOperationOnEmptyValueFieldsLayout.Controls.Add(this.failOperationOnEmptyValueFields);
            this.failOperationOnEmptyValueFieldsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.failOperationOnEmptyValueFieldsLayout.Location = new System.Drawing.Point(4, 161);
            this.failOperationOnEmptyValueFieldsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValueFieldsLayout.Name = "failOperationOnEmptyValueFieldsLayout";
            this.failOperationOnEmptyValueFieldsLayout.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.failOperationOnEmptyValueFieldsLayout.Size = new System.Drawing.Size(620, 123);
            this.failOperationOnEmptyValueFieldsLayout.TabIndex = 3;
            // 
            // failOperationOnEmptyValueFields
            // 
            this.failOperationOnEmptyValueFields.AllowUserToAddRows = false;
            this.failOperationOnEmptyValueFields.AllowUserToDeleteRows = false;
            this.failOperationOnEmptyValueFields.AllowUserToResizeColumns = false;
            this.failOperationOnEmptyValueFields.AllowUserToResizeRows = false;
            this.failOperationOnEmptyValueFields.AutoGenerateColumns = false;
            this.failOperationOnEmptyValueFields.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.failOperationOnEmptyValueFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.failOperationOnEmptyValueFields.ColumnHeadersVisible = false;
            this.failOperationOnEmptyValueFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.emptyValueFieldEnabledColumn,
            this.emptyValueFieldNameColumn});
            this.failOperationOnEmptyValueFields.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.profileBindingSource, "FailOperationOnEmptyValue", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValueFields.DataSource = this.failOperationOnEmptyValueFieldsBindingSource;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.failOperationOnEmptyValueFields.DefaultCellStyle = dataGridViewCellStyle1;
            this.failOperationOnEmptyValueFields.Dock = System.Windows.Forms.DockStyle.Left;
            this.failOperationOnEmptyValueFields.GridColor = System.Drawing.SystemColors.Window;
            this.failOperationOnEmptyValueFields.Location = new System.Drawing.Point(20, 0);
            this.failOperationOnEmptyValueFields.MultiSelect = false;
            this.failOperationOnEmptyValueFields.Name = "failOperationOnEmptyValueFields";
            this.failOperationOnEmptyValueFields.RowHeadersVisible = false;
            this.failOperationOnEmptyValueFields.RowHeadersWidth = 51;
            this.failOperationOnEmptyValueFields.RowTemplate.Height = 24;
            this.failOperationOnEmptyValueFields.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.failOperationOnEmptyValueFields.ShowEditingIcon = false;
            this.failOperationOnEmptyValueFields.Size = new System.Drawing.Size(221, 123);
            this.failOperationOnEmptyValueFields.TabIndex = 3;
            this.failOperationOnEmptyValueFields.EnabledChanged += new System.EventHandler(this.failOperationOnEmptyValueFields_EnabledChanged);
            // 
            // emptyValueFieldEnabledColumn
            // 
            this.emptyValueFieldEnabledColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.emptyValueFieldEnabledColumn.DataPropertyName = "Enabled";
            this.emptyValueFieldEnabledColumn.HeaderText = "Enabled";
            this.emptyValueFieldEnabledColumn.MinimumWidth = 6;
            this.emptyValueFieldEnabledColumn.Name = "emptyValueFieldEnabledColumn";
            this.emptyValueFieldEnabledColumn.Width = 6;
            // 
            // emptyValueFieldNameColumn
            // 
            this.emptyValueFieldNameColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.emptyValueFieldNameColumn.DataPropertyName = "Name";
            this.emptyValueFieldNameColumn.HeaderText = "Name";
            this.emptyValueFieldNameColumn.MinimumWidth = 6;
            this.emptyValueFieldNameColumn.Name = "emptyValueFieldNameColumn";
            this.emptyValueFieldNameColumn.ReadOnly = true;
            // 
            // failOperationOnEmptyValueFieldsBindingSource
            // 
            this.failOperationOnEmptyValueFieldsBindingSource.DataMember = "FailOperationOnEmptyValueFields";
            this.failOperationOnEmptyValueFieldsBindingSource.DataSource = this.profileBindingSource;
            // 
            // failOperationOnEmptyValue
            // 
            this.failOperationOnEmptyValue.AutoSize = true;
            this.failOperationOnEmptyValue.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.profileBindingSource, "FailOperationOnEmptyValue", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.failOperationOnEmptyValue.Dock = System.Windows.Forms.DockStyle.Top;
            this.failOperationOnEmptyValue.Location = new System.Drawing.Point(4, 116);
            this.failOperationOnEmptyValue.Margin = new System.Windows.Forms.Padding(4);
            this.failOperationOnEmptyValue.Name = "failOperationOnEmptyValue";
            this.failOperationOnEmptyValue.Padding = new System.Windows.Forms.Padding(8, 25, 0, 0);
            this.failOperationOnEmptyValue.Size = new System.Drawing.Size(620, 45);
            this.failOperationOnEmptyValue.TabIndex = 2;
            this.failOperationOnEmptyValue.Text = "If any of the selected fields are empty then mark the operation as failed";
            this.failOperationOnEmptyValue.UseVisualStyleBackColor = true;
            // 
            // emptyFieldReplacementLayout
            // 
            this.emptyFieldReplacementLayout.AutoSize = true;
            this.emptyFieldReplacementLayout.ColumnCount = 4;
            this.emptyFieldReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.emptyFieldReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.emptyFieldReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.emptyFieldReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.emptyFieldReplacementLayout.Controls.Add(this.emptyFieldReplacementLabel2, 0, 0);
            this.emptyFieldReplacementLayout.Controls.Add(this.emptyFieldReplacementSelector, 1, 0);
            this.emptyFieldReplacementLayout.Controls.Add(this.emptyFieldReplacementLabel3, 2, 0);
            this.emptyFieldReplacementLayout.Controls.Add(this.emptyFieldReplacement, 3, 0);
            this.emptyFieldReplacementLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.emptyFieldReplacementLayout.Location = new System.Drawing.Point(4, 84);
            this.emptyFieldReplacementLayout.Margin = new System.Windows.Forms.Padding(4);
            this.emptyFieldReplacementLayout.Name = "emptyFieldReplacementLayout";
            this.emptyFieldReplacementLayout.RowCount = 1;
            this.emptyFieldReplacementLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.emptyFieldReplacementLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
            this.emptyFieldReplacementLayout.Size = new System.Drawing.Size(620, 32);
            this.emptyFieldReplacementLayout.TabIndex = 2;
            // 
            // emptyFieldReplacementLabel2
            // 
            this.emptyFieldReplacementLabel2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.emptyFieldReplacementLabel2.AutoSize = true;
            this.emptyFieldReplacementLabel2.Location = new System.Drawing.Point(4, 8);
            this.emptyFieldReplacementLabel2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emptyFieldReplacementLabel2.Name = "emptyFieldReplacementLabel2";
            this.emptyFieldReplacementLabel2.Size = new System.Drawing.Size(37, 16);
            this.emptyFieldReplacementLabel2.TabIndex = 1;
            this.emptyFieldReplacementLabel2.Text = "Field";
            // 
            // emptyFieldReplacementSelector
            // 
            this.emptyFieldReplacementSelector.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.emptyFieldReplacementSelector.DataSource = this.emptyFieldReplacementsBindingSource;
            this.emptyFieldReplacementSelector.DisplayMember = "Field";
            this.emptyFieldReplacementSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.emptyFieldReplacementSelector.FormattingEnabled = true;
            this.emptyFieldReplacementSelector.Location = new System.Drawing.Point(49, 4);
            this.emptyFieldReplacementSelector.Margin = new System.Windows.Forms.Padding(4);
            this.emptyFieldReplacementSelector.Name = "emptyFieldReplacementSelector";
            this.emptyFieldReplacementSelector.Size = new System.Drawing.Size(161, 24);
            this.emptyFieldReplacementSelector.TabIndex = 11;
            this.emptyFieldReplacementSelector.ValueMember = "Field";
            // 
            // emptyFieldReplacementsBindingSource
            // 
            this.emptyFieldReplacementsBindingSource.DataMember = "EmptyFieldReplacements";
            this.emptyFieldReplacementsBindingSource.DataSource = this.profileBindingSource;
            // 
            // emptyFieldReplacementLabel3
            // 
            this.emptyFieldReplacementLabel3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.emptyFieldReplacementLabel3.AutoSize = true;
            this.emptyFieldReplacementLabel3.Location = new System.Drawing.Point(218, 8);
            this.emptyFieldReplacementLabel3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emptyFieldReplacementLabel3.Name = "emptyFieldReplacementLabel3";
            this.emptyFieldReplacementLabel3.Size = new System.Drawing.Size(73, 16);
            this.emptyFieldReplacementLabel3.TabIndex = 12;
            this.emptyFieldReplacementLabel3.Text = "substitution";
            // 
            // emptyFieldReplacement
            // 
            this.emptyFieldReplacement.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.emptyFieldReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.emptyFieldReplacementsBindingSource, "Replacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.emptyFieldReplacement.Location = new System.Drawing.Point(299, 5);
            this.emptyFieldReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.emptyFieldReplacement.Name = "emptyFieldReplacement";
            this.emptyFieldReplacement.Size = new System.Drawing.Size(317, 22);
            this.emptyFieldReplacement.TabIndex = 13;
            // 
            // emptyFieldReplacementLabel
            // 
            this.emptyFieldReplacementLabel.AutoSize = true;
            this.emptyFieldReplacementLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.emptyFieldReplacementLabel.Location = new System.Drawing.Point(4, 50);
            this.emptyFieldReplacementLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emptyFieldReplacementLabel.Name = "emptyFieldReplacementLabel";
            this.emptyFieldReplacementLabel.Padding = new System.Windows.Forms.Padding(4, 18, 0, 0);
            this.emptyFieldReplacementLabel.Size = new System.Drawing.Size(291, 34);
            this.emptyFieldReplacementLabel.TabIndex = 0;
            this.emptyFieldReplacementLabel.Text = "When a field is empty substitue the follow value:";
            // 
            // emptyFolderNameReplacementLabel2
            // 
            this.emptyFolderNameReplacementLabel2.AutoSize = true;
            this.emptyFolderNameReplacementLabel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.emptyFolderNameReplacementLabel2.Location = new System.Drawing.Point(4, 34);
            this.emptyFolderNameReplacementLabel2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emptyFolderNameReplacementLabel2.Name = "emptyFolderNameReplacementLabel2";
            this.emptyFolderNameReplacementLabel2.Padding = new System.Windows.Forms.Padding(4, 0, 0, 0);
            this.emptyFolderNameReplacementLabel2.Size = new System.Drawing.Size(236, 16);
            this.emptyFolderNameReplacementLabel2.TabIndex = 2;
            this.emptyFolderNameReplacementLabel2.Text = "Leave empty to remove empty folders";
            // 
            // emptyFolderNameReplacementLayout
            // 
            this.emptyFolderNameReplacementLayout.AutoSize = true;
            this.emptyFolderNameReplacementLayout.ColumnCount = 2;
            this.emptyFolderNameReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.emptyFolderNameReplacementLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.emptyFolderNameReplacementLayout.Controls.Add(this.emptyFolderNameReplacement, 1, 0);
            this.emptyFolderNameReplacementLayout.Controls.Add(this.emptyFolderNameReplacementLabel, 0, 0);
            this.emptyFolderNameReplacementLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.emptyFolderNameReplacementLayout.Location = new System.Drawing.Point(4, 4);
            this.emptyFolderNameReplacementLayout.Margin = new System.Windows.Forms.Padding(4);
            this.emptyFolderNameReplacementLayout.Name = "emptyFolderNameReplacementLayout";
            this.emptyFolderNameReplacementLayout.RowCount = 1;
            this.emptyFolderNameReplacementLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.emptyFolderNameReplacementLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.emptyFolderNameReplacementLayout.Size = new System.Drawing.Size(620, 30);
            this.emptyFolderNameReplacementLayout.TabIndex = 2;
            // 
            // emptyFolderNameReplacement
            // 
            this.emptyFolderNameReplacement.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.emptyFolderNameReplacement.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.profileBindingSource, "EmptyFolderNameReplacement", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.emptyFolderNameReplacement.Location = new System.Drawing.Point(220, 4);
            this.emptyFolderNameReplacement.Margin = new System.Windows.Forms.Padding(4);
            this.emptyFolderNameReplacement.Name = "emptyFolderNameReplacement";
            this.emptyFolderNameReplacement.Size = new System.Drawing.Size(396, 22);
            this.emptyFolderNameReplacement.TabIndex = 1;
            // 
            // emptyFolderNameReplacementLabel
            // 
            this.emptyFolderNameReplacementLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.emptyFolderNameReplacementLabel.AutoSize = true;
            this.emptyFolderNameReplacementLabel.Location = new System.Drawing.Point(4, 7);
            this.emptyFolderNameReplacementLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emptyFolderNameReplacementLabel.Name = "emptyFolderNameReplacementLabel";
            this.emptyFolderNameReplacementLabel.Size = new System.Drawing.Size(208, 16);
            this.emptyFolderNameReplacementLabel.TabIndex = 0;
            this.emptyFolderNameReplacementLabel.Text = "Replace empty folder names with:";
            // 
            // rulesPage
            // 
            this.rulesPage.Controls.Add(this.metadataRulesTabPage);
            this.rulesPage.Controls.Add(this.folderRulesTabPage);
            this.rulesPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rulesPage.Location = new System.Drawing.Point(0, 12);
            this.rulesPage.Margin = new System.Windows.Forms.Padding(4);
            this.rulesPage.Name = "rulesPage";
            this.rulesPage.SelectedIndex = 0;
            this.rulesPage.Size = new System.Drawing.Size(636, 614);
            this.rulesPage.TabIndex = 9;
            // 
            // metadataRulesTabPage
            // 
            this.metadataRulesTabPage.Controls.Add(this.metadataRulesActionsLayout);
            this.metadataRulesTabPage.Controls.Add(this.metadataRulesControlsLayout);
            this.metadataRulesTabPage.Location = new System.Drawing.Point(4, 25);
            this.metadataRulesTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesTabPage.Name = "metadataRulesTabPage";
            this.metadataRulesTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.metadataRulesTabPage.Size = new System.Drawing.Size(628, 585);
            this.metadataRulesTabPage.TabIndex = 0;
            this.metadataRulesTabPage.Text = "Metadata Rules";
            this.metadataRulesTabPage.UseVisualStyleBackColor = true;
            // 
            // metadataRulesActionsLayout
            // 
            this.metadataRulesActionsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metadataRulesActionsLayout.Location = new System.Drawing.Point(4, 76);
            this.metadataRulesActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesActionsLayout.Name = "metadataRulesActionsLayout";
            this.metadataRulesActionsLayout.Size = new System.Drawing.Size(620, 505);
            this.metadataRulesActionsLayout.TabIndex = 1;
            // 
            // metadataRulesControlsLayout
            // 
            this.metadataRulesControlsLayout.AutoSize = true;
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesAction);
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesActionLabel1);
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesMode);
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesActionLabel2);
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesAddGroup);
            this.metadataRulesControlsLayout.Controls.Add(this.metadataRulesAddRule);
            this.metadataRulesControlsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.metadataRulesControlsLayout.Location = new System.Drawing.Point(4, 4);
            this.metadataRulesControlsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesControlsLayout.Name = "metadataRulesControlsLayout";
            this.metadataRulesControlsLayout.Size = new System.Drawing.Size(620, 72);
            this.metadataRulesControlsLayout.TabIndex = 0;
            // 
            // metadataRulesAction
            // 
            this.metadataRulesAction.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesAction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.metadataRulesAction.FormattingEnabled = true;
            this.metadataRulesAction.Location = new System.Drawing.Point(4, 6);
            this.metadataRulesAction.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesAction.Name = "metadataRulesAction";
            this.metadataRulesAction.Size = new System.Drawing.Size(77, 24);
            this.metadataRulesAction.TabIndex = 0;
            // 
            // metadataRulesActionLabel1
            // 
            this.metadataRulesActionLabel1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesActionLabel1.AutoSize = true;
            this.metadataRulesActionLabel1.Location = new System.Drawing.Point(89, 10);
            this.metadataRulesActionLabel1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.metadataRulesActionLabel1.Name = "metadataRulesActionLabel1";
            this.metadataRulesActionLabel1.Size = new System.Drawing.Size(145, 16);
            this.metadataRulesActionLabel1.TabIndex = 1;
            this.metadataRulesActionLabel1.Text = "move books that match";
            // 
            // metadataRulesMode
            // 
            this.metadataRulesMode.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.metadataRulesMode.FormattingEnabled = true;
            this.metadataRulesMode.Location = new System.Drawing.Point(242, 6);
            this.metadataRulesMode.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesMode.Name = "metadataRulesMode";
            this.metadataRulesMode.Size = new System.Drawing.Size(60, 24);
            this.metadataRulesMode.TabIndex = 2;
            // 
            // metadataRulesActionLabel2
            // 
            this.metadataRulesActionLabel2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesActionLabel2.AutoSize = true;
            this.metadataRulesActionLabel2.Location = new System.Drawing.Point(310, 10);
            this.metadataRulesActionLabel2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.metadataRulesActionLabel2.Name = "metadataRulesActionLabel2";
            this.metadataRulesActionLabel2.Size = new System.Drawing.Size(126, 16);
            this.metadataRulesActionLabel2.TabIndex = 3;
            this.metadataRulesActionLabel2.Text = "of the following rules";
            // 
            // metadataRulesAddGroup
            // 
            this.metadataRulesAddGroup.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesAddGroup.Location = new System.Drawing.Point(444, 4);
            this.metadataRulesAddGroup.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesAddGroup.Name = "metadataRulesAddGroup";
            this.metadataRulesAddGroup.Size = new System.Drawing.Size(100, 28);
            this.metadataRulesAddGroup.TabIndex = 4;
            this.metadataRulesAddGroup.Text = "Add Group";
            this.metadataRulesAddGroup.UseVisualStyleBackColor = true;
            // 
            // metadataRulesAddRule
            // 
            this.metadataRulesAddRule.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.metadataRulesAddRule.Location = new System.Drawing.Point(4, 40);
            this.metadataRulesAddRule.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesAddRule.Name = "metadataRulesAddRule";
            this.metadataRulesAddRule.Size = new System.Drawing.Size(100, 28);
            this.metadataRulesAddRule.TabIndex = 5;
            this.metadataRulesAddRule.Text = "Add Rule";
            this.metadataRulesAddRule.UseVisualStyleBackColor = true;
            // 
            // folderRulesTabPage
            // 
            this.folderRulesTabPage.Controls.Add(this.excludedFoldersList);
            this.folderRulesTabPage.Controls.Add(this.folderRulesActionsLayout);
            this.folderRulesTabPage.Controls.Add(this.excludedFolderLabel);
            this.folderRulesTabPage.Location = new System.Drawing.Point(4, 25);
            this.folderRulesTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.folderRulesTabPage.Name = "folderRulesTabPage";
            this.folderRulesTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.folderRulesTabPage.Size = new System.Drawing.Size(628, 585);
            this.folderRulesTabPage.TabIndex = 1;
            this.folderRulesTabPage.Text = "Folder Rules";
            this.folderRulesTabPage.UseVisualStyleBackColor = true;
            // 
            // excludedFoldersList
            // 
            this.excludedFoldersList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.excludedFoldersList.HideSelection = false;
            this.excludedFoldersList.Location = new System.Drawing.Point(4, 32);
            this.excludedFoldersList.Margin = new System.Windows.Forms.Padding(4);
            this.excludedFoldersList.Name = "excludedFoldersList";
            this.excludedFoldersList.Size = new System.Drawing.Size(512, 549);
            this.excludedFoldersList.TabIndex = 1;
            this.excludedFoldersList.UseCompatibleStateImageBehavior = false;
            // 
            // folderRulesActionsLayout
            // 
            this.folderRulesActionsLayout.AutoSize = true;
            this.folderRulesActionsLayout.Controls.Add(this.addExcludedFolder);
            this.folderRulesActionsLayout.Controls.Add(this.removeExcludedFolder);
            this.folderRulesActionsLayout.Dock = System.Windows.Forms.DockStyle.Right;
            this.folderRulesActionsLayout.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.folderRulesActionsLayout.Location = new System.Drawing.Point(516, 32);
            this.folderRulesActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderRulesActionsLayout.Name = "folderRulesActionsLayout";
            this.folderRulesActionsLayout.Size = new System.Drawing.Size(108, 549);
            this.folderRulesActionsLayout.TabIndex = 2;
            // 
            // addExcludedFolder
            // 
            this.addExcludedFolder.Location = new System.Drawing.Point(4, 4);
            this.addExcludedFolder.Margin = new System.Windows.Forms.Padding(4);
            this.addExcludedFolder.Name = "addExcludedFolder";
            this.addExcludedFolder.Size = new System.Drawing.Size(100, 28);
            this.addExcludedFolder.TabIndex = 0;
            this.addExcludedFolder.Text = "Add";
            this.addExcludedFolder.UseVisualStyleBackColor = true;
            // 
            // removeExcludedFolder
            // 
            this.removeExcludedFolder.Location = new System.Drawing.Point(4, 40);
            this.removeExcludedFolder.Margin = new System.Windows.Forms.Padding(4);
            this.removeExcludedFolder.Name = "removeExcludedFolder";
            this.removeExcludedFolder.Size = new System.Drawing.Size(100, 28);
            this.removeExcludedFolder.TabIndex = 1;
            this.removeExcludedFolder.Text = "Remove";
            this.removeExcludedFolder.UseVisualStyleBackColor = true;
            // 
            // excludedFolderLabel
            // 
            this.excludedFolderLabel.AutoSize = true;
            this.excludedFolderLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.excludedFolderLabel.Location = new System.Drawing.Point(4, 4);
            this.excludedFolderLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.excludedFolderLabel.Name = "excludedFolderLabel";
            this.excludedFolderLabel.Padding = new System.Windows.Forms.Padding(0, 6, 0, 6);
            this.excludedFolderLabel.Size = new System.Drawing.Size(365, 28);
            this.excludedFolderLabel.TabIndex = 0;
            this.excludedFolderLabel.Text = "Do not move books if they are located in the following folders";
            // 
            // folderStructurePage
            // 
            this.folderStructurePage.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.folderStructurePage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.folderStructurePage.Controls.Add(this.folderInsertControlsPanel);
            this.folderStructurePage.Controls.Add(this.folderStructureActionsLayout);
            this.folderStructurePage.Controls.Add(this.folderStructurePreviewLayout);
            this.folderStructurePage.Controls.Add(this.folderStructurePanel);
            this.folderStructurePage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.folderStructurePage.Location = new System.Drawing.Point(0, 12);
            this.folderStructurePage.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePage.Name = "folderStructurePage";
            this.folderStructurePage.Size = new System.Drawing.Size(636, 614);
            this.folderStructurePage.TabIndex = 1;
            // 
            // folderInsertControlsPanel
            // 
            this.folderInsertControlsPanel.Controls.Add(this.tabControl1);
            this.folderInsertControlsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.folderInsertControlsPanel.Location = new System.Drawing.Point(0, 130);
            this.folderInsertControlsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.folderInsertControlsPanel.Name = "folderInsertControlsPanel";
            this.folderInsertControlsPanel.Size = new System.Drawing.Size(634, 482);
            this.folderInsertControlsPanel.TabIndex = 8;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(634, 482);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage3.Size = new System.Drawing.Size(626, 453);
            this.tabPage3.TabIndex = 0;
            this.tabPage3.Text = "tabPage3";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage4.Size = new System.Drawing.Size(626, 453);
            this.tabPage4.TabIndex = 1;
            this.tabPage4.Text = "tabPage4";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // folderStructureActionsLayout
            // 
            this.folderStructureActionsLayout.AutoSize = true;
            this.folderStructureActionsLayout.Controls.Add(this.folderSpaceAutomatically);
            this.folderStructureActionsLayout.Controls.Add(this.insertFolderSeparator);
            this.folderStructureActionsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.folderStructureActionsLayout.Location = new System.Drawing.Point(0, 90);
            this.folderStructureActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructureActionsLayout.Name = "folderStructureActionsLayout";
            this.folderStructureActionsLayout.Padding = new System.Windows.Forms.Padding(7, 0, 0, 0);
            this.folderStructureActionsLayout.Size = new System.Drawing.Size(634, 40);
            this.folderStructureActionsLayout.TabIndex = 5;
            // 
            // folderSpaceAutomatically
            // 
            this.folderSpaceAutomatically.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.folderSpaceAutomatically.AutoSize = true;
            this.folderSpaceAutomatically.Location = new System.Drawing.Point(11, 10);
            this.folderSpaceAutomatically.Margin = new System.Windows.Forms.Padding(4);
            this.folderSpaceAutomatically.Name = "folderSpaceAutomatically";
            this.folderSpaceAutomatically.Size = new System.Drawing.Size(237, 20);
            this.folderSpaceAutomatically.TabIndex = 7;
            this.folderSpaceAutomatically.Text = "Space inserted fields automatically";
            this.folderSpaceAutomatically.UseVisualStyleBackColor = true;
            // 
            // insertFolderSeparator
            // 
            this.insertFolderSeparator.AutoSize = true;
            this.insertFolderSeparator.Location = new System.Drawing.Point(256, 4);
            this.insertFolderSeparator.Margin = new System.Windows.Forms.Padding(4);
            this.insertFolderSeparator.Name = "insertFolderSeparator";
            this.insertFolderSeparator.Size = new System.Drawing.Size(159, 32);
            this.insertFolderSeparator.TabIndex = 8;
            this.insertFolderSeparator.Text = "Folder Seperator";
            this.insertFolderSeparator.UseVisualStyleBackColor = true;
            // 
            // folderStructurePreviewLayout
            // 
            this.folderStructurePreviewLayout.AutoSize = true;
            this.folderStructurePreviewLayout.ColumnCount = 4;
            this.folderStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.folderStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.folderStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.folderStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.folderStructurePreviewLayout.Controls.Add(this.folderPreviewLabel, 0, 0);
            this.folderStructurePreviewLayout.Controls.Add(this.folderPreview, 1, 0);
            this.folderStructurePreviewLayout.Controls.Add(this.folderPreviewPrevious, 2, 0);
            this.folderStructurePreviewLayout.Controls.Add(this.folderPreviewNext, 3, 0);
            this.folderStructurePreviewLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.folderStructurePreviewLayout.Location = new System.Drawing.Point(0, 56);
            this.folderStructurePreviewLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePreviewLayout.Name = "folderStructurePreviewLayout";
            this.folderStructurePreviewLayout.RowCount = 1;
            this.folderStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.folderStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.folderStructurePreviewLayout.Size = new System.Drawing.Size(634, 34);
            this.folderStructurePreviewLayout.TabIndex = 6;
            // 
            // folderPreviewLabel
            // 
            this.folderPreviewLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.folderPreviewLabel.AutoSize = true;
            this.folderPreviewLabel.Location = new System.Drawing.Point(4, 9);
            this.folderPreviewLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.folderPreviewLabel.Name = "folderPreviewLabel";
            this.folderPreviewLabel.Size = new System.Drawing.Size(55, 16);
            this.folderPreviewLabel.TabIndex = 2;
            this.folderPreviewLabel.Text = "Preview";
            // 
            // folderPreview
            // 
            this.folderPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.folderPreview.AutoSize = true;
            this.folderPreview.Location = new System.Drawing.Point(67, 9);
            this.folderPreview.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.folderPreview.Name = "folderPreview";
            this.folderPreview.Size = new System.Drawing.Size(499, 16);
            this.folderPreview.TabIndex = 1;
            this.folderPreview.Text = "label2";
            // 
            // folderPreviewPrevious
            // 
            this.folderPreviewPrevious.AutoSize = true;
            this.folderPreviewPrevious.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.folderPreviewPrevious.Location = new System.Drawing.Point(574, 4);
            this.folderPreviewPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.folderPreviewPrevious.Name = "folderPreviewPrevious";
            this.folderPreviewPrevious.Size = new System.Drawing.Size(24, 26);
            this.folderPreviewPrevious.TabIndex = 4;
            this.folderPreviewPrevious.Text = "<";
            this.folderPreviewPrevious.UseVisualStyleBackColor = true;
            // 
            // folderPreviewNext
            // 
            this.folderPreviewNext.AutoSize = true;
            this.folderPreviewNext.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.folderPreviewNext.Location = new System.Drawing.Point(606, 4);
            this.folderPreviewNext.Margin = new System.Windows.Forms.Padding(4);
            this.folderPreviewNext.Name = "folderPreviewNext";
            this.folderPreviewNext.Size = new System.Drawing.Size(24, 26);
            this.folderPreviewNext.TabIndex = 3;
            this.folderPreviewNext.Text = ">";
            this.folderPreviewNext.UseVisualStyleBackColor = true;
            // 
            // folderStructurePanel
            // 
            this.folderStructurePanel.AutoSize = true;
            this.folderStructurePanel.Controls.Add(this.folderStructureLabel);
            this.folderStructurePanel.Controls.Add(this.folderStructure);
            this.folderStructurePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.folderStructurePanel.Location = new System.Drawing.Point(0, 0);
            this.folderStructurePanel.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePanel.Name = "folderStructurePanel";
            this.folderStructurePanel.Size = new System.Drawing.Size(634, 56);
            this.folderStructurePanel.TabIndex = 4;
            // 
            // folderStructureLabel
            // 
            this.folderStructureLabel.AutoSize = true;
            this.folderStructureLabel.Location = new System.Drawing.Point(7, 21);
            this.folderStructureLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.folderStructureLabel.Name = "folderStructureLabel";
            this.folderStructureLabel.Size = new System.Drawing.Size(104, 16);
            this.folderStructureLabel.TabIndex = 0;
            this.folderStructureLabel.Text = "Folder Structure:";
            // 
            // folderStructure
            // 
            this.folderStructure.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.folderStructure.Location = new System.Drawing.Point(128, 4);
            this.folderStructure.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructure.Multiline = true;
            this.folderStructure.Name = "folderStructure";
            this.folderStructure.Size = new System.Drawing.Size(501, 48);
            this.folderStructure.TabIndex = 1;
            // 
            // fileStructurePage
            // 
            this.fileStructurePage.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.fileStructurePage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.fileStructurePage.Controls.Add(this.fileStructureInsertControlsPanel);
            this.fileStructurePage.Controls.Add(this.fileSpaceAutomatically);
            this.fileStructurePage.Controls.Add(this.fileStructurePreviewLayout);
            this.fileStructurePage.Controls.Add(this.fileStructurePanel);
            this.fileStructurePage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileStructurePage.Location = new System.Drawing.Point(0, 12);
            this.fileStructurePage.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePage.Name = "fileStructurePage";
            this.fileStructurePage.Size = new System.Drawing.Size(636, 614);
            this.fileStructurePage.TabIndex = 0;
            // 
            // fileStructureInsertControlsPanel
            // 
            this.fileStructureInsertControlsPanel.Controls.Add(this.insertControlsTabPanel);
            this.fileStructureInsertControlsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileStructureInsertControlsPanel.Location = new System.Drawing.Point(0, 118);
            this.fileStructureInsertControlsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructureInsertControlsPanel.Name = "fileStructureInsertControlsPanel";
            this.fileStructureInsertControlsPanel.Size = new System.Drawing.Size(634, 494);
            this.fileStructureInsertControlsPanel.TabIndex = 6;
            // 
            // insertControlsTabPanel
            // 
            this.insertControlsTabPanel.Controls.Add(this.tabPage1);
            this.insertControlsTabPanel.Controls.Add(this.tabPage2);
            this.insertControlsTabPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.insertControlsTabPanel.Location = new System.Drawing.Point(0, 0);
            this.insertControlsTabPanel.Margin = new System.Windows.Forms.Padding(4);
            this.insertControlsTabPanel.Name = "insertControlsTabPanel";
            this.insertControlsTabPanel.SelectedIndex = 0;
            this.insertControlsTabPanel.Size = new System.Drawing.Size(634, 494);
            this.insertControlsTabPanel.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(626, 465);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage2.Size = new System.Drawing.Size(626, 465);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // fileSpaceAutomatically
            // 
            this.fileSpaceAutomatically.AutoSize = true;
            this.fileSpaceAutomatically.Dock = System.Windows.Forms.DockStyle.Top;
            this.fileSpaceAutomatically.Location = new System.Drawing.Point(0, 90);
            this.fileSpaceAutomatically.Margin = new System.Windows.Forms.Padding(4);
            this.fileSpaceAutomatically.Name = "fileSpaceAutomatically";
            this.fileSpaceAutomatically.Padding = new System.Windows.Forms.Padding(7, 4, 0, 4);
            this.fileSpaceAutomatically.Size = new System.Drawing.Size(634, 28);
            this.fileSpaceAutomatically.TabIndex = 2;
            this.fileSpaceAutomatically.Text = "Space inserted fields automatically";
            this.fileSpaceAutomatically.UseVisualStyleBackColor = true;
            // 
            // fileStructurePreviewLayout
            // 
            this.fileStructurePreviewLayout.AutoSize = true;
            this.fileStructurePreviewLayout.ColumnCount = 4;
            this.fileStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.fileStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.fileStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.fileStructurePreviewLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.fileStructurePreviewLayout.Controls.Add(this.fileStructurePreviewLabel, 0, 0);
            this.fileStructurePreviewLayout.Controls.Add(this.fileStructurePreview, 1, 0);
            this.fileStructurePreviewLayout.Controls.Add(this.fileStructurePreviewPrevious, 2, 0);
            this.fileStructurePreviewLayout.Controls.Add(this.fileStructurePreviewNext, 3, 0);
            this.fileStructurePreviewLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.fileStructurePreviewLayout.Location = new System.Drawing.Point(0, 56);
            this.fileStructurePreviewLayout.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePreviewLayout.Name = "fileStructurePreviewLayout";
            this.fileStructurePreviewLayout.RowCount = 1;
            this.fileStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.fileStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.fileStructurePreviewLayout.Size = new System.Drawing.Size(634, 34);
            this.fileStructurePreviewLayout.TabIndex = 5;
            // 
            // fileStructurePreviewLabel
            // 
            this.fileStructurePreviewLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.fileStructurePreviewLabel.AutoSize = true;
            this.fileStructurePreviewLabel.Location = new System.Drawing.Point(4, 9);
            this.fileStructurePreviewLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.fileStructurePreviewLabel.Name = "fileStructurePreviewLabel";
            this.fileStructurePreviewLabel.Size = new System.Drawing.Size(55, 16);
            this.fileStructurePreviewLabel.TabIndex = 2;
            this.fileStructurePreviewLabel.Text = "Preview";
            // 
            // fileStructurePreview
            // 
            this.fileStructurePreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.fileStructurePreview.AutoSize = true;
            this.fileStructurePreview.Location = new System.Drawing.Point(67, 9);
            this.fileStructurePreview.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.fileStructurePreview.Name = "fileStructurePreview";
            this.fileStructurePreview.Size = new System.Drawing.Size(499, 16);
            this.fileStructurePreview.TabIndex = 1;
            this.fileStructurePreview.Text = "label2";
            // 
            // fileStructurePreviewPrevious
            // 
            this.fileStructurePreviewPrevious.AutoSize = true;
            this.fileStructurePreviewPrevious.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.fileStructurePreviewPrevious.Location = new System.Drawing.Point(574, 4);
            this.fileStructurePreviewPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePreviewPrevious.Name = "fileStructurePreviewPrevious";
            this.fileStructurePreviewPrevious.Size = new System.Drawing.Size(24, 26);
            this.fileStructurePreviewPrevious.TabIndex = 4;
            this.fileStructurePreviewPrevious.Text = "<";
            this.fileStructurePreviewPrevious.UseVisualStyleBackColor = true;
            // 
            // fileStructurePreviewNext
            // 
            this.fileStructurePreviewNext.AutoSize = true;
            this.fileStructurePreviewNext.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.fileStructurePreviewNext.Location = new System.Drawing.Point(606, 4);
            this.fileStructurePreviewNext.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePreviewNext.Name = "fileStructurePreviewNext";
            this.fileStructurePreviewNext.Size = new System.Drawing.Size(24, 26);
            this.fileStructurePreviewNext.TabIndex = 3;
            this.fileStructurePreviewNext.Text = ">";
            this.fileStructurePreviewNext.UseVisualStyleBackColor = true;
            // 
            // fileStructurePanel
            // 
            this.fileStructurePanel.AutoSize = true;
            this.fileStructurePanel.Controls.Add(this.fileStructureLabel);
            this.fileStructurePanel.Controls.Add(this.fileStructureTextBox);
            this.fileStructurePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.fileStructurePanel.Location = new System.Drawing.Point(0, 0);
            this.fileStructurePanel.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePanel.Name = "fileStructurePanel";
            this.fileStructurePanel.Size = new System.Drawing.Size(634, 56);
            this.fileStructurePanel.TabIndex = 3;
            // 
            // fileStructureLabel
            // 
            this.fileStructureLabel.AutoSize = true;
            this.fileStructureLabel.Location = new System.Drawing.Point(7, 21);
            this.fileStructureLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.fileStructureLabel.Name = "fileStructureLabel";
            this.fileStructureLabel.Size = new System.Drawing.Size(87, 16);
            this.fileStructureLabel.TabIndex = 0;
            this.fileStructureLabel.Text = "File Structure:";
            // 
            // fileStructureTextBox
            // 
            this.fileStructureTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fileStructureTextBox.Location = new System.Drawing.Point(111, 4);
            this.fileStructureTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructureTextBox.Multiline = true;
            this.fileStructureTextBox.Name = "fileStructureTextBox";
            this.fileStructureTextBox.Size = new System.Drawing.Size(518, 48);
            this.fileStructureTextBox.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(208, 666);
            this.panel2.TabIndex = 1;
            // 
            // formActionsLayout
            // 
            this.formActionsLayout.AutoSize = true;
            this.formActionsLayout.Controls.Add(this.okayButton);
            this.formActionsLayout.Controls.Add(this.cancelButton);
            this.formActionsLayout.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.formActionsLayout.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.formActionsLayout.Location = new System.Drawing.Point(208, 630);
            this.formActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.formActionsLayout.Name = "formActionsLayout";
            this.formActionsLayout.Size = new System.Drawing.Size(649, 36);
            this.formActionsLayout.TabIndex = 0;
            // 
            // okayButton
            // 
            this.okayButton.Location = new System.Drawing.Point(545, 4);
            this.okayButton.Margin = new System.Windows.Forms.Padding(4);
            this.okayButton.Name = "okayButton";
            this.okayButton.Size = new System.Drawing.Size(100, 28);
            this.okayButton.TabIndex = 0;
            this.okayButton.Text = "Okay";
            this.okayButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(437, 4);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 28);
            this.cancelButton.TabIndex = 0;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(0, 0);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 16);
            this.label6.TabIndex = 2;
            this.label6.Text = "label6";
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(0, 0);
            this.checkBox2.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(95, 20);
            this.checkBox2.TabIndex = 3;
            this.checkBox2.Text = "checkBox2";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // ConfigureForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(857, 666);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.configurationPanel);
            this.Controls.Add(this.formActionsLayout);
            this.Controls.Add(this.panel2);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "ConfigureForm";
            this.Text = "Form1";
            this.configurationPanel.ResumeLayout(false);
            this.optionsPage.ResumeLayout(false);
            this.optionsTabPage.ResumeLayout(false);
            this.optionsTabPage.PerformLayout();
            this.removeEmptyFolderExclusionsLayout.ResumeLayout(false);
            this.removeEmptyFolderExclusionsLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.removeEmptyFoldersExclusionsBindingSource)).EndInit();
            this.removeEmptyFolderExclusionsActionPanel.ResumeLayout(false);
            this.illegalCharacterReplacementsLayout.ResumeLayout(false);
            this.illegalCharacterReplacementsLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.illegalCharacterReplacementsBindingSource)).EndInit();
            this.monthReplacementsLayout.ResumeLayout(false);
            this.monthReplacementsLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.monthReplacementsBindingSource)).EndInit();
            this.emptyValuesTabPage.ResumeLayout(false);
            this.emptyValuesTabPage.PerformLayout();
            this.failOperationOnEmptyValueDestinationFolderLayout.ResumeLayout(false);
            this.failOperationOnEmptyValueDestinationFolderLayout.PerformLayout();
            this.failOperationOnEmptyValueFieldsLayout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.failOperationOnEmptyValueFields)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.failOperationOnEmptyValueFieldsBindingSource)).EndInit();
            this.emptyFieldReplacementLayout.ResumeLayout(false);
            this.emptyFieldReplacementLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.emptyFieldReplacementsBindingSource)).EndInit();
            this.emptyFolderNameReplacementLayout.ResumeLayout(false);
            this.emptyFolderNameReplacementLayout.PerformLayout();
            this.rulesPage.ResumeLayout(false);
            this.metadataRulesTabPage.ResumeLayout(false);
            this.metadataRulesTabPage.PerformLayout();
            this.metadataRulesControlsLayout.ResumeLayout(false);
            this.metadataRulesControlsLayout.PerformLayout();
            this.folderRulesTabPage.ResumeLayout(false);
            this.folderRulesTabPage.PerformLayout();
            this.folderRulesActionsLayout.ResumeLayout(false);
            this.folderStructurePage.ResumeLayout(false);
            this.folderStructurePage.PerformLayout();
            this.folderInsertControlsPanel.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.folderStructureActionsLayout.ResumeLayout(false);
            this.folderStructureActionsLayout.PerformLayout();
            this.folderStructurePreviewLayout.ResumeLayout(false);
            this.folderStructurePreviewLayout.PerformLayout();
            this.folderStructurePanel.ResumeLayout(false);
            this.folderStructurePanel.PerformLayout();
            this.fileStructurePage.ResumeLayout(false);
            this.fileStructurePage.PerformLayout();
            this.fileStructureInsertControlsPanel.ResumeLayout(false);
            this.insertControlsTabPanel.ResumeLayout(false);
            this.fileStructurePreviewLayout.ResumeLayout(false);
            this.fileStructurePreviewLayout.PerformLayout();
            this.fileStructurePanel.ResumeLayout(false);
            this.fileStructurePanel.PerformLayout();
            this.formActionsLayout.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.ListBox removeEmptyFolderExclusions;

        #endregion

        private System.Windows.Forms.Panel configurationPanel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel fileStructurePage;
        private System.Windows.Forms.FlowLayoutPanel formActionsLayout;
        private System.Windows.Forms.Button okayButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label fileStructureLabel;
        private System.Windows.Forms.TextBox fileStructureTextBox;
        private System.Windows.Forms.CheckBox fileSpaceAutomatically;
        private System.Windows.Forms.Button fileStructurePreviewPrevious;
        private System.Windows.Forms.Button fileStructurePreviewNext;
        private System.Windows.Forms.Label fileStructurePreview;
        private System.Windows.Forms.Panel fileStructurePanel;
        private System.Windows.Forms.TableLayoutPanel fileStructurePreviewLayout;
        private System.Windows.Forms.Label fileStructurePreviewLabel;
        private System.Windows.Forms.Panel fileStructureInsertControlsPanel;
        private System.Windows.Forms.Panel folderStructurePage;
        private System.Windows.Forms.TabControl insertControlsTabPanel;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.FlowLayoutPanel folderStructureActionsLayout;
        private System.Windows.Forms.CheckBox folderSpaceAutomatically;
        private System.Windows.Forms.Button insertFolderSeparator;
        private System.Windows.Forms.Panel folderInsertControlsPanel;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TableLayoutPanel folderStructurePreviewLayout;
        private System.Windows.Forms.Label folderPreviewLabel;
        private System.Windows.Forms.Label folderPreview;
        private System.Windows.Forms.Button folderPreviewPrevious;
        private System.Windows.Forms.Button folderPreviewNext;
        private System.Windows.Forms.Panel folderStructurePanel;
        private System.Windows.Forms.Label folderStructureLabel;
        private System.Windows.Forms.TextBox folderStructure;
        private System.Windows.Forms.TabControl rulesPage;
        private System.Windows.Forms.TabPage metadataRulesTabPage;
        private System.Windows.Forms.FlowLayoutPanel metadataRulesControlsLayout;
        private System.Windows.Forms.ComboBox metadataRulesAction;
        private System.Windows.Forms.Label metadataRulesActionLabel1;
        private System.Windows.Forms.ComboBox metadataRulesMode;
        private System.Windows.Forms.Label metadataRulesActionLabel2;
        private System.Windows.Forms.Button metadataRulesAddGroup;
        private System.Windows.Forms.Button metadataRulesAddRule;
        private System.Windows.Forms.TabPage folderRulesTabPage;
        private System.Windows.Forms.FlowLayoutPanel metadataRulesActionsLayout;
        private System.Windows.Forms.ListView excludedFoldersList;
        private System.Windows.Forms.FlowLayoutPanel folderRulesActionsLayout;
        private System.Windows.Forms.Button addExcludedFolder;
        private System.Windows.Forms.Button removeExcludedFolder;
        private System.Windows.Forms.Label excludedFolderLabel;
        private System.Windows.Forms.TabControl optionsPage;
        private System.Windows.Forms.TabPage optionsTabPage;
        private System.Windows.Forms.TabPage emptyValuesTabPage;
        private System.Windows.Forms.CheckBox autoSelectSingleMultiValueField;
        private System.Windows.Forms.CheckBox copyReadPercentageToReplacement;
        private System.Windows.Forms.CheckBox normalizeMultipleSpaces;
        private System.Windows.Forms.FlowLayoutPanel illegalCharacterReplacementsLayout;
        private System.Windows.Forms.FlowLayoutPanel monthReplacementsLayout;
        private System.Windows.Forms.Label monthReplacementsMonth;
        private System.Windows.Forms.ComboBox monthReplacementsMonthSelector;
        private System.Windows.Forms.Label monthReplacementsWith;
        private System.Windows.Forms.TextBox monthReplacementsReplacement;
        private System.Windows.Forms.Label illegalCharacterReplacementsLabel;
        private System.Windows.Forms.ComboBox illegalCharacterReplacementsCharacterSelector;
        private System.Windows.Forms.Label illegalCharacterReplacementsWith;
        private System.Windows.Forms.Button addIllegalCharacterReplacement;
        private System.Windows.Forms.Button removeIllegalCharacterReplacement;
        private System.Windows.Forms.Panel removeEmptyFolderExclusionsLayout;
        private System.Windows.Forms.Label removeEmptyFoldersLabel;
        private System.Windows.Forms.CheckBox removeEmptyFolders;
        private System.Windows.Forms.FlowLayoutPanel removeEmptyFolderExclusionsActionPanel;
        private System.Windows.Forms.Button addEmptyFolderExclusion;
        private System.Windows.Forms.Button removeEmptyFolderExclusion;
        private System.Windows.Forms.Label emptyFieldReplacementLabel;
        private System.Windows.Forms.Label emptyFolderNameReplacementLabel2;
        private System.Windows.Forms.TextBox emptyFolderNameReplacement;
        private System.Windows.Forms.Label emptyFolderNameReplacementLabel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox emptyFieldReplacement;
        private System.Windows.Forms.Label emptyFieldReplacementLabel3;
        private System.Windows.Forms.ComboBox emptyFieldReplacementSelector;
        private System.Windows.Forms.Label emptyFieldReplacementLabel2;
        private System.Windows.Forms.TableLayoutPanel emptyFieldReplacementLayout;
        private System.Windows.Forms.CheckBox failOperationOnEmptyValueUseDestinationFolder;
        private System.Windows.Forms.CheckBox failOperationOnEmptyValue;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.TableLayoutPanel failOperationOnEmptyValueDestinationFolderLayout;
        private System.Windows.Forms.Panel failOperationOnEmptyValueFieldsLayout;
        private System.Windows.Forms.Button failOperationOnEmptyValueDestinationFolderBrowse;
        private System.Windows.Forms.TextBox failOperationOnEmptyValueDestinationFolder;
        private System.Windows.Forms.TableLayoutPanel emptyFolderNameReplacementLayout;
        private System.Windows.Forms.TextBox illegalCharacterReplacementsReplacement;
        private System.Windows.Forms.DataGridView failOperationOnEmptyValueFields;
        private System.Windows.Forms.BindingSource profileBindingSource;
        private System.Windows.Forms.BindingSource monthReplacementsBindingSource;
        private System.Windows.Forms.BindingSource illegalCharacterReplacementsBindingSource;
        private System.Windows.Forms.BindingSource removeEmptyFoldersExclusionsBindingSource;
        private System.Windows.Forms.BindingSource emptyFieldReplacementsBindingSource;
        private System.Windows.Forms.BindingSource failOperationOnEmptyValueFieldsBindingSource;
        private System.Windows.Forms.DataGridViewCheckBoxColumn emptyValueFieldEnabledColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn emptyValueFieldNameColumn;
    }
}

