using LibraryOrganizer.Controls;

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
            System.Windows.Forms.ToolStrip toolStrip;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfigureForm));
            System.Windows.Forms.ToolStripLabel profileLabel;
            System.Windows.Forms.ToolStripSeparator profileSeparator;
            this.overviewButton = new System.Windows.Forms.ToolStripButton();
            this.filesButton = new System.Windows.Forms.ToolStripButton();
            this.foldersButton = new System.Windows.Forms.ToolStripButton();
            this.rulesButton = new System.Windows.Forms.ToolStripButton();
            this.optionsButton = new System.Windows.Forms.ToolStripButton();
            this.profileActions = new System.Windows.Forms.ToolStripDropDownButton();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.profileSelector = new System.Windows.Forms.ToolStripComboBox();
            this.profileBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.configurationPanel = new System.Windows.Forms.Panel();
            this.optionsPage = new System.Windows.Forms.TabControl();
            this.optionsTabPage = new System.Windows.Forms.TabPage();
            this.optionsConfig = new LibraryOrganizer.Controls.OptionsConfigControl();
            this.emptyValuesTabPage = new System.Windows.Forms.TabPage();
            this.emptyValuesConfig = new LibraryOrganizer.Controls.EmptyValuesConfigurationControl();
            this.overviewConfig = new LibraryOrganizer.Controls.OverviewConfigControl();
            this.rulesPage = new System.Windows.Forms.TabControl();
            this.metadataRulesTabPage = new System.Windows.Forms.TabPage();
            this.profileMetadataRules = new LibraryOrganizer.Controls.ProfileMatcherGroupControl();
            this.folderRulesTabPage = new System.Windows.Forms.TabPage();
            this.excludedFoldersList = new System.Windows.Forms.ListView();
            this.folderRulesActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.addExcludedFolder = new System.Windows.Forms.Button();
            this.removeExcludedFolder = new System.Windows.Forms.Button();
            this.excludedFolderLabel = new System.Windows.Forms.Label();
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
            this.folderStructurePage = new System.Windows.Forms.Panel();
            this.folderInsertControlsPanel = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.folderStructurePreviewLayout = new System.Windows.Forms.TableLayoutPanel();
            this.folderPreviewLabel = new System.Windows.Forms.Label();
            this.folderPreview = new System.Windows.Forms.Label();
            this.folderPreviewPrevious = new System.Windows.Forms.Button();
            this.folderPreviewNext = new System.Windows.Forms.Button();
            this.folderStructureActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.folderSpaceAutomatically = new System.Windows.Forms.CheckBox();
            this.insertFolderSeparator = new System.Windows.Forms.Button();
            this.folderStructurePanel = new System.Windows.Forms.Panel();
            this.folderStructureLabel = new System.Windows.Forms.Label();
            this.folderStructure = new System.Windows.Forms.TextBox();
            this.configFormViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.formActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.okayButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            toolStrip = new System.Windows.Forms.ToolStrip();
            profileLabel = new System.Windows.Forms.ToolStripLabel();
            profileSeparator = new System.Windows.Forms.ToolStripSeparator();
            toolStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileBindingSource)).BeginInit();
            this.configurationPanel.SuspendLayout();
            this.optionsPage.SuspendLayout();
            this.optionsTabPage.SuspendLayout();
            this.emptyValuesTabPage.SuspendLayout();
            this.rulesPage.SuspendLayout();
            this.metadataRulesTabPage.SuspendLayout();
            this.folderRulesTabPage.SuspendLayout();
            this.folderRulesActionsLayout.SuspendLayout();
            this.fileStructurePage.SuspendLayout();
            this.fileStructureInsertControlsPanel.SuspendLayout();
            this.insertControlsTabPanel.SuspendLayout();
            this.fileStructurePreviewLayout.SuspendLayout();
            this.fileStructurePanel.SuspendLayout();
            this.folderStructurePage.SuspendLayout();
            this.folderInsertControlsPanel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.folderStructurePreviewLayout.SuspendLayout();
            this.folderStructureActionsLayout.SuspendLayout();
            this.folderStructurePanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.configFormViewModelBindingSource)).BeginInit();
            this.formActionsLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip
            // 
            toolStrip.AutoSize = false;
            toolStrip.Dock = System.Windows.Forms.DockStyle.Left;
            toolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            toolStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.overviewButton,
            this.filesButton,
            this.foldersButton,
            this.rulesButton,
            this.optionsButton,
            this.profileActions,
            this.profileSelector,
            profileLabel,
            profileSeparator});
            toolStrip.Location = new System.Drawing.Point(0, 0);
            toolStrip.Name = "toolStrip";
            toolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            toolStrip.ShowItemToolTips = false;
            toolStrip.Size = new System.Drawing.Size(130, 641);
            toolStrip.TabIndex = 1;
            toolStrip.Text = "toolStrip1";
            // 
            // overviewButton
            // 
            this.overviewButton.Checked = true;
            this.overviewButton.CheckOnClick = true;
            this.overviewButton.CheckState = System.Windows.Forms.CheckState.Checked;
            this.overviewButton.Image = ((System.Drawing.Image)(resources.GetObject("overviewButton.Image")));
            this.overviewButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.overviewButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.overviewButton.Margin = new System.Windows.Forms.Padding(10);
            this.overviewButton.Name = "overviewButton";
            this.overviewButton.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
            this.overviewButton.Size = new System.Drawing.Size(108, 76);
            this.overviewButton.Text = "Overview";
            this.overviewButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.overviewButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.overviewButton.Click += new System.EventHandler(this.PageButton_Click);
            // 
            // filesButton
            // 
            this.filesButton.Image = ((System.Drawing.Image)(resources.GetObject("filesButton.Image")));
            this.filesButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.filesButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.filesButton.Margin = new System.Windows.Forms.Padding(10);
            this.filesButton.Name = "filesButton";
            this.filesButton.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
            this.filesButton.Size = new System.Drawing.Size(108, 76);
            this.filesButton.Text = "Files";
            this.filesButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.filesButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.filesButton.Click += new System.EventHandler(this.PageButton_Click);
            // 
            // foldersButton
            // 
            this.foldersButton.Image = ((System.Drawing.Image)(resources.GetObject("foldersButton.Image")));
            this.foldersButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.foldersButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.foldersButton.Margin = new System.Windows.Forms.Padding(10);
            this.foldersButton.Name = "foldersButton";
            this.foldersButton.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
            this.foldersButton.Size = new System.Drawing.Size(108, 76);
            this.foldersButton.Text = "Folders";
            this.foldersButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.foldersButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.foldersButton.Click += new System.EventHandler(this.PageButton_Click);
            // 
            // rulesButton
            // 
            this.rulesButton.Image = ((System.Drawing.Image)(resources.GetObject("rulesButton.Image")));
            this.rulesButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.rulesButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.rulesButton.Margin = new System.Windows.Forms.Padding(10);
            this.rulesButton.Name = "rulesButton";
            this.rulesButton.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
            this.rulesButton.Size = new System.Drawing.Size(108, 76);
            this.rulesButton.Text = "Rules";
            this.rulesButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.rulesButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.rulesButton.Click += new System.EventHandler(this.PageButton_Click);
            // 
            // optionsButton
            // 
            this.optionsButton.Image = ((System.Drawing.Image)(resources.GetObject("optionsButton.Image")));
            this.optionsButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.optionsButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.optionsButton.Margin = new System.Windows.Forms.Padding(10);
            this.optionsButton.Name = "optionsButton";
            this.optionsButton.Padding = new System.Windows.Forms.Padding(0, 10, 0, 10);
            this.optionsButton.Size = new System.Drawing.Size(108, 76);
            this.optionsButton.Text = "Options";
            this.optionsButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.optionsButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.optionsButton.Click += new System.EventHandler(this.PageButton_Click);
            // 
            // profileActions
            // 
            this.profileActions.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.profileActions.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.profileActions.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem});
            this.profileActions.Image = ((System.Drawing.Image)(resources.GetObject("profileActions.Image")));
            this.profileActions.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.profileActions.Name = "profileActions";
            this.profileActions.Size = new System.Drawing.Size(128, 24);
            this.profileActions.Text = "Profile Action";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("newToolStripMenuItem.Image")));
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(122, 26);
            this.newToolStripMenuItem.Text = "New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // profileSelector
            // 
            this.profileSelector.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.profileSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.profileSelector.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
            this.profileSelector.Name = "profileSelector";
            this.profileSelector.Size = new System.Drawing.Size(126, 28);
            // 
            // profileLabel
            // 
            profileLabel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            profileLabel.Name = "profileLabel";
            profileLabel.Size = new System.Drawing.Size(128, 20);
            profileLabel.Text = "Profile";
            // 
            // profileSeparator
            // 
            profileSeparator.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            profileSeparator.Name = "profileSeparator";
            profileSeparator.Size = new System.Drawing.Size(128, 6);
            // 
            // profileBindingSource
            // 
            this.profileBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ProfileViewModel);
            // 
            // configurationPanel
            // 
            this.configurationPanel.AutoSize = true;
            this.configurationPanel.Controls.Add(this.rulesPage);
            this.configurationPanel.Controls.Add(this.overviewConfig);
            this.configurationPanel.Controls.Add(this.optionsPage);
            this.configurationPanel.Controls.Add(this.fileStructurePage);
            this.configurationPanel.Controls.Add(this.folderStructurePage);
            this.configurationPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.configurationPanel.Location = new System.Drawing.Point(130, 0);
            this.configurationPanel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.configurationPanel.Name = "configurationPanel";
            this.configurationPanel.Padding = new System.Windows.Forms.Padding(0, 12, 13, 4);
            this.configurationPanel.Size = new System.Drawing.Size(761, 605);
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
            this.optionsPage.Size = new System.Drawing.Size(748, 589);
            this.optionsPage.TabIndex = 10;
            // 
            // optionsTabPage
            // 
            this.optionsTabPage.Controls.Add(this.optionsConfig);
            this.optionsTabPage.Location = new System.Drawing.Point(4, 25);
            this.optionsTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.optionsTabPage.Name = "optionsTabPage";
            this.optionsTabPage.Padding = new System.Windows.Forms.Padding(9);
            this.optionsTabPage.Size = new System.Drawing.Size(740, 560);
            this.optionsTabPage.TabIndex = 0;
            this.optionsTabPage.Text = "Options";
            this.optionsTabPage.UseVisualStyleBackColor = true;
            // 
            // optionsConfig
            // 
            this.optionsConfig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.optionsConfig.Location = new System.Drawing.Point(9, 9);
            this.optionsConfig.Name = "optionsConfig";
            this.optionsConfig.ProfileViewModel = null;
            this.optionsConfig.Size = new System.Drawing.Size(722, 542);
            this.optionsConfig.TabIndex = 7;
            // 
            // emptyValuesTabPage
            // 
            this.emptyValuesTabPage.Controls.Add(this.emptyValuesConfig);
            this.emptyValuesTabPage.Location = new System.Drawing.Point(4, 25);
            this.emptyValuesTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.emptyValuesTabPage.Name = "emptyValuesTabPage";
            this.emptyValuesTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.emptyValuesTabPage.Size = new System.Drawing.Size(740, 560);
            this.emptyValuesTabPage.TabIndex = 1;
            this.emptyValuesTabPage.Text = "Empty Values";
            this.emptyValuesTabPage.UseVisualStyleBackColor = true;
            // 
            // emptyValuesConfig
            // 
            this.emptyValuesConfig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.emptyValuesConfig.Location = new System.Drawing.Point(4, 4);
            this.emptyValuesConfig.Name = "emptyValuesConfig";
            this.emptyValuesConfig.ProfileViewModel = null;
            this.emptyValuesConfig.Size = new System.Drawing.Size(732, 552);
            this.emptyValuesConfig.TabIndex = 0;
            // 
            // overviewConfig
            // 
            this.overviewConfig.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.overviewConfig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.overviewConfig.Location = new System.Drawing.Point(0, 12);
            this.overviewConfig.Name = "overviewConfig";
            this.overviewConfig.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.overviewConfig.Size = new System.Drawing.Size(748, 589);
            this.overviewConfig.TabIndex = 4;
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
            this.rulesPage.Size = new System.Drawing.Size(748, 589);
            this.rulesPage.TabIndex = 9;
            // 
            // metadataRulesTabPage
            // 
            this.metadataRulesTabPage.Controls.Add(this.profileMetadataRules);
            this.metadataRulesTabPage.Location = new System.Drawing.Point(4, 25);
            this.metadataRulesTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.metadataRulesTabPage.Name = "metadataRulesTabPage";
            this.metadataRulesTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.metadataRulesTabPage.Size = new System.Drawing.Size(740, 560);
            this.metadataRulesTabPage.TabIndex = 0;
            this.metadataRulesTabPage.Text = "Metadata Rules";
            this.metadataRulesTabPage.UseVisualStyleBackColor = true;
            // 
            // profileMetadataRules
            // 
            this.profileMetadataRules.DataBindings.Add(new System.Windows.Forms.Binding("Matcher", this.profileBindingSource, "Matchers", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.profileMetadataRules.Dock = System.Windows.Forms.DockStyle.Fill;
            this.profileMetadataRules.Location = new System.Drawing.Point(4, 4);
            this.profileMetadataRules.Matcher = null;
            this.profileMetadataRules.Name = "profileMetadataRules";
            this.profileMetadataRules.Size = new System.Drawing.Size(732, 552);
            this.profileMetadataRules.TabIndex = 0;
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
            this.folderRulesTabPage.Size = new System.Drawing.Size(740, 560);
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
            this.excludedFoldersList.Size = new System.Drawing.Size(624, 524);
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
            this.folderRulesActionsLayout.Location = new System.Drawing.Point(628, 32);
            this.folderRulesActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderRulesActionsLayout.Name = "folderRulesActionsLayout";
            this.folderRulesActionsLayout.Size = new System.Drawing.Size(108, 524);
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
            // fileStructurePage
            // 
            this.fileStructurePage.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.fileStructurePage.Controls.Add(this.fileStructureInsertControlsPanel);
            this.fileStructurePage.Controls.Add(this.fileSpaceAutomatically);
            this.fileStructurePage.Controls.Add(this.fileStructurePreviewLayout);
            this.fileStructurePage.Controls.Add(this.fileStructurePanel);
            this.fileStructurePage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileStructurePage.ForeColor = System.Drawing.SystemColors.ControlText;
            this.fileStructurePage.Location = new System.Drawing.Point(0, 12);
            this.fileStructurePage.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructurePage.Name = "fileStructurePage";
            this.fileStructurePage.Size = new System.Drawing.Size(748, 589);
            this.fileStructurePage.TabIndex = 0;
            // 
            // fileStructureInsertControlsPanel
            // 
            this.fileStructureInsertControlsPanel.Controls.Add(this.insertControlsTabPanel);
            this.fileStructureInsertControlsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileStructureInsertControlsPanel.Location = new System.Drawing.Point(0, 118);
            this.fileStructureInsertControlsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.fileStructureInsertControlsPanel.Name = "fileStructureInsertControlsPanel";
            this.fileStructureInsertControlsPanel.Size = new System.Drawing.Size(748, 471);
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
            this.insertControlsTabPanel.Size = new System.Drawing.Size(748, 471);
            this.insertControlsTabPanel.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(740, 442);
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
            this.tabPage2.Size = new System.Drawing.Size(740, 442);
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
            this.fileSpaceAutomatically.Size = new System.Drawing.Size(748, 28);
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
            this.fileStructurePreviewLayout.Size = new System.Drawing.Size(748, 34);
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
            this.fileStructurePreview.Size = new System.Drawing.Size(613, 16);
            this.fileStructurePreview.TabIndex = 1;
            this.fileStructurePreview.Text = "label2";
            // 
            // fileStructurePreviewPrevious
            // 
            this.fileStructurePreviewPrevious.AutoSize = true;
            this.fileStructurePreviewPrevious.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.fileStructurePreviewPrevious.Location = new System.Drawing.Point(688, 4);
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
            this.fileStructurePreviewNext.Location = new System.Drawing.Point(720, 4);
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
            this.fileStructurePanel.Size = new System.Drawing.Size(748, 56);
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
            this.fileStructureTextBox.Size = new System.Drawing.Size(632, 48);
            this.fileStructureTextBox.TabIndex = 1;
            // 
            // folderStructurePage
            // 
            this.folderStructurePage.BackColor = System.Drawing.SystemColors.ControlLight;
            this.folderStructurePage.Controls.Add(this.folderInsertControlsPanel);
            this.folderStructurePage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.folderStructurePage.Location = new System.Drawing.Point(0, 12);
            this.folderStructurePage.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePage.Name = "folderStructurePage";
            this.folderStructurePage.Size = new System.Drawing.Size(748, 589);
            this.folderStructurePage.TabIndex = 1;
            // 
            // folderInsertControlsPanel
            // 
            this.folderInsertControlsPanel.Controls.Add(this.panel1);
            this.folderInsertControlsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.folderInsertControlsPanel.Location = new System.Drawing.Point(0, 0);
            this.folderInsertControlsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.folderInsertControlsPanel.Name = "folderInsertControlsPanel";
            this.folderInsertControlsPanel.Padding = new System.Windows.Forms.Padding(1);
            this.folderInsertControlsPanel.Size = new System.Drawing.Size(748, 589);
            this.folderInsertControlsPanel.TabIndex = 8;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.panel1.Controls.Add(this.folderStructurePreviewLayout);
            this.panel1.Controls.Add(this.folderStructureActionsLayout);
            this.panel1.Controls.Add(this.folderStructurePanel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(746, 587);
            this.panel1.TabIndex = 2;
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
            this.folderStructurePreviewLayout.Location = new System.Drawing.Point(0, 96);
            this.folderStructurePreviewLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePreviewLayout.Name = "folderStructurePreviewLayout";
            this.folderStructurePreviewLayout.RowCount = 1;
            this.folderStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.folderStructurePreviewLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.folderStructurePreviewLayout.Size = new System.Drawing.Size(746, 34);
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
            this.folderPreview.Size = new System.Drawing.Size(611, 16);
            this.folderPreview.TabIndex = 1;
            this.folderPreview.Text = "label2";
            // 
            // folderPreviewPrevious
            // 
            this.folderPreviewPrevious.AutoSize = true;
            this.folderPreviewPrevious.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.folderPreviewPrevious.Location = new System.Drawing.Point(686, 4);
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
            this.folderPreviewNext.Location = new System.Drawing.Point(718, 4);
            this.folderPreviewNext.Margin = new System.Windows.Forms.Padding(4);
            this.folderPreviewNext.Name = "folderPreviewNext";
            this.folderPreviewNext.Size = new System.Drawing.Size(24, 26);
            this.folderPreviewNext.TabIndex = 3;
            this.folderPreviewNext.Text = ">";
            this.folderPreviewNext.UseVisualStyleBackColor = true;
            // 
            // folderStructureActionsLayout
            // 
            this.folderStructureActionsLayout.AutoSize = true;
            this.folderStructureActionsLayout.Controls.Add(this.folderSpaceAutomatically);
            this.folderStructureActionsLayout.Controls.Add(this.insertFolderSeparator);
            this.folderStructureActionsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.folderStructureActionsLayout.Location = new System.Drawing.Point(0, 56);
            this.folderStructureActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructureActionsLayout.Name = "folderStructureActionsLayout";
            this.folderStructureActionsLayout.Padding = new System.Windows.Forms.Padding(7, 0, 0, 0);
            this.folderStructureActionsLayout.Size = new System.Drawing.Size(746, 40);
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
            // folderStructurePanel
            // 
            this.folderStructurePanel.AutoSize = true;
            this.folderStructurePanel.Controls.Add(this.folderStructureLabel);
            this.folderStructurePanel.Controls.Add(this.folderStructure);
            this.folderStructurePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.folderStructurePanel.Location = new System.Drawing.Point(0, 0);
            this.folderStructurePanel.Margin = new System.Windows.Forms.Padding(4);
            this.folderStructurePanel.Name = "folderStructurePanel";
            this.folderStructurePanel.Size = new System.Drawing.Size(746, 56);
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
            this.folderStructure.Size = new System.Drawing.Size(613, 48);
            this.folderStructure.TabIndex = 1;
            // 
            // configFormViewModelBindingSource
            // 
            this.configFormViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ConfigFormViewModel);
            // 
            // formActionsLayout
            // 
            this.formActionsLayout.AutoSize = true;
            this.formActionsLayout.Controls.Add(this.okayButton);
            this.formActionsLayout.Controls.Add(this.cancelButton);
            this.formActionsLayout.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.formActionsLayout.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.formActionsLayout.Location = new System.Drawing.Point(130, 605);
            this.formActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.formActionsLayout.Name = "formActionsLayout";
            this.formActionsLayout.Size = new System.Drawing.Size(761, 36);
            this.formActionsLayout.TabIndex = 0;
            // 
            // okayButton
            // 
            this.okayButton.Location = new System.Drawing.Point(657, 4);
            this.okayButton.Margin = new System.Windows.Forms.Padding(4);
            this.okayButton.Name = "okayButton";
            this.okayButton.Size = new System.Drawing.Size(100, 28);
            this.okayButton.TabIndex = 0;
            this.okayButton.Text = "Okay";
            this.okayButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(549, 4);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 28);
            this.cancelButton.TabIndex = 0;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // ConfigureForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(891, 641);
            this.Controls.Add(this.configurationPanel);
            this.Controls.Add(this.formActionsLayout);
            this.Controls.Add(toolStrip);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "ConfigureForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.ConfigureForm_Load);
            this.ResizeBegin += new System.EventHandler(this.ConfigureForm_ResizeBegin);
            this.ResizeEnd += new System.EventHandler(this.ConfigureForm_ResizeEnd);
            toolStrip.ResumeLayout(false);
            toolStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileBindingSource)).EndInit();
            this.configurationPanel.ResumeLayout(false);
            this.optionsPage.ResumeLayout(false);
            this.optionsTabPage.ResumeLayout(false);
            this.emptyValuesTabPage.ResumeLayout(false);
            this.rulesPage.ResumeLayout(false);
            this.metadataRulesTabPage.ResumeLayout(false);
            this.folderRulesTabPage.ResumeLayout(false);
            this.folderRulesTabPage.PerformLayout();
            this.folderRulesActionsLayout.ResumeLayout(false);
            this.fileStructurePage.ResumeLayout(false);
            this.fileStructurePage.PerformLayout();
            this.fileStructureInsertControlsPanel.ResumeLayout(false);
            this.insertControlsTabPanel.ResumeLayout(false);
            this.fileStructurePreviewLayout.ResumeLayout(false);
            this.fileStructurePreviewLayout.PerformLayout();
            this.fileStructurePanel.ResumeLayout(false);
            this.fileStructurePanel.PerformLayout();
            this.folderStructurePage.ResumeLayout(false);
            this.folderInsertControlsPanel.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.folderStructurePreviewLayout.ResumeLayout(false);
            this.folderStructurePreviewLayout.PerformLayout();
            this.folderStructureActionsLayout.ResumeLayout(false);
            this.folderStructureActionsLayout.PerformLayout();
            this.folderStructurePanel.ResumeLayout(false);
            this.folderStructurePanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.configFormViewModelBindingSource)).EndInit();
            this.formActionsLayout.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel configurationPanel;
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
        private System.Windows.Forms.TabPage folderRulesTabPage;
        private System.Windows.Forms.ListView excludedFoldersList;
        private System.Windows.Forms.FlowLayoutPanel folderRulesActionsLayout;
        private System.Windows.Forms.Button addExcludedFolder;
        private System.Windows.Forms.Button removeExcludedFolder;
        private System.Windows.Forms.Label excludedFolderLabel;
        private System.Windows.Forms.TabControl optionsPage;
        private System.Windows.Forms.TabPage optionsTabPage;
        private System.Windows.Forms.TabPage emptyValuesTabPage;
        private System.Windows.Forms.BindingSource profileBindingSource;
        private System.Windows.Forms.ToolStripButton overviewButton;
        private System.Windows.Forms.ToolStripButton filesButton;
        private System.Windows.Forms.ToolStripButton foldersButton;
        private System.Windows.Forms.ToolStripButton optionsButton;
        private System.Windows.Forms.ToolStripButton rulesButton;
        private System.Windows.Forms.ToolStripDropDownButton profileActions;
        private System.Windows.Forms.ToolStripComboBox profileSelector;
        private System.Windows.Forms.BindingSource configFormViewModelBindingSource;
        private System.Windows.Forms.Panel panel1;
        private ProfileMatcherGroupControl profileMetadataRules;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private OverviewConfigControl overviewConfig;
        private OptionsConfigControl optionsConfig;
        private EmptyValuesConfigurationControl emptyValuesConfig;
    }
}

