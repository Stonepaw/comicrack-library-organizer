namespace LibraryOrganizer.Controls
{
    partial class MatcherGroupControl
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
            System.Windows.Forms.Label matchLabel2;
            this.matchOperation = new System.Windows.Forms.ComboBox();
            this.actionsButton = new System.Windows.Forms.Button();
            this.matcherGroupActions = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.addMatcherAction = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.matchLabel = new System.Windows.Forms.Label();
            this.matchersPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.configPanel = new System.Windows.Forms.Panel();
            this.groupMatcherViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            matchLabel2 = new System.Windows.Forms.Label();
            this.matcherGroupActions.SuspendLayout();
            this.configPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupMatcherViewModelBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // matchLabel2
            // 
            matchLabel2.AutoSize = true;
            matchLabel2.Location = new System.Drawing.Point(121, 6);
            matchLabel2.Name = "matchLabel2";
            matchLabel2.Size = new System.Drawing.Size(94, 16);
            matchLabel2.TabIndex = 3;
            matchLabel2.Text = "of the following";
            // 
            // matchOperation
            // 
            this.matchOperation.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.groupMatcherViewModelBindingSource, "Mode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.matchOperation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.matchOperation.FormattingEnabled = true;
            this.matchOperation.Items.AddRange(new object[] {
            "All",
            "Any"});
            this.matchOperation.Location = new System.Drawing.Point(52, 2);
            this.matchOperation.Name = "matchOperation";
            this.matchOperation.Size = new System.Drawing.Size(63, 24);
            this.matchOperation.TabIndex = 1;
            // 
            // actionsButton
            // 
            this.actionsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsButton.Location = new System.Drawing.Point(274, 1);
            this.actionsButton.Margin = new System.Windows.Forms.Padding(0, 3, 3, 3);
            this.actionsButton.Name = "actionsButton";
            this.actionsButton.Size = new System.Drawing.Size(26, 26);
            this.actionsButton.TabIndex = 2;
            this.actionsButton.Text = "▼";
            this.actionsButton.UseVisualStyleBackColor = true;
            this.actionsButton.Click += new System.EventHandler(this.matchActionsButton_Click);
            // 
            // matcherGroupActions
            // 
            this.matcherGroupActions.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.matcherGroupActions.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addGroupAction,
            this.addMatcherAction,
            this.deleteGroupAction});
            this.matcherGroupActions.Name = "contextMenuStrip1";
            this.matcherGroupActions.Size = new System.Drawing.Size(152, 76);
            // 
            // addGroupAction
            // 
            this.addGroupAction.Name = "addGroupAction";
            this.addGroupAction.Size = new System.Drawing.Size(151, 24);
            this.addGroupAction.Text = "Add Group";
            this.addGroupAction.Click += new System.EventHandler(this.addGroupAction_Click);
            // 
            // addMatcherAction
            // 
            this.addMatcherAction.Name = "addMatcherAction";
            this.addMatcherAction.Size = new System.Drawing.Size(151, 24);
            this.addMatcherAction.Text = "Add Rule";
            this.addMatcherAction.Click += new System.EventHandler(this.addMatcherAction_Click);
            // 
            // deleteGroupAction
            // 
            this.deleteGroupAction.Name = "deleteGroupAction";
            this.deleteGroupAction.Size = new System.Drawing.Size(151, 24);
            this.deleteGroupAction.Text = "Delete";
            this.deleteGroupAction.Click += new System.EventHandler(this.deleteGroupAction_Click);
            // 
            // matchLabel
            // 
            this.matchLabel.AutoSize = true;
            this.matchLabel.Location = new System.Drawing.Point(3, 6);
            this.matchLabel.Name = "matchLabel";
            this.matchLabel.Size = new System.Drawing.Size(43, 16);
            this.matchLabel.TabIndex = 0;
            this.matchLabel.Text = "Match";
            // 
            // matchersPanel
            // 
            this.matchersPanel.AutoSize = true;
            this.matchersPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.matchersPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.matchersPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.matchersPanel.Location = new System.Drawing.Point(0, 30);
            this.matchersPanel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.matchersPanel.MinimumSize = new System.Drawing.Size(200, 0);
            this.matchersPanel.Name = "matchersPanel";
            this.matchersPanel.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.matchersPanel.Size = new System.Drawing.Size(300, 0);
            this.matchersPanel.TabIndex = 0;
            this.matchersPanel.WrapContents = false;
            // 
            // configPanel
            // 
            this.configPanel.AutoSize = true;
            this.configPanel.Controls.Add(this.matchLabel);
            this.configPanel.Controls.Add(this.matchOperation);
            this.configPanel.Controls.Add(matchLabel2);
            this.configPanel.Controls.Add(this.actionsButton);
            this.configPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.configPanel.Location = new System.Drawing.Point(0, 0);
            this.configPanel.MinimumSize = new System.Drawing.Size(200, 0);
            this.configPanel.Name = "configPanel";
            this.configPanel.Size = new System.Drawing.Size(300, 30);
            this.configPanel.TabIndex = 4;
            // 
            // groupMatcherViewModelBindingSource
            // 
            this.groupMatcherViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.GroupMatcherViewModel);
            // 
            // MatcherGroupControl
            // 
            this.Controls.Add(this.matchersPanel);
            this.Controls.Add(this.configPanel);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.MinimumSize = new System.Drawing.Size(300, 0);
            this.Name = "MatcherGroupControl";
            this.Size = new System.Drawing.Size(300, 30);
            this.matcherGroupActions.ResumeLayout(false);
            this.configPanel.ResumeLayout(false);
            this.configPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupMatcherViewModelBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox matchOperation;
        private System.Windows.Forms.Button actionsButton;
        private System.Windows.Forms.ContextMenuStrip matcherGroupActions;
        private System.Windows.Forms.ToolStripMenuItem addGroupAction;
        private System.Windows.Forms.ToolStripMenuItem deleteGroupAction;
        private System.Windows.Forms.Label matchLabel;
        private System.Windows.Forms.FlowLayoutPanel matchersPanel;
        private System.Windows.Forms.Panel configPanel;
        private System.Windows.Forms.ToolStripMenuItem addMatcherAction;
        private System.Windows.Forms.BindingSource groupMatcherViewModelBindingSource;
    }
}
