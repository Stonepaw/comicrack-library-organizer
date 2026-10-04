namespace LibraryOrganizer.Controls
{
    internal partial class MatcherControl
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
            this.field = new System.Windows.Forms.ComboBox();
            this.matcherViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.matcherGroupActions = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.addMatcherAction = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.negationButton = new System.Windows.Forms.CheckBox();
            this.actionsButton = new System.Windows.Forms.Button();
            this.matcherValuePanel = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.matcherViewModelBindingSource)).BeginInit();
            this.matcherGroupActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // field
            // 
            this.field.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.matcherViewModelBindingSource, "Field", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.field.FormattingEnabled = true;
            this.field.Location = new System.Drawing.Point(29, 1);
            this.field.Margin = new System.Windows.Forms.Padding(3, 1, 3, 0);
            this.field.Name = "field";
            this.field.Size = new System.Drawing.Size(144, 24);
            this.field.TabIndex = 0;
            // 
            // matcherViewModelBindingSource
            // 
            this.matcherViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.BookFieldMatcherViewModel);
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
            // negationButton
            // 
            this.negationButton.Appearance = System.Windows.Forms.Appearance.Button;
            this.negationButton.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.negationButton.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.matcherViewModelBindingSource, "Negated", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.negationButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.negationButton.Location = new System.Drawing.Point(0, 0);
            this.negationButton.Margin = new System.Windows.Forms.Padding(0);
            this.negationButton.MinimumSize = new System.Drawing.Size(20, 20);
            this.negationButton.Name = "negationButton";
            this.negationButton.Size = new System.Drawing.Size(26, 26);
            this.negationButton.TabIndex = 5;
            this.negationButton.Text = "!";
            this.negationButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.negationButton.UseVisualStyleBackColor = true;
            // 
            // actionsButton
            // 
            this.actionsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsButton.Location = new System.Drawing.Point(979, 0);
            this.actionsButton.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.actionsButton.Name = "actionsButton";
            this.actionsButton.Size = new System.Drawing.Size(26, 26);
            this.actionsButton.TabIndex = 4;
            this.actionsButton.Text = "▼";
            this.actionsButton.UseVisualStyleBackColor = true;
            this.actionsButton.Click += new System.EventHandler(this.actionsButton_Click);
            // 
            // matcherValuePanel
            // 
            this.matcherValuePanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.matcherValuePanel.Location = new System.Drawing.Point(179, 0);
            this.matcherValuePanel.Name = "matcherValuePanel";
            this.matcherValuePanel.Padding = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.matcherValuePanel.Size = new System.Drawing.Size(794, 26);
            this.matcherValuePanel.TabIndex = 6;
            // 
            // MatcherControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.negationButton);
            this.Controls.Add(this.matcherValuePanel);
            this.Controls.Add(this.field);
            this.Controls.Add(this.actionsButton);
            this.Margin = new System.Windows.Forms.Padding(0, 0, 0, 3);
            this.Name = "MatcherControl";
            this.Size = new System.Drawing.Size(1005, 29);
            ((System.ComponentModel.ISupportInitialize)(this.matcherViewModelBindingSource)).EndInit();
            this.matcherGroupActions.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ComboBox field;
        private System.Windows.Forms.Button actionsButton;
        private System.Windows.Forms.ContextMenuStrip matcherGroupActions;
        private System.Windows.Forms.ToolStripMenuItem addGroupAction;
        private System.Windows.Forms.ToolStripMenuItem addMatcherAction;
        private System.Windows.Forms.ToolStripMenuItem deleteGroupAction;
        private System.Windows.Forms.CheckBox negationButton;
        private System.Windows.Forms.Panel matcherValuePanel;
        private System.Windows.Forms.BindingSource matcherViewModelBindingSource;
    }
}
