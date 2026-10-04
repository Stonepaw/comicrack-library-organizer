namespace LibraryOrganizer.Control
{
    partial class StringMatcherControl
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
            this.bookFieldStringMatcherViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mode = new System.Windows.Forms.ComboBox();
            this.value = new System.Windows.Forms.TextBox();
            this.actionsButton = new System.Windows.Forms.Button();
            this.matcherGroupActions = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.addMatcherAction = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.bookFieldStringMatcherViewModelBindingSource)).BeginInit();
            this.matcherGroupActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // field
            // 
            this.field.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.bookFieldStringMatcherViewModelBindingSource, "Field", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.field.FormattingEnabled = true;
            this.field.Location = new System.Drawing.Point(3, 3);
            this.field.Name = "field";
            this.field.Size = new System.Drawing.Size(121, 24);
            this.field.TabIndex = 0;
            // 
            // bookFieldStringMatcherViewModelBindingSource
            // 
            this.bookFieldStringMatcherViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.BookFieldStringMatcherViewModel);
            // 
            // mode
            // 
            this.mode.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.bookFieldStringMatcherViewModelBindingSource, "Mode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.mode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mode.FormattingEnabled = true;
            this.mode.Location = new System.Drawing.Point(130, 3);
            this.mode.Name = "mode";
            this.mode.Size = new System.Drawing.Size(121, 24);
            this.mode.TabIndex = 1;
            // 
            // value
            // 
            this.value.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.value.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bookFieldStringMatcherViewModelBindingSource, "Value", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.value.Location = new System.Drawing.Point(257, 4);
            this.value.Name = "value";
            this.value.Size = new System.Drawing.Size(555, 22);
            this.value.TabIndex = 2;
            // 
            // actionsButton
            // 
            this.actionsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsButton.AutoSize = true;
            this.actionsButton.Location = new System.Drawing.Point(818, 2);
            this.actionsButton.Name = "actionsButton";
            this.actionsButton.Size = new System.Drawing.Size(25, 26);
            this.actionsButton.TabIndex = 4;
            this.actionsButton.Text = "▼";
            this.actionsButton.UseVisualStyleBackColor = true;
            this.actionsButton.Click += new System.EventHandler(this.actionsButton_Click);
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
            // 
            // addMatcherAction
            // 
            this.addMatcherAction.Name = "addMatcherAction";
            this.addMatcherAction.Size = new System.Drawing.Size(151, 24);
            this.addMatcherAction.Text = "Add Rule";
            // 
            // deleteGroupAction
            // 
            this.deleteGroupAction.Name = "deleteGroupAction";
            this.deleteGroupAction.Size = new System.Drawing.Size(151, 24);
            this.deleteGroupAction.Text = "Delete";
            this.deleteGroupAction.Click += new System.EventHandler(this.deleteGroupAction_Click);
            // 
            // StringMatcherControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.field);
            this.Controls.Add(this.mode);
            this.Controls.Add(this.value);
            this.Controls.Add(this.actionsButton);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "StringMatcherControl";
            this.Size = new System.Drawing.Size(846, 31);
            ((System.ComponentModel.ISupportInitialize)(this.bookFieldStringMatcherViewModelBindingSource)).EndInit();
            this.matcherGroupActions.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox field;
        private System.Windows.Forms.ComboBox mode;
        private System.Windows.Forms.TextBox value;
        private System.Windows.Forms.Button actionsButton;
        private System.Windows.Forms.ContextMenuStrip matcherGroupActions;
        private System.Windows.Forms.ToolStripMenuItem addGroupAction;
        private System.Windows.Forms.ToolStripMenuItem addMatcherAction;
        private System.Windows.Forms.ToolStripMenuItem deleteGroupAction;
        private System.Windows.Forms.BindingSource bookFieldStringMatcherViewModelBindingSource;
    }
}
