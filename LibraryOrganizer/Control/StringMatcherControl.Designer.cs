namespace LibraryOrganizer.Control
{
    internal partial class StringMatcherControl
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
            this.mode = new System.Windows.Forms.ComboBox();
            this.value = new System.Windows.Forms.TextBox();
            this.matcherGroupActions = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.addMatcherAction = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteGroupAction = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.actionsButton = new System.Windows.Forms.Button();
            this.bookFieldStringMatcherViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.matcherGroupActions.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bookFieldStringMatcherViewModelBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // field
            // 
            this.field.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.field.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.bookFieldStringMatcherViewModelBindingSource, "Field", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.field.FormattingEnabled = true;
            this.field.Location = new System.Drawing.Point(29, 1);
            this.field.Margin = new System.Windows.Forms.Padding(3, 1, 3, 0);
            this.field.Name = "field";
            this.field.Size = new System.Drawing.Size(124, 24);
            this.field.TabIndex = 0;
            // 
            // mode
            // 
            this.mode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.mode.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.bookFieldStringMatcherViewModelBindingSource, "Mode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.mode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mode.FormattingEnabled = true;
            this.mode.Location = new System.Drawing.Point(159, 1);
            this.mode.Margin = new System.Windows.Forms.Padding(3, 1, 3, 0);
            this.mode.Name = "mode";
            this.mode.Size = new System.Drawing.Size(124, 24);
            this.mode.TabIndex = 1;
            // 
            // value
            // 
            this.value.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.value.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bookFieldStringMatcherViewModelBindingSource, "Value", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.value.Location = new System.Drawing.Point(289, 2);
            this.value.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.value.Name = "value";
            this.value.Size = new System.Drawing.Size(546, 22);
            this.value.TabIndex = 2;
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
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 5;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.checkBox1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.field, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.mode, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.value, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.actionsButton, 4, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(867, 26);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox1.Appearance = System.Windows.Forms.Appearance.Button;
            this.checkBox1.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBox1.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.bookFieldStringMatcherViewModelBindingSource, "Negated", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.checkBox1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.checkBox1.Location = new System.Drawing.Point(0, 0);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(0);
            this.checkBox1.MinimumSize = new System.Drawing.Size(20, 20);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(26, 26);
            this.checkBox1.TabIndex = 5;
            this.checkBox1.Text = "!";
            this.checkBox1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // actionsButton
            // 
            this.actionsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsButton.Location = new System.Drawing.Point(841, 0);
            this.actionsButton.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.actionsButton.Name = "actionsButton";
            this.actionsButton.Size = new System.Drawing.Size(26, 26);
            this.actionsButton.TabIndex = 4;
            this.actionsButton.Text = "▼";
            this.actionsButton.UseVisualStyleBackColor = true;
            this.actionsButton.Click += new System.EventHandler(this.actionsButton_Click);
            // 
            // bookFieldStringMatcherViewModelBindingSource
            // 
            this.bookFieldStringMatcherViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.BookFieldStringMatcherViewModel);
            // 
            // StringMatcherControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "StringMatcherControl";
            this.Size = new System.Drawing.Size(867, 29);
            this.matcherGroupActions.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bookFieldStringMatcherViewModelBindingSource)).EndInit();
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
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}
