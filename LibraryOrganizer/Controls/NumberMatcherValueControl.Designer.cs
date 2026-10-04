namespace LibraryOrganizer.Controls
{
    partial class NumberMatcherValueControl
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
            this.valueTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.value = new System.Windows.Forms.TextBox();
            this.numberMatcherValueViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.value2 = new System.Windows.Forms.TextBox();
            this.mode = new System.Windows.Forms.ComboBox();
            this.valueTableLayoutPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numberMatcherValueViewModelBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // valueTableLayoutPanel
            // 
            this.valueTableLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.valueTableLayoutPanel.ColumnCount = 2;
            this.valueTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.valueTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.valueTableLayoutPanel.Controls.Add(this.value, 0, 0);
            this.valueTableLayoutPanel.Controls.Add(this.value2, 1, 0);
            this.valueTableLayoutPanel.Location = new System.Drawing.Point(127, 0);
            this.valueTableLayoutPanel.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.valueTableLayoutPanel.Name = "valueTableLayoutPanel";
            this.valueTableLayoutPanel.RowCount = 1;
            this.valueTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.valueTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.valueTableLayoutPanel.Size = new System.Drawing.Size(623, 24);
            this.valueTableLayoutPanel.TabIndex = 0;
            // 
            // value
            // 
            this.value.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.value.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.numberMatcherValueViewModelBindingSource, "Value", true));
            this.value.Location = new System.Drawing.Point(0, 1);
            this.value.Margin = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.value.Name = "value";
            this.value.Size = new System.Drawing.Size(308, 22);
            this.value.TabIndex = 0;
            // 
            // numberMatcherValueViewModelBindingSource
            // 
            this.numberMatcherValueViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.NumberMatcherValueViewModel);
            // 
            // value2
            // 
            this.value2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.value2.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.numberMatcherValueViewModelBindingSource, "Value2", true));
            this.value2.Location = new System.Drawing.Point(314, 1);
            this.value2.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.value2.Name = "value2";
            this.value2.Size = new System.Drawing.Size(309, 22);
            this.value2.TabIndex = 1;
            // 
            // mode
            // 
            this.mode.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.numberMatcherValueViewModelBindingSource, "Mode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.mode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mode.FormattingEnabled = true;
            this.mode.Location = new System.Drawing.Point(0, 0);
            this.mode.Margin = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.mode.Name = "mode";
            this.mode.Size = new System.Drawing.Size(121, 24);
            this.mode.TabIndex = 0;
            // 
            // NumberMatcherValueControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.mode);
            this.Controls.Add(this.valueTableLayoutPanel);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "NumberMatcherValueControl";
            this.Size = new System.Drawing.Size(750, 24);
            this.valueTableLayoutPanel.ResumeLayout(false);
            this.valueTableLayoutPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numberMatcherValueViewModelBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel valueTableLayoutPanel;
        private System.Windows.Forms.ComboBox mode;
        private System.Windows.Forms.TextBox value2;
        private System.Windows.Forms.TextBox value;
        private System.Windows.Forms.BindingSource numberMatcherValueViewModelBindingSource;
    }
}
