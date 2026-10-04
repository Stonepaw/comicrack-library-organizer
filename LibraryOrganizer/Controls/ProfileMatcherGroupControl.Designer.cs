using System.ComponentModel;

namespace LibraryOrganizer.Controls
{
    partial class ProfileMatcherGroupControl
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            System.Windows.Forms.Label matcherConfigLabel1;
            System.Windows.Forms.Label matcherConfigLabel2;
            this.matcherNegation = new System.Windows.Forms.ComboBox();
            this.matcherMode = new System.Windows.Forms.ComboBox();
            this.matchersPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.configPanel = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.matcherGroupActionMenuButton = new System.Windows.Forms.Button();
            this.matcherGroupActionMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addGroup = new System.Windows.Forms.ToolStripMenuItem();
            this.addMatcher = new System.Windows.Forms.ToolStripMenuItem();
            matcherConfigLabel1 = new System.Windows.Forms.Label();
            matcherConfigLabel2 = new System.Windows.Forms.Label();
            this.configPanel.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.matcherGroupActionMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // matcherConfigLabel1
            // 
            matcherConfigLabel1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            matcherConfigLabel1.AutoSize = true;
            matcherConfigLabel1.Location = new System.Drawing.Point(89, 8);
            matcherConfigLabel1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            matcherConfigLabel1.Name = "matcherConfigLabel1";
            matcherConfigLabel1.Size = new System.Drawing.Size(145, 16);
            matcherConfigLabel1.TabIndex = 1;
            matcherConfigLabel1.Text = "move books that match";
            // 
            // matcherConfigLabel2
            // 
            matcherConfigLabel2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            matcherConfigLabel2.AutoSize = true;
            matcherConfigLabel2.Location = new System.Drawing.Point(327, 8);
            matcherConfigLabel2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            matcherConfigLabel2.Name = "matcherConfigLabel2";
            matcherConfigLabel2.Size = new System.Drawing.Size(126, 16);
            matcherConfigLabel2.TabIndex = 3;
            matcherConfigLabel2.Text = "of the following rules";
            // 
            // matcherNegation
            // 
            this.matcherNegation.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.matcherNegation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.matcherNegation.FormattingEnabled = true;
            this.matcherNegation.Location = new System.Drawing.Point(4, 4);
            this.matcherNegation.Margin = new System.Windows.Forms.Padding(4);
            this.matcherNegation.Name = "matcherNegation";
            this.matcherNegation.Size = new System.Drawing.Size(77, 24);
            this.matcherNegation.TabIndex = 0;
            // 
            // matcherMode
            // 
            this.matcherMode.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.matcherMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.matcherMode.FormattingEnabled = true;
            this.matcherMode.Location = new System.Drawing.Point(242, 4);
            this.matcherMode.Margin = new System.Windows.Forms.Padding(4);
            this.matcherMode.Name = "matcherMode";
            this.matcherMode.Size = new System.Drawing.Size(77, 24);
            this.matcherMode.TabIndex = 2;
            // 
            // matchersPanel
            // 
            this.matchersPanel.AutoScroll = true;
            this.matchersPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.matchersPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.matchersPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.matchersPanel.Location = new System.Drawing.Point(0, 38);
            this.matchersPanel.Name = "matchersPanel";
            this.matchersPanel.Padding = new System.Windows.Forms.Padding(6, 0, 10, 0);
            this.matchersPanel.Size = new System.Drawing.Size(700, 87);
            this.matchersPanel.TabIndex = 6;
            this.matchersPanel.WrapContents = false;
            // 
            // configPanel
            // 
            this.configPanel.AutoSize = true;
            this.configPanel.Controls.Add(this.flowLayoutPanel1);
            this.configPanel.Controls.Add(this.matcherGroupActionMenuButton);
            this.configPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.configPanel.Location = new System.Drawing.Point(0, 0);
            this.configPanel.Margin = new System.Windows.Forms.Padding(0);
            this.configPanel.Name = "configPanel";
            this.configPanel.Size = new System.Drawing.Size(700, 38);
            this.configPanel.TabIndex = 6;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flowLayoutPanel1.Controls.Add(this.matcherNegation);
            this.flowLayoutPanel1.Controls.Add(matcherConfigLabel1);
            this.flowLayoutPanel1.Controls.Add(this.matcherMode);
            this.flowLayoutPanel1.Controls.Add(matcherConfigLabel2);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(457, 32);
            this.flowLayoutPanel1.TabIndex = 7;
            // 
            // matcherGroupActionMenuButton
            // 
            this.matcherGroupActionMenuButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.matcherGroupActionMenuButton.Location = new System.Drawing.Point(655, 5);
            this.matcherGroupActionMenuButton.Margin = new System.Windows.Forms.Padding(0);
            this.matcherGroupActionMenuButton.Name = "matcherGroupActionMenuButton";
            this.matcherGroupActionMenuButton.Size = new System.Drawing.Size(26, 26);
            this.matcherGroupActionMenuButton.TabIndex = 6;
            this.matcherGroupActionMenuButton.Text = "▼";
            this.matcherGroupActionMenuButton.UseVisualStyleBackColor = true;
            this.matcherGroupActionMenuButton.Click += new System.EventHandler(this.matcherGroupActionMenuButton_Click);
            // 
            // matcherGroupActionMenu
            // 
            this.matcherGroupActionMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.matcherGroupActionMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addGroup,
            this.addMatcher});
            this.matcherGroupActionMenu.Name = "contextMenuStrip1";
            this.matcherGroupActionMenu.Size = new System.Drawing.Size(152, 52);
            // 
            // addGroup
            // 
            this.addGroup.Name = "addGroup";
            this.addGroup.Size = new System.Drawing.Size(151, 24);
            this.addGroup.Text = "Add Group";
            this.addGroup.Click += new System.EventHandler(this.addGroup_Click);
            // 
            // addMatcher
            // 
            this.addMatcher.Name = "addMatcher";
            this.addMatcher.Size = new System.Drawing.Size(151, 24);
            this.addMatcher.Text = "Add Rule";
            this.addMatcher.Click += new System.EventHandler(this.addRule_Click);
            // 
            // ProfileMatcherGroupControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.matchersPanel);
            this.Controls.Add(this.configPanel);
            this.Name = "ProfileMatcherGroupControl";
            this.Size = new System.Drawing.Size(700, 125);
            this.Resize += new System.EventHandler(this.ProfileMatcherGroupControl_Resize);
            this.configPanel.ResumeLayout(false);
            this.configPanel.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.matcherGroupActionMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.ComboBox matcherNegation;
        private System.Windows.Forms.ComboBox matcherMode;

        #endregion

        private System.Windows.Forms.FlowLayoutPanel matchersPanel;
        private System.Windows.Forms.Panel configPanel;
        private System.Windows.Forms.ContextMenuStrip matcherGroupActionMenu;
        private System.Windows.Forms.ToolStripMenuItem addGroup;
        private System.Windows.Forms.ToolStripMenuItem addMatcher;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button matcherGroupActionMenuButton;
    }
}

