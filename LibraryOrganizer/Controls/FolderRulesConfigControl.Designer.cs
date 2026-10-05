namespace LibraryOrganizer.Controls
{
    partial class FolderRulesConfigControl
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
            this.folderRulesActionsLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.addExcludedFolder = new System.Windows.Forms.Button();
            this.removeExcludedFolder = new System.Windows.Forms.Button();
            this.excludedFolderLabel = new System.Windows.Forms.Label();
            this.excludedFoldersList = new System.Windows.Forms.ListBox();
            this.profileViewModelBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.excludeFoldersBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.folderRulesActionsLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.excludeFoldersBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // folderRulesActionsLayout
            // 
            this.folderRulesActionsLayout.AutoSize = true;
            this.folderRulesActionsLayout.Controls.Add(this.addExcludedFolder);
            this.folderRulesActionsLayout.Controls.Add(this.removeExcludedFolder);
            this.folderRulesActionsLayout.Dock = System.Windows.Forms.DockStyle.Right;
            this.folderRulesActionsLayout.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.folderRulesActionsLayout.Location = new System.Drawing.Point(489, 28);
            this.folderRulesActionsLayout.Margin = new System.Windows.Forms.Padding(4);
            this.folderRulesActionsLayout.Name = "folderRulesActionsLayout";
            this.folderRulesActionsLayout.Size = new System.Drawing.Size(108, 464);
            this.folderRulesActionsLayout.TabIndex = 5;
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
            this.addExcludedFolder.Click += new System.EventHandler(this.addExcludedFolder_Click);
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
            this.removeExcludedFolder.Click += new System.EventHandler(this.removeExcludedFolder_Click);
            // 
            // excludedFolderLabel
            // 
            this.excludedFolderLabel.AutoSize = true;
            this.excludedFolderLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.excludedFolderLabel.Location = new System.Drawing.Point(0, 0);
            this.excludedFolderLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.excludedFolderLabel.Name = "excludedFolderLabel";
            this.excludedFolderLabel.Padding = new System.Windows.Forms.Padding(0, 6, 0, 6);
            this.excludedFolderLabel.Size = new System.Drawing.Size(365, 28);
            this.excludedFolderLabel.TabIndex = 3;
            this.excludedFolderLabel.Text = "Do not move books if they are located in the following folders";
            // 
            // excludedFoldersList
            // 
            this.excludedFoldersList.DataSource = this.excludeFoldersBindingSource;
            this.excludedFoldersList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.excludedFoldersList.FormattingEnabled = true;
            this.excludedFoldersList.ItemHeight = 16;
            this.excludedFoldersList.Location = new System.Drawing.Point(0, 28);
            this.excludedFoldersList.Name = "excludedFoldersList";
            this.excludedFoldersList.Size = new System.Drawing.Size(489, 464);
            this.excludedFoldersList.TabIndex = 6;
            // 
            // profileViewModelBindingSource
            // 
            this.profileViewModelBindingSource.DataSource = typeof(LibraryOrganizer.ViewModel.ProfileViewModel);
            // 
            // excludeFoldersBindingSource
            // 
            this.excludeFoldersBindingSource.DataMember = "ExcludeFolders";
            this.excludeFoldersBindingSource.DataSource = this.profileViewModelBindingSource;
            // 
            // FolderRulesConfigControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.excludedFoldersList);
            this.Controls.Add(this.folderRulesActionsLayout);
            this.Controls.Add(this.excludedFolderLabel);
            this.Name = "FolderRulesConfigControl";
            this.Size = new System.Drawing.Size(597, 492);
            this.folderRulesActionsLayout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.profileViewModelBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.excludeFoldersBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.FlowLayoutPanel folderRulesActionsLayout;
        private System.Windows.Forms.Button addExcludedFolder;
        private System.Windows.Forms.Button removeExcludedFolder;
        private System.Windows.Forms.Label excludedFolderLabel;
        private System.Windows.Forms.ListBox excludedFoldersList;
        private System.Windows.Forms.BindingSource profileViewModelBindingSource;
        private System.Windows.Forms.BindingSource excludeFoldersBindingSource;
    }
}
