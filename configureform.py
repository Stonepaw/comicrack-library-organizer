"""
configureform.py

Contains the Configure form.

Copyright 2010-2012 Stonepaw

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""





import clr

import System

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

import pyevent

from configformcontrols import *

from locommon import Mode, ExcludeGroup, ExcludeRule, field_to_name, name_to_field

from loforms import NewIllegalCharacterDialog, GetProfileNameDialog

from locommon import SCRIPTDIRECTORY, VERSION, ICON, check_excluded_folders, check_metadata_rules

import losettings

from losettings import Profile

from lobookmover import PathMaker

failed_items = System.Array[str](["Age Rating", "Alternate Count", "Alternate Number", "Alternate Series", "Black And White", "Characters", "Colorist", "Count", "Cover Artist", 
                "Editor", "Format", "Genre", "Imprint", "Inker", "Language", "Letterer", "Locations", "Main Character Or Team", "Manga", "Month", "Notes", "Number", "Penciller", "Publisher", 
                "Rating", "Read Percentage", "Review", "Scan Information", "Series", "Series Complete", "Series Group", "Start Month", "Start Year", "Story Arc", "Tags", "Teams", "Title", "Volume", "Web", "Writer", "Year"])

empty_substitution_items = System.Array[str](["Age Rating", "Alternate Count", "Alternate Number", "Alternate Series", "Black And White", "Characters", "Colorist", "Count", "Cover Artist", 
                          "Editor", "Format", "First Letter", "Genre", "Imprint", "Inker", "Language", "Letterer", "Locations", "Main Character Or Team", "Manga", "Month", "Number", "Penciller", "Publisher", 
                          "Rating", "Read Percentage", "Scan Information", "Series", "Series Complete", "Series Group", "Start Month", "Start Year", "Story Arc", "Tags", "Teams", "Title", "Volume", "Writer", "Year"])


class ConfigureForm(Form):

    def __init__(self, profiles, last_used_profile, books):
        print "Starting to load controls"
        self._insert_controls_dict = {}
        self._text_insert_controls_list = {}
        self._number_insert_controls_list = {}
        self._yes_no_insert_controls_list = {}
        self._multiple_value_insert_controls_list = {}
        self._calculated_insert_controls_list = {}

        

        self.initialize_component()

        print "Done the initialize function"


        print "done creating controls"

        #fields = System.Array[str]([control.Name for control in self._insert_controls_dict.values() if control.Name not in ("Month Number", "Alternate Series Multi")])


        InsertControl.Insert += self.insert_control_clicked

        print "Done creating controls"

        try:
            self.profile = profiles[last_used_profile]
        except KeyError:
            self.profile = profiles[profiles.keys()[0]]

        self.profiles = profiles

        self.load_profile()

        self._profile_selector.Items.AddRange(System.Array[object](self.profiles.keys()))
        self._profile_selector.SelectedItem = self.profile.Name
        
        self.path_maker = PathMaker(self, self.profile)

        self._preview_books = books
        self._preview_book = self._preview_books[0]

        self._preview_book_selector.Maximum = len(books) -1

        self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def initialize_component(self):
        self._toolstrip = System.Windows.Forms.ToolStrip()
        self._overview_page = System.Windows.Forms.Panel()
        self._okay = System.Windows.Forms.Button()
        self._cancel = System.Windows.Forms.Button()
        self._files_page = System.Windows.Forms.Panel()
        self._folders_page = System.Windows.Forms.Panel()
        self._rules_page = System.Windows.Forms.TabControl()
        self._options_page = System.Windows.Forms.TabControl()
        self._insert_controls = System.Windows.Forms.TabControl()
        self._preview_book_selector = System.Windows.Forms.VScrollBar()
        self._space_automatically = System.Windows.Forms.CheckBox()
        self._folder_browser_dialog = System.Windows.Forms.FolderBrowserDialog()
        self._toolstrip.SuspendLayout()
        self._overview_page.SuspendLayout()
        self._files_page.SuspendLayout()
        self._folders_page.SuspendLayout()
        self._rules_page.SuspendLayout()
        self._options_page.SuspendLayout()
        self._insert_controls.SuspendLayout()
        self.SuspendLayout()
        # 
        # okay
        # 
        self._okay.DialogResult = System.Windows.Forms.DialogResult.OK
        self._okay.Location = System.Drawing.Point(474, 433)
        self._okay.Name = "okay"
        self._okay.Size = System.Drawing.Size(75, 23)
        self._okay.TabIndex = 2
        self._okay.Text = "OK"
        self._okay.UseVisualStyleBackColor = True
        self._okay.Click += self.okay_clicked
        # 
        # cancel
        # 
        self._cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        self._cancel.Location = System.Drawing.Point(555, 433)
        self._cancel.Name = "cancel"
        self._cancel.Size = System.Drawing.Size(75, 23)
        self._cancel.TabIndex = 3
        self._cancel.Text = "Cancel"
        self._cancel.UseVisualStyleBackColor = True
        # 
        # space_automatically
        # 
        self._space_automatically.AutoSize = True
        self._space_automatically.BackColor = System.Drawing.Color.Transparent
        self._space_automatically.Checked = True
        self._space_automatically.CheckState = System.Windows.Forms.CheckState.Checked
        self._space_automatically.Location = System.Drawing.Point(5, 107)
        self._space_automatically.Name = "space_automatically"
        self._space_automatically.Size = System.Drawing.Size(188, 17)
        self._space_automatically.TabIndex = 8
        self._space_automatically.Text = "Space inserted fields automatically"
        self._space_automatically.UseVisualStyleBackColor = False
        # 
        # preview_book_selector
        # 
        self._preview_book_selector.Location = System.Drawing.Point(472, 60)
        self._preview_book_selector.Name = "preview_book_selector"
        self._preview_book_selector.Size = System.Drawing.Size(17, 40)
        self._preview_book_selector.TabIndex = 7
        self._preview_book_selector.SmallChange = 1
        self._preview_book_selector.LargeChange = 1
        self._preview_book_selector.Minimum = 0
        self._preview_book_selector.Value = 0
        self._preview_book_selector.ValueChanged += self.change_preview_book
        # 
        # files_page
        # 
        self._files_page.BackColor = System.Drawing.SystemColors.ControlLightLight
        self._files_page.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        self._files_page.Location = System.Drawing.Point(130, 10)
        self._files_page.Name = "files_page"
        self._files_page.Size = System.Drawing.Size(500, 420)
        self._files_page.TabIndex = 10
        self._files_page.Visible = False
        # 
        # folders_page
        # 
        self._folders_page.BackColor = System.Drawing.SystemColors.ControlLightLight
        self._folders_page.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        self._folders_page.Location = System.Drawing.Point(130, 10)
        self._folders_page.Name = "folders_page"
        self._folders_page.Size = System.Drawing.Size(500, 420)
        self._folders_page.TabIndex = 0
        self._folders_page.Visible = False
        #
        # options_page
        #
        self._options_page.Location = System.Drawing.Point(130, 10)
        self._options_page.Name = "options_page"
        self._options_page.SelectedIndex = 0
        self._options_page.Size = System.Drawing.Size(500, 420)
        self._options_page.TabIndex = 13
        self._options_page.Visible = False
        # 
        # rules_page
        # 
        self._rules_page.Location = System.Drawing.Point(130, 10)
        self._rules_page.Name = "rules_page"
        self._rules_page.SelectedIndex = 0
        self._rules_page.Size = System.Drawing.Size(500, 420)
        self._rules_page.TabIndex = 1
        self._rules_page.Visible = False
        #
        # Create the other controls
        #
        self.create_toolbar()
        self.create_overview_page()
        #
        # Icons
        #
        self._overview_button.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\home_32.png")
        self._files_button.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\page_text_32.png")   
        self._folders_button.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\folder_32.png")    
        self._rules_button.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\chart_32.png")   
        self._options_button.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\tools_32.png")
        self._profile_action.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\tools_32.png")
        self._profile_action_delete.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\close_16.png")
        self._profile_action_new.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\add_16.png")
        self._profile_action_rename.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\pencil_32.png")
        self._profile_action_duplicate.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\save_32.png")
        self._profile_action_import.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\arrow_down_16.png")
        self._profile_action_export_menu.Image = Bitmap.FromFile(SCRIPTDIRECTORY + "\\blue_arrow_up_32.png")
        # 
        # MainForm
        # 
        self.AcceptButton = self._okay
        self.AutoSize = True
        self.CancelButton = self._cancel
        self.ClientSize = System.Drawing.Size(639, 462)
        self.Controls.Add(self._options_page)
        self.Controls.Add(self._files_page)
        self.Controls.Add(self._folders_page)
        self.Controls.Add(self._overview_page)
        self.Controls.Add(self._rules_page)
        self.Controls.Add(self._cancel)
        self.Controls.Add(self._okay)
        self.Controls.Add(self._toolstrip)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterParent
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Name = "MainForm"
        self.Text = "Library Organizer " + str(VERSION)
        self.Icon = Icon(ICON)
        self._toolstrip.ResumeLayout(False)
        self._toolstrip.PerformLayout()
        self._overview_page.ResumeLayout(False)
        self._overview_page.PerformLayout()
        self._files_page.ResumeLayout(False)
        self._files_page.PerformLayout()
        self._folders_page.ResumeLayout(False)
        self._folders_page.PerformLayout()
        self._rules_page.ResumeLayout(False)
        self._options_page.ResumeLayout(False)
        self._insert_controls.ResumeLayout(False)
        self.ResumeLayout(False)


    #Methods to create controls. Most are called only when that tab/panel is accessed the first time

    def create_toolbar(self):
        """Creates the controls on the toolbar."""
        self._overview_button = System.Windows.Forms.ToolStripButton()
        self._files_button = System.Windows.Forms.ToolStripButton()
        self._folders_button = System.Windows.Forms.ToolStripButton()
        self._rules_button = System.Windows.Forms.ToolStripButton()
        self._options_button = System.Windows.Forms.ToolStripButton()
        self._profile_selector = System.Windows.Forms.ToolStripComboBox()
        self._profile_label = System.Windows.Forms.ToolStripLabel()
        self._profile_action = System.Windows.Forms.ToolStripDropDownButton()
        self._profile_action_new = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_delete = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_duplicate = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_import = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_export_menu = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_export = System.Windows.Forms.ToolStripMenuItem()
        self._profile_action_export_all = System.Windows.Forms.ToolStripMenuItem()
        self._toolstrip_seperator = System.Windows.Forms.ToolStripSeparator()
        self._profile_action_rename = System.Windows.Forms.ToolStripMenuItem()
        # 
        # toolstrip
        # 
        self._toolstrip.AutoSize = False
        self._toolstrip.BackColor = System.Drawing.SystemColors.Control
        self._toolstrip.Dock = System.Windows.Forms.DockStyle.Left
        self._toolstrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        self._toolstrip.Items.AddRange(System.Array[System.Windows.Forms.ToolStripItem](
            [self._overview_button,
            self._files_button,
            self._folders_button,
            self._rules_button,
            self._options_button,
            self._profile_action,
            self._profile_selector,
            self._profile_label,
            self._toolstrip_seperator]))
        self._toolstrip.Location = System.Drawing.Point(0, 0)
        self._toolstrip.Name = "toolstrip"
        self._toolstrip.Padding = System.Windows.Forms.Padding(10)
        self._toolstrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        self._toolstrip.ShowItemToolTips = False
        self._toolstrip.Size = System.Drawing.Size(130, 462)
        self._toolstrip.TabIndex = 0
        self._toolstrip.Text = "toolStrip1"
        # 
        # overview_button
        # 
        self._overview_button.AutoSize = False
        self._overview_button.Checked = True
        self._overview_button.CheckState = System.Windows.Forms.CheckState.Checked
        self._overview_button.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        self._overview_button.ImageTransparentColor = System.Drawing.Color.Magenta
        self._overview_button.Name = "overview_button"
        self._overview_button.Size = System.Drawing.Size(109, 70)
        self._overview_button.Tag = self._overview_page
        self._overview_button.Text = "Overview"
        self._overview_button.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        self._overview_button.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        self._overview_button.Click += self.change_page
        # 
        # files_button
        # 
        self._files_button.AutoSize = False
        self._files_button.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        self._files_button.ImageTransparentColor = System.Drawing.Color.Magenta
        self._files_button.Name = "files_button"
        self._files_button.Size = System.Drawing.Size(109, 70)
        self._files_button.Tag = self._files_page
        self._files_button.Text = "Files"
        self._files_button.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        self._files_button.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        self._files_button.Click += self.change_page
        # 
        # folders_button
        # 
        self._folders_button.AutoSize = False
        self._folders_button.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        self._folders_button.ImageTransparentColor = System.Drawing.Color.Magenta
        self._folders_button.Name = "folders_button"
        self._folders_button.Size = System.Drawing.Size(109, 70)
        self._folders_button.Tag = self._folders_page
        self._folders_button.Text = "Folders"
        self._folders_button.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        self._folders_button.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        self._folders_button.Click += self.change_page
        # 
        # rules_button
        # 
        self._rules_button.AutoSize = False
        self._rules_button.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        self._rules_button.ImageTransparentColor = System.Drawing.Color.Magenta
        self._rules_button.Name = "rules_button"
        self._rules_button.Size = System.Drawing.Size(109, 70)
        self._rules_button.Tag = self._rules_page
        self._rules_button.Text = "Rules"
        self._rules_button.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        self._rules_button.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        self._rules_button.Click += self.change_page
        # 
        # options_button
        # 
        self._options_button.AutoSize = False
        self._options_button.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        self._options_button.ImageTransparentColor = System.Drawing.Color.Magenta
        self._options_button.Name = "options_button"
        self._options_button.Size = System.Drawing.Size(109, 70)
        self._options_button.Tag = self._options_page
        self._options_button.Text = "Options"
        self._options_button.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        self._options_button.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        self._options_button.Click += self.change_page
        # 
        # profile_selector
        # 
        self._profile_selector.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        self._profile_selector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._profile_selector.FlatStyle = System.Windows.Forms.FlatStyle.System
        self._profile_selector.Name = "profile_selector"
        self._profile_selector.Sorted = True
        self._profile_selector.MaxDropDownItems = 5
        self._profile_selector.IntegralHeight = False
        self._profile_selector.Size = System.Drawing.Size(107, 23)
        self._profile_selector.ComboBox.SelectionChangeCommitted += self.profile_selector_selection_change_commited
        # 
        # profile_label
        # 
        self._profile_label.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        self._profile_label.Name = "profile_label"
        self._profile_label.Size = System.Drawing.Size(109, 15)
        self._profile_label.Text = "Profile"
        # 
        # profile_action
        # 
        self._profile_action.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        self._profile_action.DropDownItems.AddRange(System.Array[System.Windows.Forms.ToolStripItem](
            [self._profile_action_new,
            self._profile_action_rename,
            self._profile_action_delete,
            self._profile_action_duplicate,
            self._profile_action_import,
            self._profile_action_export_menu]))
        self._profile_action.ImageTransparentColor = System.Drawing.Color.Magenta
        self._profile_action.Name = "profile_action"
        self._profile_action.Size = System.Drawing.Size(109, 19)
        self._profile_action.Text = "Profile Action"
        # 
        # profile_action_new
        # 
        self._profile_action_new.Name = "profile_action_new"
        self._profile_action_new.Size = System.Drawing.Size(117, 22)
        self._profile_action_new.Text = "New"
        self._profile_action_new.Click += self.add_new_profile
        # 
        # profile_action_rename
        # 
        self._profile_action_rename.Name = "profile_action_rename"
        self._profile_action_rename.Size = System.Drawing.Size(117, 22)
        self._profile_action_rename.Text = "Rename"
        self._profile_action_rename.Click += self.rename_profile
        # 
        # profile_action_delete
        # 
        self._profile_action_delete.Name = "profile_action_delete"
        self._profile_action_delete.Size = System.Drawing.Size(117, 22)
        self._profile_action_delete.Text = "Delete"
        self._profile_action_delete.Click += self.delete_profile
        # 
        # profile_action_duplicate
        # 
        self._profile_action_duplicate.Name = "profile_action_duplicate"
        self._profile_action_duplicate.Size = System.Drawing.Size(117, 22)
        self._profile_action_duplicate.Text = "Duplicate"
        self._profile_action_duplicate.Click += self.duplicate_profile
        # 
        # profile_action_import
        # 
        self._profile_action_import.Name = "profile_action_import"
        self._profile_action_import.Size = System.Drawing.Size(117, 22)
        self._profile_action_import.Text = "Import"
        self._profile_action_import.Click += self.import_profile
        # 
        # profile_action_export_menu
        # 
        self._profile_action_export_menu.DropDownItems.AddRange(System.Array[System.Windows.Forms.ToolStripItem](
            [self._profile_action_export,
            self._profile_action_export_all]))
        self._profile_action_export_menu.Name = "profile_action_export_menu"
        self._profile_action_export_menu.Size = System.Drawing.Size(117, 22)
        self._profile_action_export_menu.Text = "Export"
        # 
        # profile_action_export
        # 
        self._profile_action_export.Name = "profile_action_export"
        self._profile_action_export.Size = System.Drawing.Size(124, 22)
        self._profile_action_export.Text = "Export"
        self._profile_action_export.Click += self.export_profile
        # 
        # profile_action_export_all
        # 
        self._profile_action_export_all.Name = "profile_action_export_all"
        self._profile_action_export_all.Size = System.Drawing.Size(124, 22)
        self._profile_action_export_all.Text = "Export All"
        self._profile_action_export_all.Click += self.export_all_profiles
        # 
        # toolstrip_seperator
        # 
        self._toolstrip_seperator.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        self._toolstrip_seperator.Name = "toolstrip_seperator"
        self._toolstrip_seperator.Size = System.Drawing.Size(109, 6)


    def create_overview_page(self):
        """Creates the controls for the overview page."""        
        self._mode_move = System.Windows.Forms.RadioButton()
        self._mode_copy = System.Windows.Forms.RadioButton()
        self._mode_simulate = System.Windows.Forms.RadioButton()
        self._use_file_organization = System.Windows.Forms.CheckBox()
        self._use_folder_organization = System.Windows.Forms.CheckBox()
        self._label_simulate = System.Windows.Forms.Label()
        self._mode_groupbox = System.Windows.Forms.GroupBox()
        self._copy_option = System.Windows.Forms.CheckBox()
        self._label_base_folder = System.Windows.Forms.Label()
        self._base_folder_path = System.Windows.Forms.TextBox()
        self._browse_button = System.Windows.Forms.Button()
        self._copy_fileless_custom_thumbnails = System.Windows.Forms.CheckBox()
        self._fileless_image_format = System.Windows.Forms.ComboBox()
        self._fileless_image_format_label = System.Windows.Forms.Label()
        self._mode_groupbox.SuspendLayout()
        # 
        # overview_page
        # 
        self._overview_page.BackColor = System.Drawing.SystemColors.ControlLightLight
        self._overview_page.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        self._overview_page.Controls.Add(self._fileless_image_format_label)
        self._overview_page.Controls.Add(self._fileless_image_format)
        self._overview_page.Controls.Add(self._copy_fileless_custom_thumbnails)
        self._overview_page.Controls.Add(self._browse_button)
        self._overview_page.Controls.Add(self._base_folder_path)
        self._overview_page.Controls.Add(self._label_base_folder)
        self._overview_page.Controls.Add(self._mode_groupbox)
        self._overview_page.Controls.Add(self._use_folder_organization)
        self._overview_page.Controls.Add(self._use_file_organization)
        self._overview_page.Location = System.Drawing.Point(130, 10)
        self._overview_page.Margin = System.Windows.Forms.Padding(0)
        self._overview_page.Name = "overview_page"
        self._overview_page.Size = System.Drawing.Size(500, 420)
        self._overview_page.TabIndex = 1
        # 
        # mode_move
        # 
        self._mode_move.AutoSize = True
        self._mode_move.Checked = True
        self._mode_move.Location = System.Drawing.Point(15, 19)
        self._mode_move.Name = "mode_move"
        self._mode_move.Size = System.Drawing.Size(52, 17)
        self._mode_move.TabIndex = 0
        self._mode_move.TabStop = True
        self._mode_move.Text = "Move"
        self._mode_move.UseVisualStyleBackColor = True
        # 
        # mode_copy
        # 
        self._mode_copy.AutoSize = True
        self._mode_copy.Location = System.Drawing.Point(188, 19)
        self._mode_copy.Name = "mode_copy"
        self._mode_copy.Size = System.Drawing.Size(49, 17)
        self._mode_copy.TabIndex = 1
        self._mode_copy.Text = "Copy"
        self._mode_copy.UseVisualStyleBackColor = True
        self._mode_copy.CheckedChanged += self.mode_copy_checked_changed
        # 
        # mode_simulate
        # 
        self._mode_simulate.AutoSize = True
        self._mode_simulate.Location = System.Drawing.Point(361, 19)
        self._mode_simulate.Name = "mode_simulate"
        self._mode_simulate.TabIndex = 2
        self._mode_simulate.Text = "Simulate"
        self._mode_simulate.UseVisualStyleBackColor = True
        # 
        # use_file_organization
        # 
        self._use_file_organization.AutoSize = True
        self._use_file_organization.Checked = True
        self._use_file_organization.CheckState = System.Windows.Forms.CheckState.Checked
        self._use_file_organization.Location = System.Drawing.Point(20, 172)
        self._use_file_organization.Name = "use_file_organization"
        self._use_file_organization.Size = System.Drawing.Size(121, 17)
        self._use_file_organization.TabIndex = 3
        self._use_file_organization.Tag = self._files_button
        self._use_file_organization.Text = "Use file organization"
        self._use_file_organization.UseVisualStyleBackColor = True
        self._use_file_organization.CheckedChanged += self.use_organization_check_changed
        # 
        # use_folder_organization
        # 
        self._use_folder_organization.AutoSize = True
        self._use_folder_organization.Checked = True
        self._use_folder_organization.CheckState = System.Windows.Forms.CheckState.Checked
        self._use_folder_organization.Location = System.Drawing.Point(193, 172)
        self._use_folder_organization.Name = "use_folder_organization"
        self._use_folder_organization.Size = System.Drawing.Size(134, 17)
        self._use_folder_organization.TabIndex = 4
        self._use_folder_organization.Tag = self._folders_button
        self._use_folder_organization.Text = "Use folder organization"
        self._use_folder_organization.UseVisualStyleBackColor = True
        self._use_folder_organization.CheckedChanged += self.use_organization_check_changed
        # 
        # label_simulate
        # 
        self._label_simulate.AutoSize = True
        self._label_simulate.Location = System.Drawing.Point(330, 39)
        self._label_simulate.Name = "label_simulate"
        self._label_simulate.Size = System.Drawing.Size(122, 26)
        self._label_simulate.TabIndex = 5
        self._label_simulate.Text = """no files touched\ncomplete log file created"""
        self._label_simulate.TextAlign = System.Drawing.ContentAlignment.TopCenter
        # 
        # mode_groupbox
        # 
        self._mode_groupbox.Controls.Add(self._copy_option)
        self._mode_groupbox.Controls.Add(self._mode_move)
        self._mode_groupbox.Controls.Add(self._label_simulate)
        self._mode_groupbox.Controls.Add(self._mode_copy)
        self._mode_groupbox.Controls.Add(self._mode_simulate)
        self._mode_groupbox.Location = System.Drawing.Point(5, 5)
        self._mode_groupbox.Margin = System.Windows.Forms.Padding(0)
        self._mode_groupbox.Name = "mode_groupbox"
        self._mode_groupbox.Size = System.Drawing.Size(490, 75)
        self._mode_groupbox.TabIndex = 6
        self._mode_groupbox.TabStop = False
        self._mode_groupbox.Text = "Mode"
        # 
        # copy_option
        # 
        self._copy_option.AutoSize = True
        self._copy_option.Enabled = False
        self._copy_option.Location = System.Drawing.Point(138, 44)
        self._copy_option.Name = "copy_option"
        self._copy_option.Size = System.Drawing.Size(149, 17)
        self._copy_option.TabIndex = 6
        self._copy_option.Text = "Add copied book to library"
        self._copy_option.UseVisualStyleBackColor = True
        # 
        # label_base_folder
        # 
        self._label_base_folder.AutoSize = True
        self._label_base_folder.Location = System.Drawing.Point(7, 112)
        self._label_base_folder.Name = "label_base_folder"
        self._label_base_folder.Size = System.Drawing.Size(66, 13)
        self._label_base_folder.TabIndex = 7
        self._label_base_folder.Text = "Base Folder:"
        # 
        # base_folder_path
        # 
        self._base_folder_path.Location = System.Drawing.Point(79, 108)
        self._base_folder_path.Name = "base_folder_path"
        self._base_folder_path.Size = System.Drawing.Size(327, 20)
        self._base_folder_path.TabIndex = 8
        self._base_folder_path.Text = ""
        # 
        # browse_button
        # 
        self._browse_button.Location = System.Drawing.Point(410, 107)
        self._browse_button.Name = "browse_button"
        self._browse_button.Size = System.Drawing.Size(75, 23)
        self._browse_button.TabIndex = 9
        self._browse_button.Text = "Browse"
        self._browse_button.UseVisualStyleBackColor = True
        self._browse_button.Tag = self._base_folder_path
        self._browse_button.Click += self.add_folder_path_to_text_box
        # 
        # copy_fileless_custom_thumbnails
        # 
        self._copy_fileless_custom_thumbnails.Location = System.Drawing.Point(20, 237)
        self._copy_fileless_custom_thumbnails.Name = "copy_fileless_custom_thumbnails"
        self._copy_fileless_custom_thumbnails.Size = System.Drawing.Size(324, 35)
        self._copy_fileless_custom_thumbnails.TabIndex = 10
        self._copy_fileless_custom_thumbnails.Text = "Copy fileless book custom thumbnails to the calculated path. (Does not affect the original image)"
        self._copy_fileless_custom_thumbnails.UseVisualStyleBackColor = True
        # 
        # fileless_image_format
        # 
        self._fileless_image_format.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._fileless_image_format.FormattingEnabled = True
        self._fileless_image_format.Items.AddRange(System.Array[System.Object](
            [".bmp",
            ".jpg",
            ".png"]))
        self._fileless_image_format.Location = System.Drawing.Point(366, 253)
        self._fileless_image_format.Name = "fileless_image_format"
        self._fileless_image_format.Size = System.Drawing.Size(66, 21)
        self._fileless_image_format.TabIndex = 11
        # 
        # fileless_image_format_label
        # 
        self._fileless_image_format_label.AutoSize = True
        self._fileless_image_format_label.Location = System.Drawing.Point(361, 234)
        self._fileless_image_format_label.Name = "fileless_image_format_label"
        self._fileless_image_format_label.Size = System.Drawing.Size(71, 13)
        self._fileless_image_format_label.TabIndex = 12
        self._fileless_image_format_label.Text = "Image format:"
        self._mode_groupbox.ResumeLayout(False)


    def create_files_page(self):
        """Creates the controls on the files page."""

        self._files_page.SuspendLayout()

        self._file_structure = System.Windows.Forms.TextBox()
        self._label_file_structure = System.Windows.Forms.Label()
        self._label_file_preview = System.Windows.Forms.Label()
        self._file_preview = System.Windows.Forms.Label()
        # 
        # file_structure
        # 
        self._file_structure.HideSelection = False
        self._file_structure.Location = System.Drawing.Point(83, 10)
        self._file_structure.Multiline = True
        self._file_structure.Name = "file_structure"
        self._file_structure.Size = System.Drawing.Size(407, 40)
        self._file_structure.TabIndex = 1
        self._file_structure.TextChanged += self.update_template_text
        # 
        # label_file_structure
        # 
        self._label_file_structure.AutoSize = True
        self._label_file_structure.Location = System.Drawing.Point(5, 24)
        self._label_file_structure.Name = "label_file_structure"
        self._label_file_structure.Size = System.Drawing.Size(72, 13)
        self._label_file_structure.TabIndex = 0
        self._label_file_structure.Text = "File Structure:"
        # 
        # label_file_preview
        # 
        self._label_file_preview.AutoSize = True
        self._label_file_preview.Location = System.Drawing.Point(5, 60)
        self._label_file_preview.Name = "label_file_preview"
        self._label_file_preview.Size = System.Drawing.Size(48, 13)
        self._label_file_preview.TabIndex = 2
        self._label_file_preview.Text = "Preview:"
        # 
        # file_preview
        # 
        self._file_preview.Location = System.Drawing.Point(59, 60)
        self._file_preview.Name = "file_preview"
        self._file_preview.Size = System.Drawing.Size(410, 40)
        self._file_preview.TabIndex = 3

        self._files_page.Controls.Add(self._preview_book_selector)
        self._files_page.Controls.Add(self._file_preview)
        self._files_page.Controls.Add(self._label_file_preview)
        self._files_page.Controls.Add(self._label_file_structure)
        self._files_page.Controls.Add(self._file_structure)

        self.load_files_page_settings()

        self._files_page.ResumeLayout()


    def create_folders_page(self):
        """Creates the controls on the folders page."""
        
        self._folders_page.SuspendLayout()

        self._folder_structure = System.Windows.Forms.TextBox()
        self._label_folder_structure = System.Windows.Forms.Label()
        self._label_folder_preview = System.Windows.Forms.Label()
        self._folder_preview = System.Windows.Forms.Label()
        self._insert_folder_seperator = System.Windows.Forms.Button()
        #
        # 
        #
        self._folders_page.Controls.Add(self._folder_preview)
        self._folders_page.Controls.Add(self._label_folder_preview)
        self._folders_page.Controls.Add(self._label_folder_structure)
        self._folders_page.Controls.Add(self._folder_structure)
        self._folders_page.Controls.Add(self._insert_folder_seperator)
        # 
        # folder_structure
        # 
        self._folder_structure.HideSelection = False
        self._folder_structure.Location = System.Drawing.Point(94, 10)
        self._folder_structure.Multiline = True
        self._folder_structure.Name = "folder_structure"
        self._folder_structure.Size = System.Drawing.Size(395, 40)
        self._folder_structure.TabIndex = 1
        self._folder_structure.TextChanged += self.update_template_text
        # 
        # label_folder_structure
        # 
        self._label_folder_structure.AutoSize = True
        self._label_folder_structure.Location = System.Drawing.Point(3, 24)
        self._label_folder_structure.Name = "label_folder_structure"
        self._label_folder_structure.Size = System.Drawing.Size(85, 13)
        self._label_folder_structure.TabIndex = 0
        self._label_folder_structure.Text = "Folder Structure:"
        # 
        # label_folder_preview
        # 
        self._label_folder_preview.AutoSize = True
        self._label_folder_preview.Location = System.Drawing.Point(5, 60)
        self._label_folder_preview.Name = "label_folder_preview"
        self._label_folder_preview.Size = System.Drawing.Size(48, 13)
        self._label_folder_preview.TabIndex = 2
        self._label_folder_preview.Text = "Preview:"
        # 
        # folder_preview
        # 
        self._folder_preview.Location = System.Drawing.Point(59, 60)
        self._folder_preview.Name = "folder_preview"
        self._folder_preview.Size = System.Drawing.Size(410, 40)
        self._folder_preview.TabIndex = 3
        # 
        # insert_folder_seperator
        # 
        self._insert_folder_seperator.Location = System.Drawing.Point(220, 104)
        self._insert_folder_seperator.Name = "insert_folder_seperator"
        self._insert_folder_seperator.Size = System.Drawing.Size(98, 23)
        self._insert_folder_seperator.TabIndex = 5
        self._insert_folder_seperator.Text = "Folder Seperator"
        self._insert_folder_seperator.UseVisualStyleBackColor = True
        self._insert_folder_seperator.Click += self.insert_folder_seperator_clicked

        self.load_folders_page_settings()

        self._folders_page.ResumeLayout()


    def create_rules_page(self):
        """Creates the controls in the rule page."""
        self._metadata_rules_page = System.Windows.Forms.TabPage()
        self._folder_rules_page = System.Windows.Forms.TabPage()
        self._add_excluded_folder = System.Windows.Forms.Button()
        self._remove_excluded_folder = System.Windows.Forms.Button()
        self._excluded_folders_list = System.Windows.Forms.ListBox()
        self._excluded_folder_label = System.Windows.Forms.Label()
        self._metadata_rules_container = System.Windows.Forms.FlowLayoutPanel()
        self._metadata_rules_label1 = System.Windows.Forms.Label()
        self._metadata_rules_mode = System.Windows.Forms.ComboBox()
        self._metadata_rules_operator = System.Windows.Forms.ComboBox()
        self._metadata_rules_label2 = System.Windows.Forms.Label()
        self._metadata_rules_add_group = System.Windows.Forms.Button()
        self._metadata_rules_add_rule = System.Windows.Forms.Button()

        self._rules_page.SuspendLayout()
        self._metadata_rules_page.SuspendLayout()
        self._folder_rules_page.SuspendLayout()
        #
        # rules_page
        #
        self._rules_page.Controls.Add(self._metadata_rules_page)
        self._rules_page.Controls.Add(self._folder_rules_page)
        # 
        # metadata_rules_page
        # 
        self._metadata_rules_page.Controls.Add(self._metadata_rules_add_rule)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_add_group)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_label2)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_operator)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_mode)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_label1)
        self._metadata_rules_page.Controls.Add(self._metadata_rules_container)
        self._metadata_rules_page.Location = System.Drawing.Point(4, 22)
        self._metadata_rules_page.Name = "metadata_rules_page"
        self._metadata_rules_page.Size = System.Drawing.Size(492, 394)
        self._metadata_rules_page.TabIndex = 0
        self._metadata_rules_page.Text = "Metadata Rules"
        self._metadata_rules_page.UseVisualStyleBackColor = True
        # 
        # folder_rules_page
        # 
        self._folder_rules_page.Controls.Add(self._excluded_folder_label)
        self._folder_rules_page.Controls.Add(self._excluded_folders_list)
        self._folder_rules_page.Controls.Add(self._remove_excluded_folder)
        self._folder_rules_page.Controls.Add(self._add_excluded_folder)
        self._folder_rules_page.Location = System.Drawing.Point(4, 22)
        self._folder_rules_page.Name = "folder_rules_page"
        self._folder_rules_page.Size = System.Drawing.Size(492, 394)
        self._folder_rules_page.TabIndex = 1
        self._folder_rules_page.Text = "Folder Rules"
        self._folder_rules_page.UseVisualStyleBackColor = True
        # 
        # add_excluded_folder
        # 
        self._add_excluded_folder.Location = System.Drawing.Point(403, 26)
        self._add_excluded_folder.Name = "add_excluded_folder"
        self._add_excluded_folder.Size = System.Drawing.Size(75, 23)
        self._add_excluded_folder.TabIndex = 2
        self._add_excluded_folder.Tag = self._excluded_folders_list
        self._add_excluded_folder.Text = "Add"
        self._add_excluded_folder.UseVisualStyleBackColor = True
        self._add_excluded_folder.Click += self.add_folder_path_to_list
        # 
        # remove_excluded_folder
        # 
        self._remove_excluded_folder.Location = System.Drawing.Point(403, 62)
        self._remove_excluded_folder.Name = "remove_excluded_folder"
        self._remove_excluded_folder.Size = System.Drawing.Size(75, 23)
        self._remove_excluded_folder.TabIndex = 3
        self._remove_excluded_folder.Tag = self._excluded_folders_list
        self._remove_excluded_folder.Text = "Remove"
        self._remove_excluded_folder.UseVisualStyleBackColor = True
        self._remove_excluded_folder.Click += self.remove_folder_path_from_list
        # 
        # excluded_folders_list
        # 
        self._excluded_folders_list.FormattingEnabled = True
        self._excluded_folders_list.Location = System.Drawing.Point(8, 26)
        self._excluded_folders_list.Name = "excluded_folders_list"
        self._excluded_folders_list.Size = System.Drawing.Size(389, 355)
        self._excluded_folders_list.Sorted = True
        self._excluded_folders_list.TabIndex = 1
        # 
        # excluded_folder_label
        # 
        self._excluded_folder_label.AutoSize = True
        self._excluded_folder_label.Location = System.Drawing.Point(8, 7)
        self._excluded_folder_label.Name = "excluded_folder_label"
        self._excluded_folder_label.Size = System.Drawing.Size(294, 13)
        self._excluded_folder_label.TabIndex = 0
        self._excluded_folder_label.Text = "Do not move books if they are located in the following folders"
        # 
        # metadata_rules_container
        # 
        self._metadata_rules_container.AutoScroll = True
        self._metadata_rules_container.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        self._metadata_rules_container.Location = System.Drawing.Point(0, 49)
        self._metadata_rules_container.Name = "metadata_rules_container"
        self._metadata_rules_container.Size = System.Drawing.Size(492, 345)
        self._metadata_rules_container.TabIndex = 6
        self._metadata_rules_container.WrapContents = False
        # 
        # metadata_rules_label1
        # 
        self._metadata_rules_label1.AutoSize = True
        self._metadata_rules_label1.Location = System.Drawing.Point(66, 16)
        self._metadata_rules_label1.Name = "metadata_rules_label1"
        self._metadata_rules_label1.Size = System.Drawing.Size(118, 13)
        self._metadata_rules_label1.TabIndex = 1
        self._metadata_rules_label1.Text = "move books that match"
        # 
        # metadata_rules_mode
        # 
        self._metadata_rules_mode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._metadata_rules_mode.FormattingEnabled = True
        self._metadata_rules_mode.Items.AddRange(System.Array[System.Object](
            ["Do not",
            "Only"]))
        self._metadata_rules_mode.Location = System.Drawing.Point(8, 12)
        self._metadata_rules_mode.Name = "metadata_rules_mode"
        self._metadata_rules_mode.Size = System.Drawing.Size(55, 21)
        self._metadata_rules_mode.TabIndex = 0
        # 
        # metadata_rules_operator
        # 
        self._metadata_rules_operator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._metadata_rules_operator.FormattingEnabled = True
        self._metadata_rules_operator.Items.AddRange(System.Array[System.Object](
            ["All",
            "Any"]))
        self._metadata_rules_operator.Location = System.Drawing.Point(188, 12)
        self._metadata_rules_operator.Name = "metadata_rules_operator"
        self._metadata_rules_operator.Size = System.Drawing.Size(43, 21)
        self._metadata_rules_operator.TabIndex = 2
        # 
        # metadata_rules_label2
        # 
        self._metadata_rules_label2.AutoSize = True
        self._metadata_rules_label2.Location = System.Drawing.Point(235, 15)
        self._metadata_rules_label2.Name = "metadata_rules_label2"
        self._metadata_rules_label2.Size = System.Drawing.Size(106, 13)
        self._metadata_rules_label2.TabIndex = 3
        self._metadata_rules_label2.Text = "of the following rules."
        # 
        # metadata_rules_add_group
        # 
        self._metadata_rules_add_group.AutoSize = True
        self._metadata_rules_add_group.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._metadata_rules_add_group.Location = System.Drawing.Point(347, 11)
        self._metadata_rules_add_group.Name = "metadata_rules_add_group"
        self._metadata_rules_add_group.Size = System.Drawing.Size(68, 23)
        self._metadata_rules_add_group.TabIndex = 4
        self._metadata_rules_add_group.Text = "Add Group"
        self._metadata_rules_add_group.UseVisualStyleBackColor = True
        self._metadata_rules_add_group.Click += self.add_metadata_rule_group
        # 
        # metadata_rules_add_rule
        # 
        self._metadata_rules_add_rule.AutoSize = True
        self._metadata_rules_add_rule.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._metadata_rules_add_rule.Location = System.Drawing.Point(421, 11)
        self._metadata_rules_add_rule.Name = "metadata_rules_add_rule"
        self._metadata_rules_add_rule.Size = System.Drawing.Size(61, 23)
        self._metadata_rules_add_rule.TabIndex = 5
        self._metadata_rules_add_rule.Text = "Add Rule"
        self._metadata_rules_add_rule.UseVisualStyleBackColor = True
        self._metadata_rules_add_rule.Click += self.add_metadata_rule

        self.load_rules_page_settings()

        self._metadata_rules_page.ResumeLayout()
        self._folder_rules_page.ResumeLayout()
        self._rules_page.ResumeLayout()


    def create_options_page(self):
        """Creates the controls on the options page."""
        self._options_page_options_tab = System.Windows.Forms.TabPage()
        self._options_page_empty_values_tab = System.Windows.Forms.TabPage()
        self._replace_multiple_spaces = System.Windows.Forms.CheckBox()
        self._insert_multiple_value_field_when_one = System.Windows.Forms.CheckBox()
        self._remove_empty_folders = System.Windows.Forms.CheckBox()
        self._copy_read_percentage = System.Windows.Forms.CheckBox()
        self._month_label1 = System.Windows.Forms.Label()
        self._month_number = System.Windows.Forms.ComboBox()
        self._month_name = System.Windows.Forms.TextBox()
        self._illegal_character_label1 = System.Windows.Forms.Label()
        self._remove_illegal_character = System.Windows.Forms.Button()
        self._add_illegal_character = System.Windows.Forms.Button()
        self._month_label2 = System.Windows.Forms.Label()
        self._illegal_character_label2 = System.Windows.Forms.Label()
        self._illegal_character_replacement = System.Windows.Forms.TextBox()
        self._illegal_character_selector = System.Windows.Forms.ComboBox()
        self._empty_folder_exceptions_list = System.Windows.Forms.ListBox()
        self._remove_empty_folders_label = System.Windows.Forms.Label()
        self._add_empty_folder_exception = System.Windows.Forms.Button()
        self._remove_empty_folder_exception = System.Windows.Forms.Button()
        self._empty_folder_name_label = System.Windows.Forms.Label()
        self._empty_folder_name = System.Windows.Forms.TextBox()
        self._failed_empty_selection = System.Windows.Forms.CheckedListBox()
        self._failed_empty_checkbox = System.Windows.Forms.CheckBox()
        self._failed_empty_folder = System.Windows.Forms.TextBox()
        self._failed_empty_browse = System.Windows.Forms.Button()
        self._empty_substitution_label = System.Windows.Forms.Label()
        self._empty_substitution_field = System.Windows.Forms.ComboBox()
        self._empty_substitution_label1 = System.Windows.Forms.Label()
        self._empty_substitution_value = System.Windows.Forms.TextBox()
        self._empty_substitution_label2 = System.Windows.Forms.Label()
        self._move_failed_empty = System.Windows.Forms.CheckBox()

        self._options_page.SuspendLayout()
        self._options_page_options_tab.SuspendLayout()
        self._options_page_empty_values_tab.SuspendLayout()
        # 
        # options_page
        # 
        self._options_page.Controls.Add(self._options_page_options_tab)
        self._options_page.Controls.Add(self._options_page_empty_values_tab)
        # 
        # options_page_options_tab
        # 
        self._options_page_options_tab.Controls.Add(self._remove_empty_folder_exception)
        self._options_page_options_tab.Controls.Add(self._add_empty_folder_exception)
        self._options_page_options_tab.Controls.Add(self._copy_read_percentage)
        self._options_page_options_tab.Controls.Add(self._remove_empty_folders_label)
        self._options_page_options_tab.Controls.Add(self._empty_folder_exceptions_list)
        self._options_page_options_tab.Controls.Add(self._illegal_character_selector)
        self._options_page_options_tab.Controls.Add(self._illegal_character_replacement)
        self._options_page_options_tab.Controls.Add(self._illegal_character_label2)
        self._options_page_options_tab.Controls.Add(self._month_label2)
        self._options_page_options_tab.Controls.Add(self._add_illegal_character)
        self._options_page_options_tab.Controls.Add(self._remove_illegal_character)
        self._options_page_options_tab.Controls.Add(self._illegal_character_label1)
        self._options_page_options_tab.Controls.Add(self._month_name)
        self._options_page_options_tab.Controls.Add(self._month_number)
        self._options_page_options_tab.Controls.Add(self._month_label1)
        self._options_page_options_tab.Controls.Add(self._remove_empty_folders)
        self._options_page_options_tab.Controls.Add(self._insert_multiple_value_field_when_one)
        self._options_page_options_tab.Controls.Add(self._replace_multiple_spaces)
        self._options_page_options_tab.Location = System.Drawing.Point(4, 22)
        self._options_page_options_tab.Name = "options_page_options_tab"
        self._options_page_options_tab.Size = System.Drawing.Size(492, 394)
        self._options_page_options_tab.TabIndex = 0
        self._options_page_options_tab.Text = "Options"
        self._options_page_options_tab.UseVisualStyleBackColor = True
        # 
        # options_page_empty_values_tab
        # 
        self._options_page_empty_values_tab.Controls.Add(self._move_failed_empty)
        self._options_page_empty_values_tab.Controls.Add(self._empty_substitution_label2)
        self._options_page_empty_values_tab.Controls.Add(self._empty_substitution_value)
        self._options_page_empty_values_tab.Controls.Add(self._empty_substitution_label1)
        self._options_page_empty_values_tab.Controls.Add(self._empty_substitution_field)
        self._options_page_empty_values_tab.Controls.Add(self._empty_substitution_label)
        self._options_page_empty_values_tab.Controls.Add(self._failed_empty_browse)
        self._options_page_empty_values_tab.Controls.Add(self._failed_empty_folder)
        self._options_page_empty_values_tab.Controls.Add(self._failed_empty_checkbox)
        self._options_page_empty_values_tab.Controls.Add(self._failed_empty_selection)
        self._options_page_empty_values_tab.Controls.Add(self._empty_folder_name)
        self._options_page_empty_values_tab.Controls.Add(self._empty_folder_name_label)
        self._options_page_empty_values_tab.Location = System.Drawing.Point(4, 22)
        self._options_page_empty_values_tab.Name = "options_page_empty_values_tab"
        self._options_page_empty_values_tab.Size = System.Drawing.Size(492, 394)
        self._options_page_empty_values_tab.TabIndex = 1
        self._options_page_empty_values_tab.Text = "Empty values"
        self._options_page_empty_values_tab.UseVisualStyleBackColor = True
        # 
        # replace_multiple_spaces
        # 
        self._replace_multiple_spaces.AutoSize = True
        self._replace_multiple_spaces.Location = System.Drawing.Point(17, 28)
        self._replace_multiple_spaces.Name = "replace_multiple_spaces"
        self._replace_multiple_spaces.Size = System.Drawing.Size(234, 17)
        self._replace_multiple_spaces.TabIndex = 0
        self._replace_multiple_spaces.Text = "Replace multiple spaces with a single space."
        self._replace_multiple_spaces.UseVisualStyleBackColor = True
        #
        # copy_read_percentage
        #
        self._copy_read_percentage.Location = System.Drawing.Point(17, 63)
        self._copy_read_percentage.Size = System.Drawing.Size(411, 24)
        self._copy_read_percentage.Name = "copy_read_percentage"
        self._copy_read_percentage.Text = "When overwriting an existing file, copy the read percentage to the new file."
        self._copy_read_percentage.TabIndex = 1
        self._copy_read_percentage.UseVisualStyleBackColor = True
        # 
        # insert_multiple_value_field_when_one
        # 
        self._insert_multiple_value_field_when_one.AutoSize = True
        self._insert_multiple_value_field_when_one.Location = System.Drawing.Point(17, 105)
        self._insert_multiple_value_field_when_one.Name = "insert_multiple_value_field_when_one"
        self._insert_multiple_value_field_when_one.Size = System.Drawing.Size(381, 17)
        self._insert_multiple_value_field_when_one.TabIndex = 2
        self._insert_multiple_value_field_when_one.Text = "If there is only one value in a multiple value field then insert it without asking."
        self._insert_multiple_value_field_when_one.UseVisualStyleBackColor = True
        # 
        # remove_empty_folders
        # 
        self._remove_empty_folders.AutoSize = True
        self._remove_empty_folders.Checked = True
        self._remove_empty_folders.CheckState = System.Windows.Forms.CheckState.Checked
        self._remove_empty_folders.Location = System.Drawing.Point(17, 200)
        self._remove_empty_folders.Name = "remove_empty_folders"
        self._remove_empty_folders.Size = System.Drawing.Size(134, 17)
        self._remove_empty_folders.TabIndex = 13
        self._remove_empty_folders.Text = " Remove empty folders"
        self._remove_empty_folders.UseVisualStyleBackColor = True
        self._remove_empty_folders.CheckedChanged += self.remove_empty_folders_checked_changed
        # 
        # month_label1
        # 
        self._month_label1.AutoSize = True
        self._month_label1.Location = System.Drawing.Point(17, 140)
        self._month_label1.Name = "month_label1"
        self._month_label1.Size = System.Drawing.Size(37, 13)
        self._month_label1.TabIndex = 3
        self._month_label1.Text = "Month"
        # 
        # month_number
        # 
        self._month_number.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._month_number.FormattingEnabled = True
        self._month_number.Location = System.Drawing.Point(60, 136)
        self._month_number.Name = "month_number"
        self._month_number.Size = System.Drawing.Size(38, 21)
        self._month_number.TabIndex = 4
        self._month_number.SelectedIndexChanged += self.month_number_selected_index_changed
        # 
        # month_name
        # 
        self._month_name.Location = System.Drawing.Point(120, 137)
        self._month_name.Name = "month_name"
        self._month_name.Size = System.Drawing.Size(131, 20)
        self._month_name.TabIndex = 6
        self._month_name.Leave += self.month_name_leave
        # 
        # illegal_character_label1
        # 
        self._illegal_character_label1.AutoSize = True
        self._illegal_character_label1.Location = System.Drawing.Point(17, 171)
        self._illegal_character_label1.Name = "illegal_character_label1"
        self._illegal_character_label1.Size = System.Drawing.Size(124, 13)
        self._illegal_character_label1.TabIndex = 7
        self._illegal_character_label1.Text = "Replace illegal character"
        # 
        # remove_illegal_character
        # 
        self._remove_illegal_character.AutoSize = True
        self._remove_illegal_character.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._remove_illegal_character.Location = System.Drawing.Point(292, 166)
        self._remove_illegal_character.Name = "remove_illegal_character"
        self._remove_illegal_character.Size = System.Drawing.Size(20, 23)
        self._remove_illegal_character.TabIndex = 12
        self._remove_illegal_character.Text = "-"
        self._remove_illegal_character.UseVisualStyleBackColor = True
        self._remove_illegal_character.Click += self.remove_illegal_character
        # 
        # add_illegal_character
        # 
        self._add_illegal_character.AutoSize = True
        self._add_illegal_character.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._add_illegal_character.Location = System.Drawing.Point(263, 166)
        self._add_illegal_character.Name = "add_illegal_character"
        self._add_illegal_character.Size = System.Drawing.Size(23, 23)
        self._add_illegal_character.TabIndex = 11
        self._add_illegal_character.Text = "+"
        self._add_illegal_character.UseVisualStyleBackColor = True
        self._add_illegal_character.Click += self.add_illegal_character
        # 
        # month_label2
        # 
        self._month_label2.AutoSize = True
        self._month_label2.Location = System.Drawing.Point(100, 139)
        self._month_label2.Name = "month_label2"
        self._month_label2.Size = System.Drawing.Size(14, 13)
        self._month_label2.TabIndex = 5
        self._month_label2.Text = "is"
        # 
        # illegal_character_label2
        # 
        self._illegal_character_label2.AutoSize = True
        self._illegal_character_label2.Location = System.Drawing.Point(185, 171)
        self._illegal_character_label2.Name = "illegal_character_label2"
        self._illegal_character_label2.Size = System.Drawing.Size(26, 13)
        self._illegal_character_label2.TabIndex = 9
        self._illegal_character_label2.Text = "with"
        # 
        # illegal_character_replacement
        # 
        self._illegal_character_replacement.Location = System.Drawing.Point(215, 167)
        self._illegal_character_replacement.Name = "illegal_character_replacement"
        self._illegal_character_replacement.Size = System.Drawing.Size(37, 20)
        self._illegal_character_replacement.TabIndex = 10
        self._illegal_character_replacement.KeyPress += self.illegal_character_replacement_keypress
        self._illegal_character_replacement.Leave += self.illegal_character_replacement_leave
        # 
        # illegal_character_selector
        # 
        self._illegal_character_selector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._illegal_character_selector.FormattingEnabled = True
        self._illegal_character_selector.Location = System.Drawing.Point(149, 167)
        self._illegal_character_selector.Name = "illegal_character_selector"
        self._illegal_character_selector.Size = System.Drawing.Size(32, 21)
        self._illegal_character_selector.TabIndex = 8
        self._illegal_character_selector.SelectedIndexChanged += self.illegal_character_selector_selected_index_changed
        # 
        # empty_folder_exceptions_list
        # 
        self._empty_folder_exceptions_list.FormattingEnabled = True
        self._empty_folder_exceptions_list.Location = System.Drawing.Point(17, 248)
        self._empty_folder_exceptions_list.Name = "empty_folder_exceptions_list"
        self._empty_folder_exceptions_list.Size = System.Drawing.Size(411, 134)
        self._empty_folder_exceptions_list.TabIndex = 15
        # 
        # remove_empty_folders_label
        # 
        self._remove_empty_folders_label.AutoSize = True
        self._remove_empty_folders_label.Location = System.Drawing.Point(36, 226)
        self._remove_empty_folders_label.Name = "remove_empty_folders_label"
        self._remove_empty_folders_label.Size = System.Drawing.Size(193, 13)
        self._remove_empty_folders_label.TabIndex = 14
        self._remove_empty_folders_label.Text = "But do not remove the following folders:"
        # 
        # add_empty_folder_exception
        # 
        self._add_empty_folder_exception.Location = System.Drawing.Point(432, 277)
        self._add_empty_folder_exception.Name = "add_empty_folder_exception"
        self._add_empty_folder_exception.Size = System.Drawing.Size(57, 23)
        self._add_empty_folder_exception.TabIndex = 16
        self._add_empty_folder_exception.Tag = self._empty_folder_exceptions_list
        self._add_empty_folder_exception.Text = "Add"
        self._add_empty_folder_exception.UseVisualStyleBackColor = True
        self._add_empty_folder_exception.Click += self.add_folder_path_to_list
        # 
        # remove_empty_folder_exception
        # 
        self._remove_empty_folder_exception.Location = System.Drawing.Point(432, 328)
        self._remove_empty_folder_exception.Name = "remove_empty_folder_exception"
        self._remove_empty_folder_exception.Size = System.Drawing.Size(57, 23)
        self._remove_empty_folder_exception.TabIndex = 17
        self._remove_empty_folder_exception.Tag = self._empty_folder_exceptions_list
        self._remove_empty_folder_exception.Text = "Remove"
        self._remove_empty_folder_exception.UseVisualStyleBackColor = True
        self._remove_empty_folder_exception.Click += self.remove_folder_path_from_list
        # 
        # empty_folder_name_label
        # 
        self._empty_folder_name_label.Location = System.Drawing.Point(8, 16)
        self._empty_folder_name_label.Name = "empty_folder_name_label"
        self._empty_folder_name_label.Size = System.Drawing.Size(201, 29)
        self._empty_folder_name_label.TabIndex = 0
        self._empty_folder_name_label.Text = "Replace empty folder names with: (Leave empty to remove empty folders)"
        # 
        # empty_folder_name
        # 
        self._empty_folder_name.Location = System.Drawing.Point(208, 20)
        self._empty_folder_name.Name = "empty_folder_name"
        self._empty_folder_name.Size = System.Drawing.Size(274, 20)
        self._empty_folder_name.TabIndex = 1
        # 
        # failed_empty_selection
        # 
        self._failed_empty_selection.CheckOnClick = True
        self._failed_empty_selection.FormattingEnabled = True
        self._failed_empty_selection.HorizontalScrollbar = True
        self._failed_empty_selection.Location = System.Drawing.Point(26, 219)
        self._failed_empty_selection.Name = "failed_empty_selection"
        self._failed_empty_selection.Size = System.Drawing.Size(232, 109)
        self._failed_empty_selection.Sorted = True
        self._failed_empty_selection.TabIndex = 9
        self._failed_empty_selection.Items.AddRange(failed_items)
        # 
        # failed_empty_checkbox
        # 
        self._failed_empty_checkbox.Checked = True
        self._failed_empty_checkbox.CheckState = System.Windows.Forms.CheckState.Checked
        self._failed_empty_checkbox.Location = System.Drawing.Point(8, 182)
        self._failed_empty_checkbox.Name = "failed_empty_checkbox"
        self._failed_empty_checkbox.Size = System.Drawing.Size(474, 31)
        self._failed_empty_checkbox.TabIndex = 8
        self._failed_empty_checkbox.Text = "If any of the selected fields are empty then mark the operation as failed."
        self._failed_empty_checkbox.CheckedChanged += self.failed_empty_checkbox_checked_changed
        # 
        # failed_empty_folder
        # 
        self._failed_empty_folder.Location = System.Drawing.Point(26, 357)
        self._failed_empty_folder.Name = "failed_empty_folder"
        self._failed_empty_folder.Size = System.Drawing.Size(377, 20)
        self._failed_empty_folder.TabIndex = 11
        self._failed_empty_folder.Enabled = False
        # 
        # failed_empty_browse
        # 
        self._failed_empty_browse.Location = System.Drawing.Point(409, 355)
        self._failed_empty_browse.Name = "failed_empty_browse"
        self._failed_empty_browse.Size = System.Drawing.Size(75, 23)
        self._failed_empty_browse.TabIndex = 12
        self._failed_empty_browse.Text = "Browse"
        self._failed_empty_browse.UseVisualStyleBackColor = True
        self._failed_empty_browse.Tag = self._failed_empty_folder
        self._failed_empty_browse.Click += self.add_folder_path_to_text_box
        self._failed_empty_browse.Enabled = False
        # 
        # empty_substitution_label
        # 
        self._empty_substitution_label.AutoSize = True
        self._empty_substitution_label.Location = System.Drawing.Point(8, 97)
        self._empty_substitution_label.Name = "empty_substitution_label"
        self._empty_substitution_label.Size = System.Drawing.Size(250, 13)
        self._empty_substitution_label.TabIndex = 3
        self._empty_substitution_label.Text = "When a field is empty substitute the following value:"
        # 
        # empty_substitution_field
        # 
        self._empty_substitution_field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._empty_substitution_field.FormattingEnabled = True
        self._empty_substitution_field.IntegralHeight = False
        self._empty_substitution_field.Location = System.Drawing.Point(43, 115)
        self._empty_substitution_field.MaxDropDownItems = 15
        self._empty_substitution_field.Name = "empty_substitution_field"
        self._empty_substitution_field.Size = System.Drawing.Size(121, 21)
        self._empty_substitution_field.Sorted = True
        self._empty_substitution_field.TabIndex = 5
        self._empty_substitution_field.Items.AddRange(empty_substitution_items)
        self._empty_substitution_field.SelectedIndex = 0
        self._empty_substitution_field.SelectedIndexChanged += self.empty_substitution_field_selected_index_changed
        # 
        # empty_substitution_label1
        # 
        self._empty_substitution_label1.AutoSize = True
        self._empty_substitution_label1.Location = System.Drawing.Point(8, 118)
        self._empty_substitution_label1.Name = "empty_substitution_label1"
        self._empty_substitution_label1.Size = System.Drawing.Size(29, 13)
        self._empty_substitution_label1.TabIndex = 4
        self._empty_substitution_label1.Text = "Field"
        # 
        # empty_substitution_value
        # 
        self._empty_substitution_value.Location = System.Drawing.Point(237, 115)
        self._empty_substitution_value.Name = "empty_substitution_value"
        self._empty_substitution_value.Size = System.Drawing.Size(245, 20)
        self._empty_substitution_value.TabIndex = 7
        self._empty_substitution_value.Leave += self.empty_subsititution_value_leave
        # 
        # empty_substitution_label2
        # 
        self._empty_substitution_label2.AutoSize = True
        self._empty_substitution_label2.Location = System.Drawing.Point(171, 118)
        self._empty_substitution_label2.Name = "empty_substitution_label2"
        self._empty_substitution_label2.Size = System.Drawing.Size(60, 13)
        self._empty_substitution_label2.TabIndex = 6
        self._empty_substitution_label2.Text = "substitution"
        # 
        # moved_failed_empty
        # 
        self._move_failed_empty.AutoSize = True
        self._move_failed_empty.Location = System.Drawing.Point(26, 334)
        self._move_failed_empty.Name = "moved_failed_empty"
        self._move_failed_empty.Size = System.Drawing.Size(162, 17)
        self._move_failed_empty.TabIndex = 10
        self._move_failed_empty.Text = "and move/copy them to this folder:"
        self._move_failed_empty.UseVisualStyleBackColor = True
        self._move_failed_empty.CheckedChanged += self.move_failed_empty_check_changed

        self.load_options_page_settings()

        self._options_page_empty_values_tab.ResumeLayout()        
        self._options_page_options_tab.ResumeLayout()
        self._options_page.ResumeLayout()


    def create_insert_controls(self):
        """Creates the insert controls."""
        self._text_insert_controls = System.Windows.Forms.TabPage()
        self._yes_no_insert_controls = System.Windows.Forms.TabPage()
        self._multiple_value_insert_controls = System.Windows.Forms.TabPage()
        self._calculated_insert_controls = System.Windows.Forms.TabPage()
        self._search_insert_controls = System.Windows.Forms.TabPage()
        self._number_insert_controls = System.Windows.Forms.TabPage()

        self._insert_controls.SuspendLayout()
        self._yes_no_insert_controls.SuspendLayout()
        self._multiple_value_insert_controls.SuspendLayout()
        self._calculated_insert_controls.SuspendLayout()
        self._search_insert_controls.SuspendLayout()
        self._number_insert_controls.SuspendLayout()
      
        # 
        # insert_controls
        # 
        self._insert_controls.Controls.Add(self._text_insert_controls)
        self._insert_controls.Controls.Add(self._number_insert_controls)
        self._insert_controls.Controls.Add(self._yes_no_insert_controls)
        self._insert_controls.Controls.Add(self._multiple_value_insert_controls)
        self._insert_controls.Controls.Add(self._calculated_insert_controls)
        self._insert_controls.Controls.Add(self._search_insert_controls)
        self._insert_controls.Location = System.Drawing.Point(-1, 131)
        self._insert_controls.Name = "insert_controls"
        self._insert_controls.SelectedIndex = 0
        self._insert_controls.Size = System.Drawing.Size(503, 289)
        self._insert_controls.TabIndex = 6
        self._insert_controls.Selecting += self.insert_controls_selecting
        self._insert_controls.Selected += self.insert_controls_selected
        self._insert_controls.Deselected += self.insert_controls_deselected
        # 
        # text_insert_controls
        # 
        self._text_insert_controls.Location = System.Drawing.Point(4, 22)
        self._text_insert_controls.Name = "text_insert_controls"
        self._text_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._text_insert_controls.Size = System.Drawing.Size(495, 263)
        self._text_insert_controls.TabIndex = 0
        self._text_insert_controls.Text = "Text Fields"
        self._text_insert_controls.UseVisualStyleBackColor = True
        self._text_insert_controls.AutoScroll = True
        # 
        # yes_no_insert_controls
        # 
        self._yes_no_insert_controls.Location = System.Drawing.Point(4, 22)
        self._yes_no_insert_controls.Name = "yes_no_insert_controls"
        self._yes_no_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._yes_no_insert_controls.Size = System.Drawing.Size(495, 263)
        self._yes_no_insert_controls.TabIndex = 1
        self._yes_no_insert_controls.Text = "Yes/No Fields"
        self._yes_no_insert_controls.UseVisualStyleBackColor = True
        # 
        # multiple_value_insert_controls
        # 
        self._multiple_value_insert_controls.AutoScroll = True
        self._multiple_value_insert_controls.Location = System.Drawing.Point(4, 22)
        self._multiple_value_insert_controls.Name = "multiple_value_insert_controls"
        self._multiple_value_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._multiple_value_insert_controls.Size = System.Drawing.Size(495, 263)
        self._multiple_value_insert_controls.TabIndex = 2
        self._multiple_value_insert_controls.Text = "Multiple Value Fields"
        self._multiple_value_insert_controls.UseVisualStyleBackColor = True
        # 
        # calculated_insert_controls
        # 
        self._calculated_insert_controls.Location = System.Drawing.Point(4, 22)
        self._calculated_insert_controls.Name = "calculated_insert_controls"
        self._calculated_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._calculated_insert_controls.Size = System.Drawing.Size(495, 263)
        self._calculated_insert_controls.TabIndex = 3
        self._calculated_insert_controls.Text = "Calculated Values"
        self._calculated_insert_controls.UseVisualStyleBackColor = True
        # 
        # search_insert_controls
        # 
        self._search_insert_controls.Location = System.Drawing.Point(4, 22)
        self._search_insert_controls.Name = "search_insert_controls"
        self._search_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._search_insert_controls.Size = System.Drawing.Size(495, 263)
        self._search_insert_controls.TabIndex = 4
        self._search_insert_controls.Text = "Search"
        self._search_insert_controls.UseVisualStyleBackColor = True
        # 
        # number_insert_controls
        # 
        self._number_insert_controls.Location = System.Drawing.Point(4, 22)
        self._number_insert_controls.Name = "number_insert_controls"
        self._number_insert_controls.Padding = System.Windows.Forms.Padding(3)
        self._number_insert_controls.Size = System.Drawing.Size(495, 263)
        self._number_insert_controls.TabIndex = 5
        self._number_insert_controls.Text = "Number Fields"
        self._number_insert_controls.UseVisualStyleBackColor = True

        self.create_text_insert_controls()

        self._insert_controls_dict.update(self._text_insert_controls_list)

        self.load_insert_controls_settings()

        self._insert_controls.ResumeLayout()
        self._yes_no_insert_controls.ResumeLayout()
        self._multiple_value_insert_controls.ResumeLayout()
        self._calculated_insert_controls.ResumeLayout()
        self._search_insert_controls.ResumeLayout()
        self._number_insert_controls.ResumeLayout()


    #These five methods create the controls in the insert_control. 
    #Remember to add the location to the tag property of the control.
    def create_text_insert_controls(self):
        
        self.AgeRating = InsertControl()
        self.AgeRating.SetTemplate("ageRating", "Age Rate.")
        self.AgeRating.Name = "Age Rating"
        self.AgeRating.SetLabels("Prefix", "", "Suffix")
        self.AgeRating.Location = Point(4, 10)
        self.AgeRating.Tag = self.AgeRating.Location
        self.AgeRating.TabIndex = 0
        self._text_insert_controls_list["age rating"] = self.AgeRating
        
        self.AlternateSeries = InsertControl()
        self.AlternateSeries.SetTemplate("altSeries", "Alt. Series")
        self.AlternateSeries.Name = "Alternate Series"
        self.AlternateSeries.SetLabels("Prefix", "", "Suffix")
        self.AlternateSeries.Location = Point(248, 10)
        self.AlternateSeries.Tag = self.AlternateSeries.Location
        self.AlternateSeries.TabIndex = 1
        self._text_insert_controls_list["alternate series"] = self.AlternateSeries
        
        self.Format = InsertControl()
        self.Format.SetTemplate("format", "Format")
        self.Format.Name = "Format"
        self.Format.Location = Point(4, 55)
        self.Format.Tag = self.Format.Location
        self.Format.TabIndex = 2
        self._text_insert_controls_list["format"] = self.Format

        self.Imprint = InsertControl()
        self.Imprint.SetTemplate("imprint", "Imprint")
        self.Imprint.Name = "Imprint"
        self.Imprint.Location = Point(248, 55)
        self.Imprint.Tag = self.Imprint.Location
        self.Imprint.TabIndex = 3
        self._text_insert_controls_list["imprint"] = self.Imprint
        
        self.Language = InsertControl()
        self.Language.SetTemplate("language", "Language")
        self.Language.Name = "Language"
        self.Language.Location = Point(4, 100)
        self.Language.Tag = self.Language.Location
        self.Language.TabIndex = 4
        self._text_insert_controls_list["language"] = self.Language

        self.MainCharacterOrTeam = InsertControl()
        self.MainCharacterOrTeam.SetTemplate("maincharacter", "Main Character")
        self.MainCharacterOrTeam.Name = "Main Character Or Team"
        self.MainCharacterOrTeam.Location = Point(248, 100)
        self.MainCharacterOrTeam.Tag = self.MainCharacterOrTeam.Location
        self.MainCharacterOrTeam.TabIndex = 5
        self._text_insert_controls_list["main character or team"] = self.MainCharacterOrTeam

        self.Month = InsertControl()
        self.Month.SetTemplate("month", "Month")
        self.Month.Name = "Month"
        self.Month.Location = Point(4, 145)
        self.Month.Tag = self.Month.Location
        self.Month.TabIndex = 6
        self._text_insert_controls_list["month"] = self.Month

        self.Publisher = InsertControl()
        self.Publisher.SetTemplate("publisher", "Publisher")
        self.Publisher.Name = "Publisher"
        self.Publisher.Location = Point(248, 145)
        self.Publisher.Tag = self.Publisher.Location
        self.Publisher.TabIndex = 7
        self._text_insert_controls_list["publisher"] = self.Publisher
        
        self.Series = InsertControl()
        self.Series.SetTemplate("series", "Series")
        self.Series.Name = "Series"
        self.Series.Location = Point(4, 190)
        self.Series.Tag = self.Series.Location
        self.Series.TabIndex = 8
        self._text_insert_controls_list["series"] = self.Series

        self.SeriesGroup = InsertControl()
        self.SeriesGroup.SetTemplate("seriesgroup", "Series Group")
        self.SeriesGroup.Name = "Series Group"
        self.SeriesGroup.Location = Point(248, 190)
        self.SeriesGroup.Tag = self.SeriesGroup.Location
        self.SeriesGroup.TabIndex = 9
        self._text_insert_controls_list["series group"] = self.SeriesGroup

        self.StoryArc = InsertControl()
        self.StoryArc.SetTemplate("storyarc", "Story Arc")
        self.StoryArc.Name = "Story Arc"
        self.StoryArc.Location = Point(4, 235)
        self.StoryArc.Tag = self.StoryArc.Location
        self.StoryArc.TabIndex = 10
        self._text_insert_controls_list["story arc"] = self.StoryArc
        
        self.Title = InsertControl()
        self.Title.SetTemplate("title", "Title")
        self.Title.Name = "Title"
        self.Title.Location = Point(248, 235)
        self.Title.Tag = self.Title.Location
        self.Title.TabIndex = 11
        self._text_insert_controls_list["title"] = self.Title

        self._text_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._text_insert_controls_list.Values))
        
        
    def create_number_insert_controls(self):
        
        self._insert_controls.SuspendLayout()
        self._number_insert_controls.SuspendLayout()

        self.AlternateCount = InsertControlNumber()
        self.AlternateCount.SetTemplate("altCount", "Alt. Count")
        self.AlternateCount.Name = "Alternate Count"
        self.AlternateCount.SetLabels("Prefix", "", "Suffix", "Padding")
        self.AlternateCount.Location = Point(4, 10)
        self.AlternateCount.Tag = self.AlternateCount.Location
        self._number_insert_controls_list["alternate count"] = self.AlternateCount
        
        self.AlternateNumber = InsertControlNumber()
        self.AlternateNumber.SetTemplate("altNumber", "Alt. Num.")
        self.AlternateNumber.Name = "Alternate Number"
        self.AlternateNumber.SetLabels("Prefix", "", "Suffix", "Padding")
        self.AlternateNumber.Location = Point(248, 10)
        self.AlternateNumber.Tag = self.AlternateNumber.Location
        self._number_insert_controls_list["alternate number"] = self.AlternateNumber
        
        self.Count = InsertControlNumber()
        self.Count.SetTemplate("count", "Count")
        self.Count.Name = "Count"
        self.Count.Location = Point(4, 55)
        self.Count.Tag = self.Count.Location
        self._number_insert_controls_list["count"] = self.Count
        
        self.MonthNumber = InsertControlNumber()
        self.MonthNumber.SetTemplate("month#", "Month#")
        self.MonthNumber.Name = "Month Number"
        self.MonthNumber.Location = Point(248, 55)
        self.MonthNumber.Tag = self.MonthNumber.Location
        self._number_insert_controls_list["month number"] = self.MonthNumber
        
        self.Number = InsertControlNumber()
        self.Number.SetTemplate("number", "Number")
        self.Number.Name = "Number"
        self.Number.Location = Point(4, 100)
        self.Number.Tag = self.Number.Location
        self._number_insert_controls_list["number"] = self.Number
        
        self.Volume = InsertControlNumber()
        self.Volume.SetTemplate("volume", "Volume")
        self.Volume.Name = "Volume"
        self.Volume.Location = Point(248, 100)
        self.Volume.Tag = self.Volume.Location
        self._number_insert_controls_list["volume"] = self.Volume
        
        self.Year = InsertControlNumber()
        self.Year.SetTemplate("year", "Year")
        self.Year.Name = "Year"
        self.Year.Location = Point(4, 145)
        self.Year.Tag = self.Year.Location
        self._number_insert_controls_list["year"] = self.Year

        self._number_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._number_insert_controls_list.Values))

        self._insert_controls_dict.update(self._number_insert_controls_list)

        self.load_insert_controls_settings()

        self._number_insert_controls.ResumeLayout()
        self._insert_controls.ResumeLayout()
        
        
    def create_yes_no_insert_controls(self):

        self._insert_controls.SuspendLayout()
        self._yes_no_insert_controls.SuspendLayout()

        self._yes_no_insert_controls_instructions = System.Windows.Forms.Label()
        self.Manga = InsertControlYesNo()
        self.Manga.SetTemplate("manga", "Manga")
        self.Manga.Name = "Manga"
        self.Manga.SetLabels("Prefix", "", "Suffix", "Text")
        self.Manga.Location = Point(4, 10)
        self.Manga.Tag = self.Manga.Location
        self._yes_no_insert_controls_list["manga"] = self.Manga
        
        self.SeriesComplete = InsertControlYesNo()
        self.SeriesComplete.SetTemplate("seriesComplete", "Series Complete")
        self.SeriesComplete.InsertButton.Width += 30
        self.SeriesComplete.SetLabels("Prefix", "", "Suffix", "Text")
        self.SeriesComplete.Name = "Series Complete"
        self.SeriesComplete.Location = Point(4, 55)
        self.SeriesComplete.Tag = self.SeriesComplete.Location
        self._yes_no_insert_controls_list["series complete"] = self.SeriesComplete

        # 
        # yes_no_insert_controls_instructions
        # 
        self._yes_no_insert_controls_instructions.Location = System.Drawing.Point(3, 146)
        self._yes_no_insert_controls_instructions.Name = "yes_no_insert_controls_instructions"
        self._yes_no_insert_controls_instructions.Size = System.Drawing.Size(490, 113)
        self._yes_no_insert_controls_instructions.TabIndex = 0
        self._yes_no_insert_controls_instructions.Text = "The text in the \"Text\" textbox will be inserted when the field is Yes.\n\nPress the \"!\" button before inserting the field to make the \"Text\" be inserted when the field is No."

        self._yes_no_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._yes_no_insert_controls_list.Values))
        self._yes_no_insert_controls.Controls.Add(self._yes_no_insert_controls_instructions)

        self._insert_controls_dict.update(self._yes_no_insert_controls_list)

        self.load_insert_controls_settings()

        self._yes_no_insert_controls.ResumeLayout()
        self._insert_controls.ResumeLayout()
    
    
    def create_multiple_value_insert_controls(self):

        self._insert_controls.SuspendLayout()
        self._multiple_value_insert_controls.SuspendLayout()

        self._multiple_value_insert_controls_instructions = System.Windows.Forms.Label()

        self.AlternateSeriesMulti = InsertControlMultipleValue()
        self.AlternateSeriesMulti.Location = Point(4, 10)
        self.AlternateSeriesMulti.Name = "Alternate Series Multi"
        self.AlternateSeriesMulti.SetLabels("Prefix", "", "Suffix", "Seperator", "")
        self.AlternateSeriesMulti.SetTemplate("altSeries", "AltSeries")
        self.AlternateSeriesMulti.Tag   = self.AlternateSeriesMulti.Location
        self.AlternateSeriesMulti.TabIndex = 0
        self._multiple_value_insert_controls_list["alternate series multi"] = self.AlternateSeriesMulti

        self.Characters = InsertControlMultipleValue()
        self.Characters.Location = Point(244, 10)
        self.Characters.Name = "Characters"
        self.Characters.SetLabels("Prefix", "", "Suffix", "Seperator", "")
        self.Characters.SetTemplate("characters", "Character")
        self.Characters.Tag = self.Characters.Location
        self.Characters.TabIndex = 1
        self._multiple_value_insert_controls_list["characters"] = self.Characters

        self.Colorist = InsertControlMultipleValue()
        self.Colorist.SetTemplate("colorist", "Colorist")
        self.Colorist.Name = "Colorist"
        self.Colorist.Location = Point(4, 55)
        self.Colorist.Tag  = self.Colorist.Location
        self.Colorist.TabIndex = 2
        self._multiple_value_insert_controls_list["colorist"] = self.Colorist

        self.CoverArtist = InsertControlMultipleValue()
        self.CoverArtist.SetTemplate("coverartist", "Cover Artist")
        self.CoverArtist.Name = "Cover Artist"
        self.CoverArtist.Location = Point(244, 55)
        self.CoverArtist.Tag  = self.CoverArtist.Location
        self.CoverArtist.TabIndex = 3
        self._multiple_value_insert_controls_list["cover artist"] = self.CoverArtist

        self.Editor = InsertControlMultipleValue()
        self.Editor.SetTemplate("editor", "Editor")
        self.Editor.Name = "Editor"
        self.Editor.Location = Point(4, 100)
        self.Editor.Tag  = self.Editor.Location
        self.Editor.TabIndex = 4
        self._multiple_value_insert_controls_list["editor"] = self.Editor

        self.Genre = InsertControlMultipleValue()
        self.Genre.Location = Point(244, 100)
        self.Genre.Name = "Genre"
        self.Genre.SetTemplate("genre", "Genre")
        self.Genre.Tag  = self.Genre.Location
        self.Genre.TabIndex = 5
        self._multiple_value_insert_controls_list["genre"] = self.Genre

        self.Inker = InsertControlMultipleValue()
        self.Inker.SetTemplate("inker", "Inker")
        self.Inker.Name = "Inker"
        self.Inker.Location = Point(4, 145)
        self.Inker.Tag  = self.Inker.Location
        self.Inker.TabIndex = 6
        self._multiple_value_insert_controls_list["inker"] = self.Inker

        self.Letterer = InsertControlMultipleValue()
        self.Letterer.SetTemplate("letterer", "Letterer")
        self.Letterer.Name = "Letterer"
        self.Letterer.Location = Point(244, 145)
        self.Letterer.Tag  = self.Letterer.Location
        self.Letterer.TabIndex = 7
        self._multiple_value_insert_controls_list["letterer"] = self.Letterer

        self.Locations = InsertControlMultipleValue()
        self.Locations.Location = Point(4, 190)
        self.Locations.Name = "Locations"
        self.Locations.SetTemplate("locations", "Locations")
        self.Locations.Tag  = self.Locations.Location
        self.Locations.TabIndex = 8
        self._multiple_value_insert_controls_list["locations"] = self.Locations

        self.Penciller = InsertControlMultipleValue()
        self.Penciller.SetTemplate("penciller", "Penciller")
        self.Penciller.Name = "Penciller"
        self.Penciller.Location = Point(244, 190)
        self.Penciller.Tag  = self.Penciller.Location
        self.Penciller.TabIndex = 9
        self._multiple_value_insert_controls_list["penciller"] = self.Penciller

        self.ScanInformation = InsertControlMultipleValue()
        self.ScanInformation.SetTemplate("scaninfo", "Scan Info.")
        self.ScanInformation.Name = "Scan Information"
        self.ScanInformation.Location = Point(4, 235)
        self.ScanInformation.Tag    = self.ScanInformation.Location
        self.ScanInformation.TabIndex = 10
        self._multiple_value_insert_controls_list["scan information"] = self.ScanInformation

        self.Tags = InsertControlMultipleValue()
        self.Tags.Location = Point(244, 235)
        self.Tags.Name = "Tags"
        self.Tags.SetTemplate("tags", "Tags")
        self.Tags.Tag   = self.Tags.Location
        self.Tags.TabIndex = 11
        self._multiple_value_insert_controls_list["tags"] = self.Tags

        self.Teams = InsertControlMultipleValue()
        self.Teams.SetTemplate("teams", "Team")
        self.Teams.Name = "Teams"
        self.Teams.Location = Point(4, 280)
        self.Teams.Tag  = self.Teams.Location
        self.Teams.TabIndex = 12
        self._multiple_value_insert_controls_list["teams"] = self.Teams
        
        self.Writer = InsertControlMultipleValue()
        self.Writer.Location = Point(244, 280)
        self.Writer.Name = "Writer"
        self.Writer.SetTemplate("writer", "Writer")
        self.Writer.Tag = self.Writer.Location
        self.Writer.TabIndex = 13
        self._multiple_value_insert_controls_list["writer"] = self.Writer
        # 
        # multiple_value_insert_controls_instructions
        # 
        self._multiple_value_insert_controls_instructions.Location = System.Drawing.Point(0, 345)
        self._multiple_value_insert_controls_instructions.Name = "multiple_value_insert_controls_instructions"
        self._multiple_value_insert_controls_instructions.Size = System.Drawing.Size(470, 82)
        self._multiple_value_insert_controls_instructions.TabIndex = 1
        self._multiple_value_insert_controls_instructions.Text = "For these fields that can have multiple entries the script will ask which ones you would like to use.\nIf you want to select for all the issues in the series at once, check the checkbox beside the field before inserting.\nIn the selction box is an option to have each entry as its own folder"
        
        self._multiple_value_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._multiple_value_insert_controls_list.Values))
        self._multiple_value_insert_controls.Controls.Add(self._multiple_value_insert_controls_instructions)

        self._insert_controls_dict.update(self._multiple_value_insert_controls_list)

        self.load_insert_controls_settings()

        self._multiple_value_insert_controls.ResumeLayout()
        self._insert_controls.ResumeLayout()
    
    
    def create_calculated_insert_controls(self):
        
        self._insert_controls.SuspendLayout()
        self._calculated_insert_controls.SuspendLayout()

        self._calculated_insert_controls_start_year_information = System.Windows.Forms.Label()

        self.StartYear = InsertControl()
        self.StartYear.Location = Point(4, 10)
        self.StartYear.Tag = self.StartYear.Location
        self.StartYear.SetLabels("Prefix", "", "Suffix")
        self.StartYear.Name = "Start Year"
        self.StartYear.SetTemplate("startyear", "Start Year")
        self._calculated_insert_controls_list["start year"] = self.StartYear

        self.StartMonth = InsertControlStartMonth("Use month names")
        self.StartMonth.InsertButton.Width += 10
        self.StartMonth.Location = Point(4, 55)
        self.StartMonth.Tag = self.StartMonth.Location
        self.StartMonth.SetLabels("Prefix", "", "Suffix", "Padding")
        self.StartMonth.SetTemplate("startmonth", "Start Month")
        self.StartMonth.Name = "Start Month"
        self._calculated_insert_controls_list["start month"] = self.StartMonth

        self.FirstLetter = InsertControlFirstLetter()
        self.FirstLetter.SetTemplate("first", "FirstLetter")
        self.FirstLetter.Location = Point(4, 100)
        self.FirstLetter.Tag = self.FirstLetter.Location
        self.FirstLetter.SetLabels("Prefix", "", "Suffix", "Field")
        self.FirstLetter.Name = "First Letter"
        self.FirstLetter.SetComboBoxItems(["Series", "Publisher", "Imprint", "AlternateSeries"])
        self._calculated_insert_controls_list["first letter"] = self.FirstLetter

        self.Counter = InsertControlCounter()
        self.Counter.SetTemplate("counter", "Counter")
        self.Counter.SetLabels("Prefix", "", "Suffix", "Start", "Increment", "Pad")
        self.Counter.Location = Point(4, 145)
        self.Counter.Name = "Counter"
        self.Counter.Tag = self.Counter.Location
        self._calculated_insert_controls_list["counter"] = self.Counter

        self.Read = InsertControlReadPercentage()
        self.Read.SetTemplate("read", "Read %")
        self.Read.SetLabels("Prefix", "", "Suffix", "Text", "Operator", "Percent")
        self.Read.Location = Point(4, 190)
        self.Read.Tag = self.Read.Location
        self.Read.Name = "Read Percentage"
        self._calculated_insert_controls_list["read percentage"] = self.Read

        # 
        # calculated_insert_controls_start_year_information
        # 
        self._calculated_insert_controls_start_year_information.Location = System.Drawing.Point(225, 20)
        self._calculated_insert_controls_start_year_information.Name = "calculated_insert_controls_start_year_information"
        self._calculated_insert_controls_start_year_information.Size = System.Drawing.Size(253, 35)
        self._calculated_insert_controls_start_year_information.TabIndex = 1
        self._calculated_insert_controls_start_year_information.Text = "Start year and start month are calculated from the earliest issue of each series in your library."
        
        self._calculated_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._calculated_insert_controls_list.Values))
        self._calculated_insert_controls.Controls.Add(self._calculated_insert_controls_start_year_information)
        
        self._insert_controls_dict.update(self._calculated_insert_controls_list)

        self.load_insert_controls_settings()

        self._calculated_insert_controls.ResumeLayout()
        self._insert_controls.ResumeLayout()


    def create_search_insert_controls(self):
        self._insert_controls.SuspendLayout()
        self._search_insert_controls.SuspendLayout()
        self._search_insert_controls_name = System.Windows.Forms.TextBox()
        self._search_insert_controls_label = System.Windows.Forms.Label()
        self._search_insert_controls_layoutpanel = System.Windows.Forms.FlowLayoutPanel()

        self._search_insert_controls.Controls.Add(self._search_insert_controls_layoutpanel)
        self._search_insert_controls.Controls.Add(self._search_insert_controls_label)
        self._search_insert_controls.Controls.Add(self._search_insert_controls_name)

        # 
        # search_insert_controls_name
        # 
        self._search_insert_controls_name.Location = System.Drawing.Point(51, 6)
        self._search_insert_controls_name.Name = "search_insert_controls_name"
        self._search_insert_controls_name.Size = System.Drawing.Size(435, 20)
        self._search_insert_controls_name.TabIndex = 1
        self._search_insert_controls_name.TextChanged += self.search_insert_controls_name_text_changed
        # 
        # search_insert_controls_label
        # 
        self._search_insert_controls_label.AutoSize = True
        self._search_insert_controls_label.Location = System.Drawing.Point(7, 9)
        self._search_insert_controls_label.Name = "search_insert_controls_label"
        self._search_insert_controls_label.Size = System.Drawing.Size(38, 13)
        self._search_insert_controls_label.TabIndex = 3
        self._search_insert_controls_label.Text = "Name:"
        # 
        # search_insert_controls_layoutpanel
        # 
        self._search_insert_controls_layoutpanel.AutoScroll = True
        self._search_insert_controls_layoutpanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        self._search_insert_controls_layoutpanel.Location = System.Drawing.Point(0, 32)
        self._search_insert_controls_layoutpanel.Name = "search_insert_controls_layoutpanel"
        self._search_insert_controls_layoutpanel.Padding = System.Windows.Forms.Padding(30, 0, 0, 0)
        self._search_insert_controls_layoutpanel.Size = System.Drawing.Size(493, 229)
        self._search_insert_controls_layoutpanel.TabIndex = 4
        self._search_insert_controls_layoutpanel.WrapContents = False

        self._search_insert_controls.ResumeLayout()
        self._insert_controls.ResumeLayout()
    
          
    def change_page(self, sender, e):
        """
        When a toolbar button is clicked this function changes the visable page
        The toolbar button Tag property should contain the related page.
        The insert_controls are moved to the visible page in the case of the file or folders page
        """
        if sender.Checked:
            return
        
        self.SuspendLayout()

        #Save the metadata rules when switching away from the rules page. 
        #This caches it so it doesn't have to be rebuit everytime the preview text is updated
        if self._rules_button.Checked:
            self.save_rules_page_settings()

        elif self._options_button.Checked:
            self.save_options_page_settings()

        elif self._overview_button.Checked:
            self.save_overview_page_settings()

        elif self._files_button.Checked:
            self.save_files_page_settings()

        elif self._folders_button.Checked:
            self.save_folders_page_settings()
        
            
        if sender.Tag is self._folders_page:
            if self._folders_page.Controls.Count == 0:
                self.create_folders_page()
            if self._insert_controls.Controls.Count == 0:
                self.create_insert_controls()
            self._folders_page.Controls.Add(self._insert_controls)
            self._folders_page.Controls.Add(self._space_automatically)
            self._folders_page.Controls.Add(self._preview_book_selector)
        elif sender.Tag is self._files_page:
            if self._files_page.Controls.Count == 0:
                self.create_files_page()
            if self._insert_controls.Controls.Count == 0:
                self.create_insert_controls()
            self._files_page.Controls.Add(self._insert_controls)
            self._files_page.Controls.Add(self._space_automatically)
            self._files_page.Controls.Add(self._preview_book_selector)

        elif sender.Tag is self._rules_page:
            if self._rules_page.Controls.Count == 0:
                self.create_rules_page()

        elif sender.Tag is self._options_page:
            if self._options_page.Controls.Count == 0:
                self.create_options_page()

        for control in self._toolstrip.Items:
            if type(control) is System.Windows.Forms.ToolStripButton:
                control.Checked = False
                control.Tag.Visible = False
        
        sender.Checked = True
                
        sender.Tag.Visible = True
        
        self.update_template_text()

        self.ResumeLayout()
        

    #These five methods adjust which controls are visible when a checkbox changes

    def use_organization_check_changed(self, sender, e):
        """
        Enables and renables controls when the use organization controls are checked/unchecked
        The tag of each checkbox should contain the related toolbar button
        """
        sender.Tag.Enabled = sender.Checked
        
        if sender is self._use_folder_organization:
            self._base_folder_path.Enabled = sender.Checked
            self._browse_button.Enabled = sender.Checked
            self._label_base_folder.Enabled = sender.Checked
            self._fileless_image_format.Enabled = sender.Checked
            self._fileless_image_format_label.Enabled = sender.Checked
            self._copy_fileless_custom_thumbnails.Enabled = sender.Checked


    def failed_empty_checkbox_checked_changed(self, sender, e):
        self._move_failed_empty.Enabled = sender.Checked
        if self._move_failed_empty.Checked or not sender.Checked:
            self._failed_empty_browse.Enabled = sender.Checked
            self._failed_empty_folder.Enabled = sender.Checked
        self._failed_empty_selection.Enabled = sender.Checked
        

    def remove_empty_folders_checked_changed(self, sender, e):
        self._remove_empty_folder_exception.Enabled = sender.Checked
        self._remove_empty_folders_label.Enabled = sender.Checked
        self._add_empty_folder_exception.Enabled = sender.Checked
        self._empty_folder_exceptions_list.Enabled = sender.Checked


    def mode_copy_checked_changed(self, sender, e):
        self._copy_option.Enabled = sender.Checked


    def move_failed_empty_check_changed(self, sender, e):
        if sender.Enabled:
            self._failed_empty_browse.Enabled = sender.Checked
            self._failed_empty_folder.Enabled = sender.Checked
        else:
            self._failed_empty_browse.Enabled = False
            self._failed_empty_folder.Enabled = False


    #These two methods manage the month selectors
    def month_number_selected_index_changed(self, sender, e):
        """Loads the month from the profile using the selected number."""
        self._month_name.Text = self.profile.Months[int(self._month_number.SelectedItem)]


    def month_name_leave(self, sender, e):
        """Saves the month text to the profile."""
        self.profile.Months[int(self._month_number.SelectedItem)] = self._month_name.Text


    #These two methods manage the empty substitutions
    def empty_substitution_field_selected_index_changed(self, sender, e):
        """Loads the empty substitution from the profile."""
        if self._empty_substitution_field.SelectedItem in name_to_field and name_to_field[self._empty_substitution_field.SelectedItem] in self.profile.EmptyData:
            self._empty_substitution_value.Text = self.profile.EmptyData[name_to_field[self._empty_substitution_field.SelectedItem]]
        else:
            self._empty_substitution_value.Text = ""


    def empty_subsititution_value_leave(self, sender, e):
        """Save the empty substitution to the profile."""
        self.profile.EmptyData[name_to_field[self._empty_substitution_field.SelectedItem]] = self._empty_substitution_value.Text


    #These five methods manage the illegal characters
    def illegal_character_selector_selected_index_changed(self, sender, e):
        """Loads the entered replacement for the selected illegal character."""
        self._illegal_character_replacement.Text = self.profile.IllegalCharacters[self._illegal_character_selector.SelectedItem]


    def illegal_character_replacement_leave(self, sender, e):
        """Saves the illegal character replacement to the profile."""
        self.profile.IllegalCharacters[self._illegal_character_selector.SelectedItem] = self._illegal_character_replacement.Text


    def illegal_character_replacement_keypress(self, sender, e):
        """Stops any illegal character from being entered as an illegal character replacement."""
        if e.KeyChar in ("\\", "/", "|", "*", "<", ">", "?", '"', ":"):
            e.Handled = True


    def add_illegal_character(self, sender, e):
        """Adds a new character replacement into the form and profile"""
        dialog = NewIllegalCharacterDialog(self.profile.IllegalCharacters.keys())

        result = dialog.ShowDialog()

        if result == DialogResult.OK:
            character = dialog.GetCharacter()
            self.profile.IllegalCharacters[character] = ""
            self._illegal_character_selector.Items.Add(character)
            self._illegal_character_selector.SelectedItem = character

        dialog.Dispose()


    def remove_illegal_character(self, sender, e):
        """Removes an illegal character from the profile. The character is only removed as long as it is not one of the required ones."""
        character = self._illegal_character_selector.SelectedItem

        #Do not remove any of the essential chracters. Otherwise there is a large chance of errors
        if character not in ("?", "/", "\\", "*", ":", "<", ">", "|", "\""):
            index = self._illegal_character_selector.SelectedIndex
            self._illegal_character_selector.Items.Remove(character)
            del(self.profile.IllegalCharacters[character])
            l = len(self._illegal_character_selector.Items)
            if index < len(self._illegal_character_selector.Items) -1:
                self._illegal_character_selector.SelectedIndex = index
            else:
                self._illegal_character_selector.SelectedIndex = index - 1


    #These next four methods are for handling the insert controls with the search tab and creating controls when required.
    def search_insert_controls_name_text_changed(self, sender, e):
        """Selects the controls that match the text in the search text box"""
        text = sender.Text.lower()
        
        if not text:
            self._search_insert_controls_layoutpanel.SuspendLayout()
            self._search_insert_controls_layoutpanel.Controls.Clear()
            self._search_insert_controls_layoutpanel.ResumeLayout()
            return
        
        controls = [self._insert_controls_dict[controlname] for controlname in sorted(self._insert_controls_dict.keys()) if text in controlname]
        
        self._search_insert_controls_layoutpanel.SuspendLayout()
        self._search_insert_controls_layoutpanel.Controls.Clear()
        self._search_insert_controls_layoutpanel.Controls.AddRange(System.Array[System.Windows.Forms.Control](controls))
        self._search_insert_controls_layoutpanel.ResumeLayout()


    def insert_controls_deselected(self, sender, e):
        """
        Adds the controls back into the other tabs once the search tab is deselected
        
        Because the controls location's are now incorrect it is reset from the value in the control tag
        """
        if e.TabPage is self._search_insert_controls:
            self._search_insert_controls_layoutpanel.Controls.Clear()

            for control in self._insert_controls_dict.itervalues():
                control.Location = control.Tag

            self._text_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._text_insert_controls_list.values()))
            self._number_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._number_insert_controls_list.values()))
            self._yes_no_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._yes_no_insert_controls_list.values()))
            self._yes_no_insert_controls.Controls.Add(self._yes_no_insert_controls_instructions)
            self._multiple_value_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._multiple_value_insert_controls_list.values()))
            self._multiple_value_insert_controls.Controls.Add(self._multiple_value_insert_controls_instructions)
            self._calculated_insert_controls.Controls.AddRange(System.Array[System.Windows.Forms.Control](self._calculated_insert_controls_list.values()))
            self._calculated_insert_controls.Controls.Add(self._calculated_insert_controls_start_year_information)


    def insert_controls_selecting(self, sender, e):
        """When switching to the tabpage, clears the controls from all the other tabpages in the insert controls"""

        self._insert_controls.SuspendLayout()

        if e.TabPage is self._yes_no_insert_controls and self._yes_no_insert_controls.Controls.Count == 0 and self._search_insert_controls.Controls.Count == 0:
            self.create_yes_no_insert_controls()

        elif e.TabPage is self._number_insert_controls and self._number_insert_controls.Controls.Count == 0 and self._search_insert_controls.Controls.Count == 0:
            self.create_number_insert_controls()

        elif e.TabPage is self._multiple_value_insert_controls and self._multiple_value_insert_controls.Controls.Count == 0 and self._search_insert_controls.Controls.Count == 0:
            self.create_multiple_value_insert_controls()

        elif e.TabPage is self._calculated_insert_controls and self._calculated_insert_controls.Controls.Count == 0 and self._search_insert_controls.Controls.Count == 0:
            self.create_calculated_insert_controls()

        elif e.TabPage is self._search_insert_controls:
            if self._search_insert_controls.Controls.Count == 0:
                self.create_search_insert_controls()

            self._text_insert_controls.Controls.Clear()

            if self._number_insert_controls.Controls.Count == 0:
                self.create_number_insert_controls()
            self._number_insert_controls.Controls.Clear()

            if self._yes_no_insert_controls.Controls.Count == 0:
                self.create_yes_no_insert_controls()
            self._yes_no_insert_controls.Controls.Clear()

            if self._multiple_value_insert_controls.Controls.Count == 0:
                self.create_multiple_value_insert_controls()
            self._multiple_value_insert_controls.Controls.Clear()

            if self._calculated_insert_controls.Controls.Count== 0:
                self.create_calculated_insert_controls()
            self._calculated_insert_controls.Controls.Clear()

        self._insert_controls.ResumeLayout()


    def insert_controls_selected(self, sender, e):
        """Once the tabpage is selected, make sure that the text in the textbox is immediately searched"""
        if e.TabPage is self._search_insert_controls:
            self.search_insert_controls_name_text_changed(self._search_insert_controls_name, None)


    #These four methods are for inserting the template text into the correct textboxes and updating the preview when needed.

    def insert_control_clicked(self, sender, e):
        """Gets the template text from the clicked insert control then passes it to the function that adds it to the correct textbox"""
        template = sender.GetTemplateText(self._space_automatically.Checked)

        if self._files_page.Visible:
            self.insert_template_text(template, self._file_structure)

        else:
            self.insert_template_text(template, self._folder_structure)


    def insert_folder_seperator_clicked(self, sender, e):
        self.insert_template_text("\\", self._folder_structure)


    def insert_template_text(self, string, textbox):
        """Adds the string into the textbox, honoring where the insertion point is and what is selected"""
        start = textbox.SelectionStart
        length = textbox.SelectionLength
        s = textbox.Text
        if length > 0:
            s = s.Remove(start, length)
            length = len(string)
            textbox.Text = s.Insert(start, string)
            textbox.SelectionStart = start
            textbox.SelectionLength = length
        else:
            textbox.Text = s.Insert(start, string)
            textbox.SelectionStart = start + len(string)


    def update_template_text(self, sender=None, e=None):
        if self._files_page.Visible:
            if check_excluded_folders(self._preview_book.FilePath, self.profile) and check_metadata_rules(self._preview_book, self.profile):
                folder_path, file_name, failed = self.path_maker.make_path(self._preview_book, self.profile.FolderTemplate, self._file_structure.Text)
                self._file_preview.Text = file_name
            else:
                self._file_preview.Text = self._preview_book.FileNameWithExtension
        elif self._folders_page.Visible:
            if check_excluded_folders(self._preview_book.FilePath, self.profile) and check_metadata_rules(self._preview_book, self.profile):
                folder_path, file_name, failed = self.path_maker.make_path(self._preview_book, self._folder_structure.Text, self.profile.FileTemplate)
                self._folder_preview.Text = folder_path
            else:
                self._folder_preview.Text = self._preview_book.FileDirectory


    #These three methods manage adding and removing folder paths in the correct controls
    def add_folder_path_to_list(self, sender, e):
        """Show the folder browser dialog and then adds that folder to the listbox in the tag of the sender"""
        result = self._folder_browser_dialog.ShowDialog()
        if result is not DialogResult.Cancel:
            sender.Tag.Items.Add(self._folder_browser_dialog.SelectedPath)


    def remove_folder_path_from_list(self, sender, e):
        """Removed the selected item from the excluded folders list"""
        sender.Tag.Items.Remove(sender.Tag.SelectedItem)


    def add_folder_path_to_text_box(self, sender, e):
        """Show the folder browser dialog and then adds that folderpath into the textbox in the tag of the sender"""
        result = self._folder_browser_dialog.ShowDialog()
        if result is not DialogResult.Cancel:
            sender.Tag.Text = self._folder_browser_dialog.SelectedPath


    #These three methods manage the metadata rules.
    
    def add_metadata_rule(self, sender, e, exclude_rule=None):
        """Creates a new metadata rule and adds it into metadata rules container"""
        rule = MetadataExcludeRuleControl(self.remove_metadata_rule, exclude_rule)
        self._metadata_rules_container.Controls.Add(rule)
        self._metadata_rules_container.ScrollControlIntoView(rule)


    def add_metadata_rule_group(self, sender, e, exclude_rule_group=None):
        """Creates a new metadata group and adds it into the metadata rules container"""
        group = MetadataExcludeGroupControl(self.remove_metadata_rule, exclude_rule_group)
        self._metadata_rules_container.Controls.Add(group)
        self._metadata_rules_container.ScrollControlIntoView(group)


    def remove_metadata_rule(self, sender, e):
        """Removes a metadata rule or group from the metadata rules container"""
        rule = sender.Tag
        self._metadata_rules_container.Controls.Remove(rule)
        rule.Dispose()


    #These methods save and load the profiles
    def load_profile(self):
        """Loads the settings from the currently selected profile"""
        
        self.SuspendLayout()

        self.load_overview_page_settings()

        self._space_automatically.Checked = self.profile.AutoSpaceFields

        
        if self._rules_page.Controls.Count > 0:
            self.load_rules_page_settings()

        if self._options_page.Controls.Count > 0:
            self.load_options_page_settings()

        if self._insert_controls.Controls.Count > 0:
            self.load_insert_controls_settings()



        #Note the strucuctre text boxes are loaded last, otherwise the previews will be updated incorrectly.
        if self._files_page.Controls.Count > 0:        
            self.load_files_page_settings()

        if self._folders_page.Controls.Count > 0:
            self.load_folders_page_settings()
        

        #If one the file or folder template page and that function is disabled for the current profile, switch to the overview page.
        if not self.profile.UseFileName and self._files_page.Visible:
            self.change_page(self._overview_button, None)

        if not self.profile.UseFolder and self._folders_page.Visible:
            self.change_page(self._overview_button, None)

        self.ResumeLayout()


    def load_overview_page_settings(self):
        #Mode
        if self.profile.Mode == Mode.Move:
            self._mode_move.Checked = True

        elif self.profile.Mode == Mode.Copy:
            self._mode_copy.Checked = True

        else:
            self._mode_simulate.Checked = True

        self._copy_option.Checked = self.profile.CopyMode
        self._use_file_organization.Checked = self.profile.UseFileName
        self._use_folder_organization.Checked = self.profile.UseFolder
        self._copy_fileless_custom_thumbnails.Checked = self.profile.MoveFileless
        self._fileless_image_format.SelectedItem = self.profile.FilelessFormat
        self._base_folder_path.Text = self.profile.BaseFolder


    def load_files_page_settings(self):
        self._file_structure.Text = self.profile.FileTemplate
        self._file_structure.SelectionStart = len(self._file_structure.Text)


    def load_folders_page_settings(self):
        self._folder_structure.Text = self.profile.FolderTemplate
        self._folder_structure.SelectionStart = len(self._folder_structure.Text)


    def load_insert_controls_settings(self):

        for insertcontrol in self._insert_controls_dict.itervalues():
            name = insertcontrol.Name
            if name in self.profile.Prefix:
                insertcontrol.SetPrefixText(self.profile.Prefix[name])
            else:
                insertcontrol.SetPrefixText("")
                
            if name in self.profile.Postfix: 
                insertcontrol.SetPostfixText(self.profile.Postfix[name])
            else:
                insertcontrol.SetPostfixText("")

            if type(insertcontrol) is InsertControlMultipleValue:
                if name in self.profile.Seperator:
                    insertcontrol.SetSeperatorText(self.profile.Seperator[name])
                else:
                    insertcontrol.SetSeperatorText("")


            if type(insertcontrol) in (InsertControlYesNo, InsertControlReadPercentage):
                if name in self.profile.TextBox:
                    insertcontrol.SetTextBoxText(self.profile.TextBox[name])
                else:
                    insertcontrol.SetTextBoxText("")


    def load_rules_page_settings(self):
        self._metadata_rules_mode.SelectedItem = self.profile.ExcludeMode
        self._metadata_rules_operator.SelectedItem = self.profile.ExcludeOperator

        self._metadata_rules_container.Controls.Clear()

        for rule in self.profile.ExcludeRules:
            if type(rule) == ExcludeGroup:
                self.add_metadata_rule_group(None, None, rule)
            else:
                self.add_metadata_rule(None, None, rule)

        #Listboxes
        self._excluded_folders_list.Items.Clear()
        self._excluded_folders_list.Items.AddRange(System.Array[System.String](self.profile.ExcludeFolders))


    def load_options_page_settings(self):

        self._replace_multiple_spaces.Checked = self.profile.ReplaceMultipleSpaces
        self._insert_multiple_value_field_when_one.Checked = self.profile.DontAskWhenMultiOne
        self._remove_empty_folders.Checked = self.profile.RemoveEmptyFolder
        self._failed_empty_checkbox.Checked = self.profile.FailEmptyValues
        self._move_failed_empty.Checked = self.profile.MoveFailed
        self._copy_read_percentage.Checked = self.profile.CopyReadPercentage

        self._failed_empty_folder.Text = self.profile.FailedFolder
        self._empty_folder_name.Text = self.profile.EmptyFolder

        self._illegal_character_selector.Items.Clear()
        self._illegal_character_selector.Items.AddRange(System.Array[System.String](self.profile.IllegalCharacters.keys()))
        self._illegal_character_selector.SelectedIndex = 0

        self._month_number.Items.Clear()
        self._month_number.Items.AddRange(System.Array[System.Object](self.profile.Months.keys()))
        self._month_number.SelectedIndex = 0

        self._empty_folder_exceptions_list.Items.Clear()
        self._empty_folder_exceptions_list.Items.AddRange(System.Array[System.String](self.profile.ExcludedEmptyFolder))

        for i in self._failed_empty_selection.CheckedIndices:
            self._failed_empty_selection.SetItemChecked(i, False)

        for field in self.profile.FailedFields:
            self._failed_empty_selection.SetItemChecked(self._failed_empty_selection.Items.IndexOf(field_to_name[field]), True)


    def save_profile(self):
        """Saves all the current settings to the currently selected profile."""

        self.save_overview_page_settings()

        self.profile.AutoSpaceFields = self._space_automatically.Checked

        if self._options_page.Controls.Count > 0:
            self.save_options_page_settings()

        if self._rules_page.Controls.Count > 0:
            self.save_rules_page_settings()

        
        #Insert controls
        for insertcontrol in self._insert_controls_dict.itervalues():
            self.profile.Prefix[insertcontrol.Name] = insertcontrol.GetPrefixText()
            self.profile.Postfix[insertcontrol.Name] = insertcontrol.GetPostfixText()

            if type(insertcontrol) is InsertControlMultipleValue:
                self.profile.Seperator[insertcontrol.Name] = insertcontrol.GetSeperatorText()

            if type(insertcontrol) in (InsertControlYesNo, InsertControlReadPercentage):
                self.profile.TextBox[insertcontrol.Name] = insertcontrol.GetTextBoxText()

        if self._files_page.Controls.Count > 0:
            self.save_files_page_settings()

        if self._folders_page.Controls.Count > 0:
            self.save_folders_page_settings()


    def save_files_page_settings(self):
        self.profile.FileTemplate = self._file_structure.Text


    def save_folders_page_settings(self):
        self.profile.FolderTemplate = self._folder_structure.Text


    def save_rules_page_settings(self):
        """Saves the settings on the rules page to the profile."""
        self.profile.ExcludeRules = [rule.get_rule() for rule in self._metadata_rules_container.Controls]
        self.profile.ExcludeMode = self._metadata_rules_mode.SelectedItem
        self.profile.ExcludeOperator = self._metadata_rules_operator.SelectedItem
        self.profile.ExcludeFolders = list(self._excluded_folders_list.Items)


    def save_overview_page_settings(self):
        """Saves the settings on the rules page to the profile."""
        self.profile.BaseFolder = self._base_folder_path.Text

        if self._mode_move.Checked:
            self.profile.Mode = Mode.Move
        elif self._mode_copy.Checked:
            self.profile.Mode = Mode.Copy
            self.profile.CopyMode = self._copy_option.Checked
        else:
            self.profile.Mode = Mode.Simulate
        self.profile.UseFileName = self._use_file_organization.Checked
        self.profile.UseFolder = self._use_folder_organization.Checked
        self.profile.MoveFileless = self._copy_fileless_custom_thumbnails.Checked
        self.profile.FilelessFormat = self._fileless_image_format.SelectedItem


    def save_options_page_settings(self):
        """Saves the settings on the options page."""
        self.profile.EmptyFolder = self._empty_folder_name.Text
        self.profile.FailedFolder = self._failed_empty_folder.Text
        self.profile.ExcludedEmptyFolder = list(self._empty_folder_exceptions_list.Items)
        self.profile.FailedFields = [name_to_field[item] for item in self._failed_empty_selection.CheckedItems]
        self.profile.ReplaceMultipleSpaces = self._replace_multiple_spaces.Checked
        self.profile.DontAskWhenMultiOne = self._insert_multiple_value_field_when_one.Checked
        self.profile.RemoveEmptyFolder = self._remove_empty_folders.Checked
        self.profile.FailEmptyValues = self._failed_empty_checkbox.Checked
        self.profile.MoveFailed = self._move_failed_empty.Checked
        self.profile.CopyReadPercentage = self._copy_read_percentage.Checked


    #These 11 methods manage the profile actions
    def add_new_profile(self, sender, e):
        """Creates a new profile. Checks that all current settings are valid before starting to create a new profile.

        Asks the user for a non-duplicate profile name and then creates that profile and switches to it.
        """

        if not self.check_profile_settings():
            return

        self.save_profile()

        name = self.get_profile_name("Enter the name for the new profile.")

        if name is not None:
            self.profile = Profile()
            self.profile.Name = name
            self.profiles[name] = self.profile
            self._profile_selector.Items.Add(name)
            self._profile_selector.SelectedItem = name
            self.path_maker.profile = self.profile
            self.load_profile()
            self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def delete_profile(self, sender, e):
        """Deletes the selected profile as long as there is more than one profile."""
        if len(self.profiles) == 1:
            return

        #Since the combobox will revert to an index of -1 once an item is deleted a new index has to be found
        selected_index = self._profile_selector.SelectedIndex

        #The new index will be the previous item
        if selected_index == len(self.profiles) - 1:
            new_index = selected_index -1

        else:
            new_index = selected_index


        del(self.profiles[self._profile_selector.SelectedItem])
        self._profile_selector.Items.Remove(self._profile_selector.SelectedItem)
        self._profile_selector.SelectedIndex = new_index

        self.profile = self.profiles[self._profile_selector.SelectedItem]
        self.path_maker.profile = self.profile
        self.load_profile()
        self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def duplicate_profile(self, sender, e):
        """Duplicates the current profile """
        if not self.check_profile_settings():
            return

        self.save_profile()

        name = self.get_profile_name("Enter the name for the duplicate of the profile.")

        if name is not None:
            newprofile = self.profile.duplicate()
            newprofile.Name = name
            self.profiles[name] = newprofile
            self.profile = newprofile
            self._profile_selector.Items.Add(name)
            self._profile_selector.SelectedItem = name
            self.path_maker.profile = self.profile
            self.load_profile()
            self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def rename_profile(self, sender, e):
        """Shows a dialog to the user to enter a new name and then renames the currently selected profile with the submitted name."""

        name = self.get_profile_name("Enter the new name of the profile.", self._profile_selector.SelectedItem)

        if name is not None:
            oldname = self._profile_selector.SelectedItem
            self._profile_selector.Items.Remove(oldname)
            self.profile.Name = name
            self.profiles[name] = self.profile
            del(self.profiles[oldname])
            self._profile_selector.Items.Add(name)
            self._profile_selector.SelectedItem = name
            self.path_maker.profile = self.profile

            self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def export_profile(self, sender, e):
        """Exports the currently selected profile to a file."""
        if not self.check_profile_settings():
            return

        self.save_profile()

        file_path = self.get_export_profile_file_path()

        losettings.save_profile(file_path, self.profile)


    def export_all_profiles(self, sender, e):
        """Exports all the profiles to a file"""
        if not self.check_profile_settings():
            return

        self.save_profile()

        file_path = self.get_export_profile_file_path()

        losettings.save_profiles(file_path, self.profiles)


    def get_export_profile_file_path(self):
        """Shows a file dialog and returns the file path or None if the user canceled."""
        file_dialog = System.Windows.Forms.SaveFileDialog()
        file_dialog.Filter = "Library Organizer Profile (*.lop)|*.lop"
        dialog_result = file_dialog.ShowDialog()

        if dialog_result == DialogResult.Cancel:
            return None

        return file_dialog.FileName


    def get_profile_name(self, label_text, existing_name=""):
        """Gets a profile name from the user. 

        label_text->The label text in the GetProfileNameDialog.
        existing_name->Existing profile name to put in the textbox.

        Returns the new name or None if the user canceled the dialog"""
        profile_name_dialog = GetProfileNameDialog(self.profiles.keys(), label_text, existing_name)

        dialog_result = profile_name_dialog.ShowDialog()

        if dialog_result == DialogResult.Cancel:
            return None

        return profile_name_dialog.get_name()


    def import_profile(self, sender, e):
        """Asks the user for a profile file and imports the contents into the profiles. Asks to rename duplicates."""
        file_dialog = System.Windows.Forms.OpenFileDialog()
        file_dialog.Filter = "Library Organizer Profile (*.lop)|*.lop"
        dialog_result = file_dialog.ShowDialog()

        if dialog_result == DialogResult.Cancel:
            return

        imported_profiles = losettings.import_profiles(file_dialog.FileName)

        if not imported_profiles:
            return

        for profile_name in imported_profiles:
            if profile_name in self.profiles:
                newname = self.get_profile_name("The name of the imported profile already exists. Enter a different name.", profile_name)
                if newname is not None:
                    self.profiles[newname] = imported_profiles[profile_name]
                    self._profile_selector.Items.Add(newname)
            else:
                self.profiles[profile_name] = imported_profiles[profile_name]
                self._profile_selector.Items.Add(profile_name)
        
        self.adjust_combo_box_drop_down_width(self._profile_selector.ComboBox)


    def check_profile_settings(self):
        """
        Checks the currently selected settings to make sure there are no errors. If there are errors a messagebox will be shown to the user.

        Returns False if errors.
        Returns True if no errors.
        """
        errors = []
        if not self._base_folder_path.Text and self._use_folder_organization.Checked:
            errors.append("The base folder cannot be empty when using folder organization.")
        
        if self._files_page.Visible:
            if self._use_file_organization.Checked and not self._file_structure.Text.strip():
                errors.append("The File Structure cannot be empty when using file organization")
        else:
            if self._use_file_organization.Checked and not self.profile.FileTemplate:
                errors.append("The File Structure cannot be empty when using file organization")
        
        if not self._use_file_organization.Checked and not self._use_folder_organization.Checked:
            errors.append("You must enable the Folder or File naming functions for the script to actually do anything.")

        if self._options_page.Visible:
            if self._move_failed_empty.Checked and not self._failed_empty_folder.Text.strip():
                errors.append("The folder for failed files cannot be empty when moved failed files is checked.")
        else:
            if self.profile.MoveFailed and not self.profile.FailedFolder:
                errors.append("The folder for failed files cannot be empty when moved failed files is checked.")

        if errors:
            errors.insert(0, "Please correct the following errors:")
            MessageBox.Show("\n\n".join(errors), "Incompatiable settings", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        else:
            return True


    def profile_selector_selection_change_commited(self, sender, e):
        """
        Occurs when the user changes the profile. The function checks and makes sure the current profile has no errors and then
        save the current profile and loads the selected profile.
        """
        if self.check_profile_settings():
            self.save_profile()
            self.profile = self.profiles[sender.SelectedItem]
            self.path_maker.profile = self.profile
            self.load_profile()            

        else:
            sender.SelectedItem = self.profile.Name


    def change_preview_book(self, sender, e):
        self._preview_book = self._preview_books[sender.Value]
        self.update_template_text()


    def okay_clicked(self, sender, e):
        if not self.check_profile_settings():
            self.DialogResult = DialogResult.None


    def adjust_combo_box_drop_down_width(self, combobox):
        """Adjusts a combobox drop down width to fit all items along with a scroll bar if required.
        combobox->The ComboBox to adjust.
        """
        width = combobox.DropDownWidth
        g = combobox.CreateGraphics()
        font = combobox.Font
        scrollbarwidth = 0
        if combobox.Items.Count > combobox.MaxDropDownItems:
            scrollbarwidth = SystemInformation.VerticalScrollBarWidth

        newwidth = 0
        for string in combobox.Items:
            newwidth = g.MeasureString(string, font).Width + scrollbarwidth
            if newwidth > width:
                width = newwidth

        combobox.DropDownWidth = width
        g.Dispose()