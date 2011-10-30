"""loforms.py
This file contains various dialogs and forms required by the Library Orgainizer


Version 1.7.5

		Added Custom Chracter form

Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.
"""

import clr

import re

import System

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import ScrollBars, ListBox, Button, CheckBox, Form, FormBorderStyle, FormStartPosition, DialogResult, Label, TextBox, ErrorProvider, MessageBox

clr.AddReference("System.Drawing")

from System.Drawing import Point, Size, ContentAlignment

from System.IO import FileInfo

import locommon
from locommon import ICON

class SelectionForm(Form):
	def __init__(self, args):

		self.InitializeComponent()
		self.Icon = System.Drawing.Icon(ICON)

		if args.FieldText.endswith("s"):
			field = args.FieldText[:-1]
			fieldplural = args.FieldText
		else:
			field = args.FieldText
			fieldplural = args.FieldText + "s"

		self.Text = "Choose which " + fieldplural + " you would like to use"




		self.UseRegardless.Visible = args.Series

		self.UseRegardless.Text = "Use every " + field + " in the selection list for this series, even if the issue does not have that " + field

		if args.Series:
			self.Label.Text = "Select which " + fieldplural + " you would like to use for each issue in the series " + args.BookText
			self.AlwaysUse.Visible = False
			self.AlwaysUseDontAsk.Visible = False
		else:
			self.Label.Text = "Select which " + fieldplural + " you would like to use in the issue " + args.BookText

		self.SLabel.Text = "Selected " + fieldplural
		self.ILabel.Text = fieldplural
		self.UseAsFolder.Text = "Use each " + field + " as a folder"
		self.AlwaysUse.Text = "Use the " + field + "(s) in the selection list for every issue in this operation that has that " + field + "."
		self.AlwaysUseDontAsk.Text = "Do not ask if there are additional " + fieldplural + "."
		
		for i in args.SelectedItems:
			args.Items.Remove(i)
		self.Items.Items.AddRange(System.Array[System.String](args.Items))
		self.Selection.Items.AddRange(System.Array[System.String](args.SelectedItems))

	def InitializeComponent(self):
		self.Label = Label()
		self.Items = ListBox()
		self.Selection = ListBox()
		self.Okay = Button()
		self.Up = Button()
		self.Down = Button()
		self.UseRegardless = CheckBox()
		self.UseAsFolder = CheckBox()
		self.AlwaysUseDontAsk = CheckBox()
		self.Add = Button()
		self.Remove = Button()
		self.SLabel = Label()
		self.ILabel = Label()
		#
		# Label
		#
		self.Label.Location = Point(12, 8)
		self.Label.Size = Size(430, 30)
		#
		# Items Label
		#
		self.ILabel.Location = Point(12, 38)
		self.ILabel.TextAlign = ContentAlignment.MiddleCenter
		self.ILabel.Size = Size(159, 14)
		# 
		# Items
		# 
		self.Items.Location = Point(12, 55)
		self.Items.Name = "Items"
		self.Items.Size = Size(159, 134)
		self.Items.TabIndex = 0
		self.Items.Sorted = True
		self.Items.DoubleClick += self.AddItem
		#
		# Selection Label
		#
		self.SLabel.Location = Point(211, 38)
		self.SLabel.TextAlign = ContentAlignment.MiddleCenter
		self.SLabel.Size = Size(159, 14)
		# 
		# Selection
		# 
		self.Selection.Location = Point(211, 55)
		self.Selection.Name = "Selection"
		self.Selection.Size = Size(159, 134)
		self.Selection.TabIndex = 3
		self.Selection.DoubleClick += self.RemoveItem
		# 
		# Up
		# 
		self.Up.Location = Point(378, 70)
		self.Up.Name = "Up"
		self.Up.Size = Size(47, 23)
		self.Up.TabIndex = 4
		self.Up.Text = "Up"
		self.Up.UseVisualStyleBackColor = True
		self.Up.Click += self.MoveUp
		# 
		# Down
		# 
		self.Down.Location = Point(378, 130)
		self.Down.Name = "Down"
		self.Down.Size = Size(47, 23)
		self.Down.TabIndex = 5
		self.Down.Text = "Down"
		self.Down.UseVisualStyleBackColor = True
		self.Down.Click += self.MoveDown
		# 
		# UseRegardless
		# 
		self.UseRegardless.Location = Point(12, 190)
		self.UseRegardless.Name = "UseRegardless"
		self.UseRegardless.Size = Size(421, 34)
		self.UseRegardless.TabIndex = 7
		self.UseRegardless.UseVisualStyleBackColor = True
		self.UseRegardless.Enabled = True
		#
		# Always use
		#
		self.AlwaysUse = CheckBox()
		self.AlwaysUse.Location = Point(12, 190)
		self.AlwaysUse.Size = Size(413, 35)
		self.AlwaysUse.CheckedChanged += self.AlwaysUseCheckChanged
		#
		# Always use don't ask
		#
		self.AlwaysUseDontAsk.Location = Point(30, 230)
		self.AlwaysUseDontAsk.AutoSize = True
		self.AlwaysUseDontAsk.Checked = True
		self.AlwaysUseDontAsk.Enabled = False

		# 
		# UseAsFolder
		# 
		self.UseAsFolder.Location = Point(12, 260)
		self.UseAsFolder.Name = "UseAsFolder"
		self.UseAsFolder.Size = Size(252, 24)
		self.UseAsFolder.TabIndex = 8
		self.UseAsFolder.Text = "Seperate each item with a folder"
		self.UseAsFolder.UseVisualStyleBackColor = True
		# 
		# Okay
		# 
		self.Okay.Location = Point(267, 260)
		self.Okay.Name = "Okay"
		self.Okay.Size = Size(75, 23)
		self.Okay.TabIndex = 9
		self.Okay.Text = "Okay"
		self.Okay.UseVisualStyleBackColor = True
		self.Okay.DialogResult = DialogResult.OK
		#
		# Cancel
		#
		self.Cancel = Button()
		self.Cancel.Location = Point(350, 260)
		self.Cancel.Size = Size(75, 23)
		self.Cancel.Text = "Cancel"
		self.Cancel.DialogResult = DialogResult.Cancel
		# 
		# Add
		# 
		self.Add.Location = Point(179, 70)
		self.Add.Name = "Add"
		self.Add.Size = Size(24, 23)
		self.Add.TabIndex = 1
		self.Add.Text = "->"
		self.Add.UseVisualStyleBackColor = True
		self.Add.Click += self.AddItem
		# 
		# Remove
		# 
		self.Remove.Location = Point(179, 130)
		self.Remove.Name = "Remove"
		self.Remove.Size = Size(24, 23)
		self.Remove.TabIndex = 2
		self.Remove.Text = "<-"
		self.Remove.UseVisualStyleBackColor = True
		self.Remove.Click += self.RemoveItem
		# 
		# SelectionForm
		# 
		self.ClientSize = Size(437, 290)
		self.Controls.Add(self.Label)
		self.Controls.Add(self.UseAsFolder)
		self.Controls.Add(self.UseRegardless)
		self.Controls.Add(self.Remove)
		self.Controls.Add(self.Add)
		self.Controls.Add(self.Down)
		self.Controls.Add(self.Up)
		self.Controls.Add(self.Okay)
		self.Controls.Add(self.Selection)
		self.Controls.Add(self.AlwaysUse)
		self.Controls.Add(self.AlwaysUseDontAsk)
		self.Controls.Add(self.Cancel)
		self.Controls.Add(self.Items)
		self.Controls.Add(self.SLabel)
		self.Controls.Add(self.ILabel)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.ShowIcon = True
		self.AcceptButton = self.Okay
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Icon = System.Drawing.Icon(ICON)
		self.StartPosition = FormStartPosition.CenterParent


	def AlwaysUseCheckChanged(self, sender, e):
		self.AlwaysUseDontAsk.Enabled = self.AlwaysUse.Checked


	def AddItem(self, sender, e):
		if not self.Items.SelectedIndex == -1:
			index = self.Selection.SelectedIndex
			iindex = self.Items.SelectedIndex
			if iindex > self.Items.Items.Count -2: #Minus two 1 to get a 0 based index and 1 for the item to be removed
				iindex -= 1
			item = self.Items.SelectedItem
			self.Items.Items.Remove(item)
			self.Selection.Items.Insert(index + 1, item)
			self.Selection.SelectedIndex = index + 1
			self.Items.SelectedIndex = iindex

	def RemoveItem(self, sender, e):
		index = self.Selection.SelectedIndex
		if index != -1:
			if index > self.Selection.Items.Count - 2:
				index -= 1
			item = self.Selection.SelectedItem
			self.Selection.Items.Remove(item)
			self.Items.Items.Add(item)
			self.Selection.SelectedIndex = index

	def MoveUp(self, sender, e):
		if self.Selection.SelectedIndex != -1:
			if self.Selection.SelectedIndex > 0:
				item = self.Selection.SelectedItem
				i = self.Selection.SelectedIndex
				self.Selection.Items.RemoveAt(i)
				self.Selection.Items.Insert(i - 1, item)
				self.Selection.SelectedIndex = i - 1
                
	def MoveDown(self, sender, e):
		if self.Selection.SelectedIndex != -1:
			if self.Selection.SelectedIndex < self.Selection.Items.Count - 1:
				item = self.Selection.SelectedItem
				i = self.Selection.SelectedIndex
				self.Selection.Items.RemoveAt(i)
				self.Selection.Items.Insert(i + 1, item)
				self.Selection.SelectedIndex = i + 1

	def CheckChanged(self, sender, e):
		self.UseRegardless.Enabled = self.UseEveryIssue.Checked

	def GetResults(self):
		if self.DialogResult == DialogResult.OK:
			return SelectionFormResult(list(self.Selection.Items), self.UseRegardless.Checked, self.UseAsFolder.Checked, self.AlwaysUse.Checked, self.GetDontAsk())
		else:
			return SelectionFormResult([], self.UseRegardless.Checked, self.UseAsFolder.Checked, self.AlwaysUse.Checked, self.GetDontAsk())

	def GetDontAsk(self):
		if self.AlwaysUseDontAsk.Enabled:
			return self.AlwaysUseDontAsk.Checked
		else:
			return False
class SelectionFormResult(object):
	
	def __init__(self, selection, every = False, folder = False, alwaysuse = False, dontask = False):
		self.Selection = selection
		self.EveryIssue = every
		self.Folder = folder
		self.AlwaysUse = alwaysuse
		self.AlwaysUseDontAsk = dontask

class SelectionFormArgs(object):

	def __init__(self, items, selecteditems, fieldtext, booktext, series):
		self.Items = items
		self.FieldText = fieldtext
		self.BookText = booktext
		self.Series = series
		self.SelectedItems = selecteditems

class InputBox(Form):
	def __init__(self):
		self.TextBox = TextBox()
		self.TextBox.Size = Size(250, 20)
		self.TextBox.Location = Point(15, 12)
		self.TextBox.TabIndex = 1
		
		self.OK = Button()
		self.OK.Text = "OK"
		self.OK.Size = Size(75, 23)
		self.OK.Location = Point(109, 38)
		self.OK.DialogResult = DialogResult.OK
		self.OK.Click += self.CheckTextBox
		
		self.Cancel = Button()
		self.Cancel.Size = Size(75, 23)
		self.Cancel.Text = "Cancel"
		self.Cancel.Location = Point(190, 38)
		self.Cancel.DialogResult = DialogResult.Cancel
		
		self.Size = Size(300, 100)
		self.Text = "Please enter the profile name"
		self.Controls.Add(self.OK)
		self.Controls.Add(self.Cancel)
		self.Controls.Add(self.TextBox)
		self.AcceptButton = self.OK
		self.CancelButton = self.Cancel
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.StartPosition = FormStartPosition.CenterParent
		self.Icon = System.Drawing.Icon(ICON)
		self.ActiveControl = self.TextBox
		
	def FindName(self):
		if self.DialogResult == DialogResult.OK:
			return self.TextBox.Text.strip()
		else:
			return None
		
	def CheckTextBox(self, sender, e):
		if not self.TextBox.Text.strip():
			MessageBox.Show("Please enter a name into the textbox")
			self.DialogResult = DialogResult.None
		
		if self.TextBox.Text.strip() in self.Owner.allsettings:
			MessageBox.Show("The entered name is already in use. Please enter another")
			self.DialogResult = DialogResult.None

class IllegalCharacterInputBox(Form):
	def __init__(self, chracters):
		self.existingchracters = chracters
		self.TextBox = TextBox()
		self.TextBox.Size = Size(250, 20)
		self.TextBox.Location = Point(15, 12)
		self.TextBox.TabIndex = 1
		
		self.OK = Button()
		self.OK.Text = "OK"
		self.OK.Size = Size(75, 23)
		self.OK.Location = Point(109, 38)
		self.OK.DialogResult = DialogResult.OK
		self.OK.Click += self.CheckTextBox
		
		self.Cancel = Button()
		self.Cancel.Size = Size(75, 23)
		self.Cancel.Text = "Cancel"
		self.Cancel.Location = Point(190, 38)
		self.Cancel.DialogResult = DialogResult.Cancel
		
		self.Size = Size(300, 100)
		self.Text = "Please enter the character(s)"
		self.Controls.Add(self.OK)
		self.Controls.Add(self.Cancel)
		self.Controls.Add(self.TextBox)
		self.AcceptButton = self.OK
		self.CancelButton = self.Cancel
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.StartPosition = FormStartPosition.CenterParent
		self.Icon = System.Drawing.Icon(ICON)
		self.ActiveControl = self.TextBox

	def CheckTextBox(self, sender, e):
		if len(self.TextBox.Text) == 0:
			MessageBox.Show("Please enter a character into the textbox")
			self.DialogResult = DialogResult.None
		
		if self.TextBox.Text in self.existingchracters:
			MessageBox.Show("The entered character is already in use. Please enter another")
			self.DialogResult = DialogResult.None

	def GetCharacter(self):
		return self.TextBox.Text


class PathTooLongForm(Form):
	def __init__(self, path):
		self.InitializeComponent()
		self._Path.Text = path
		self.CheckPathLength(None, None)

	def InitializeComponent(self):
		self._components = System.ComponentModel.Container()
		self._label = System.Windows.Forms.Label()
		self._Okay = System.Windows.Forms.Button()
		self._Cancel = System.Windows.Forms.Button()
		self._Path = System.Windows.Forms.TextBox()
		self._errorProvider = System.Windows.Forms.ErrorProvider(self._components)
		self._errorProvider.BeginInit()
		self.SuspendLayout()
		# 
		# label
		# 
		self._label.Location = System.Drawing.Point(12, 9)
		self._label.Name = "label"
		self._label.Size = System.Drawing.Size(254, 23)
		self._label.Text = "The path created was too long. Please shorten it."
		# 
		# Okay
		# 
		self._Okay.DialogResult = System.Windows.Forms.DialogResult.OK
		self._Okay.Location = System.Drawing.Point(382, 120)
		self._Okay.Name = "Okay"
		self._Okay.Size = System.Drawing.Size(75, 23)
		self._Okay.TabIndex = 1
		self._Okay.Text = "Okay"
		self._Okay.UseVisualStyleBackColor = True
		# 
		# Cancel
		# 
		self._Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		self._Cancel.Location = System.Drawing.Point(463, 120)
		self._Cancel.Name = "Cancel"
		self._Cancel.Size = System.Drawing.Size(75, 23)
		self._Cancel.TabIndex = 2
		self._Cancel.Text = "Cancel"
		self._Cancel.UseVisualStyleBackColor = True
		# 
		# Path
		# 
		self._Path.Location = System.Drawing.Point(12, 35)
		self._Path.Name = "Path"
		self._Path.Size = System.Drawing.Size(500, 75)
		self._Path.TabIndex = 0
		self._Path.Multiline = True
		self._Path.TextChanged += self.CheckPathLength
		# 
		# errorProvider
		# 
		self._errorProvider.ContainerControl = self
		# 
		# Form1
		# 
		self.AcceptButton = self._Okay
		self.CancelButton = self._Cancel
		self.ClientSize = System.Drawing.Size(550, 150)
		self.Controls.Add(self._Path)
		self.Controls.Add(self._Cancel)
		self.Controls.Add(self._Okay)
		self.Controls.Add(self._label)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Name = "Form1"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		self.Text = "The path created is too long"
		self._errorProvider.EndInit()
		self.ResumeLayout(False)
		self.PerformLayout()
		self.Icon = System.Drawing.Icon(ICON)

	def CheckPathLength(self, sender, e):
		if self._Path.Text.Length >= 260:
			self._errorProvider.SetError(self._Path, "The entire path has to be less then 260 characters. Current path size is: " + str(self._Path.Text.Length))
			self.DialogResult = DialogResult.None
			return

		else:
			f = FileInfo(self._Path.Text)
			if len(f.DirectoryName) >= 248:
				self._errorProvider.SetError(self._Path, "The folder path has to be less then 248 characters. Current path size is: " + str(len(f.DirectoryName)))
				del(f)
				self.DialogResult = DialogResult.None
				return

		#Check for illegal characters. Note that there may be a : at the beginning. eg C:\
		path = self._Path.Text

		if len(path) > 1:
			if path[1] == ":":
				path = path [2:]
		r = re.search("[<>:|\*\?\"]", path)
		if r != None:
			self._errorProvider.SetError(self._Path, "The path cannot contain any of < > : | * ? \" ")
			self.DialogResult = DialogResult.None
			return
		self._errorProvider.SetError(self._Path, "")
		return

	def OkayClick(self, sender, e):
		
		self.CheckPathLength(self, sender, e)


class ReportForm(Form):

	def __init__(self):
		
		self.Size = Size(780, 400)
		self.StartPosition = FormStartPosition.CenterParent

		NumberColumn = System.Windows.Forms.DataGridViewTextBoxColumn()
		NumberColumn.HeaderText = "#"
		NumberColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
		
		ActionColumn = System.Windows.Forms.DataGridViewTextBoxColumn()
		ActionColumn.HeaderText = "Action"
		ActionColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells

		PathColumn = System.Windows.Forms.DataGridViewTextBoxColumn()
		PathColumn.HeaderText = "Path"
		PathColumn.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True
		PathColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill

		MessageColumn = System.Windows.Forms.DataGridViewTextBoxColumn()
		MessageColumn.HeaderText = "Message"
		MessageColumn.MinimumWidth = 200
		MessageColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
		MessageColumn.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True

		self.DataGrid = System.Windows.Forms.DataGridView()
		self.DataGrid.Height = 320
		self.DataGrid.Width = self.ClientSize.Width
		self.DataGrid.Columns.Add(NumberColumn)
		self.DataGrid.Columns.Add(ActionColumn)
		self.DataGrid.Columns.Add(PathColumn)
		self.DataGrid.Columns.Add(MessageColumn)
		self.DataGrid.RowHeadersVisible = False
		self.DataGrid.ReadOnly = True
		self.DataGrid.AllowUserToResizeRows = False
		self.DataGrid.AllowUserToResizeColumns = False
		self.DataGrid.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self.DataGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells

		
		self.Controls.Add(self.DataGrid)

		
		self.Okay = Button()
		self.Okay.Text = "OK"
		self.Okay.Location = Point(680, 330)
		self.Okay.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self.Okay.DialogResult = DialogResult.OK

		self.Save = Button()
		self.Save.Text = "Save"
		self.Save.Location = Point(600, 330)
		self.Save.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self.Save.DialogResult = DialogResult.Yes

		self.Controls.Add(self.Save)
		self.Controls.Add(self.Okay)
		self.Icon = System.Drawing.Icon(ICON)
		self.Text = "Library Organizer Report"
		self.AcceptButton = self.Okay
		
	def CheckDataGrid(self):
		for row in self.DataGrid.Rows:
			if row.Cells[1].Value == "Failed":
				row.DefaultCellStyle.ForeColor = System.Drawing.Color.Red
		

	def LoadData(self, data):
		"""
		Data should be an array containing arrays of strings.
		"""
		count = 1
		for row in data:
			self.DataGrid.Rows.Add(System.Array[System.String]([str(count)] + row))
			count += 1

		self.CheckDataGrid()