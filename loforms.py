"""loforms.py
This file contains various dialogs and forms required by the Library Orgainizer


Version 1.4
	-Moved various dialogs into here
	-Added SelectionFrom

Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.
"""

import clr

import System

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import ListBox, Button, CheckBox, Form, FormBorderStyle, FormStartPosition, DialogResult, Label, TextBox

clr.AddReference("System.Drawing")

from System.Drawing import Point, Size, ContentAlignment

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

		self.UseRegardless.Text = "Use every " + field + " in the selection list, even if the issue does not have that " + field

		if args.Series:
			self.Label.Text = "Select which " + fieldplural + " you would like to use for each issue in the series " + args.BookText
		else:
			self.Label.Text = "Select which " + fieldplural + " you would like to use in the issue " + args.BookText

		self.SLabel.Text = "Selected " + fieldplural
		self.ILabel.Text = fieldplural
		self.UseAsFolder.Text = "Use each " + field + " as a folder"

		self.Items.Items.AddRange(System.Array[System.String](args.Items))

	def InitializeComponent(self):
		self.Label = Label()
		self.Items = ListBox()
		self.Selection = ListBox()
		self.Okay = Button()
		self.Up = Button()
		self.Down = Button()
		self.UseRegardless = CheckBox()
		self.UseAsFolder = CheckBox()
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
		self.UseRegardless.Size = Size(400, 40)
		self.UseRegardless.TabIndex = 7
		self.UseRegardless.UseVisualStyleBackColor = True
		self.UseRegardless.Enabled = True
		# 
		# UseAsFolder
		# 
		self.UseAsFolder.Location = Point(12, 230)
		self.UseAsFolder.Name = "UseAsFolder"
		self.UseAsFolder.Size = Size(200, 24)
		self.UseAsFolder.TabIndex = 8
		self.UseAsFolder.Text = "Seperate each item with a folder"
		self.UseAsFolder.UseVisualStyleBackColor = True
		# 
		# Okay
		# 
		self.Okay.Location = Point(350, 230)
		self.Okay.Name = "Okay"
		self.Okay.Size = Size(75, 23)
		self.Okay.TabIndex = 9
		self.Okay.Text = "Okay"
		self.Okay.UseVisualStyleBackColor = True
		self.Okay.DialogResult = DialogResult.OK
		self.Okay.BringToFront()
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
		self.ClientSize = Size(437, 265)
		self.ControlBox = False
		self.Controls.Add(self.Label)
		self.Controls.Add(self.UseAsFolder)
		self.Controls.Add(self.UseRegardless)
		self.Controls.Add(self.Remove)
		self.Controls.Add(self.Add)
		self.Controls.Add(self.Down)
		self.Controls.Add(self.Up)
		self.Controls.Add(self.Okay)
		self.Controls.Add(self.Selection)
		self.Controls.Add(self.Items)
		self.Controls.Add(self.SLabel)
		self.Controls.Add(self.ILabel)
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.ShowIcon = True
		self.AcceptButton = self.Okay

		self.StartPosition = FormStartPosition.CenterParent



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
		return SelectionFormResult(list(self.Selection.Items), self.UseRegardless.Checked, self.UseAsFolder.Checked)

class SelectionFormResult(object):
	
	def __init__(self, selection, every = False, folder = False):
		self.Selection = selection
		self.EveryIssue = every
		self.Folder = folder

class SelectionFormArgs(object):

	def __init__(self, items, fieldtext, booktext, series):
		self.Items = items
		self.FieldText = fieldtext
		self.BookText = booktext
		self.Series = series

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