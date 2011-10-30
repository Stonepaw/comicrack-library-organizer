"""
loworkerform.py

Contains the worker form. All the copying is done in the background worker of this form.

Author: Stonepaw

Version 1.6
			added cancel button
		
			Fixed progress bar percenage calculation for when proccessing file in amounts greated than 100

			

Changes: 	Added undo worker form.

Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.
"""


import clr
import System
import System.Drawing
clr.AddReference("System.Windows.Forms")

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
from System.Windows.Interop import WindowInteropHelper

import System.IO
from System.IO import StreamWriter, TextWriter

import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

import loduplicate
from loduplicate import DuplicateForm

from locommon import ICON, Mode

from lobookmover import BookMover, OverwriteAction, UndoMover

import loforms

class WorkerForm(Form):
	def __init__(self, b, s):
		#print "Intializing compents"
		self.InitializeComponent()
		self.books = b
		self.settings = s
		#The percentage to raise to progress bar of one book.
		self.percentage = 1.0/len(b)*100
		self.progress = 0.0
		self._DuplicateForm = None
		#pass the ComicRack object
		loduplicate.ComicRack = ComicRack
		
		if self.settings.Mode == Mode.Copy:
			self.Text = "Copying Files..."

		elif self.settings.Mode == Mode.Test:
			self.Text = "Simulating Moving Files..."


	def InitializeComponent(self):
		self._progress = System.Windows.Forms.ProgressBar()
		self._Worker = System.ComponentModel.BackgroundWorker()
		self._Cancel = System.Windows.Forms.Button()
		self.SuspendLayout()
		# 
		# progress
		# 
		self._progress.Dock = System.Windows.Forms.DockStyle.Top
		self._progress.Location = System.Drawing.Point(0, 0)
		self._progress.Name = "progress"
		self._progress.Size = System.Drawing.Size(350, 23)
		self._progress.TabIndex = 0
		#
		# Cancel
		#
		self._Cancel.Location = System.Drawing.Point(141, 29)
		self._Cancel.Size = System.Drawing.Size(75, 23)
		self._Cancel.TabIndex = 1
		self._Cancel.Text = "Cancel"
		self._Cancel.Click += self.WorkerFormFormClosing
		# 
		# Worker
		# 
		self._Worker.WorkerReportsProgress = True
		self._Worker.WorkerSupportsCancellation = True
		self._Worker.DoWork += self.WorkerDoWork
		self._Worker.ProgressChanged += self.WorkerProgressChanged
		self._Worker.RunWorkerCompleted += self.WorkerRunWorkerCompleted
		# 
		# WorkerForm
		# 
		self.ClientSize = System.Drawing.Size(350, 55)
		self.Controls.Add(self._progress)
		self.Controls.Add(self._Cancel)
		self.CancelButton = self._Cancel
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Icon = System.Drawing.Icon(ICON)
		self.Name = "WorkerForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Moving Files..."
		self.FormClosing += self.WorkerFormFormClosing
		self.Shown += self.WorkerFormLoad
		self.ResumeLayout(False)


	def WorkerDoWork(self, sender, e):
		books = e.Argument[0]
		settings = e.Argument[1]
		
		mover = BookMover(sender, self, books, settings)
		result = mover.MoveBooks()
		e.Result = result

	def WorkerProgressChanged(self, sender, e):
		#self.Text = "Moving book " + str(e.ProgressPercentage) + " of " + self.total
		self.progress += self.percentage
		self._progress.Value = int(round(self.progress))

	def WorkerRunWorkerCompleted(self, sender, e):
		#print "Thread completed"
		if not e.Error:
			if self.settings.Mode == Mode.Test:
				report = loforms.ReportForm()
				report.LoadData(e.Result[2].ToArray())
				r = report.ShowDialog()
				if r == DialogResult.Yes:
					save = SaveFileDialog()
					save.AddExtension = True
					save.Filter = "Text files (*.txt)|*.txt"
					if save.ShowDialog() == DialogResult.OK:
						e.Result[2].SaveLog(save.FileName)
				report.Dispose()
			else:
				if e.Result[0] > 0:
					result = MessageBox.Show(e.Result[1] + "\n\nYou you like to see a detailed report of the failed or skipped files?", "View full report?", MessageBoxButtons.YesNo)
					if result == DialogResult.Yes:
						report = loforms.ReportForm()
						report.LoadData(e.Result[2].ToArray())
						r = report.ShowDialog()
						if r == DialogResult.Yes:
							save = SaveFileDialog()
							save.AddExtension = True
							save.Filter = "Text files (*.txt)|*.txt"
							if save.ShowDialog() == DialogResult.OK:
								e.Result[2].SaveLog(save.FileName)
						report.Dispose()

				else:
					MessageBox.Show(e.Result[1])

		else:
			MessageBox.Show(str(e.Error))
		self.Close()

	def WorkerFormLoad(self, sender, e):
		self._Worker.RunWorkerAsync([self.books, self.settings])
	
	def WorkerFormFormClosing(self, sender, e):
		if self._Worker.IsBusy:
			if self._Worker.CancellationPending == False:
				self._Worker.CancelAsync()
			self.DialogResult = DialogResult.None

	def ShowDuplicateForm(self, newbook, oldbook, renamefile, count):

		if self._DuplicateForm == None:
			self._DuplicateForm = DuplicateForm(self.settings.Mode)
			helper = WindowInteropHelper(self._DuplicateForm.win)
			helper.Owner = self.Handle

		return self._DuplicateForm.ShowForm(newbook, oldbook, renamefile, count)


class ProfileSelector(Form):
	
	def __init__(self, names, selected):
		self.label1 = System.Windows.Forms.Label()
		self.Profile = System.Windows.Forms.ComboBox()
		self.OK = System.Windows.Forms.Button()
		# 
		# label1
		# 
		self.label1.Location = System.Drawing.Point(12, 9)
		self.label1.Name = "label1"
		self.label1.Size = System.Drawing.Size(250, 19)
		self.label1.TabIndex = 0
		self.label1.Text = "Please select which profile you would like to use"
		# 
		# Profile
		# 
		self.Profile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self.Profile.FormattingEnabled = True
		self.Profile.Location = System.Drawing.Point(12, 33)
		self.Profile.Name = "Profile"
		self.Profile.Size = System.Drawing.Size(156, 21)
		self.Profile.TabIndex = 1
		self.Profile.Items.AddRange(System.Array[str](names))
		self.Profile.SelectedItem = selected
		# 
		# OK
		# 
		self.OK.Location = System.Drawing.Point(190, 31)
		self.OK.Name = "OK"
		self.OK.Size = System.Drawing.Size(62, 23)
		self.OK.TabIndex = 2
		self.OK.Text = "OK"
		self.OK.UseVisualStyleBackColor = True
		self.OK.DialogResult = DialogResult.OK
		# 
		# Form1
		# 
		self.AcceptButton = self.OK
		self.ClientSize = System.Drawing.Size(270, 61)
		self.Controls.Add(self.OK)
		self.Controls.Add(self.Profile)
		self.Controls.Add(self.label1)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Name = "Profile Selector"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		self.Text = "Profile Selector"
		self.Icon = System.Drawing.Icon(ICON)
		
class ReportForm(Form):
	def __init__(self, text):
		self.InitializeComponent()
		self._report.Text = text
	
	def InitializeComponent(self):
		self._button1 = System.Windows.Forms.Button()
		self._report = System.Windows.Forms.RichTextBox()
		self.SuspendLayout()
		# 
		# button1
		# 
		self._button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._button1.DialogResult = System.Windows.Forms.DialogResult.OK
		self._button1.Location = System.Drawing.Point(370, 399)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(75, 23)
		self._button1.TabIndex = 1
		self._button1.Text = "OK"
		self._button1.UseVisualStyleBackColor = True
		# 
		# report
		# 
		self._report.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._report.Location = System.Drawing.Point(12, 12)
		self._report.Name = "report"
		self._report.ReadOnly = True
		self._report.Size = System.Drawing.Size(425, 379)
		self._report.TabIndex = 4
		self._report.Text = ""
		# 
		# lomessagebox
		# 
		self.AcceptButton = self._button1
		self.ClientSize = System.Drawing.Size(453, 435)
		self.Controls.Add(self._report)
		self.Controls.Add(self._button1)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Icon = System.Drawing.Icon(ICON)
		self.Name = "ReportForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Report"
		self.ResumeLayout(False)

class WorkerFormUndo(WorkerForm):
	"""
	Can use pretty much all the code from the WorkerForm. Just have to override a couple of things
	"""
	
	def __init__(self, b, s):
		super(WorkerFormUndo,self).__init__(b, s)
		self.Text = "Undoing last move..."

	def WorkerRunWorkerCompleted(self, sender, e):
		#print "Thread completed"
		if not e.Error:
			if e.Result[0] > 0:
				result = MessageBox.Show(e.Result[1] + "\n\nYou you like to see a detailed report of the failed or skipped files?", "View full report?", MessageBoxButtons.YesNo)
				if result == DialogResult.Yes:
					report = ReportForm(e.Result[2])
					report.ShowDialog()
					report.Dispose()

			else:
				MessageBox.Show(e.Result[1])

		else:
			MessageBox.Show(str(e.Error))
		self.Close()

	def WorkerDoWork(self, sender, e):
		dict = e.Argument[0]
		settings = e.Argument[1]
		
		mover = UndoMover(sender, self, dict, settings)
		result = mover.MoveBooks()
		e.Result = result