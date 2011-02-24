"""
loduplicate.py

Author: Stonepaw

Version 1.0

Description: Contains a Form for displaying two comic images and various information


Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.

"""



import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

from locommon import ICON

class DuplicateForm(Form):
	def __init__(self, existingbook, movingbook, ComicRack, existingpath):
		self.InitializeComponent()
		
		self.Icon = System.Drawing.Icon(ICON)
		
		if existingbook:
			self._existingcomicpath.Text = existingbook.FilePath
			
			self._existinginfo.Text = "Series: %s \n\nVolume: %s \n\nNumber: %s \n\nPublished Date: %s, %s\n\n# of pages: %s\n\nAdded to library: %s\n\nFile Size: %s" % \
										(existingbook.ShadowSeries, existingbook.ShadowVolume, existingbook.ShadowNumber, existingbook.Month, existingbook.Year, \
										existingbook.PageCount, existingbook.AddedTime, existingbook.FileSizeAsText)
			self._existingcover.Image = ComicRack.App.GetComicThumbnail(existingbook, existingbook.PreferredFrontCover)
		else:
			self._existingcomicpath.Text = existingpath
			self._existinginfo.Text = "Comic not in library"

		self._movingcomicpath.Text = movingbook.FilePath
		self._movinginfo.Text = "Series: %s \n\nVolume: %s \n\nNumber: %s \n\nPublished Date: %s, %s\n\n# of pages: %s\n\nAdded to library: %s\n\nFile Size: %s" % \
									(movingbook.ShadowSeries, movingbook.ShadowVolume, movingbook.ShadowNumber, movingbook.Month, movingbook.Year, \
									movingbook.PageCount, movingbook.AddedTime, movingbook.FileSizeAsText)
									
		
		self._movingcover.Image = ComicRack.App.GetComicThumbnail(movingbook, movingbook.PreferredFrontCover)
	
	def InitializeComponent(self):
		self._label1 = System.Windows.Forms.Label()
		self._always = System.Windows.Forms.CheckBox()
		self._cancel = System.Windows.Forms.Button()
		self._rename = System.Windows.Forms.Button()
		self._replace = System.Windows.Forms.Button()
		self._movingcover = System.Windows.Forms.PictureBox()
		self._existingcover = System.Windows.Forms.PictureBox()
		self._existinginfo = System.Windows.Forms.Label()
		self._movinginfo = System.Windows.Forms.Label()
		self._movingcomicpath = System.Windows.Forms.Label()
		self._existingcomicpath = System.Windows.Forms.Label()
		self._gbexisting = System.Windows.Forms.GroupBox()
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._movingcover.BeginInit()
		self._existingcover.BeginInit()
		self._gbexisting.SuspendLayout()
		self._groupBox1.SuspendLayout()
		self.SuspendLayout()
		# 
		# label1
		# 
		self._label1.Location = System.Drawing.Point(12, 9)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(516, 28)
		self._label1.TabIndex = 0
		self._label1.Text = "A comic with the same name already exists in the calculated location. What would you like to do?"
		# 
		# always
		# 
		self._always.Location = System.Drawing.Point(12, 316)
		self._always.Name = "always"
		self._always.Size = System.Drawing.Size(245, 34)
		self._always.TabIndex = 2
		self._always.Text = "Always do this action for the rest of this operation"
		self._always.UseVisualStyleBackColor = True
		# 
		# cancel
		# 
		self._cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		self._cancel.Location = System.Drawing.Point(374, 322)
		self._cancel.Name = "cancel"
		self._cancel.Size = System.Drawing.Size(75, 23)
		self._cancel.TabIndex = 3
		self._cancel.Text = "Don't Move"
		self._cancel.UseVisualStyleBackColor = True
		# 
		# rename
		# 
		self._rename.DialogResult = System.Windows.Forms.DialogResult.Retry
		self._rename.Location = System.Drawing.Point(455, 322)
		self._rename.Name = "rename"
		self._rename.Size = System.Drawing.Size(187, 23)
		self._rename.TabIndex = 5
		self._rename.Text = "Move but rename to filename (1)"
		self._rename.UseVisualStyleBackColor = True
		# 
		# replace
		# 
		self._replace.DialogResult = System.Windows.Forms.DialogResult.Yes
		self._replace.Location = System.Drawing.Point(263, 322)
		self._replace.Name = "replace"
		self._replace.Size = System.Drawing.Size(105, 23)
		self._replace.TabIndex = 4
		self._replace.Text = "Move and replace"
		self._replace.UseVisualStyleBackColor = True
		# 
		# movingcover
		# 
		self._movingcover.Location = System.Drawing.Point(6, 19)
		self._movingcover.Name = "movingcover"
		self._movingcover.Size = System.Drawing.Size(156, 200)
		self._movingcover.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
		self._movingcover.TabIndex = 6
		self._movingcover.TabStop = False
		# 
		# existingcover
		# 
		self._existingcover.Location = System.Drawing.Point(6, 19)
		self._existingcover.Name = "existingcover"
		self._existingcover.Size = System.Drawing.Size(154, 200)
		self._existingcover.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
		self._existingcover.TabIndex = 1
		self._existingcover.TabStop = False
		# 
		# existinginfo
		# 
		self._existinginfo.Location = System.Drawing.Point(166, 19)
		self._existinginfo.Name = "existinginfo"
		self._existinginfo.Size = System.Drawing.Size(139, 200)
		self._existinginfo.TabIndex = 9
		self._existinginfo.Text = "label4"
		# 
		# movinginfo
		# 
		self._movinginfo.Location = System.Drawing.Point(168, 19)
		self._movinginfo.Name = "movinginfo"
		self._movinginfo.Size = System.Drawing.Size(137, 200)
		self._movinginfo.TabIndex = 10
		self._movinginfo.Text = "label5"
		# 
		# movingcomicpath
		# 
		self._movingcomicpath.Location = System.Drawing.Point(6, 222)
		self._movingcomicpath.Name = "movingcomicpath"
		self._movingcomicpath.Size = System.Drawing.Size(286, 44)
		self._movingcomicpath.TabIndex = 11
		self._movingcomicpath.Text = "label6"
		# 
		# existingcomicpath
		# 
		self._existingcomicpath.Location = System.Drawing.Point(6, 222)
		self._existingcomicpath.MaximumSize = System.Drawing.Size(299, 87)
		self._existingcomicpath.Name = "existingcomicpath"
		self._existingcomicpath.Size = System.Drawing.Size(299, 44)
		self._existingcomicpath.TabIndex = 12
		self._existingcomicpath.Text = "label7"
		# 
		# gbexisting
		# 
		self._gbexisting.AutoSize = True
		self._gbexisting.Controls.Add(self._existingcover)
		self._gbexisting.Controls.Add(self._existingcomicpath)
		self._gbexisting.Controls.Add(self._existinginfo)
		self._gbexisting.Location = System.Drawing.Point(12, 28)
		self._gbexisting.MaximumSize = System.Drawing.Size(311, 319)
		self._gbexisting.Name = "gbexisting"
		self._gbexisting.Size = System.Drawing.Size(311, 282)
		self._gbexisting.TabIndex = 13
		self._gbexisting.TabStop = False
		self._gbexisting.Text = "Existing Comic"
		# 
		# groupBox1
		# 
		self._groupBox1.BackColor = System.Drawing.SystemColors.Window
		self._groupBox1.Controls.Add(self._movingcover)
		self._groupBox1.Controls.Add(self._movinginfo)
		self._groupBox1.Controls.Add(self._movingcomicpath)
		self._groupBox1.Location = System.Drawing.Point(331, 28)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(311, 282)
		self._groupBox1.TabIndex = 14
		self._groupBox1.TabStop = False
		self._groupBox1.Text = "Comic being moved"
		# 
		# DuplicateForm
		# 
		self.AcceptButton = self._replace
		self.BackColor = System.Drawing.SystemColors.Window
		self.CancelButton = self._cancel
		self.ClientSize = System.Drawing.Size(654, 352)
		self.Controls.Add(self._groupBox1)
		self.Controls.Add(self._gbexisting)
		self.Controls.Add(self._rename)
		self.Controls.Add(self._replace)
		self.Controls.Add(self._cancel)
		self.Controls.Add(self._always)
		self.Controls.Add(self._label1)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Name = "DuplicateForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		self.Text = "Duplicate found"
		self._movingcover.EndInit()
		self._existingcover.EndInit()
		self._gbexisting.ResumeLayout(False)
		self._groupBox1.ResumeLayout(False)
		self.ResumeLayout(False)
		self.PerformLayout()

