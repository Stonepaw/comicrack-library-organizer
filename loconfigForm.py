"""
loconfigform.py

Version: 1.6

			Misc. GUI tweaks
		
Contains the config form. Most functions are related to makeing the GUI work. Several functions are related to settings.


Author: Stonepaw

Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.
"""

import clr
import System
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *


import losettings

from locommon import ICON, ExcludeRule, ExcludeGroup, Mode

import lobookmover
from lobookmover import PathMaker, ExcludePath, ExcludeMeta

import pyevent

import loforms
from loforms import InputBox, SelectionForm

class ConfigForm(Form):
	def __init__(self, books, allsettings, lastused):

		self.Icon = System.Drawing.Icon(ICON)
		InsertControl.Insert += self.InsertItem
		self.InitializeComponent()
		
		#Books for displaying sample text
		self.allbooks = books
		self.samplebook = self.allbooks[self._vsbBookSelector.Value]
		self._vsbBookSelector.Maximum = len(self.allbooks)
		
		#Holds all the profiles
		self.allsettings = allsettings
		
		#Sample text creator
		self.PathCreator = PathMaker(self)
		
		#In the case that lastused is none
		try:
			self.settings = allsettings[lastused]
		except KeyError:
			self.settings = allsettings[allsettings.keys()[0]]
			
		self._cmbEmptyData.SelectedIndex = 0
			
		self.LoadSettings()
		
		for i in allsettings:
			self._cmbProfiles.Items.Add(i)
		self._cmbProfiles.SelectedItem = self.settings.Name
		
		#For saving the previous index.
		self._cmbProfiles.Tag = self._cmbProfiles.SelectedIndex
	
		
	def InitializeComponent(self):
		self._tabs = System.Windows.Forms.TabControl()
		self._tpDirectory = System.Windows.Forms.TabPage()
		self._tpFilename = System.Windows.Forms.TabPage()
		self._txbDirStruct = System.Windows.Forms.TextBox()
		self._lblDirStruct = System.Windows.Forms.Label()
		self._txbBaseDir = System.Windows.Forms.TextBox()
		self._lblBaseDir = System.Windows.Forms.Label()
		self._btnBrowse = System.Windows.Forms.Button()
		self._ckbDirectory = System.Windows.Forms.CheckBox()
		self._gbInsertButtons = System.Windows.Forms.GroupBox()
		self._btnSep = System.Windows.Forms.Button()
		self._ckbFileNaming = System.Windows.Forms.CheckBox()
		self._txbFileStruct = System.Windows.Forms.TextBox()
		self._lblFileStruct = System.Windows.Forms.Label()
		self._FolderBrowser = System.Windows.Forms.FolderBrowserDialog()
		self._ckbSpace = System.Windows.Forms.CheckBox()
		self._label12 = System.Windows.Forms.Label()
		self._sampleTextDir = System.Windows.Forms.Label()
		self._sampleTextFile = System.Windows.Forms.Label()
		self._label14 = System.Windows.Forms.Label()
		self._ok = System.Windows.Forms.Button()
		self._cancel = System.Windows.Forms.Button()
		self._tpExcludes = System.Windows.Forms.TabPage()
		self._button1 = System.Windows.Forms.Button()
		self._btnProNew = System.Windows.Forms.Button()
		self._btnProSaveAs = System.Windows.Forms.Button()
		self._btnProDelete = System.Windows.Forms.Button()
		self._cmbProfiles = System.Windows.Forms.ComboBox()
		self._vsbBookSelector = System.Windows.Forms.VScrollBar()
		self._tpOptions = System.Windows.Forms.TabPage()
		self._insertTabs = System.Windows.Forms.TabControl()
		self._tpInsertBasic = System.Windows.Forms.TabPage()
		self._tpInsertAdvanced = System.Windows.Forms.TabPage()
		self._tpInsertAdvanced2 = System.Windows.Forms.TabPage()
		self._tabs.SuspendLayout()
		self._tpDirectory.SuspendLayout()
		self._tpFilename.SuspendLayout()
		self._gbInsertButtons.SuspendLayout()
		self._tpExcludes.SuspendLayout()
		self._tpOptions.SuspendLayout()
		self._insertTabs.SuspendLayout()
		self._tpInsertBasic.SuspendLayout()
		self._tpInsertAdvanced.SuspendLayout()
		self.SuspendLayout()
		# 
		# tabs
		# 
		self._tabs.Controls.Add(self._tpDirectory)
		self._tabs.Controls.Add(self._tpFilename)
		self._tabs.Controls.Add(self._tpExcludes)
		self._tabs.Controls.Add(self._tpOptions)
		self._tabs.Location = System.Drawing.Point(10, 6)
		self._tabs.Name = "tabs"
		self._tabs.SelectedIndex = 0
		self._tabs.Size = System.Drawing.Size(526, 452)
		self._tabs.TabIndex = 0
		self._tabs.SelectedIndexChanged += self.TabsSelectedIndexChanged

		self.CreateOptionsPage()
		self.CreateBasicInsertControls()
		self.CreateAdvancedInsertControls()
		self.CreateAdvanced2InsertControls()
		self.CreateRulesPage()
		# 
		# tpDirectory
		# 
		self._tpDirectory.Controls.Add(self._vsbBookSelector)
		self._tpDirectory.Controls.Add(self._sampleTextDir)
		self._tpDirectory.Controls.Add(self._label12)
		self._tpDirectory.Controls.Add(self._ckbDirectory)
		self._tpDirectory.Controls.Add(self._btnBrowse)
		self._tpDirectory.Controls.Add(self._lblBaseDir)
		self._tpDirectory.Controls.Add(self._txbBaseDir)
		self._tpDirectory.Controls.Add(self._lblDirStruct)
		self._tpDirectory.Controls.Add(self._txbDirStruct)
		self._tpDirectory.Controls.Add(self._gbInsertButtons)
		self._tpDirectory.Location = System.Drawing.Point(4, 22)
		self._tpDirectory.Name = "tpDirOrganize"
		self._tpDirectory.Padding = System.Windows.Forms.Padding(3)
		self._tpDirectory.Size = System.Drawing.Size(518, 426)
		self._tpDirectory.TabIndex = 0
		self._tpDirectory.Text = "Directories"
		self._tpDirectory.UseVisualStyleBackColor = True
		# 
		# txbDirStruct
		# 
		self._txbDirStruct.Location = System.Drawing.Point(113, 71)
		self._txbDirStruct.Name = "txbDirStruct"
		self._txbDirStruct.Size = System.Drawing.Size(397, 20)
		self._txbDirStruct.TabIndex = 5
		self._txbDirStruct.HideSelection = False
		# 
		# lblDirStruct
		# 
		self._lblDirStruct.Location = System.Drawing.Point(8, 74)
		self._lblDirStruct.Name = "lblDirStruct"
		self._lblDirStruct.Size = System.Drawing.Size(100, 14)
		self._lblDirStruct.TabIndex = 4
		self._lblDirStruct.Text = "Directory Structure:"
		# 
		# txbBaseDir
		# 
		self._txbBaseDir.Location = System.Drawing.Point(113, 38)
		self._txbBaseDir.Name = "txbBaseDir"
		self._txbBaseDir.ReadOnly = True
		self._txbBaseDir.Size = System.Drawing.Size(300, 20)
		self._txbBaseDir.TabIndex = 2
		# 
		# lblBaseDir
		# 
		self._lblBaseDir.Location = System.Drawing.Point(8, 41)
		self._lblBaseDir.Name = "lblBaseDir"
		self._lblBaseDir.Size = System.Drawing.Size(100, 17)
		self._lblBaseDir.TabIndex = 1
		self._lblBaseDir.Text = "Base Directory:"
		# 
		# btnBrowse
		# 
		self._btnBrowse.Location = System.Drawing.Point(419, 36)
		self._btnBrowse.Name = "btnBrowse"
		self._btnBrowse.Size = System.Drawing.Size(91, 23)
		self._btnBrowse.TabIndex = 3
		self._btnBrowse.Text = "Browse"
		self._btnBrowse.UseVisualStyleBackColor = True
		self._btnBrowse.Click += self.BtnBrowseClick
		# 
		# FolderBrowser
		# 
		self._FolderBrowser.Description = "Select a folder"
		# 
		# ckbDirectory
		# 
		self._ckbDirectory.Checked = True
		self._ckbDirectory.CheckState = System.Windows.Forms.CheckState.Checked
		self._ckbDirectory.Location = System.Drawing.Point(8, 6)
		self._ckbDirectory.Name = "ckbDirectory"
		self._ckbDirectory.Size = System.Drawing.Size(287, 24)
		self._ckbDirectory.TabIndex = 0
		self._ckbDirectory.Text = "Use directory organization"
		self._ckbDirectory.UseVisualStyleBackColor = True
		self._ckbDirectory.CheckedChanged += self.CkbDirectoryCheckedChanged
		# 
		# label12
		# 
		self._label12.Location = System.Drawing.Point(7, 98)
		self._label12.Name = "label12"
		self._label12.Size = System.Drawing.Size(62, 19)
		self._label12.TabIndex = 6
		self._label12.Text = "Example:"
		# 
		# sampleTextDir
		# 
		self._sampleTextDir.Location = System.Drawing.Point(63, 98)
		self._sampleTextDir.Name = "sampleTextDir"
		self._sampleTextDir.Size = System.Drawing.Size(424, 26)
		self._sampleTextDir.TabIndex = 7
		# 
		# groupbox insert buttons
		# 
		self._gbInsertButtons.Controls.Add(self._ckbSpace)
		self._gbInsertButtons.Controls.Add(self._btnSep)
		self._gbInsertButtons.Controls.Add(self._insertTabs)
		self._gbInsertButtons.Dock = System.Windows.Forms.DockStyle.Bottom
		self._gbInsertButtons.Location = System.Drawing.Point(3, 121)
		self._gbInsertButtons.Name = "groupBox1"
		self._gbInsertButtons.Size = System.Drawing.Size(512, 302)
		self._gbInsertButtons.TabIndex = 8
		self._gbInsertButtons.TabStop = False
		self._gbInsertButtons.Text = "Metadata"
		# 
		# btnSep
		# 
		self._btnSep.Location = System.Drawing.Point(185, 12)
		self._btnSep.Name = "btnSep"
		self._btnSep.Size = System.Drawing.Size(115, 23)
		self._btnSep.TabIndex = 0
		self._btnSep.Text = "Directory Seperator"
		self._btnSep.UseVisualStyleBackColor = True
		self._btnSep.Click += self.BtnSepClick
		# 
		# ckbSpace
		# 
		self._ckbSpace.AutoSize = True
		self._ckbSpace.Location = System.Drawing.Point(311, 16)
		self._ckbSpace.Name = "ckbSpace"
		self._ckbSpace.Size = System.Drawing.Size(188, 17)
		self._ckbSpace.TabIndex = 1
		self._ckbSpace.Text = "Space inserted fields automatically"
		self._ckbSpace.UseVisualStyleBackColor = True
		# 
		# tpFileNames
		# 
		self._tpFilename.Controls.Add(self._sampleTextFile)
		self._tpFilename.Controls.Add(self._label14)
		self._tpFilename.Controls.Add(self._lblFileStruct)
		self._tpFilename.Controls.Add(self._txbFileStruct)
		self._tpFilename.Controls.Add(self._ckbFileNaming)
		self._tpFilename.Location = System.Drawing.Point(4, 22)
		self._tpFilename.Name = "tpFileNames"
		self._tpFilename.Padding = System.Windows.Forms.Padding(3)
		self._tpFilename.Size = System.Drawing.Size(518, 426)
		self._tpFilename.TabIndex = 1
		self._tpFilename.Text = "File Names"
		self._tpFilename.UseVisualStyleBackColor = True
		# 
		# ckbFileNaming
		# 
		self._ckbFileNaming.Checked = True
		self._ckbFileNaming.CheckState = System.Windows.Forms.CheckState.Checked
		self._ckbFileNaming.Location = System.Drawing.Point(9, 7)
		self._ckbFileNaming.Name = "ckbFileNaming"
		self._ckbFileNaming.Size = System.Drawing.Size(432, 24)
		self._ckbFileNaming.TabIndex = 0
		self._ckbFileNaming.Text = "Use file naming"
		self._ckbFileNaming.UseVisualStyleBackColor = True
		self._ckbFileNaming.CheckedChanged += self.ChkFileNamingCheckedChanged
		# 
		# txbFileStruct
		# 
		self._txbFileStruct.Location = System.Drawing.Point(86, 37)
		self._txbFileStruct.Name = "txbFileStruct"
		self._txbFileStruct.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
		self._txbFileStruct.Size = System.Drawing.Size(424, 20)
		self._txbFileStruct.TabIndex = 1
		self._txbFileStruct.HideSelection = False
		# 
		# lblFileStruct
		# 
		self._lblFileStruct.Location = System.Drawing.Point(9, 34)
		self._lblFileStruct.Name = "lblFileStruct"
		self._lblFileStruct.Size = System.Drawing.Size(71, 35)
		self._lblFileStruct.TabIndex = 2
		self._lblFileStruct.Text = "File Structure"
		# 
		# sampleTextFile
		# 
		self._sampleTextFile.Location = System.Drawing.Point(66, 69)
		self._sampleTextFile.Name = "sampleTextFile"
		self._sampleTextFile.Size = System.Drawing.Size(421, 35)
		self._sampleTextFile.TabIndex = 10
		# 
		# label14
		# 
		self._label14.Location = System.Drawing.Point(10, 69)
		self._label14.Name = "label14"
		self._label14.Size = System.Drawing.Size(62, 19)
		self._label14.TabIndex = 9
		self._label14.Text = "Example:"
		# 
		# vsbBookSelector
		# 
		self._vsbBookSelector.LargeChange = 2
		self._vsbBookSelector.Location = System.Drawing.Point(490, 98)
		self._vsbBookSelector.Maximum = 1
		self._vsbBookSelector.Name = "vsbBookSelector"
		self._vsbBookSelector.Size = System.Drawing.Size(17, 26)
		self._vsbBookSelector.TabIndex = 10
		self._vsbBookSelector.ValueChanged += self.BookIndexChanged
		# 
		# insertTabs
		# 
		self._insertTabs.Controls.Add(self._tpInsertBasic)
		self._insertTabs.Controls.Add(self._tpInsertAdvanced)
		self._insertTabs.Controls.Add(self._tpInsertAdvanced2)
		self._insertTabs.Location = System.Drawing.Point(3, 19)
		self._insertTabs.Name = "insertTabs"
		self._insertTabs.SelectedIndex = 0
		self._insertTabs.Size = System.Drawing.Size(506, 280)
		self._insertTabs.TabIndex = 59

		# 
		# button1
		# 
		self._button1.Location = System.Drawing.Point(10, 491)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(75, 23)
		self._button1.TabIndex = 3
		self._button1.Text = "Save"
		self._button1.UseVisualStyleBackColor = True
		# 
		# btnProNew
		# 
		self._btnProNew.Location = System.Drawing.Point(172, 464)
		self._btnProNew.Name = "btnProNew"
		self._btnProNew.Size = System.Drawing.Size(75, 23)
		self._btnProNew.TabIndex = 2
		self._btnProNew.Text = "New"
		self._btnProNew.UseVisualStyleBackColor = True
		self._btnProNew.Click += self.BtnProNewClick
		# 
		# btnProSaveAs
		# 
		self._btnProSaveAs.Location = System.Drawing.Point(91, 491)
		self._btnProSaveAs.Name = "btnProSaveAs"
		self._btnProSaveAs.Size = System.Drawing.Size(75, 23)
		self._btnProSaveAs.TabIndex = 4
		self._btnProSaveAs.Text = "Save As"
		self._btnProSaveAs.UseVisualStyleBackColor = True
		self._btnProSaveAs.Click += self.BtnProSaveAsClick
		# 
		# btnProDelete
		# 
		self._btnProDelete.Location = System.Drawing.Point(172, 491)
		self._btnProDelete.Name = "btnProDelete"
		self._btnProDelete.Size = System.Drawing.Size(75, 23)
		self._btnProDelete.TabIndex = 5
		self._btnProDelete.Text = "Delete"
		self._btnProDelete.UseVisualStyleBackColor = True
		self._btnProDelete.Click += self.BtnProDeleteClick
		# 
		# cmbProfiles
		# 
		self._cmbProfiles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbProfiles.FormattingEnabled = True
		self._cmbProfiles.Location = System.Drawing.Point(10, 464)
		self._cmbProfiles.Name = "cmbProfiles"
		self._cmbProfiles.Size = System.Drawing.Size(156, 21)
		self._cmbProfiles.Sorted = True
		self._cmbProfiles.TabIndex = 1
		self._cmbProfiles.SelectionChangeCommitted += self.CmbProfilesItemChanged
		# 
		# ok
		# 
		self._ok.DialogResult = System.Windows.Forms.DialogResult.OK
		self._ok.Location = System.Drawing.Point(380, 484)
		self._ok.Name = "ok"
		self._ok.Size = System.Drawing.Size(75, 23)
		self._ok.TabIndex = 6
		self._ok.Text = "OK"
		self._ok.UseVisualStyleBackColor = True
		self._ok.Click += self.OkClick
		# 
		# cancel
		# 
		self._cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		self._cancel.Location = System.Drawing.Point(461, 484)
		self._cancel.Name = "cancel"
		self._cancel.Size = System.Drawing.Size(75, 23)
		self._cancel.TabIndex = 7
		self._cancel.Text = "Cancel"
		self._cancel.UseVisualStyleBackColor = True
		# 
		# ConfigForm
		# 
		self.AcceptButton = self._ok
		self.CancelButton = self._cancel
		self.ClientSize = System.Drawing.Size(542, 519)
		self.Shown += self.FormShown
		self.Controls.Add(self._btnProNew)
		self.Controls.Add(self._btnProDelete)
		self.Controls.Add(self._cmbProfiles)
		self.Controls.Add(self._btnProSaveAs)
		self.Controls.Add(self._cancel)
		self.Controls.Add(self._ok)
		self.Controls.Add(self._button1)
		self.Controls.Add(self._tabs)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.Name = "ConfigForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Library Organizer"
		self._tabs.ResumeLayout(False)
		self._tpDirectory.ResumeLayout(False)
		self._tpDirectory.PerformLayout()
		self._tpFilename.ResumeLayout(False)
		self._tpFilename.PerformLayout()
		self._gbInsertButtons.ResumeLayout(False)
		self._gbInsertButtons.PerformLayout()
		self._tpExcludes.ResumeLayout(False)
		self._tpOptions.ResumeLayout(False)
		self._tpOptions.PerformLayout()
		self._gbMode.ResumeLayout(False)
		self._gbMode.PerformLayout()
		self._insertTabs.ResumeLayout(False)
		self._tpInsertBasic.ResumeLayout(False)
		self._tpInsertBasic.PerformLayout()
		self._tpInsertAdvanced.ResumeLayout(False)
		self._tpInsertAdvanced.PerformLayout()
		self.ResumeLayout(False)

	def CreateOptionsPage(self):
		# 
		# rdbModeMove
		# 
		self._rdbModeMove = RadioButton()
		self._rdbModeMove.AutoSize = True
		self._rdbModeMove.Location = System.Drawing.Point(11, 19)
		self._rdbModeMove.Size = System.Drawing.Size(52, 17)
		self._rdbModeMove.TabIndex = 0
		self._rdbModeMove.TabStop = True
		self._rdbModeMove.Tag = "Move"
		self._rdbModeMove.Text = "Move"
		self._rdbModeMove.CheckedChanged += self.ModeChange
		# 
		# rdbModeCopy
		# 
		self._rdbModeCopy = RadioButton()
		self._rdbModeCopy.AutoSize = True
		self._rdbModeCopy.Location = System.Drawing.Point(158, 19)
		self._rdbModeCopy.Size = System.Drawing.Size(49, 17)
		self._rdbModeCopy.TabIndex = 1
		self._rdbModeCopy.TabStop = True
		self._rdbModeCopy.Tag = "Copy"
		self._rdbModeCopy.Text = "Copy"
		self._rdbModeCopy.CheckedChanged += self.ModeChange
		# 
		# rdbModeTest
		# 
		self._rdbModeTest = RadioButton()
		self._rdbModeTest.AutoSize = True
		self._rdbModeTest.Location = System.Drawing.Point(354, 19)
		self._rdbModeTest.TabIndex = 2
		self._rdbModeTest.TabStop = True
		self._rdbModeTest.Tag = "Test"
		self._rdbModeTest.Text = "Simulate"
		self._rdbModeTest.CheckedChanged += self.ModeChange
		#
		#
		#
		labelSimulate = Label()
		labelSimulate.Location = Point(282, 47)
		labelSimulate.AutoSize = True
		labelSimulate.Text = "(no files touched, complete log file created)"
		#
		# ckbCopyMode
		#
		self._ckbCopyMode = CheckBox()
		self._ckbCopyMode.Location = System.Drawing.Point(104, 41)
		self._ckbCopyMode.Size = System.Drawing.Size(172, 24)
		self._ckbCopyMode.Text = "Add copied book to Library"
		# 
		# gbMode
		# 
		self._gbMode = GroupBox()
		self._gbMode.Controls.Add(self._rdbModeTest)
		self._gbMode.Controls.Add(self._rdbModeCopy)
		self._gbMode.Controls.Add(self._ckbCopyMode)
		self._gbMode.Controls.Add(self._rdbModeMove)
		self._gbMode.Controls.Add(labelSimulate)
		self._gbMode.Location = System.Drawing.Point(8, 6)
		self._gbMode.Size = System.Drawing.Size(497, 71)
		self._gbMode.TabIndex = 0
		self._gbMode.TabStop = False
		self._gbMode.Text = "Mode"
		# 
		# label
		# 
		label8 = Label()
		label8.AutoSize = True
		label8.Location = System.Drawing.Point(3, 32)
		label8.Size = System.Drawing.Size(204, 13)
		label8.Text = "(Leave empty to remove empty directories)"
		# 
		# label7
		# 
		label7 = Label()
		label7.AutoSize = True
		label7.Location = System.Drawing.Point(11, 19)
		label7.Size = System.Drawing.Size(181, 13)
		label7.Text = "Replace empty directory names with: "
		# 
		# txbEmptyDir
		# 
		self._txbEmptyDir = TextBox()
		self._txbEmptyDir.Location = System.Drawing.Point(207, 19)
		self._txbEmptyDir.Size = System.Drawing.Size(284, 20)
		self._txbEmptyDir.TabIndex = 1
		# 
		# label9
		# 
		label9 = Label()
		label9.AutoSize = True
		label9.Location = System.Drawing.Point(6, 57)
		label9.Size = System.Drawing.Size(441, 13)
		label9.Text = "When a field is empty, substitute the entered value:"
		# 
		# label10
		# 
		label10 = Label()
		label10.AutoSize = True
		label10.Location = System.Drawing.Point(6, 80)
		label10.Size = System.Drawing.Size(32, 13)
		label10.Text = "Field:"
		# 
		# cmbEmptyData
		# 
		self._cmbEmptyData = ComboBox()
		self._cmbEmptyData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbEmptyData.Items.AddRange(System.Array[System.Object](
			["Age Rating",
			"Alternate Count",
			"Alternate Number",
			"Alternate Series",
			"Characters",
			"Count",
			"Format",
			"Genre",
			"Imprint",
			"Manga",
			"Month",
			"Number",
			"Publisher",
			"Scan Information"
			"Series",
			"Series Complete", 
			"Start Year",
			"Tags",
			"Title",
			"Volume",
			"Writer",
			"Year"]))
		self._cmbEmptyData.Location = System.Drawing.Point(44, 76)
		self._cmbEmptyData.Size = System.Drawing.Size(121, 21)
		self._cmbEmptyData.Sorted = True
		self._cmbEmptyData.TabIndex = 2
		self._cmbEmptyData.SelectedIndexChanged += self.CmbEmptyDataSelectedIndexChanged
		# 
		# label11
		# 
		label11 = Label()
		label11.AutoSize = True
		label11.Location = System.Drawing.Point(171, 80)
		label11.Size = System.Drawing.Size(32, 13)
		label11.Text = "Substitution:"
		# 
		# txbEmptyData
		# 
		self._txbEmptyData = TextBox()
		self._txbEmptyData.Location = System.Drawing.Point(242, 76)
		self._txbEmptyData.Size = System.Drawing.Size(249, 20)
		self._txbEmptyData.TabIndex = 3
		self._txbEmptyData.Leave += self.TxbEmptyDataLeave
		#
		#
		#
		gbEmpty = GroupBox()
		gbEmpty.Controls.Add(label7)
		gbEmpty.Controls.Add(label8)
		gbEmpty.Controls.Add(self._txbEmptyDir)
		gbEmpty.Controls.Add(label9)
		gbEmpty.Controls.Add(label10)
		gbEmpty.Controls.Add(self._cmbEmptyData)
		gbEmpty.Controls.Add(label11)
		gbEmpty.Controls.Add(self._txbEmptyData)
		gbEmpty.Location = Point(8, 84)
		gbEmpty.Size = Size(497, 116)
		gbEmpty.Text = "Empty substitutions"
		# 
		# ckbRemoveEmptyDir
		# 
		self._ckbRemoveEmptyDir = CheckBox()
		self._ckbRemoveEmptyDir.Checked = True
		self._ckbRemoveEmptyDir.Location = System.Drawing.Point(8, 206)
		self._ckbRemoveEmptyDir.Size = System.Drawing.Size(261, 24)
		self._ckbRemoveEmptyDir.TabIndex = 4
		self._ckbRemoveEmptyDir.Text = "Remove empty directories"
		self._ckbRemoveEmptyDir.CheckedChanged += self.CkbRemoveEmptyDirCheckedChanged
		# 
		# label18
		# 
		label18 = Label()
		label18.Location = System.Drawing.Point(39, 229)
		label18.Size = System.Drawing.Size(230, 18)
		label18.Text = "...but never remove the following directories:"
		# 
		# lbRemoveEmptyDir
		# 
		self._lbRemoveEmptyDir = ListBox()
		self._lbRemoveEmptyDir.Location = System.Drawing.Point(52, 250)
		self._lbRemoveEmptyDir.Size = System.Drawing.Size(383, 69)
		self._lbRemoveEmptyDir.TabIndex = 5
		# 
		# btnAddEmptyDir
		# 
		self._btnAddEmptyDir = Button()
		self._btnAddEmptyDir.Location = System.Drawing.Point(444, 250)
		self._btnAddEmptyDir.Size = System.Drawing.Size(68, 23)
		self._btnAddEmptyDir.TabIndex = 6
		self._btnAddEmptyDir.Text = "Add"
		self._btnAddEmptyDir.Click += self.BtnAddEmptyDirClick
		# 
		# btnRemoveEmptyDir
		# 
		self._btnRemoveEmptyDir = Button()
		self._btnRemoveEmptyDir.Location = System.Drawing.Point(441, 296)
		self._btnRemoveEmptyDir.Size = System.Drawing.Size(69, 23)
		self._btnRemoveEmptyDir.TabIndex = 7
		self._btnRemoveEmptyDir.Text = "Remove"
		self._btnRemoveEmptyDir.Click += self.BtnRemoveEmptyDirClick
		#
		# ckbMultiOneDontAsk
		#
		self._ckbDontAskWhenMultiOne = CheckBox()
		self._ckbDontAskWhenMultiOne.Location = Point(8, 328)
		self._ckbDontAskWhenMultiOne.AutoSize = True
		self._ckbDontAskWhenMultiOne.Text = "If there is only one character, genre, tag, team, scanner or writer, then insert it without asking."
		# 
		# ckbFileless
		# 
		self._ckbFileless = CheckBox()
		self._ckbFileless.Location = System.Drawing.Point(8, 358)
		self._ckbFileless.Size = System.Drawing.Size(502, 41)
		self._ckbFileless.TabIndex = 8
		self._ckbFileless.Text = "Copy fileless comic's custom thumbnail image to the calaculated path. (Does not affect the source image at all)"
		self._ckbFileless.CheckedChanged += self.CkbFilelessCheckedChanged
		# 
		# cmbImageFormat
		# 
		self._cmbImageFormat = ComboBox()
		self._cmbImageFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbImageFormat.Enabled = False
		self._cmbImageFormat.Items.AddRange(System.Array[System.Object](
			[".bmp",
			".jpg",
			".png"]))
		self._cmbImageFormat.Location = System.Drawing.Point(190, 388)
		self._cmbImageFormat.Size = System.Drawing.Size(54, 21)
		self._cmbImageFormat.TabIndex = 9
		# 
		# label19
		# 
		label19 = Label()
		label19.Location = System.Drawing.Point(97, 391)
		label19.Size = System.Drawing.Size(87, 21)
		label19.Text = "Save image as:"
		# 
		# options tag page
		# 
		self._tpOptions.Controls.Add(self._gbMode)
		self._tpOptions.Controls.Add(gbEmpty)
		self._tpOptions.Controls.Add(self._ckbRemoveEmptyDir)
		self._tpOptions.Controls.Add(self._lbRemoveEmptyDir)
		self._tpOptions.Controls.Add(self._btnAddEmptyDir)
		self._tpOptions.Controls.Add(self._btnRemoveEmptyDir)
		self._tpOptions.Controls.Add(self._ckbDontAskWhenMultiOne)
		self._tpOptions.Controls.Add(self._cmbImageFormat)
		self._tpOptions.Controls.Add(label18)
		self._tpOptions.Controls.Add(label19)
		self._tpOptions.Controls.Add(self._ckbFileless)
		self._tpOptions.Location = System.Drawing.Point(4, 22)
		self._tpOptions.Padding = System.Windows.Forms.Padding(3)
		self._tpOptions.Size = System.Drawing.Size(518, 426)
		self._tpOptions.UseVisualStyleBackColor = True
		self._tpOptions.TabIndex = 4
		self._tpOptions.Text = "Options"

	def CreateRulesPage(self):
		# 
		# lbExFolder
		# 
		self._lbExFolder = ListBox()
		self._lbExFolder.FormattingEnabled = True
		self._lbExFolder.HorizontalScrollbar = True
		self._lbExFolder.Location = System.Drawing.Point(16, 42)
		self._lbExFolder.Size = System.Drawing.Size(391, 82)
		self._lbExFolder.Sorted = True
		self._lbExFolder.TabIndex = 1
		# 
		# btnAddExFolder
		# 
		self._btnAddExFolder = Button()
		self._btnAddExFolder.Location = System.Drawing.Point(413, 42)
		self._btnAddExFolder.Size = System.Drawing.Size(75, 23)
		self._btnAddExFolder.TabIndex = 2
		self._btnAddExFolder.Text = "Add"
		self._btnAddExFolder.Click += self.BtnAddExFolderClick
		# 
		# btnRemoveExFolder
		# 
		self._btnRemoveExFolder = Button()
		self._btnRemoveExFolder.Location = System.Drawing.Point(413, 101)
		self._btnRemoveExFolder.Name = "btnRemoveExFolder"
		self._btnRemoveExFolder.Size = System.Drawing.Size(75, 23)
		self._btnRemoveExFolder.TabIndex = 3
		self._btnRemoveExFolder.Text = "Remove"
		self._btnRemoveExFolder.Click += self.BtnRemoveExFolderClick
		# 
		# label13
		# 
		label13 = Label()
		label13.Location = System.Drawing.Point(6, 16)
		label13.Name = "label13"
		label13.Size = System.Drawing.Size(454, 23)
		label13.TabIndex = 0
		label13.Text = "Do not move eComics if they are located in any of these folders:"
		#
		# groupbox folder rules
		#
		gbfolderrules = GroupBox()
		gbfolderrules.Size = Size(498, 144)
		gbfolderrules.Location = Point(8, 6)
		gbfolderrules.Controls.Add(label13)
		gbfolderrules.Controls.Add(self._lbExFolder)
		gbfolderrules.Controls.Add(self._btnAddExFolder)
		gbfolderrules.Controls.Add(self._btnRemoveExFolder)
		gbfolderrules.Text = "Folder Rules"
		# 
		# cmbExcludeMode
		# 
		self._cmbExcludeMode = ComboBox()
		self._cmbExcludeMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbExcludeMode.FormattingEnabled = True
		self._cmbExcludeMode.Items.AddRange(System.Array[System.Object](
			["Do not",
			"Only"]))
		self._cmbExcludeMode.Location = System.Drawing.Point(6, 19)
		self._cmbExcludeMode.Name = "cmbExcludeMode"
		self._cmbExcludeMode.Size = System.Drawing.Size(57, 21)
		self._cmbExcludeMode.TabIndex = 0
		self._cmbExcludeMode.SelectedIndexChanged += self.ExcludeModeChange
		# 
		# label20
		# 
		label20 = Label()
		label20.AutoSize = True
		label20.Location = System.Drawing.Point(69, 22)
		label20.Size = System.Drawing.Size(129, 13)
		label20.Text = "move eComics that match"
		# 
		# cmbExcludeOperator
		# 
		self._cmbExcludeOperator = ComboBox()
		self._cmbExcludeOperator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbExcludeOperator.FormattingEnabled = True
		self._cmbExcludeOperator.Items.AddRange(System.Array[System.Object](
			["Any",
			"All"]))
		self._cmbExcludeOperator.Location = System.Drawing.Point(204, 19)
		self._cmbExcludeOperator.Name = "cmbExcludeOperator"
		self._cmbExcludeOperator.Size = System.Drawing.Size(43, 21)
		self._cmbExcludeOperator.TabIndex = 1
		self._cmbExcludeOperator.SelectedIndexChanged += self.ExcludeOperatorChange
		# 
		# label17
		# 
		label17 = Label()
		label17.AutoSize = True
		label17.Location = System.Drawing.Point(250, 22)
		label17.Size = System.Drawing.Size(106, 13)
		label17.Text = "of the following rules."
		# 
		# btnAddGroup
		# 
		self._btnAddGroup = Button()
		self._btnAddGroup.Location = System.Drawing.Point(362, 17)
		self._btnAddGroup.Size = System.Drawing.Size(71, 23)
		self._btnAddGroup.TabIndex = 2
		self._btnAddGroup.Text = "Add Group"
		self._btnAddGroup.Click += self.CreateRuleGroup
		# 
		# btnExMetaAdd
		# 
		self._btnExMetaAdd = Button()
		self._btnExMetaAdd.Location = System.Drawing.Point(439, 17)
		self._btnExMetaAdd.Size = System.Drawing.Size(59, 23)
		self._btnExMetaAdd.TabIndex = 3
		self._btnExMetaAdd.Text = "Add rule"
		self._btnExMetaAdd.Click += self.CreateRuleSet
		# 
		# flpExcludes
		# 
		self._flpExcludes = FlowLayoutPanel()
		self._flpExcludes.AutoSize = True
		self._flpExcludes.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
		self._flpExcludes.Location = System.Drawing.Point(3, 3)
		self._flpExcludes.Size = System.Drawing.Size(482, 40)
		self._flpExcludes.TabIndex = 4
		# 
		# ExPanel
		# 
		self._ExPanel = Panel()
		self._ExPanel.AutoScroll = True
		self._ExPanel.Controls.Add(self._flpExcludes)
		self._ExPanel.Location = System.Drawing.Point(3, 46)
		self._ExPanel.Name = "ExPanel"
		self._ExPanel.Size = System.Drawing.Size(495, 207)
		self._ExPanel.TabIndex = 4
		# 
		# groupBox2
		# 
		gbRules = GroupBox()
		gbRules.Controls.Add(self._cmbExcludeMode)
		gbRules.Controls.Add(label20)
		gbRules.Controls.Add(self._btnAddGroup)
		gbRules.Controls.Add(self._btnExMetaAdd)
		gbRules.Controls.Add(self._ExPanel)
		gbRules.Controls.Add(label17)
		gbRules.Controls.Add(self._cmbExcludeOperator)
		gbRules.Location = System.Drawing.Point(8, 156)
		gbRules.Size = System.Drawing.Size(504, 264)
		gbRules.TabIndex = 6
		gbRules.TabStop = False
		gbRules.Text = "Metadata Rules"
		# 
		# tpExcludes
		# 
		self._tpExcludes.Controls.Add(gbfolderrules)
		self._tpExcludes.Controls.Add(gbRules)
		self._tpExcludes.Name = "tpExcludes"
		self._tpExcludes.Padding = System.Windows.Forms.Padding(3)
		self._tpExcludes.Size = System.Drawing.Size(518, 426)
		self._tpExcludes.TabIndex = 3
		self._tpExcludes.Text = "Rules"
		self._tpExcludes.UseVisualStyleBackColor = True

	def CreateBasicInsertControls(self):
		#
		# Publisher
		#
		self.Publisher = InsertControl()
		self.Publisher.Location = Point(6, 25)
		self.Publisher.SetTemplate("publisher", "Publisher")
		# 
		# Imprint
		# 
		self.Imprint = InsertControl()
		self.Imprint.Location = Point(6, 53)
		self.Imprint.SetTemplate("imprint", "Imprint")
		# 
		# Series
		# 
		self.Series = InsertControl()
		self.Series.Location = Point(6, 82)
		self.Series.SetTemplate("series", "Series")
		# 
		# Title
		# 
		self.Title = InsertControl()
		self.Title.Location = Point(6, 110)
		self.Title.SetTemplate("title", "Title")
		# 
		# Format
		#
		self.Format = InsertControl()
		self.Format.Location = Point(6, 138)
		self.Format.SetTemplate("format", "Format")
		# 
		# Volume
		#
		self.Volume = InsertControlPadding()
		self.Volume.Location = Point(6, 166)
		self.Volume.SetTemplate ("volume", "Volume")
		#
		# Age Rating
		self.AgeRating = InsertControl()
		self.AgeRating.SetTemplate("ageRating", "Age Rate.")
		self.AgeRating.Location = Point(6, 222)
		#
		# 
		# AlternateSeries
		# 
		self.AlternateSeries = InsertControl()
		self.AlternateSeries.Location = Point(6, 194)
		self.AlternateSeries.SetTemplate("altSeries", "Alt. Series")
		# 
		# AlternateNumber
		# 
		self.AlternateNumber = InsertControlPadding()
		self.AlternateNumber.Location = Point(262, 194)
		self.AlternateNumber.SetTemplate("altNumber", "Alt. Num.")
		# 
		# AlternateCount
		# 
		self.AlternateCount = InsertControlPadding()
		self.AlternateCount.Location = Point(262, 166)
		self.AlternateCount.SetTemplate("altCount", "Alt. Count")
		# 
		# Year
		# 
		self.Year = InsertControl()
		self.Year.Location = Point(262, 138)
		self.Year.SetTemplate("year", "Year")
		# 
		# Month
		# 
		self.Month = InsertControl()
		self.Month.Location = Point(262, 108)
		self.Month.SetTemplate("month", "Month")
		# 
		# Month Number
		# 
		self.MonthNumber = InsertControlPadding()
		self.MonthNumber.Location = Point(262, 82)
		self.MonthNumber.SetTemplate("month#", "Month #")
		# 
		# Count
		# 
		self.Count = InsertControl()
		self.Count.Location = Point(262, 54)
		self.Count.SetTemplate("count", "Count")
		# 
		# Number
		# 
		self.Number = InsertControlPadding()
		self.Number.Location = Point(262, 26)
		self.Number.SetTemplate("number", "Number")
		# 
		# label1
		# 
		label1 = Label()
		label1.AutoSize = True
		label1.Location = System.Drawing.Point(21, 3)
		label1.Size = System.Drawing.Size(33, 13)
		label1.Text = "Prefix"
		# 
		# label2
		# 
		label2 = Label()
		label2.AutoSize = True
		label2.Location = System.Drawing.Point(148, 3)
		label2.Size = System.Drawing.Size(38, 13)
		label2.Text = "Postfix"
		# 
		# label3
		# 
		label3 = Label()
		label3.AutoSize = True
		label3.Location = System.Drawing.Point(404, 3)
		label3.Size = System.Drawing.Size(38, 13)
		label3.Text = "Postfix"
		# 
		# label4
		# 
		label4 = Label()
		label4.AutoSize = True
		label4.Location = System.Drawing.Point(275, 3)
		label4.Size = System.Drawing.Size(33, 13)
		label4.Text = "Prefix"
		# 
		# label5
		# 
		label5 = Label()
		label5.AutoSize = True
		label5.Location = System.Drawing.Point(210, 3)
		label5.Size = System.Drawing.Size(26, 13)
		label5.Text = "Pad"
		# 
		# label6
		# 
		label6 = Label()
		label6.AutoSize = True
		label6.Location = System.Drawing.Point(466, 3)
		label6.Size = System.Drawing.Size(26, 13)
		label6.Text = "Pad"

		# 
		# tpInsertBasic
		# 
		self._tpInsertBasic.Controls.Add(self.Publisher)
		self._tpInsertBasic.Controls.Add(self.Imprint)
		self._tpInsertBasic.Controls.Add(self.Series)
		self._tpInsertBasic.Controls.Add(self.Title)
		self._tpInsertBasic.Controls.Add(self.Format)
		self._tpInsertBasic.Controls.Add(self.Volume)
		self._tpInsertBasic.Controls.Add(self.AgeRating)
		self._tpInsertBasic.Controls.Add(self.AlternateSeries)
		self._tpInsertBasic.Controls.Add(self.Number)
		self._tpInsertBasic.Controls.Add(self.Count)		
		self._tpInsertBasic.Controls.Add(self.MonthNumber)
		self._tpInsertBasic.Controls.Add(self.Month)
		self._tpInsertBasic.Controls.Add(self.Year)
		self._tpInsertBasic.Controls.Add(self.AlternateCount)
		self._tpInsertBasic.Controls.Add(self.AlternateNumber)		
		self._tpInsertBasic.Controls.Add(label6)
		self._tpInsertBasic.Controls.Add(label5)
		self._tpInsertBasic.Controls.Add(label3)
		self._tpInsertBasic.Controls.Add(label4)
		self._tpInsertBasic.Controls.Add(label1)
		self._tpInsertBasic.Controls.Add(label2)
		self._tpInsertBasic.Padding = System.Windows.Forms.Padding(3)
		self._tpInsertBasic.Size = System.Drawing.Size(500, 254)
		self._tpInsertBasic.TabIndex = 0
		self._tpInsertBasic.Text = "Basic"
		self._tpInsertBasic.UseVisualStyleBackColor = True

	def CreateAdvancedInsertControls(self):
		# 
		# Start Year
		# 
		self.StartYear = InsertControl()
		self.StartYear.Location = Point(6, 25)
		self.StartYear.SetTemplate("startyear", "Start Year")
		# 
		# label15
		# 
		label15 = Label()
		label15.Location = System.Drawing.Point(220, 25)
		label15.Size = System.Drawing.Size(205, 30)
		label15.Text = "Start year is calculated from the earliest year for the series in your library"
		# 
		# Manga
		# 
		self.Manga = InsertControlTextBox()
		self.Manga.Location = Point(6, 100)
		self.Manga.SetTemplate("manga", "Manga")
		# 
		# SeriesComplete
		# 
		self.SeriesComplete = InsertControlTextBox()
		self.SeriesComplete.Location = Point(6, 145)
		self.SeriesComplete.SetTemplate("seriesComplete", "Series Complete")
		self.SeriesComplete.InsertButton.Width += 30
		self.SeriesComplete.Width += 30
		self.SeriesComplete.Postfix.Left += 30
		self.SeriesComplete.TextBox.Left += 30
		# 
		# label16
		# 
		label16 = Label()
		label16.AutoSize = True
		label16.Location = System.Drawing.Point(225, 85)
		label16.Size = System.Drawing.Size(33, 13)
		label16.Text = "Text"
		# 
		# label21
		# 
		label21 = Label()
		label21.Location = System.Drawing.Point(270, 100)
		label21.Size = System.Drawing.Size(230, 60)
		label21.Text = "Fill in the \"Text\" box for the text to be inserted when the item is marked as Yes."
		#
		# Label Pre
		lblPre = Label()
		lblPre.Size = System.Drawing.Size(33, 13)
		lblPre.Location = System.Drawing.Point(20, 3)
		lblPre.Text = "Prefix"
		#
		# Label Post
		#
		lblPost = Label()
		lblPost.Size = System.Drawing.Size(38, 13)
		lblPost.Location = System.Drawing.Point(150, 3)
		lblPost.Text = "Postfix"

		# 
		# tpInsertAdvanced
		# 
		self._tpInsertAdvanced.Controls.Add(self.StartYear)
		self._tpInsertAdvanced.Controls.Add(self.Manga)
		self._tpInsertAdvanced.Controls.Add(self.SeriesComplete)
		self._tpInsertAdvanced.Controls.Add(label21)
		self._tpInsertAdvanced.Controls.Add(label16)
		self._tpInsertAdvanced.Controls.Add(label15)
		self._tpInsertAdvanced.Controls.Add(lblPost)
		self._tpInsertAdvanced.Controls.Add(lblPre)
		self._tpInsertAdvanced.Padding = System.Windows.Forms.Padding(3)
		self._tpInsertAdvanced.Size = System.Drawing.Size(500, 254)
		self._tpInsertAdvanced.TabIndex = 1
		self._tpInsertAdvanced.Text = "Advanced"
		self._tpInsertAdvanced.UseVisualStyleBackColor = True

	def CreateAdvanced2InsertControls(self):
		# 
		# Characters
		# 
		self.Characters = InsertControlCheckBox()
		self.Characters.Location = Point(6, 20)
		self.Characters.SetTemplate("characters", "Character")
		# 
		# Genre
		# 
		self.Genre = InsertControlCheckBox()
		self.Genre.Location = Point(250, 20)
		self.Genre.SetTemplate("genre", "Genre")
		# 
		# Writer
		# 
		self.Writer = InsertControlCheckBox()
		self.Writer.Location = Point(250, 110)
		self.Writer.SetTemplate("writer", "Writer")
		# 
		# Tags
		# 
		self.Tags = InsertControlCheckBox()
		self.Tags.Location = Point(250, 65)
		self.Tags.SetTemplate("tags", "Tags")
		#
		# Team
		#
		self.Teams = InsertControlCheckBox()
		self.Teams.SetTemplate("teams", "Team")
		self.Teams.Location = Point(6, 110)
		#
		# Scan Information
		#
		self.ScanInformation = InsertControlCheckBox()
		self.ScanInformation.SetTemplate("scaninfo", "Scan Info.")
		self.ScanInformation.Location = Point(6, 65)
		#
		# Label Seperator
		#
		lblSep = Label()
		lblSep.Location = System.Drawing.Point(190, 48)
		lblSep.Size = System.Drawing.Size(58, 13)
		lblSep.Text = "Seperator"
		#
		# Label Seperator 2
		#
		lblSep2 = Label()
		lblSep2.Size = System.Drawing.Size(58, 13)
		lblSep2.Location = System.Drawing.Point(430, 48)
		lblSep2.Text = "Seperator"
		# 
		# label22
		# 
		label22 = Label()
		label22.Location = System.Drawing.Point(4, 155)
		label22.Size = System.Drawing.Size(494, 48)
		label22.Text = "For these fields that can have multiple entries the script will ask which ones you would like to use. If you want to select for all the issues in the series at once, check the checkbox beside the field before inserting. In the selction box is an option to have each entry as its own folder"
		#
		# Label Pre
		lblPre3 = Label()
		lblPre3.Size = System.Drawing.Size(33, 13)
		lblPre3.Location = System.Drawing.Point(20, 3)
		lblPre3.Text = "Prefix"
		#
		# Label Pre2
		#
		lblPre2 = Label()
		lblPre2.Size = System.Drawing.Size(33, 13)
		lblPre2.Location = System.Drawing.Point(260, 3)
		lblPre2.Text = "Prefix"
		#
		# Label Post
		#
		lblPost3 = Label()
		lblPost3.Size = System.Drawing.Size(38, 13)
		lblPost3.Location = System.Drawing.Point(150, 3)
		lblPost3.Text = "Postfix"
		#
		# Label Post
		#
		lblPost2 = Label()
		lblPost2.Size = System.Drawing.Size(38, 13)
		lblPost2.Location = System.Drawing.Point(395, 3)
		lblPost2.Text = "Postfix"

		#
		# tpInsertAdvanced2
		#
		self._tpInsertAdvanced2.Controls.Add(self.Characters)
		self._tpInsertAdvanced2.Controls.Add(self.Genre)
		self._tpInsertAdvanced2.Controls.Add(self.Tags)
		self._tpInsertAdvanced2.Controls.Add(self.ScanInformation)
		self._tpInsertAdvanced2.Controls.Add(self.Teams)
		self._tpInsertAdvanced2.Controls.Add(self.Writer)
		self._tpInsertAdvanced2.Controls.Add(label22)
		self._tpInsertAdvanced2.Controls.Add(lblSep)
		self._tpInsertAdvanced2.Controls.Add(lblSep2)
		self._tpInsertAdvanced2.Controls.Add(lblPre2)
		self._tpInsertAdvanced2.Controls.Add(lblPre3)
		self._tpInsertAdvanced2.Controls.Add(lblPost2)
		self._tpInsertAdvanced2.Controls.Add(lblPost3)
		self._tpInsertAdvanced2.Padding = System.Windows.Forms.Padding(3)
		self._tpInsertAdvanced2.Size = System.Drawing.Size(500, 254)
		self._tpInsertAdvanced2.TabIndex = 2
		self._tpInsertAdvanced2.Text = "Advanced 2"
		self._tpInsertAdvanced2.UseVisualStyleBackColor = True


	def BtnSepClick(self, sender, e):
		self.InsertText("\\", self._txbDirStruct)

	def ChkFileNamingCheckedChanged(self, sender, e):
		self._txbFileStruct.Enabled = self._ckbFileNaming.Checked
		self._lblFileStruct.Enabled = self._ckbFileNaming.Checked
		self.UpdateSampleText()

	def TabsSelectedIndexChanged(self, sender, e):
		if self._tabs.SelectedIndex == 0:
			self._btnSep.Enabled = True
			self._btnSep.Visible = True
			self._tpFilename.Controls.Remove(self._gbInsertButtons)
			self._tpFilename.Controls.Remove(self._vsbBookSelector)
			self._tpDirectory.Controls.Add(self._vsbBookSelector)
			self._vsbBookSelector.Location = System.Drawing.Point(490, 98)
			self._tpDirectory.Controls.Add(self._gbInsertButtons)

		if self._tabs.SelectedIndex == 1:
			self._btnSep.Enabled = False
			self._btnSep.Visible = False
			self._tpDirectory.Controls.Remove(self._gbInsertButtons)
			self._tpDirectory.Controls.Remove(self._vsbBookSelector)
			self._tpFilename.Controls.Add(self._vsbBookSelector)
			self._vsbBookSelector.Location = System.Drawing.Point(490, 69)
			self._tpFilename.Controls.Add(self._gbInsertButtons)
		self.UpdateSampleText()

	def CkbDirectoryCheckedChanged(self, sender, e):
		self._txbDirStruct.Enabled = self._ckbDirectory.Checked
		self._txbBaseDir.Enabled = self._ckbDirectory.Checked
		self._lblBaseDir.Enabled = self._ckbDirectory.Checked
		self._lblDirStruct.Enabled = self._ckbDirectory.Checked
		self._btnBrowse.Enabled = self._ckbDirectory.Checked
		self.UpdateSampleText()

	#The following three function insert text into the correct textboxes
	def InsertItem(self, sender, e):
		s = sender.GetTemplateText(self._ckbSpace.Checked)

		if self._tabs.SelectedIndex == 0:
			self.InsertText(s, self._txbDirStruct)

		elif self._tabs.SelectedIndex == 1:
			self.InsertText(s, self._txbFileStruct)

	def InsertText(self, string, textbox):
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

	def UpdateSampleText(self):
		#Directory Preview
		if self._tabs.SelectedIndex == 0:
			if not ExcludePath(self.samplebook, list(self._lbExFolder.Items)) and not ExcludeMeta(self.samplebook, self.settings.ExcludeRules, self.settings.ExcludeOperator, self.settings.ExcludeMode) and self._ckbDirectory.Checked:
				self._sampleTextDir.Text = self.PathCreator.CreateDirectoryPath(self.samplebook, self._txbDirStruct.Text, self._txbBaseDir.Text, self._txbEmptyDir.Text, self.settings.EmptyData, self._ckbDontAskWhenMultiOne.Checked)
			else:
				self._sampleTextDir.Text = self.samplebook.FileDirectory
				
		#Filename preview
		elif self._tabs.SelectedIndex == 1:
			if not ExcludePath(self.samplebook, list(self._lbExFolder.Items)) and not ExcludeMeta(self.samplebook, self.settings.ExcludeRules, self.settings.ExcludeOperator, self.settings.ExcludeMode) and self._ckbFileNaming.Checked:
				self._sampleTextFile.Text = self.PathCreator.CreateFileName(self.samplebook, self._txbFileStruct.Text, self.settings.EmptyData, self._cmbImageFormat.SelectedItem, self._ckbDontAskWhenMultiOne.Checked)
			else:
				self._sampleTextFile.Text = self.samplebook.FileNameWithExtension

	def BtnBrowseClick(self, sender, e):
		self._FolderBrowser.ShowDialog()
		self._txbBaseDir.Text = self._FolderBrowser.SelectedPath
		self.UpdateSampleText()		
		
	def TxbStructTextChanged(self, sender, e):
		self.UpdateSampleText()

	def BookIndexChanged(self, sender, e):
		self.samplebook = self.allbooks[self._vsbBookSelector.Value]
		self.UpdateSampleText()

	def CmbEmptyDataSelectedIndexChanged(self, sender, e):
		#For ease of adding additional values later replace the spaces in Additions values. That way they match up with the dictorary keys later
		if sender.SelectedItem.replace(" ", "") in self.settings.EmptyData:
			self._txbEmptyData.Text = self.settings.EmptyData[sender.SelectedItem.replace(" ", "")]
		else:
			self._txbEmptyData.Text = ""

	def TxbEmptyDataLeave(self, sender, e):
		self.settings.EmptyData[self._cmbEmptyData.SelectedItem.replace(" ", "")] = sender.Text
		
	def LoadSettings(self):
		
		#Checkboxes
		
		self._ckbDirectory.Checked = self.settings.UseDirectory
		self._ckbFileNaming.Checked = self.settings.UseFileName
		self._ckbDontAskWhenMultiOne.Checked = self.settings.DontAskWhenMultiOne
		
		#Move mode
		if self.settings.Mode == Mode.Move:
			self._rdbModeMove.Checked = True
		elif self.settings.Mode == Mode.Copy:
			self._rdbModeCopy.Checked = True
		elif self.settings.Mode == Mode.Test:
			self._rdbModeTest.Checked = True

		#
		self._ckbCopyMode.Checked = self.settings.CopyMode
		
		#Reload the selected index of the emptydata cmb
		self.CmbEmptyDataSelectedIndexChanged(self._cmbEmptyData, None)
		
		#Excludes
		

		self._lbExFolder.Items.Clear()
		for i in self.settings.ExcludeFolders:
			self._lbExFolder.Items.Add(i)
		

		#Exclude rules
		#If changing settings to a different one. Clear the exlcude rule  container
		
		self._flpExcludes.Controls.Clear()

		for i in self.settings.ExcludeRules:
			#Check for none in the case of switching profiles with deleted rules
			if not i == None:
				i.Remove.Click += self.RemoveRuleSet
				self._flpExcludes.Controls.Add(i.Panel)

		
		self._cmbExcludeOperator.SelectedItem = self.settings.ExcludeOperator
		self._cmbExcludeMode.SelectedItem = self.settings.ExcludeMode
		
		#Fileless

		self._ckbFileless.Checked = self.settings.MoveFileless
		self._cmbImageFormat.SelectedItem = self.settings.FilelessFormat

		#Empty directories
		self._ckbRemoveEmptyDir.Checked = self.settings.RemoveEmptyDir
		self._lbRemoveEmptyDir.Items.Clear()
		self._lbRemoveEmptyDir.Items.AddRange(System.Array[System.String](self.settings.ExcludedEmptyDir))


		#Prefix
		for i in self.settings.Prefix:
			getattr(self, i).SetPrefixText(self.settings.Prefix[i])

		#Postfix
		for i in self.settings.Postfix:
			getattr(self, i).SetPostfixText(self.settings.Postfix[i])

		#Seperator
		for i in self.settings.Seperator:
			getattr(self, i).SetSeperatorText(self.settings.Seperator[i])

		#Textbox
		for i in self.settings.TextBox:
			getattr(self, i).SetTextBoxText(self.settings.TextBox[i])

		

		#Other stuff has to be loaded before text boxes, otherwise the sample text is updated before everyother setting is set
		#Base Textboxes
		self._txbBaseDir.Text = self.settings.BaseDir
		self._txbDirStruct.Text = self.settings.DirTemplate
		self._txbFileStruct.Text = self.settings.FileTemplate
		self._txbEmptyDir.Text = self.settings.EmptyDir
		
		self._txbDirStruct.SelectionStart = len(self._txbDirStruct.Text)
		self._txbFileStruct.SelectionStart = len(self._txbFileStruct.Text)

	def SaveSettings(self):
		#Base Textboxes
		self.settings.BaseDir = self._txbBaseDir.Text
		self.settings.DirTemplate = self._txbDirStruct.Text
		self.settings.FileTemplate = self._txbFileStruct.Text
		self.settings.EmptyDir = self._txbEmptyDir.Text
		self.settings.DontAskWhenMultiOne = self._ckbDontAskWhenMultiOne.Checked
		
		self.settings.CopyMode = self._ckbCopyMode.Checked

		#Postfix and Prefix
		for i in self.settings.Prefix:
			self.settings.Prefix[i] = getattr(self, i).GetPrefixText()

		#Postfix and Prefix
		for i in self.settings.Postfix:
			self.settings.Postfix[i] = getattr(self, i).GetPostfixText()

		#Seperators:
		for i in self.settings.Seperator:
			self.settings.Seperator[i] = getattr(self, i).GetSeperatorText()

		#Textboxes:
		for i in self.settings.TextBox:
			self.settings.TextBox[i] = getattr(self, i).GetTextBoxText()

		#Checkboxes
		
		self.settings.UseDirectory = self._ckbDirectory.Checked
		self.settings.UseFileName = self._ckbFileNaming.Checked
		
		self.settings.MoveFileless = self._ckbFileless.Checked
		self.settings.FilelessFormat = self._cmbImageFormat.SelectedItem

		self.settings.RemoveEmptyDir = self._ckbRemoveEmptyDir.Checked
		self.settings.ExcludedEmptyDir = list(self._lbRemoveEmptyDir.Items)

		#Note for excludes. Have to remove the event handlers for the excluderules since they have to get added in the 
		#LoadSettings method.
		for i in self.settings.ExcludeRules:
			if not i == None:
				i.Remove.Click -= self.RemoveRuleSet

		self.settings.ExcludeFolders = list(self._lbExFolder.Items)

	def OkClick(self, sender, e):
		if not self.CheckFields():
			self.DialogResult = DialogResult.None
	
	def CheckFields(self):
		"""
		Make sure all the needed fields are filled in
		"""
		errors = ""
		if self._txbBaseDir.Text == "" and self._ckbDirectory.Checked:
			errors += "The base directory cannot be empty\n"
		
		if self._ckbFileNaming.Checked and self._txbFileStruct.Text.strip() == "":
			errors += "The File Structure cannot be empty"
		
		if not self._ckbDirectory.Checked and not self._ckbFileNaming.Checked:
			errors += "You must enable Directory or File naming functions for the script to actually do anything."

		if errors:
			errors = errors.Insert(0, "Please correct the following errors\n\n")
			MessageBox.Show(errors, "Please complete", MessageBoxButtons.OK, MessageBoxIcon.Error)
			return False
		else:
			return True
		

#The following 3 functions are for managing exclude rules
	
	def CreateRuleSet(self, sender, e):
		r = ExcludeRule()
		r.Remove.Click += self.RemoveRuleSet
		self.settings.ExcludeRules.append(r)
		self._flpExcludes.Controls.Add(self.settings.ExcludeRules[-1].Panel)
		self._ExPanel.ScrollControlIntoView(r.Panel)
	
	def CreateRuleGroup(self, sender, e):
		g = ExcludeGroup()
		g.Remove.Click += self.RemoveRuleSet
		g.CreateRule(None, None)
		self.settings.ExcludeRules.append(g)
		self._flpExcludes.Controls.Add(self.settings.ExcludeRules[-1].Panel)

	def RemoveRuleSet(self, sender, e):
		index = sender.Tag
		i = self.settings.ExcludeRules.index(index)
		self._flpExcludes.Controls.Remove(self.settings.ExcludeRules[i].Panel)
		del(self.settings.ExcludeRules[i])
		


	def CkbRemoveEmptyDirCheckedChanged(self, sender, e):
		self._lbRemoveEmptyDir.Enabled = sender.Checked
		self._btnAddEmptyDir.Enabled = sender.Checked
		self._btnRemoveEmptyDir.Enabled = sender.Checked
		
	def CkbFilelessCheckedChanged(self, sender, e):
		self._cmbImageFormat.Enabled = sender.Checked
		
	def BtnAddExFolderClick(self, sender, e):
		self._FolderBrowser.ShowDialog()
		self._lbExFolder.Items.Add(self._FolderBrowser.SelectedPath)
		
	def BtnAddEmptyDirClick(self, sender, e):
		self._FolderBrowser.ShowDialog()
		self._lbRemoveEmptyDir.Items.Add(self._FolderBrowser.SelectedPath)
	
	def BtnRemoveEmptyDirClick(self, sender, e):
		self._lbRemoveEmptyDir.Items.Remove(self._lbRemoveEmptyDir.SelectedItem)

	def BtnRemoveExFolderClick(self, sender, e):
		self._lbExFolder.Items.Remove(self._lbExFolder.SelectedItem)



	#Profiles features
	def BtnProSaveAsClick(self, sender, e):
		if self.CheckFields():	
			self.CreateNewSetting()
			self.SaveSettings()

	def BtnSaveClick(self, sender, e):
		self.SaveSettings()

	def BtnProNewClick(self, sender, e):
		if self.CheckFields():
			self.CreateNewSetting()
			self.LoadSettings()
				
	def CreateNewSetting(self):
		self.SaveSettings()
		ib = InputBox()
		ib.ShowDialog(self)
		i = ib.FindName()
		if i and ib.DialogResult == DialogResult.OK:
			self.allsettings[i] = losettings.settings()
			self.allsettings[i].Name = i
			self._cmbProfiles.Items.Add(i)
			self._cmbProfiles.SelectedItem = i
			self._cmbProfiles.Tag = self._cmbProfiles.Items.IndexOf(i)
			self.settings = self.allsettings[i]

	def BtnProDeleteClick(self, sender, e):
		if self._cmbProfiles.Items.Count > 1:
			i = self._cmbProfiles.SelectedItem
			n = self._cmbProfiles.SelectedIndex
			
			nn = n
			
			#We need to find the next item name or failing that the previous
			if n == self._cmbProfiles.Items.Count -1:
				n -= 1
			
			else: 
				n += 1
			#For the new index find the next index that exists. Note that once the one item is deleted, all the indexs will shift up
			#so try and keep the same index.
			if nn > self._cmbProfiles.Items.Count -2:
				nn -= 1

			self.settings = self.allsettings[self._cmbProfiles.Items[n]]
			self.LoadSettings()
			self._cmbProfiles.Items.Remove(i)
			#If the current selected index is the highest one the use a lower one.

			self._cmbProfiles.SelectedIndex = nn
			del(self.allsettings[i])
		
	def CmbProfilesItemChanged(self, sender, e):
		if self.CheckFields():
			self.SaveSettings()
			self.settings = self.allsettings[sender.SelectedItem]
			sender.Tag = sender.SelectedIndex
			self.LoadSettings()
			self.UpdateSampleText()
		else:
			sender.SelectedIndex = sender.Tag
			return



	def ModeChange(self, sender, e):
		self.settings.Mode = sender.Tag
	

	#For exclude rules combo boxes
	def ExcludeOperatorChange(self, sender, e):
		self.settings.ExcludeOperator = sender.SelectedItem
		
	def ExcludeModeChange(self, sender, e):
		self.settings.ExcludeMode =  sender.SelectedItem
		
	def FormShown(self, sender, e):
		#Don't tie the sample textbox textchanged events until the form is shown or there will be a nasty error in the case of 
		#one of the multiselect trying to invoke the form before it is fully created.
		self._txbDirStruct.TextChanged += self.TxbStructTextChanged	
		self._txbFileStruct.TextChanged += self.TxbStructTextChanged
	
		#This won't update the sample text since the textbox contents have already been loaded.
		self.UpdateSampleText()	



class InsertControl(Panel):
	"""
	Custom class to simplfy inserting template items
	"""	
	Insert, _insert = pyevent.make_event()

	def __init__(self):

		self.Template = ""

		self.Prefix = TextBox()
		self.InsertButton = Button()
		self.Postfix = TextBox()
		

		self.Prefix.Size = Size(58, 22)
		self.Prefix.Location = Point(0, 2)
		self.Prefix.TabIndex = 0

		self.InsertButton.Size = Size(66, 23)
		self.InsertButton.Location = Point(64, 0)
		self.InsertButton.Click += self.ButtonClick
		self.InsertButton.TabIndex = 1

		self.Postfix.Size = Size(58, 22)
		self.Postfix.Location = Point(136, 2)
		self.Postfix.TabIndex = 2

		self.Size = Size(194, 23)

		self.Controls.Add(self.Prefix)
		self.Controls.Add(self.Postfix)
		self.Controls.Add(self.InsertButton)

	def SetPrefixText(self, text):
		self.Prefix.Text = text

	def SetPostfixText(self, text):
		self.Postfix.Text = text

	def SetTemplate(self, template, text):
		"""
		sets the template field of the control. template is a string. Text sets the text of the button.
		"""
		self.Template = template
		self.InsertButton.Text = text

	def GetPrefixText(self):
		return self.Prefix.Text

	def GetPostfixText(self):
		return self.Postfix.Text

	def ButtonClick(self, sender, e):
		self._insert(self, None)


	def GetTemplateText(self, space):
		s = ""
		if space:
			s = " "
		return "{" + self.Prefix.Text + "<" + self.Template + ">" + self.Postfix.Text + s + "}"

class InsertControlPadding(InsertControl):
	
	def __init__(self):
		super(InsertControlPadding,self).__init__()
		self.Pad = NumericUpDown()
		self.Pad.Size = Size(34, 22)
		self.Pad.Location = Point(200, 2)
		self.Pad.TabIndex = 3

		self.Width = 234
		self.Controls.Add(self.Pad)
		

	def GetTemplateText(self, space):
		s = ""
		if space:
			s = " "
		return "{" + self.Prefix.Text + "<" + self.Template + str(self.Pad.Value) + ">" + self.Postfix.Text + s + "}"

class InsertControlCheckBox(InsertControl):
	
	def __init__(self):
		super(InsertControlCheckBox,self).__init__()
		self.Seperator = TextBox()
		self.Seperator.Size = Size(20, 22)
		self.Seperator.Location = Point(200, 2)
		self.Seperator.TabIndex = 3

		self.Check = CheckBox()
		self.Check.Size = Size(15, 14)
		self.Check.Location = Point(226, 4)
		self.Check.TabIndex = 4



		self.Width = 241
		self.Controls.Add(self.Seperator)
		self.Controls.Add(self.Check)

	def GetSeperatorText(self):
		return self.Seperator.Text

	def SetSeperatorText(self, text):
		self.Seperator.Text = text		

	def GetTemplateText(self, space):
		if self.Check.Checked:
			checktext = "(series)"
		else:
			checktext = "(issue)"
		sep = self.Seperator.Text
		s = ""
		if space:
			s = " "
			sep = sep.rstrip() + s
		return "{" + self.Prefix.Text + "<" + self.Template + "(" + sep + ")" + checktext + ">" + self.Postfix.Text + s + "}"

class InsertControlTextBox(InsertControl):
	
	def __init__(self):
		super(InsertControlTextBox,self).__init__()
		self.TextBox = TextBox()
		self.TextBox.Size = Size(58, 22)
		self.TextBox.Location = Point(200, 2)
		self.TextBox.TabIndex = 3

		self.Width = 258
		self.Controls.Add(self.TextBox)
		
	def SetTextBoxText(self, text):
		self.TextBox.Text = text

	def GetTextBoxText(self):
		return self.TextBox.Text

	def GetTemplateText(self, space):
		s = ""
		if space:
			s = " "
		return "{" + self.Prefix.Text + "<" + self.Template + "(" + self.TextBox.Text + ")>" + self.Postfix.Text + s + "}"