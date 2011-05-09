"""
loconfigform.py

Version: 1.1

Contains the config form. Most functions are related to makeing the GUI work. Several functions are related to settings.

Some fuctions are a bit of a mess and could be consolidated/simplified/cleaned up

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

from locommon import PathMaker, ICON, ExcludePath, ExcludeRule, ExcludeGroup, Mode, ExcludeMeta

import pyevent

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
		self.PathCreator = PathMaker()
		
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
		self._tpDirOrganize = System.Windows.Forms.TabPage()
		self._tpFileNames = System.Windows.Forms.TabPage()
		self._txbDirStruct = System.Windows.Forms.TextBox()
		self._lblDirStruct = System.Windows.Forms.Label()
		self._txbBaseDir = System.Windows.Forms.TextBox()
		self._lblBaseDir = System.Windows.Forms.Label()
		self._btnBrowse = System.Windows.Forms.Button()
		self._ckbDirectory = System.Windows.Forms.CheckBox()
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._btnSep = System.Windows.Forms.Button()
		self._ckbFileNaming = System.Windows.Forms.CheckBox()
		self._txbFileStruct = System.Windows.Forms.TextBox()
		self._lblFileStruct = System.Windows.Forms.Label()
		self._btnMonthText = System.Windows.Forms.Button()
		self.Publisher = InsertControl()
		self.Imprint = InsertControl()
		self.Series = InsertControl()
		self.Title = InsertControl()
		self.Format = InsertControl()
		self.Volume = InsertControlPadding()
		self.Count = InsertControl()
		self.Number = InsertControlPadding()
		self.Month = InsertControl()
		self.MonthNumber = InsertControlPadding()
		self.AlternateSeries = InsertControl()
		self.AlternateNumber = InsertControlPadding()
		self.AlternateCount = InsertControlPadding()
		self.Year = InsertControl()
		self._label1 = System.Windows.Forms.Label()
		self._label2 = System.Windows.Forms.Label()
		self._label3 = System.Windows.Forms.Label()
		self._label4 = System.Windows.Forms.Label()
		self._label5 = System.Windows.Forms.Label()
		self._label6 = System.Windows.Forms.Label()
		self._FolderBrowser = System.Windows.Forms.FolderBrowserDialog()
		self._ckbSpace = System.Windows.Forms.CheckBox()
		self._label12 = System.Windows.Forms.Label()
		self._sampleTextDir = System.Windows.Forms.Label()
		self._sampleTextFile = System.Windows.Forms.Label()
		self._label14 = System.Windows.Forms.Label()
		self._ok = System.Windows.Forms.Button()
		self._cancel = System.Windows.Forms.Button()
		self._tpExcludes = System.Windows.Forms.TabPage()
		self._label13 = System.Windows.Forms.Label()
		self._lbExFolder = System.Windows.Forms.ListBox()
		self._btnAddExFolder = System.Windows.Forms.Button()
		self._btnRemoveExFolder = System.Windows.Forms.Button()
		self._ExPanel = System.Windows.Forms.Panel()
		self._cmbExcludeOperator = System.Windows.Forms.ComboBox()
		self._label17 = System.Windows.Forms.Label()
		self._btnExMetaAdd = System.Windows.Forms.Button()
		self._flpExcludes = System.Windows.Forms.FlowLayoutPanel()
		self._groupBox2 = System.Windows.Forms.GroupBox()
		self._button1 = System.Windows.Forms.Button()
		self._btnProNew = System.Windows.Forms.Button()
		self._btnProSaveAs = System.Windows.Forms.Button()
		self._btnProDelete = System.Windows.Forms.Button()
		self._cmbProfiles = System.Windows.Forms.ComboBox()
		self._vsbBookSelector = System.Windows.Forms.VScrollBar()
		self._label15 = System.Windows.Forms.Label()
		self._tabPage1 = System.Windows.Forms.TabPage()
		self._txbEmptyData = System.Windows.Forms.TextBox()
		self._label11 = System.Windows.Forms.Label()
		self._label10 = System.Windows.Forms.Label()
		self._cmbEmptyData = System.Windows.Forms.ComboBox()
		self._label9 = System.Windows.Forms.Label()
		self._label8 = System.Windows.Forms.Label()
		self._label7 = System.Windows.Forms.Label()
		self._txbEmptyDir = System.Windows.Forms.TextBox()
		self._ckbRemoveEmptyDir = System.Windows.Forms.CheckBox()
		self._lbRemoveEmptyDir = System.Windows.Forms.ListBox()
		self._label18 = System.Windows.Forms.Label()
		self._ckbFileless = System.Windows.Forms.CheckBox()
		self._btnAddEmptyDir = System.Windows.Forms.Button()
		self._btnRemoveEmptyDir = System.Windows.Forms.Button()
		self._cmbImageFormat = System.Windows.Forms.ComboBox()
		self._label19 = System.Windows.Forms.Label()
		self._btnAddGroup = System.Windows.Forms.Button()
		self._label20 = System.Windows.Forms.Label()
		self._cmbExcludeMode = System.Windows.Forms.ComboBox()
		self._gbMode = System.Windows.Forms.GroupBox()
		self._rdbModeMove = System.Windows.Forms.RadioButton()
		self._rdbMoveCopy = System.Windows.Forms.RadioButton()
		self._rdbModeTest = System.Windows.Forms.RadioButton()
		self._insertTabs = System.Windows.Forms.TabControl()
		self._tabPage2 = System.Windows.Forms.TabPage()
		self._tpInsertAdvanced = System.Windows.Forms.TabPage()		
		self._label16 = System.Windows.Forms.Label()
		self._label21 = System.Windows.Forms.Label()
		self.StartYear = InsertControl()
		self.Manga = InsertControlTextBox()
		self.Character = InsertControlCheckBox()
		self.Writer = InsertControlCheckBox()
		self.Genre = InsertControlCheckBox()
		self.Tags = InsertControlCheckBox()
		self._label22 = System.Windows.Forms.Label()
		self._tabs.SuspendLayout()
		self._tpDirOrganize.SuspendLayout()
		self._tpFileNames.SuspendLayout()
		self._groupBox1.SuspendLayout()
		self._tpExcludes.SuspendLayout()
		self._ExPanel.SuspendLayout()
		self._groupBox2.SuspendLayout()
		self._tabPage1.SuspendLayout()
		self._gbMode.SuspendLayout()
		self._insertTabs.SuspendLayout()
		self._tabPage2.SuspendLayout()
		self._tpInsertAdvanced.SuspendLayout()
		self.SuspendLayout()
		# 
		# tabs
		# 
		self._tabs.Controls.Add(self._tpDirOrganize)
		self._tabs.Controls.Add(self._tpFileNames)
		self._tabs.Controls.Add(self._tpExcludes)
		self._tabs.Controls.Add(self._tabPage1)
		self._tabs.Location = System.Drawing.Point(10, 6)
		self._tabs.Name = "tabs"
		self._tabs.SelectedIndex = 0
		self._tabs.Size = System.Drawing.Size(526, 452)
		self._tabs.TabIndex = 0
		self._tabs.SelectedIndexChanged += self.TabsSelectedIndexChanged
		# 
		# tpDirOrganize
		# 
		self._tpDirOrganize.Controls.Add(self._vsbBookSelector)
		self._tpDirOrganize.Controls.Add(self._sampleTextDir)
		self._tpDirOrganize.Controls.Add(self._label12)
		self._tpDirOrganize.Controls.Add(self._ckbDirectory)
		self._tpDirOrganize.Controls.Add(self._btnBrowse)
		self._tpDirOrganize.Controls.Add(self._lblBaseDir)
		self._tpDirOrganize.Controls.Add(self._txbBaseDir)
		self._tpDirOrganize.Controls.Add(self._lblDirStruct)
		self._tpDirOrganize.Controls.Add(self._txbDirStruct)
		self._tpDirOrganize.Controls.Add(self._groupBox1)
		self._tpDirOrganize.Location = System.Drawing.Point(4, 22)
		self._tpDirOrganize.Name = "tpDirOrganize"
		self._tpDirOrganize.Padding = System.Windows.Forms.Padding(3)
		self._tpDirOrganize.Size = System.Drawing.Size(518, 426)
		self._tpDirOrganize.TabIndex = 0
		self._tpDirOrganize.Text = "Directories"
		self._tpDirOrganize.UseVisualStyleBackColor = True
		# 
		# tpFileNames
		# 
		self._tpFileNames.Controls.Add(self._sampleTextFile)
		self._tpFileNames.Controls.Add(self._label14)
		self._tpFileNames.Controls.Add(self._lblFileStruct)
		self._tpFileNames.Controls.Add(self._txbFileStruct)
		self._tpFileNames.Controls.Add(self._ckbFileNaming)
		self._tpFileNames.Location = System.Drawing.Point(4, 22)
		self._tpFileNames.Name = "tpFileNames"
		self._tpFileNames.Padding = System.Windows.Forms.Padding(3)
		self._tpFileNames.Size = System.Drawing.Size(518, 426)
		self._tpFileNames.TabIndex = 1
		self._tpFileNames.Text = "File Names"
		self._tpFileNames.UseVisualStyleBackColor = True
		# 
		# txbDirStruct
		# 
		self._txbDirStruct.Location = System.Drawing.Point(113, 71)
		self._txbDirStruct.Name = "txbDirStruct"
		self._txbDirStruct.Size = System.Drawing.Size(397, 20)
		self._txbDirStruct.TabIndex = 5
		self._txbDirStruct.TextChanged += self.TxbStructTextChanged
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
		# groupBox1
		# 
		self._groupBox1.Controls.Add(self._insertTabs)
		self._groupBox1.Controls.Add(self._ckbSpace)
		self._groupBox1.Controls.Add(self._btnSep)
		self._groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom
		self._groupBox1.Location = System.Drawing.Point(3, 121)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(512, 302)
		self._groupBox1.TabIndex = 8
		self._groupBox1.TabStop = False
		self._groupBox1.Text = "Metadata"
		# 
		# btnSep
		# 
		self._btnSep.Location = System.Drawing.Point(53, 19)
		self._btnSep.Name = "btnSep"
		self._btnSep.Size = System.Drawing.Size(115, 23)
		self._btnSep.TabIndex = 0
		self._btnSep.Text = "Directory Seperator"
		self._btnSep.UseVisualStyleBackColor = True
		self._btnSep.Click += self.BtnSepClick
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
		self._txbFileStruct.TextChanged += self.TxbStructTextChanged
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
		# Publisher
		#
		self.Publisher.Location = Point(6, 25)
		self.Publisher.SetTemplate("publisher", "Publisher")
		# 
		# Imprint
		# 
		self.Imprint.Location = Point(6, 53)
		self.Imprint.SetTemplate("imprint", "Imprint")
		# 
		# Series
		# 
		self.Series.Location = Point(6, 82)
		self.Series.SetTemplate("series", "Series")
		# 
		# Title
		# 
		self.Title.Location = Point(6, 110)
		self.Title.SetTemplate("title", "Title")
		# 
		# Format
		#
		self.Format.Location = Point(6, 138)
		self.Format.SetTemplate("format", "Format")
		# 
		# Volume
		#
		self.Volume.Location = Point(6, 166)
		self.Volume.SetTemplate ("volume", "Volume")
		# 
		# AlternateSeries
		# 
		self.AlternateSeries.Location = Point(6, 194)
		self.AlternateSeries.SetTemplate("altSeries", "Alt. Series")
		# 
		# AlternateNumber
		# 
		self.AlternateNumber.Location = Point(262, 194)
		self.AlternateNumber.SetTemplate("altNumber", "Alt. Num.")
		# 
		# AlternateCount
		# 
		self.AlternateCount.Location = Point(262, 166)
		self.AlternateCount.SetTemplate("altCount", "Alt. Count")
		# 
		# Year
		# 
		self.Year.Location = Point(262, 138)
		self.Year.SetTemplate("year", "Year")
		# 
		# Month
		# 
		self.Month.Location = Point(262, 108)
		self.Month.SetTemplate("month", "Month")
		# 
		# Month Number
		# 
		self.MonthNumber.Location = Point(262, 82)
		self.MonthNumber.SetTemplate("month#", "Month #")
		# 
		# Count
		# 
		self.Count.Location = Point(262, 54)
		self.Count.SetTemplate("count", "Count")
		# 
		# Number
		# 
		self.Number.Location = Point(262, 26)
		self.Number.SetTemplate("number", "Number")
		# 
		# label1
		# 
		self._label1.AutoSize = True
		self._label1.Location = System.Drawing.Point(21, 3)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(33, 13)
		self._label1.TabIndex = 2
		self._label1.Text = "Prefix"
		# 
		# label2
		# 
		self._label2.AutoSize = True
		self._label2.Location = System.Drawing.Point(148, 3)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(38, 13)
		self._label2.TabIndex = 3
		self._label2.Text = "Postfix"
		# 
		# label3
		# 
		self._label3.AutoSize = True
		self._label3.Location = System.Drawing.Point(404, 3)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(38, 13)
		self._label3.TabIndex = 27
		self._label3.Text = "Postfix"
		# 
		# label4
		# 
		self._label4.AutoSize = True
		self._label4.Location = System.Drawing.Point(275, 3)
		self._label4.Name = "label4"
		self._label4.Size = System.Drawing.Size(33, 13)
		self._label4.TabIndex = 26
		self._label4.Text = "Prefix"
		# 
		# label5
		# 
		self._label5.AutoSize = True
		self._label5.Location = System.Drawing.Point(210, 3)
		self._label5.Name = "label5"
		self._label5.Size = System.Drawing.Size(26, 13)
		self._label5.TabIndex = 4
		self._label5.Text = "Pad"
		# 
		# label6
		# 
		self._label6.AutoSize = True
		self._label6.Location = System.Drawing.Point(466, 3)
		self._label6.Name = "label6"
		self._label6.Size = System.Drawing.Size(26, 13)
		self._label6.TabIndex = 28
		self._label6.Text = "Pad"
		# 
		# FolderBrowser
		# 
		self._FolderBrowser.Description = "Selecte the base folder for your comic library"
		# 
		# ckbSpace
		# 
		self._ckbSpace.AutoSize = True
		self._ckbSpace.Location = System.Drawing.Point(264, 22)
		self._ckbSpace.Name = "ckbSpace"
		self._ckbSpace.Size = System.Drawing.Size(188, 17)
		self._ckbSpace.TabIndex = 1
		self._ckbSpace.Text = "Space inserted fields automatically"
		self._ckbSpace.UseVisualStyleBackColor = True
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
		# ok
		# 
		self._ok.DialogResult = System.Windows.Forms.DialogResult.OK
		self._ok.Location = System.Drawing.Point(380, 484)
		self._ok.Name = "ok"
		self._ok.Size = System.Drawing.Size(75, 23)
		self._ok.TabIndex = 1
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
		self._cancel.TabIndex = 2
		self._cancel.Text = "Cancel"
		self._cancel.UseVisualStyleBackColor = True
		# 
		# tpExcludes
		# 
		self._tpExcludes.Controls.Add(self._groupBox2)
		self._tpExcludes.Controls.Add(self._btnRemoveExFolder)
		self._tpExcludes.Controls.Add(self._btnAddExFolder)
		self._tpExcludes.Controls.Add(self._lbExFolder)
		self._tpExcludes.Controls.Add(self._label13)
		self._tpExcludes.Location = System.Drawing.Point(4, 22)
		self._tpExcludes.Name = "tpExcludes"
		self._tpExcludes.Padding = System.Windows.Forms.Padding(3)
		self._tpExcludes.Size = System.Drawing.Size(518, 426)
		self._tpExcludes.TabIndex = 3
		self._tpExcludes.Text = "Excludes"
		self._tpExcludes.UseVisualStyleBackColor = True
		# 
		# label13
		# 
		self._label13.Location = System.Drawing.Point(8, 16)
		self._label13.Name = "label13"
		self._label13.Size = System.Drawing.Size(502, 23)
		self._label13.TabIndex = 0
		self._label13.Text = "Do not move eComics if they are located in any of these folders:"
		# 
		# lbExFolder
		# 
		self._lbExFolder.FormattingEnabled = True
		self._lbExFolder.HorizontalScrollbar = True
		self._lbExFolder.Location = System.Drawing.Point(8, 42)
		self._lbExFolder.Name = "lbExFolder"
		self._lbExFolder.Size = System.Drawing.Size(391, 108)
		self._lbExFolder.Sorted = True
		self._lbExFolder.TabIndex = 1
		# 
		# btnAddExFolder
		# 
		self._btnAddExFolder.Location = System.Drawing.Point(421, 53)
		self._btnAddExFolder.Name = "btnAddExFolder"
		self._btnAddExFolder.Size = System.Drawing.Size(75, 23)
		self._btnAddExFolder.TabIndex = 2
		self._btnAddExFolder.Text = "Add"
		self._btnAddExFolder.UseVisualStyleBackColor = True
		self._btnAddExFolder.Click += self.BtnAddExFolderClick
		# 
		# btnRemoveExFolder
		# 
		self._btnRemoveExFolder.Location = System.Drawing.Point(421, 106)
		self._btnRemoveExFolder.Name = "btnRemoveExFolder"
		self._btnRemoveExFolder.Size = System.Drawing.Size(75, 23)
		self._btnRemoveExFolder.TabIndex = 3
		self._btnRemoveExFolder.Text = "Remove"
		self._btnRemoveExFolder.UseVisualStyleBackColor = True
		self._btnRemoveExFolder.Click += self.BtnRemoveExFolderClick
		# 
		# ExPanel
		# 
		self._ExPanel.AutoScroll = True
		self._ExPanel.Controls.Add(self._flpExcludes)
		self._ExPanel.Location = System.Drawing.Point(3, 46)
		self._ExPanel.Name = "ExPanel"
		self._ExPanel.Size = System.Drawing.Size(495, 207)
		self._ExPanel.TabIndex = 4
		# 
		# cmbExcludeOperator
		# 
		self._cmbExcludeOperator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbExcludeOperator.FormattingEnabled = True
		self._cmbExcludeOperator.Items.AddRange(System.Array[System.Object](
			["Any",
			"All"]))
		self._cmbExcludeOperator.Location = System.Drawing.Point(204, 19)
		self._cmbExcludeOperator.Name = "cmbExcludeOperator"
		self._cmbExcludeOperator.Size = System.Drawing.Size(43, 21)
		self._cmbExcludeOperator.TabIndex = 0
		self._cmbExcludeOperator.SelectedIndexChanged += self.ExcludeOperatorChange
		# 
		# label17
		# 
		self._label17.AutoSize = True
		self._label17.Location = System.Drawing.Point(250, 22)
		self._label17.Name = "label17"
		self._label17.Size = System.Drawing.Size(106, 13)
		self._label17.TabIndex = 2
		self._label17.Text = "of the following rules."
		# 
		# btnExMetaAdd
		# 
		self._btnExMetaAdd.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._btnExMetaAdd.Location = System.Drawing.Point(439, 17)
		self._btnExMetaAdd.Name = "btnExMetaAdd"
		self._btnExMetaAdd.Size = System.Drawing.Size(59, 23)
		self._btnExMetaAdd.TabIndex = 3
		self._btnExMetaAdd.Text = "Add rule"
		self._btnExMetaAdd.UseVisualStyleBackColor = True
		self._btnExMetaAdd.Click += self.CreateRuleSet
		# 
		# flpExcludes
		# 
		self._flpExcludes.AutoSize = True
		self._flpExcludes.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
		self._flpExcludes.Location = System.Drawing.Point(3, 3)
		self._flpExcludes.Name = "flpExcludes"
		self._flpExcludes.Size = System.Drawing.Size(482, 40)
		self._flpExcludes.TabIndex = 8
		# 
		# groupBox2
		# 
		self._groupBox2.Controls.Add(self._cmbExcludeMode)
		self._groupBox2.Controls.Add(self._label20)
		self._groupBox2.Controls.Add(self._btnAddGroup)
		self._groupBox2.Controls.Add(self._btnExMetaAdd)
		self._groupBox2.Controls.Add(self._ExPanel)
		self._groupBox2.Controls.Add(self._label17)
		self._groupBox2.Controls.Add(self._cmbExcludeOperator)
		self._groupBox2.Location = System.Drawing.Point(8, 156)
		self._groupBox2.Name = "groupBox2"
		self._groupBox2.Size = System.Drawing.Size(504, 264)
		self._groupBox2.TabIndex = 6
		self._groupBox2.TabStop = False
		self._groupBox2.Text = "Metadata Rules"
		# 
		# button1
		# 
		self._button1.Location = System.Drawing.Point(10, 491)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(75, 23)
		self._button1.TabIndex = 0
		self._button1.Text = "Save"
		self._button1.UseVisualStyleBackColor = True
		# 
		# btnProNew
		# 
		self._btnProNew.Location = System.Drawing.Point(172, 464)
		self._btnProNew.Name = "btnProNew"
		self._btnProNew.Size = System.Drawing.Size(75, 23)
		self._btnProNew.TabIndex = 1
		self._btnProNew.Text = "New"
		self._btnProNew.UseVisualStyleBackColor = True
		self._btnProNew.Click += self.BtnProNewClick
		# 
		# btnProSaveAs
		# 
		self._btnProSaveAs.Location = System.Drawing.Point(91, 491)
		self._btnProSaveAs.Name = "btnProSaveAs"
		self._btnProSaveAs.Size = System.Drawing.Size(75, 23)
		self._btnProSaveAs.TabIndex = 2
		self._btnProSaveAs.Text = "Save As"
		self._btnProSaveAs.UseVisualStyleBackColor = True
		self._btnProSaveAs.Click += self.BtnProSaveAsClick
		# 
		# btnProDelete
		# 
		self._btnProDelete.Location = System.Drawing.Point(172, 491)
		self._btnProDelete.Name = "btnProDelete"
		self._btnProDelete.Size = System.Drawing.Size(75, 23)
		self._btnProDelete.TabIndex = 3
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
		self._cmbProfiles.TabIndex = 4
		self._cmbProfiles.SelectionChangeCommitted += self.CmbProfilesItemChanged
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
		# Start Year
		# 
		self.StartYear.Location = Point(6, 23)
		self.StartYear.SetTemplate("startyear", "Start Year")
		# 
		# label15
		# 
		self._label15.Location = System.Drawing.Point(6, 48)
		self._label15.Name = "label15"
		self._label15.Size = System.Drawing.Size(205, 30)
		self._label15.TabIndex = 58
		self._label15.Text = "Start year is calculated from the earliest year for the series in your library"
		# 
		# tabPage1
		# 
		self._tabPage1.Controls.Add(self._gbMode)
		self._tabPage1.Controls.Add(self._cmbImageFormat)
		self._tabPage1.Controls.Add(self._label19)
		self._tabPage1.Controls.Add(self._btnRemoveEmptyDir)
		self._tabPage1.Controls.Add(self._btnAddEmptyDir)
		self._tabPage1.Controls.Add(self._ckbFileless)
		self._tabPage1.Controls.Add(self._label18)
		self._tabPage1.Controls.Add(self._lbRemoveEmptyDir)
		self._tabPage1.Controls.Add(self._ckbRemoveEmptyDir)
		self._tabPage1.Controls.Add(self._txbEmptyData)
		self._tabPage1.Controls.Add(self._label11)
		self._tabPage1.Controls.Add(self._label10)
		self._tabPage1.Controls.Add(self._cmbEmptyData)
		self._tabPage1.Controls.Add(self._label9)
		self._tabPage1.Controls.Add(self._label8)
		self._tabPage1.Controls.Add(self._label7)
		self._tabPage1.Controls.Add(self._txbEmptyDir)
		self._tabPage1.Location = System.Drawing.Point(4, 22)
		self._tabPage1.Name = "tabPage1"
		self._tabPage1.Padding = System.Windows.Forms.Padding(3)
		self._tabPage1.Size = System.Drawing.Size(518, 426)
		self._tabPage1.TabIndex = 4
		self._tabPage1.Text = "Options"
		self._tabPage1.UseVisualStyleBackColor = True
		# 
		# txbEmptyData
		# 
		self._txbEmptyData.Location = System.Drawing.Point(275, 149)
		self._txbEmptyData.Name = "txbEmptyData"
		self._txbEmptyData.Size = System.Drawing.Size(235, 20)
		self._txbEmptyData.TabIndex = 22
		self._txbEmptyData.Leave += self.TxbEmptyDataLeave
		# 
		# label11
		# 
		self._label11.AutoSize = True
		self._label11.Location = System.Drawing.Point(195, 152)
		self._label11.Name = "label11"
		self._label11.Size = System.Drawing.Size(74, 13)
		self._label11.TabIndex = 21
		self._label11.Text = "Default Value:"
		# 
		# label10
		# 
		self._label10.AutoSize = True
		self._label10.Location = System.Drawing.Point(8, 152)
		self._label10.Name = "label10"
		self._label10.Size = System.Drawing.Size(55, 13)
		self._label10.TabIndex = 20
		self._label10.Text = "Data field:"
		# 
		# cmbEmptyData
		# 
		self._cmbEmptyData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbEmptyData.FormattingEnabled = True
		self._cmbEmptyData.Items.AddRange(System.Array[System.Object](
			["Alternate Count",
			"Alternate Number",
			"Alternate Series",
			"Count",
			"Format",
			"Imprint",
			"Month",
			"Number",
			"Publisher",
			"Series",
			"Start Year",
			"Title",
			"Volume",
			"Year"]))
		self._cmbEmptyData.Location = System.Drawing.Point(68, 149)
		self._cmbEmptyData.Name = "cmbEmptyData"
		self._cmbEmptyData.Size = System.Drawing.Size(121, 21)
		self._cmbEmptyData.Sorted = True
		self._cmbEmptyData.TabIndex = 19
		self._cmbEmptyData.SelectedIndexChanged += self.CmbEmptyDataSelectedIndexChanged
		# 
		# label9
		# 
		self._label9.AutoSize = True
		self._label9.Location = System.Drawing.Point(8, 124)
		self._label9.Name = "label9"
		self._label9.Size = System.Drawing.Size(441, 13)
		self._label9.TabIndex = 15
		self._label9.Text = "To create a default value to be entered in the case of a blank data field, fill in the box below."
		# 
		# label8
		# 
		self._label8.AutoSize = True
		self._label8.Location = System.Drawing.Point(8, 94)
		self._label8.Name = "label8"
		self._label8.Size = System.Drawing.Size(204, 13)
		self._label8.TabIndex = 18
		self._label8.Text = "(Leave empty to remove blank directories)"
		# 
		# label7
		# 
		self._label7.AutoSize = True
		self._label7.Location = System.Drawing.Point(8, 74)
		self._label7.Name = "label7"
		self._label7.Size = System.Drawing.Size(181, 13)
		self._label7.TabIndex = 17
		self._label7.Text = "Replace blank directory names with: "
		self._label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
		# 
		# txbEmptyDir
		# 
		self._txbEmptyDir.Location = System.Drawing.Point(226, 79)
		self._txbEmptyDir.Name = "txbEmptyDir"
		self._txbEmptyDir.Size = System.Drawing.Size(284, 20)
		self._txbEmptyDir.TabIndex = 16
		# 
		# ckbRemoveEmptyDir
		# 
		self._ckbRemoveEmptyDir.Checked = True
		self._ckbRemoveEmptyDir.CheckState = System.Windows.Forms.CheckState.Checked
		self._ckbRemoveEmptyDir.Location = System.Drawing.Point(8, 177)
		self._ckbRemoveEmptyDir.Name = "ckbRemoveEmptyDir"
		self._ckbRemoveEmptyDir.Size = System.Drawing.Size(261, 24)
		self._ckbRemoveEmptyDir.TabIndex = 23
		self._ckbRemoveEmptyDir.Text = "Remove empty directories"
		self._ckbRemoveEmptyDir.UseVisualStyleBackColor = True
		self._ckbRemoveEmptyDir.CheckedChanged += self.CkbRemoveEmptyDirCheckedChanged
		# 
		# lbRemoveEmptyDir
		# 
		self._lbRemoveEmptyDir.FormattingEnabled = True
		self._lbRemoveEmptyDir.Location = System.Drawing.Point(52, 229)
		self._lbRemoveEmptyDir.Name = "lbRemoveEmptyDir"
		self._lbRemoveEmptyDir.Size = System.Drawing.Size(383, 108)
		self._lbRemoveEmptyDir.TabIndex = 24
		# 
		# label18
		# 
		self._label18.Location = System.Drawing.Point(39, 203)
		self._label18.Name = "label18"
		self._label18.Size = System.Drawing.Size(308, 23)
		self._label18.TabIndex = 25
		self._label18.Text = "...but never remove the following directories:"
		# 
		# ckbFileless
		# 
		self._ckbFileless.Location = System.Drawing.Point(8, 355)
		self._ckbFileless.Name = "ckbFileless"
		self._ckbFileless.Size = System.Drawing.Size(502, 41)
		self._ckbFileless.TabIndex = 26
		self._ckbFileless.Text = "Copy fileless comic's custom thumbnail image to the calaculated path. (Does not affect the source image at all)"
		self._ckbFileless.UseVisualStyleBackColor = True
		self._ckbFileless.CheckedChanged += self.CkbFilelessCheckedChanged
		# 
		# btnAddEmptyDir
		# 
		self._btnAddEmptyDir.Location = System.Drawing.Point(441, 246)
		self._btnAddEmptyDir.Name = "btnAddEmptyDir"
		self._btnAddEmptyDir.Size = System.Drawing.Size(68, 23)
		self._btnAddEmptyDir.TabIndex = 27
		self._btnAddEmptyDir.Text = "Add"
		self._btnAddEmptyDir.UseVisualStyleBackColor = True
		self._btnAddEmptyDir.Click += self.BtnAddEmptyDirClick
		# 
		# btnRemoveEmptyDir
		# 
		self._btnRemoveEmptyDir.Location = System.Drawing.Point(441, 296)
		self._btnRemoveEmptyDir.Name = "btnRemoveEmptyDir"
		self._btnRemoveEmptyDir.Size = System.Drawing.Size(69, 23)
		self._btnRemoveEmptyDir.TabIndex = 28
		self._btnRemoveEmptyDir.Text = "Remove"
		self._btnRemoveEmptyDir.UseVisualStyleBackColor = True
		self._btnRemoveEmptyDir.Click += self.BtnRemoveEmptyDirClick
		# 
		# cmbImageFormat
		# 
		self._cmbImageFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbImageFormat.Enabled = False
		self._cmbImageFormat.FormattingEnabled = True
		self._cmbImageFormat.Items.AddRange(System.Array[System.Object](
			[".bmp",
			".jpg",
			".png"]))
		self._cmbImageFormat.Location = System.Drawing.Point(179, 390)
		self._cmbImageFormat.Name = "cmbImageFormat"
		self._cmbImageFormat.Size = System.Drawing.Size(54, 21)
		self._cmbImageFormat.TabIndex = 29
		# 
		# label19
		# 
		self._label19.Location = System.Drawing.Point(86, 393)
		self._label19.Name = "label19"
		self._label19.Size = System.Drawing.Size(87, 21)
		self._label19.TabIndex = 30
		self._label19.Text = "Save image as:"
		# 
		# btnAddGroup
		# 
		self._btnAddGroup.Location = System.Drawing.Point(362, 17)
		self._btnAddGroup.Name = "btnAddGroup"
		self._btnAddGroup.Size = System.Drawing.Size(71, 23)
		self._btnAddGroup.TabIndex = 5
		self._btnAddGroup.Text = "Add Group"
		self._btnAddGroup.UseVisualStyleBackColor = True
		self._btnAddGroup.Click += self.CreateRuleGroup
		# 
		# label20
		# 
		self._label20.AutoSize = True
		self._label20.Location = System.Drawing.Point(69, 22)
		self._label20.Name = "label20"
		self._label20.Size = System.Drawing.Size(129, 13)
		self._label20.TabIndex = 6
		self._label20.Text = "move eComics that match"
		# 
		# cmbExcludeMode
		# 
		self._cmbExcludeMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbExcludeMode.FormattingEnabled = True
		self._cmbExcludeMode.Items.AddRange(System.Array[System.Object](
			["Do not",
			"Only"]))
		self._cmbExcludeMode.Location = System.Drawing.Point(6, 19)
		self._cmbExcludeMode.Name = "cmbExcludeMode"
		self._cmbExcludeMode.Size = System.Drawing.Size(57, 21)
		self._cmbExcludeMode.TabIndex = 7
		self._cmbExcludeMode.SelectedIndexChanged += self.ExcludeModeChange
		# 
		# gbMode
		# 
		self._gbMode.Controls.Add(self._rdbModeTest)
		self._gbMode.Controls.Add(self._rdbMoveCopy)
		self._gbMode.Controls.Add(self._rdbModeMove)
		self._gbMode.Location = System.Drawing.Point(8, 6)
		self._gbMode.Name = "gbMode"
		self._gbMode.Size = System.Drawing.Size(497, 49)
		self._gbMode.TabIndex = 31
		self._gbMode.TabStop = False
		self._gbMode.Text = "Mode"
		# 
		# rdbModeMove
		# 
		self._rdbModeMove.AutoSize = True
		self._rdbModeMove.Location = System.Drawing.Point(6, 19)
		self._rdbModeMove.Name = "rdbModeMove"
		self._rdbModeMove.Size = System.Drawing.Size(52, 17)
		self._rdbModeMove.TabIndex = 0
		self._rdbModeMove.TabStop = True
		self._rdbModeMove.Tag = "Move"
		self._rdbModeMove.Text = "Move"
		self._rdbModeMove.UseVisualStyleBackColor = True
		self._rdbModeMove.CheckedChanged += self.ModeChange
		# 
		# rdbMoveCopy
		# 
		self._rdbMoveCopy.AutoSize = True
		self._rdbMoveCopy.Location = System.Drawing.Point(124, 19)
		self._rdbMoveCopy.Name = "rdbMoveCopy"
		self._rdbMoveCopy.Size = System.Drawing.Size(49, 17)
		self._rdbMoveCopy.TabIndex = 1
		self._rdbMoveCopy.TabStop = True
		self._rdbMoveCopy.Tag = "Copy"
		self._rdbMoveCopy.Text = "Copy"
		self._rdbMoveCopy.UseVisualStyleBackColor = True
		self._rdbMoveCopy.CheckedChanged += self.ModeChange
		# 
		# rdbModeTest
		# 
		self._rdbModeTest.AutoSize = True
		self._rdbModeTest.Location = System.Drawing.Point(239, 19)
		self._rdbModeTest.Name = "rdbModeTest"
		self._rdbModeTest.Size = System.Drawing.Size(251, 17)
		self._rdbModeTest.TabIndex = 2
		self._rdbModeTest.TabStop = True
		self._rdbModeTest.Tag = "Test"
		self._rdbModeTest.Text = "Test (no files touched, complete log file created)"
		self._rdbModeTest.UseVisualStyleBackColor = True
		self._rdbModeTest.CheckedChanged += self.ModeChange
		# 
		# insertTabs
		# 
		self._insertTabs.Controls.Add(self._tabPage2)
		self._insertTabs.Controls.Add(self._tpInsertAdvanced)
		self._insertTabs.Location = System.Drawing.Point(3, 45)
		self._insertTabs.Name = "insertTabs"
		self._insertTabs.SelectedIndex = 0
		self._insertTabs.Size = System.Drawing.Size(506, 254)
		self._insertTabs.TabIndex = 59
		# 
		# tabPage2
		# 
		self._tabPage2.Controls.Add(self._label2)
		self._tabPage2.Controls.Add(self.Publisher)
		self._tabPage2.Controls.Add(self.Imprint)
		self._tabPage2.Controls.Add(self.Volume)
		self._tabPage2.Controls.Add(self.Title)
		self._tabPage2.Controls.Add(self.Format)
		self._tabPage2.Controls.Add(self.Series)
		self._tabPage2.Controls.Add(self.Number)
		self._tabPage2.Controls.Add(self.AlternateNumber)
		self._tabPage2.Controls.Add(self.Year)
		self._tabPage2.Controls.Add(self.MonthNumber)
		self._tabPage2.Controls.Add(self.Count)
		self._tabPage2.Controls.Add(self.AlternateSeries)
		self._tabPage2.Controls.Add(self._label6)
		self._tabPage2.Controls.Add(self.AlternateNumber)
		self._tabPage2.Controls.Add(self._label5)
		self._tabPage2.Controls.Add(self.AlternateCount)
		self._tabPage2.Controls.Add(self._label3)
		self._tabPage2.Controls.Add(self.Month)
		self._tabPage2.Controls.Add(self._label4)
		self._tabPage2.Controls.Add(self._label1)
		self._tabPage2.Location = System.Drawing.Point(4, 22)
		self._tabPage2.Name = "tabPage2"
		self._tabPage2.Padding = System.Windows.Forms.Padding(3)
		self._tabPage2.Size = System.Drawing.Size(500, 228)
		self._tabPage2.TabIndex = 0
		self._tabPage2.Text = "Basic"
		self._tabPage2.UseVisualStyleBackColor = True
		# 
		# tpInsertAdvanced
		# 
		self._tpInsertAdvanced.Controls.Add(self.StartYear)
		self._tpInsertAdvanced.Controls.Add(self._label22)
		self._tpInsertAdvanced.Controls.Add(self.Manga)
		self._tpInsertAdvanced.Controls.Add(self.Character)
		self._tpInsertAdvanced.Controls.Add(self.Genre)
		self._tpInsertAdvanced.Controls.Add(self.Tags)
		self._tpInsertAdvanced.Controls.Add(self.Writer)
		self._tpInsertAdvanced.Controls.Add(self._label21)
		self._tpInsertAdvanced.Controls.Add(self._label16)
		self._tpInsertAdvanced.Controls.Add(self._label15)

		self._tpInsertAdvanced.Location = System.Drawing.Point(4, 22)
		self._tpInsertAdvanced.Name = "tpInsertAdvanced"
		self._tpInsertAdvanced.Padding = System.Windows.Forms.Padding(3)
		self._tpInsertAdvanced.Size = System.Drawing.Size(500, 228)
		self._tpInsertAdvanced.TabIndex = 1
		self._tpInsertAdvanced.Text = "Advanced"
		self._tpInsertAdvanced.UseVisualStyleBackColor = True
		# 
		# Manga
		# 
		self.Manga.Location = Point(240, 23)
		self.Manga.SetTemplate("manga", "Manga")
		# 
		# label16
		# 
		self._label16.AutoSize = True
		self._label16.Location = System.Drawing.Point(453, 3)
		self._label16.Name = "label16"
		self._label16.Size = System.Drawing.Size(33, 13)
		self._label16.TabIndex = 63
		self._label16.Text = "Other"
		# 
		# label21
		# 
		self._label21.Location = System.Drawing.Point(240, 48)
		self._label21.Name = "label21"
		self._label21.Size = System.Drawing.Size(255, 44)
		self._label21.TabIndex = 64
		self._label21.Text = "Fill in the other box for the text to be inserted when the item is marked as manga. If the item is not manga, nothing will be inserted"
		# 
		# Character
		# 
		self.Character.Location = Point(6, 143)
		self.Character.SetTemplate("character", "Character")
		# 
		# Writer
		# 
		self.Writer.Location = Point(241, 183)
		self.Writer.SetTemplate("writer", "Writer")
		# 
		# Genre
		# 
		self.Genre.Location = Point(241, 144)
		self.Genre.SetTemplate("genre", "Genre")
		# 
		# Tags
		# 
		self.Tags.Location = Point(6, 182)
		self.Tags.SetTemplate("tags", "Tags")
		# 
		# label22
		# 
		self._label22.Location = System.Drawing.Point(6, 100)
		self._label22.Name = "label22"
		self._label22.Size = System.Drawing.Size(494, 40)
		self._label22.TabIndex = 77
		self._label22.Text = "For these fields that can have multiple entries the script will ask which ones you would like to use. If you want to select for all the issues in the series at once, check the checkbox beside the field before inserting. In the selction box is an option to have each entry as its own folder"
		# 
		# ConfigForm
		# 
		self.AcceptButton = self._ok
		self.CancelButton = self._cancel
		self.ClientSize = System.Drawing.Size(542, 519)
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
		self._tpDirOrganize.ResumeLayout(False)
		self._tpDirOrganize.PerformLayout()
		self._tpFileNames.ResumeLayout(False)
		self._tpFileNames.PerformLayout()
		self._groupBox1.ResumeLayout(False)
		self._groupBox1.PerformLayout()
		self._tpExcludes.ResumeLayout(False)
		self._ExPanel.ResumeLayout(False)
		self._ExPanel.PerformLayout()
		self._groupBox2.ResumeLayout(False)
		self._groupBox2.PerformLayout()
		self._tabPage1.ResumeLayout(False)
		self._tabPage1.PerformLayout()
		self._gbMode.ResumeLayout(False)
		self._gbMode.PerformLayout()
		self._insertTabs.ResumeLayout(False)
		self._tabPage2.ResumeLayout(False)
		self._tabPage2.PerformLayout()
		self._tpInsertAdvanced.ResumeLayout(False)
		self._tpInsertAdvanced.PerformLayout()
		self.ResumeLayout(False)


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
			self._tpFileNames.Controls.Remove(self._groupBox1)
			self._tpFileNames.Controls.Remove(self._vsbBookSelector)
			self._tpDirOrganize.Controls.Add(self._vsbBookSelector)
			self._vsbBookSelector.Location = System.Drawing.Point(490, 98)
			self._tpDirOrganize.Controls.Add(self._groupBox1)

		if self._tabs.SelectedIndex == 1:
			self._btnSep.Enabled = False
			self._btnSep.Visible = False
			self._tpDirOrganize.Controls.Remove(self._groupBox1)
			self._tpDirOrganize.Controls.Remove(self._vsbBookSelector)
			self._tpFileNames.Controls.Add(self._vsbBookSelector)
			self._vsbBookSelector.Location = System.Drawing.Point(490, 69)
			self._tpFileNames.Controls.Add(self._groupBox1)
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
				self._sampleTextDir.Text = self.PathCreator.CreateDirectoryPath(self.samplebook, self._txbDirStruct.Text, self._txbBaseDir.Text, self._txbEmptyDir.Text, self.settings.EmptyData)
			else:
				self._sampleTextDir.Text = self.samplebook.FileDirectory
				
		#Filename preview
		elif self._tabs.SelectedIndex == 1:
			if not ExcludePath(self.samplebook, list(self._lbExFolder.Items)) and not ExcludeMeta(self.samplebook, self.settings.ExcludeRules, self.settings.ExcludeOperator, self.settings.ExcludeMode) and self._ckbFileNaming.Checked:
				self._sampleTextFile.Text = self.PathCreator.CreateFileName(self.samplebook, self._txbFileStruct.Text, self.settings.EmptyData, self._cmbImageFormat.SelectedItem)
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
		
		if self.settings.Mode == Mode.Move:
			self._rdbModeMove.Checked = True
		elif self.settings.Mode == Mode.Copy:
			self._rdbModeCopy.Checked = True
		elif self.settings.Mode == Mode.Test:
			self._rdbModeTest.Checked = True
		
		#Reload the selected index of the emptydata cmb
		self.CmbEmptyDataSelectedIndexChanged(self._cmbEmptyData, None)
		
		#Excludes
		

		self._lbExFolder.Items.Clear()
		for i in self.settings.ExcludeFolders:
			self._lbExFolder.Items.Add(i)
		
		#If changing settings to a different one. Clear the exlcude rule  container
		
		self._flpExcludes.Controls.Clear()

		for i in self.settings.ExcludeRules:
			#Check for none in the case of switching profiles with deleted rules
			if not i == None:
				i.Remove.Click += self.RemoveRuleSet
				self._flpExcludes.Controls.Add(i.Panel)

		
		self._cmbExcludeOperator.SelectedItem = self.settings.ExcludeOperator
		self._cmbExcludeMode.SelectedItem = self.settings.ExcludeMode
		
		self._ckbFileless.Checked = self.settings.MoveFileless
		self._cmbImageFormat.SelectedItem = self.settings.FilelessFormat

		self._ckbRemoveEmptyDir.Checked = self.settings.RemoveEmptyDir
		self._lbRemoveEmptyDir.Items.Clear()
		self._lbRemoveEmptyDir.Items.AddRange(System.Array[System.String](self.settings.ExcludedEmptyDir))

		#post &  pre
		self.Publisher.SetPrefixText(self.settings.Pre["Publisher"])
		self.Publisher.SetPostfixText(self.settings.Post["Publisher"])
		self.Imprint.SetPrefixText(self.settings.Pre["Imprint"])
		self.Imprint.SetPostfixText(self.settings.Post["Imprint"])
		self.Series.SetPostfixText(self.settings.Post["Series"])
		self.Series.SetPrefixText(self.settings.Pre["Series"])
		self.Title.SetPostfixText(self.settings.Post["Title"])
		self.Title.SetPrefixText(self.settings.Pre["Title"])
		self.Format.SetPostfixText(self.settings.Post["Format"])
		self.Format.SetPrefixText(self.settings.Pre["Format"])
		self.Volume.SetPostfixText(self.settings.Post["Volume"])
		self.Volume.SetPrefixText(self.settings.Pre["Volume"])
		self.AlternateSeries.SetPostfixText(self.settings.Post["AltSeries"])
		self.AlternateSeries.SetPrefixText(self.settings.Pre["AltSeries"])
		self.AlternateNumber.SetPostfixText(self.settings.Post["AltNumber"])
		self.AlternateNumber.SetPrefixText(self.settings.Pre["AltNumber"])
		self.AlternateCount.SetPostfixText(self.settings.Post["AltCount"])
		self.AlternateCount.SetPrefixText(self.settings.Pre["AltCount"])
		self.Year.SetPostfixText(self.settings.Post["Year"])
		self.Year.SetPrefixText(self.settings.Pre["Year"])
		self.Month.SetPostfixText(self.settings.Post["Month"])
		self.Month.SetPrefixText(self.settings.Pre["Month"])
		self.MonthNumber.SetPostfixText(self.settings.Post["Month#"])
		self.MonthNumber.SetPrefixText(self.settings.Pre["Month#"])
		self.Count.SetPostfixText(self.settings.Post["Count"])
		self.Count.SetPrefixText(self.settings.Pre["Count"])
		self.Number.SetPostfixText(self.settings.Post["Number"])
		self.Number.SetPrefixText(self.settings.Pre["Number"])
		""" Temportaly disabled while switching to new insert controls

		self._txbStartYearPost.Text = self.settings.Post["StartYear"]
		self._txbStartYearPre.Text = self.settings.Pre["StartYear"]
		"""

		#Other stuff has to be loaded before text boxes, otherwise they may have the wrong calculated value.
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
		
		#post &  pre
		self.settings.Pre["Publisher"] = self.Publisher.GetPrefixText()
		self.settings.Post["Publisher"] = self.Publisher.GetPostfixText()
		self.settings.Pre["Imprint"] = self.Imprint.GetPrefixText()
		self.settings.Post["Imprint"] = self.Imprint.GetPostfixText()
		self.settings.Post["Series"] = self.Series.GetPostfixText()
		self.settings.Pre["Series"] = self.Series.GetPrefixText()
		self.settings.Post["Title"] = self.Title.GetPostfixText()
		self.settings.Pre["Title"] = self.Title.GetPrefixText()
		self.settings.Post["Format"] = self.Format.GetPostfixText()
		self.settings.Pre["Format"] = self.Format.GetPrefixText()
		self.settings.Post["Volume"] = self.Volume.GetPostfixText()
		self.settings.Pre["Volume"] = self.Volume.GetPrefixText()
		self.settings.Post["AltSeries"] = self.AlternateSeries.GetPostfixText()
		self.settings.Pre["AltSeries"] = self.AlternateSeries.GetPrefixText()
		self.settings.Post["AltNumber"] = self.AlternateNumber.GetPostfixText()
		self.settings.Pre["AltNumber"] = self.AlternateNumber.GetPrefixText()
		self.settings.Post["AltCount"] = self.AlternateCount.GetPostfixText()
		self.settings.Pre["AltCount"] = self.AlternateCount.GetPrefixText()
		self.settings.Post["Year"] = self.Year.GetPostfixText()
		self.settings.Pre["Year"] = self.Year.GetPrefixText()
		self.settings.Post["Month"] = self.Month.GetPostfixText()
		self.settings.Pre["Month"] = self.Month.GetPrefixText()
		self.settings.Post["Month#"] = self.MonthNumber.GetPostfixText()
		self.settings.Pre["Month#"] = self.MonthNumber.GetPrefixText()
		self.settings.Post["Count"] = self.Count.GetPostfixText()
		self.settings.Pre["Count"] = self.Count.GetPrefixText()
		self.settings.Post["Number"] = self.Number.GetPostfixText()
		self.settings.Pre["Number"] = self.Number.GetPrefixText()
		#self.settings.Post["StartYear"] = self._txbStartYearPost.Text
		#self.settings.Pre["StartYear"] = self._txbStartYearPre.Text
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
		#Don't delete the index of the list as that will screw up other deletion methods
		#Instead make it a null index so that it is skipped when calculated
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
		self.Check = CheckBox()
		self.Check.Size = Size(15, 14)
		self.Check.Location = Point(200, 4)
		self.Check.TabIndex = 3

		self.Width = 220
		self.Controls.Add(self.Check)
		

	def GetTemplateText(self, space):
		if self.Check.Checked:
			checktext = "(series)"
		else:
			checktext = "(issue)"
		s = ""
		if space:
			s = " "
		return "{" + self.Prefix.Text + "<" + self.Template + checktext + ">" + self.Postfix.Text + s + "}"

class InsertControlTextBox(InsertControl):
	
	def __init__(self):
		super(InsertControlTextBox,self).__init__()
		self.TextBox = TextBox()
		self.TextBox.Size = Size(58, 22)
		self.TextBox.Location = Point(200, 2)
		self.TextBox.TabIndex = 3

		self.Width = 258
		self.Controls.Add(self.TextBox)
		

	def GetTemplateText(self, space):
		s = ""
		if space:
			s = " "
		return "{" + self.Prefix.Text + "<" + self.Template + "(" + self.TextBox.Text + ")>" + self.Postfix.Text + s + "}"



