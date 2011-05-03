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

from locommon import PathMaker, ICON, ExcludePath, ExcludeRule

class ConfigForm(Form):
	def __init__(self, books, allsettings, lastused):

		self.Icon = System.Drawing.Icon(ICON)
		self.InitializeComponent()		
		self.allbooks = books
		self.samplebook = self.allbooks[self._vsbBookSelector.Value]
		self._vsbBookSelector.Maximum = len(self.allbooks)
		self.allsettings = allsettings
		self.PathCreator = PathMaker()
		#In the case that lastused is null
		try:
			self.settings = allsettings[lastused]
		except KeyError:
			self.settings = allsettings[allsettings.keys()[0]]
			
		self._cmbEmptyData.SelectedIndex = 0
			
		#Array for excludes adding of metadata rules. Each list item will contain an ExcludeRule object or ExcludeGroup object.
		self.Excludes = []
		
		self.LoadSettings()
		
		for i in allsettings:
			self._cmbProfiles.Items.Add(i)
		self._cmbProfiles.SelectedItem = self.settings.Name
		
		#For saving the previous index.
		self._cmbProfiles.Tag = self._cmbProfiles.SelectedIndex
		
		
		
	def InitializeComponent(self):
		self._components = System.ComponentModel.Container()
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
		self._btnInsertPub = System.Windows.Forms.Button()
		self._btnImprint = System.Windows.Forms.Button()
		self._btnSeries = System.Windows.Forms.Button()
		self._btnVolume = System.Windows.Forms.Button()
		self._btnNumber = System.Windows.Forms.Button()
		self._btnYear = System.Windows.Forms.Button()
		self._btnMonthNumber = System.Windows.Forms.Button()
		self._btnCount = System.Windows.Forms.Button()
		self._btnTitle = System.Windows.Forms.Button()
		self._btnFormat = System.Windows.Forms.Button()
		self._btnAltSeries = System.Windows.Forms.Button()
		self._btnAltNumber = System.Windows.Forms.Button()
		self._btnAltCount = System.Windows.Forms.Button()
		self._ckbFileNaming = System.Windows.Forms.CheckBox()
		self._txbFileStruct = System.Windows.Forms.TextBox()
		self._lblFileStruct = System.Windows.Forms.Label()
		self._btnMonthText = System.Windows.Forms.Button()
		self._txbPubPre = System.Windows.Forms.TextBox()
		self._txbPubPost = System.Windows.Forms.TextBox()
		self._txbImprintPre = System.Windows.Forms.TextBox()
		self._txbImprintPost = System.Windows.Forms.TextBox()
		self._txbSeriesPost = System.Windows.Forms.TextBox()
		self._txbSeriesPre = System.Windows.Forms.TextBox()
		self._txbTitlePost = System.Windows.Forms.TextBox()
		self._txbTitlePre = System.Windows.Forms.TextBox()
		self._txbFormatPost = System.Windows.Forms.TextBox()
		self._txbFormatPre = System.Windows.Forms.TextBox()
		self._txbVolumePost = System.Windows.Forms.TextBox()
		self._txbVolumePre = System.Windows.Forms.TextBox()
		self._txbAltSeriesPost = System.Windows.Forms.TextBox()
		self._txbAltSeriesPre = System.Windows.Forms.TextBox()
		self._txbAltNumberPost = System.Windows.Forms.TextBox()
		self._txbAltNumberPre = System.Windows.Forms.TextBox()
		self._txbAltCountPost = System.Windows.Forms.TextBox()
		self._txbAltCountPre = System.Windows.Forms.TextBox()
		self._txbYearPost = System.Windows.Forms.TextBox()
		self._txbYearPre = System.Windows.Forms.TextBox()
		self._txbMonthPost = System.Windows.Forms.TextBox()
		self._txbMonthPre = System.Windows.Forms.TextBox()
		self._txbMonthNumberPost = System.Windows.Forms.TextBox()
		self._txbMonthNumberPre = System.Windows.Forms.TextBox()
		self._txbCountPost = System.Windows.Forms.TextBox()
		self._txbCountPre = System.Windows.Forms.TextBox()
		self._txbNumberPost = System.Windows.Forms.TextBox()
		self._txbNumberPre = System.Windows.Forms.TextBox()
		self._label1 = System.Windows.Forms.Label()
		self._label2 = System.Windows.Forms.Label()
		self._label3 = System.Windows.Forms.Label()
		self._label4 = System.Windows.Forms.Label()
		self._label5 = System.Windows.Forms.Label()
		self._label6 = System.Windows.Forms.Label()
		self._nudVolume = System.Windows.Forms.NumericUpDown()
		self._nudNumber = System.Windows.Forms.NumericUpDown()
		self._nudCount = System.Windows.Forms.NumericUpDown()
		self._nudMonth = System.Windows.Forms.NumericUpDown()
		self._nudAltCount = System.Windows.Forms.NumericUpDown()
		self._nudAltNumber = System.Windows.Forms.NumericUpDown()
		self._FolderBrowser = System.Windows.Forms.FolderBrowserDialog()
		self._ckbSpace = System.Windows.Forms.CheckBox()
		self._tooltip = System.Windows.Forms.ToolTip(self._components)
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
		self._cmbExMatchType = System.Windows.Forms.ComboBox()
		self._label16 = System.Windows.Forms.Label()
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
		self._txbStartYearPost = System.Windows.Forms.TextBox()
		self._txbStartYearPre = System.Windows.Forms.TextBox()
		self._btnStartYear = System.Windows.Forms.Button()
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
		self._tabs.SuspendLayout()
		self._tpDirOrganize.SuspendLayout()
		self._tpFileNames.SuspendLayout()
		self._groupBox1.SuspendLayout()
		self._nudVolume.BeginInit()
		self._nudNumber.BeginInit()
		self._nudCount.BeginInit()
		self._nudMonth.BeginInit()
		self._nudAltCount.BeginInit()
		self._nudAltNumber.BeginInit()
		self._tpExcludes.SuspendLayout()
		self._ExPanel.SuspendLayout()
		self._groupBox2.SuspendLayout()
		self._tabPage1.SuspendLayout()
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
		self._txbDirStruct.Enter += self.TxbFocusEnter
		self._txbDirStruct.Leave += self.TxbFocusLeave
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
		self._groupBox1.Controls.Add(self._label15)
		self._groupBox1.Controls.Add(self._txbStartYearPost)
		self._groupBox1.Controls.Add(self._txbStartYearPre)
		self._groupBox1.Controls.Add(self._btnStartYear)
		self._groupBox1.Controls.Add(self._ckbSpace)
		self._groupBox1.Controls.Add(self._nudAltNumber)
		self._groupBox1.Controls.Add(self._nudAltCount)
		self._groupBox1.Controls.Add(self._nudMonth)
		self._groupBox1.Controls.Add(self._nudCount)
		self._groupBox1.Controls.Add(self._nudNumber)
		self._groupBox1.Controls.Add(self._nudVolume)
		self._groupBox1.Controls.Add(self._label6)
		self._groupBox1.Controls.Add(self._label5)
		self._groupBox1.Controls.Add(self._label3)
		self._groupBox1.Controls.Add(self._label4)
		self._groupBox1.Controls.Add(self._label2)
		self._groupBox1.Controls.Add(self._label1)
		self._groupBox1.Controls.Add(self._txbAltNumberPost)
		self._groupBox1.Controls.Add(self._txbAltNumberPre)
		self._groupBox1.Controls.Add(self._txbAltCountPost)
		self._groupBox1.Controls.Add(self._txbAltCountPre)
		self._groupBox1.Controls.Add(self._txbYearPost)
		self._groupBox1.Controls.Add(self._txbYearPre)
		self._groupBox1.Controls.Add(self._txbMonthPost)
		self._groupBox1.Controls.Add(self._txbMonthPre)
		self._groupBox1.Controls.Add(self._txbMonthNumberPost)
		self._groupBox1.Controls.Add(self._txbMonthNumberPre)
		self._groupBox1.Controls.Add(self._txbCountPost)
		self._groupBox1.Controls.Add(self._txbCountPre)
		self._groupBox1.Controls.Add(self._txbNumberPost)
		self._groupBox1.Controls.Add(self._txbNumberPre)
		self._groupBox1.Controls.Add(self._txbAltSeriesPost)
		self._groupBox1.Controls.Add(self._txbAltSeriesPre)
		self._groupBox1.Controls.Add(self._txbVolumePost)
		self._groupBox1.Controls.Add(self._txbVolumePre)
		self._groupBox1.Controls.Add(self._txbFormatPost)
		self._groupBox1.Controls.Add(self._txbFormatPre)
		self._groupBox1.Controls.Add(self._txbTitlePost)
		self._groupBox1.Controls.Add(self._txbTitlePre)
		self._groupBox1.Controls.Add(self._txbSeriesPost)
		self._groupBox1.Controls.Add(self._txbSeriesPre)
		self._groupBox1.Controls.Add(self._txbImprintPost)
		self._groupBox1.Controls.Add(self._txbImprintPre)
		self._groupBox1.Controls.Add(self._txbPubPost)
		self._groupBox1.Controls.Add(self._txbPubPre)
		self._groupBox1.Controls.Add(self._btnMonthText)
		self._groupBox1.Controls.Add(self._btnAltCount)
		self._groupBox1.Controls.Add(self._btnAltNumber)
		self._groupBox1.Controls.Add(self._btnAltSeries)
		self._groupBox1.Controls.Add(self._btnFormat)
		self._groupBox1.Controls.Add(self._btnTitle)
		self._groupBox1.Controls.Add(self._btnCount)
		self._groupBox1.Controls.Add(self._btnMonthNumber)
		self._groupBox1.Controls.Add(self._btnYear)
		self._groupBox1.Controls.Add(self._btnNumber)
		self._groupBox1.Controls.Add(self._btnVolume)
		self._groupBox1.Controls.Add(self._btnSeries)
		self._groupBox1.Controls.Add(self._btnImprint)
		self._groupBox1.Controls.Add(self._btnInsertPub)
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
		# btnInsertPub
		# 
		self._btnInsertPub.Location = System.Drawing.Point(73, 67)
		self._btnInsertPub.Name = "btnInsertPub"
		self._btnInsertPub.Size = System.Drawing.Size(75, 23)
		self._btnInsertPub.TabIndex = 6
		self._btnInsertPub.Text = "Publisher"
		self._btnInsertPub.UseVisualStyleBackColor = True
		self._btnInsertPub.Click += self.BtnInsertPubClick
		# 
		# btnImprint
		# 
		self._btnImprint.Location = System.Drawing.Point(73, 95)
		self._btnImprint.Name = "btnImprint"
		self._btnImprint.Size = System.Drawing.Size(75, 23)
		self._btnImprint.TabIndex = 9
		self._btnImprint.Text = "Imprint"
		self._btnImprint.UseVisualStyleBackColor = True
		self._btnImprint.Click += self.BtnImprintClick
		# 
		# btnSeries
		# 
		self._btnSeries.Location = System.Drawing.Point(73, 123)
		self._btnSeries.Name = "btnSeries"
		self._btnSeries.Size = System.Drawing.Size(75, 23)
		self._btnSeries.TabIndex = 12
		self._btnSeries.Text = "Series"
		self._btnSeries.UseVisualStyleBackColor = True
		self._btnSeries.Click += self.BtnSeriesClick
		# 
		# btnVolume
		# 
		self._btnVolume.Location = System.Drawing.Point(73, 207)
		self._btnVolume.Name = "btnVolume"
		self._btnVolume.Size = System.Drawing.Size(75, 23)
		self._btnVolume.TabIndex = 21
		self._btnVolume.Text = "Volume"
		self._btnVolume.UseVisualStyleBackColor = True
		self._btnVolume.Click += self.BtnVolumeClick
		# 
		# btnNumber
		# 
		self._btnNumber.Location = System.Drawing.Point(328, 67)
		self._btnNumber.Name = "btnNumber"
		self._btnNumber.Size = System.Drawing.Size(75, 23)
		self._btnNumber.TabIndex = 30
		self._btnNumber.Text = "Number"
		self._btnNumber.UseVisualStyleBackColor = True
		self._btnNumber.Click += self.BtnNumberClick
		# 
		# btnYear
		# 
		self._btnYear.Location = System.Drawing.Point(328, 179)
		self._btnYear.Name = "btnYear"
		self._btnYear.Size = System.Drawing.Size(75, 23)
		self._btnYear.TabIndex = 45
		self._btnYear.Text = "Year"
		self._btnYear.UseVisualStyleBackColor = True
		self._btnYear.Click += self.BtnYearClick
		# 
		# btnMonthNumber
		# 
		self._btnMonthNumber.Location = System.Drawing.Point(328, 123)
		self._btnMonthNumber.Name = "btnMonthNumber"
		self._btnMonthNumber.Size = System.Drawing.Size(75, 23)
		self._btnMonthNumber.TabIndex = 38
		self._btnMonthNumber.Text = "Month (#)"
		self._btnMonthNumber.UseVisualStyleBackColor = True
		self._btnMonthNumber.Click += self.BtnMonthNumberClick
		# 
		# btnCount
		# 
		self._btnCount.Location = System.Drawing.Point(328, 95)
		self._btnCount.Name = "btnCount"
		self._btnCount.Size = System.Drawing.Size(75, 23)
		self._btnCount.TabIndex = 34
		self._btnCount.Text = "Count"
		self._btnCount.UseVisualStyleBackColor = True
		self._btnCount.Click += self.BtnCountClick
		# 
		# btnTitle
		# 
		self._btnTitle.Location = System.Drawing.Point(73, 151)
		self._btnTitle.Name = "btnTitle"
		self._btnTitle.Size = System.Drawing.Size(75, 23)
		self._btnTitle.TabIndex = 15
		self._btnTitle.Text = "Title"
		self._btnTitle.UseVisualStyleBackColor = True
		self._btnTitle.Click += self.BtnTitleClick
		# 
		# btnFormat
		# 
		self._btnFormat.Location = System.Drawing.Point(73, 179)
		self._btnFormat.Name = "btnFormat"
		self._btnFormat.Size = System.Drawing.Size(75, 23)
		self._btnFormat.TabIndex = 18
		self._btnFormat.Text = "Format"
		self._btnFormat.UseVisualStyleBackColor = True
		self._btnFormat.Click += self.BtnFormatClick
		# 
		# btnAltSeries
		# 
		self._btnAltSeries.Location = System.Drawing.Point(73, 235)
		self._btnAltSeries.Name = "btnAltSeries"
		self._btnAltSeries.Size = System.Drawing.Size(75, 23)
		self._btnAltSeries.TabIndex = 24
		self._btnAltSeries.Text = "Alt. Series"
		self._btnAltSeries.UseVisualStyleBackColor = True
		self._btnAltSeries.Click += self.BtnAltSeriesClick
		# 
		# btnAltNumber
		# 
		self._btnAltNumber.Location = System.Drawing.Point(328, 235)
		self._btnAltNumber.Name = "btnAltNumber"
		self._btnAltNumber.Size = System.Drawing.Size(75, 23)
		self._btnAltNumber.TabIndex = 52
		self._btnAltNumber.Text = "Alt. Number"
		self._btnAltNumber.UseVisualStyleBackColor = True
		self._btnAltNumber.Click += self.BtnAltNumberClick
		# 
		# btnAltCount
		# 
		self._btnAltCount.Location = System.Drawing.Point(328, 207)
		self._btnAltCount.Name = "btnAltCount"
		self._btnAltCount.Size = System.Drawing.Size(75, 23)
		self._btnAltCount.TabIndex = 48
		self._btnAltCount.Text = "Alt. Count"
		self._btnAltCount.UseVisualStyleBackColor = True
		self._btnAltCount.Click += self.BtnAltCountClick
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
		self._txbFileStruct.Enter += self.TxbFocusEnter
		self._txbFileStruct.Leave += self.TxbFocusLeave
		# 
		# lblFileStruct
		# 
		self._lblFileStruct.Location = System.Drawing.Point(9, 34)
		self._lblFileStruct.Name = "lblFileStruct"
		self._lblFileStruct.Size = System.Drawing.Size(71, 35)
		self._lblFileStruct.TabIndex = 2
		self._lblFileStruct.Text = "File Structure"
		# 
		# btnMonthText
		# 
		self._btnMonthText.Location = System.Drawing.Point(328, 149)
		self._btnMonthText.Name = "btnMonthText"
		self._btnMonthText.Size = System.Drawing.Size(75, 23)
		self._btnMonthText.TabIndex = 42
		self._btnMonthText.Text = "Month"
		self._btnMonthText.UseVisualStyleBackColor = True
		self._btnMonthText.Click += self.BtnMonthTextClick
		# 
		# txbPubPre
		# 
		self._txbPubPre.Location = System.Drawing.Point(8, 68)
		self._txbPubPre.Name = "txbPubPre"
		self._txbPubPre.Size = System.Drawing.Size(58, 20)
		self._txbPubPre.TabIndex = 5
		# 
		# txbPubPost
		# 
		self._txbPubPost.Location = System.Drawing.Point(154, 68)
		self._txbPubPost.Name = "txbPubPost"
		self._txbPubPost.Size = System.Drawing.Size(58, 20)
		self._txbPubPost.TabIndex = 7
		# 
		# txbImprintPre
		# 
		self._txbImprintPre.Location = System.Drawing.Point(8, 96)
		self._txbImprintPre.Name = "txbImprintPre"
		self._txbImprintPre.Size = System.Drawing.Size(58, 20)
		self._txbImprintPre.TabIndex = 8
		# 
		# txbImprintPost
		# 
		self._txbImprintPost.Location = System.Drawing.Point(154, 96)
		self._txbImprintPost.Name = "txbImprintPost"
		self._txbImprintPost.Size = System.Drawing.Size(58, 20)
		self._txbImprintPost.TabIndex = 10
		# 
		# txbSeriesPost
		# 
		self._txbSeriesPost.Location = System.Drawing.Point(155, 124)
		self._txbSeriesPost.Name = "txbSeriesPost"
		self._txbSeriesPost.Size = System.Drawing.Size(58, 20)
		self._txbSeriesPost.TabIndex = 13
		# 
		# txbSeriesPre
		# 
		self._txbSeriesPre.Location = System.Drawing.Point(8, 124)
		self._txbSeriesPre.Name = "txbSeriesPre"
		self._txbSeriesPre.Size = System.Drawing.Size(58, 20)
		self._txbSeriesPre.TabIndex = 11
		# 
		# txbTitlePost
		# 
		self._txbTitlePost.Location = System.Drawing.Point(155, 152)
		self._txbTitlePost.Name = "txbTitlePost"
		self._txbTitlePost.Size = System.Drawing.Size(58, 20)
		self._txbTitlePost.TabIndex = 16
		# 
		# txbTitlePre
		# 
		self._txbTitlePre.Location = System.Drawing.Point(8, 152)
		self._txbTitlePre.Name = "txbTitlePre"
		self._txbTitlePre.Size = System.Drawing.Size(58, 20)
		self._txbTitlePre.TabIndex = 14
		# 
		# txbFormatPost
		# 
		self._txbFormatPost.Location = System.Drawing.Point(154, 180)
		self._txbFormatPost.Name = "txbFormatPost"
		self._txbFormatPost.Size = System.Drawing.Size(58, 20)
		self._txbFormatPost.TabIndex = 3
		# 
		# txbFormatPre
		# 
		self._txbFormatPre.Location = System.Drawing.Point(8, 180)
		self._txbFormatPre.Name = "txbFormatPre"
		self._txbFormatPre.Size = System.Drawing.Size(58, 20)
		self._txbFormatPre.TabIndex = 17
		# 
		# txbVolumePost
		# 
		self._txbVolumePost.Location = System.Drawing.Point(154, 208)
		self._txbVolumePost.Name = "txbVolumePost"
		self._txbVolumePost.Size = System.Drawing.Size(58, 20)
		self._txbVolumePost.TabIndex = 0
		# 
		# txbVolumePre
		# 
		self._txbVolumePre.Location = System.Drawing.Point(8, 208)
		self._txbVolumePre.Name = "txbVolumePre"
		self._txbVolumePre.Size = System.Drawing.Size(58, 20)
		self._txbVolumePre.TabIndex = 20
		# 
		# txbAltSeriesPost
		# 
		self._txbAltSeriesPost.Location = System.Drawing.Point(155, 236)
		self._txbAltSeriesPost.Name = "txbAltSeriesPost"
		self._txbAltSeriesPost.Size = System.Drawing.Size(58, 20)
		self._txbAltSeriesPost.TabIndex = 2
		# 
		# txbAltSeriesPre
		# 
		self._txbAltSeriesPre.Location = System.Drawing.Point(8, 236)
		self._txbAltSeriesPre.Name = "txbAltSeriesPre"
		self._txbAltSeriesPre.Size = System.Drawing.Size(58, 20)
		self._txbAltSeriesPre.TabIndex = 23
		# 
		# txbAltNumberPost
		# 
		self._txbAltNumberPost.Location = System.Drawing.Point(410, 236)
		self._txbAltNumberPost.Name = "txbAltNumberPost"
		self._txbAltNumberPost.Size = System.Drawing.Size(58, 20)
		self._txbAltNumberPost.TabIndex = 53
		# 
		# txbAltNumberPre
		# 
		self._txbAltNumberPre.Location = System.Drawing.Point(264, 236)
		self._txbAltNumberPre.Name = "txbAltNumberPre"
		self._txbAltNumberPre.Size = System.Drawing.Size(58, 20)
		self._txbAltNumberPre.TabIndex = 51
		# 
		# txbAltCountPost
		# 
		self._txbAltCountPost.Location = System.Drawing.Point(410, 208)
		self._txbAltCountPost.Name = "txbAltCountPost"
		self._txbAltCountPost.Size = System.Drawing.Size(58, 20)
		self._txbAltCountPost.TabIndex = 49
		# 
		# txbAltCountPre
		# 
		self._txbAltCountPre.Location = System.Drawing.Point(264, 208)
		self._txbAltCountPre.Name = "txbAltCountPre"
		self._txbAltCountPre.Size = System.Drawing.Size(58, 20)
		self._txbAltCountPre.TabIndex = 47
		# 
		# txbYearPost
		# 
		self._txbYearPost.Location = System.Drawing.Point(410, 180)
		self._txbYearPost.Name = "txbYearPost"
		self._txbYearPost.Size = System.Drawing.Size(58, 20)
		self._txbYearPost.TabIndex = 46
		# 
		# txbYearPre
		# 
		self._txbYearPre.Location = System.Drawing.Point(264, 180)
		self._txbYearPre.Name = "txbYearPre"
		self._txbYearPre.Size = System.Drawing.Size(58, 20)
		self._txbYearPre.TabIndex = 44
		# 
		# txbMonthPost
		# 
		self._txbMonthPost.Location = System.Drawing.Point(410, 150)
		self._txbMonthPost.Name = "txbMonthPost"
		self._txbMonthPost.Size = System.Drawing.Size(58, 20)
		self._txbMonthPost.TabIndex = 43
		# 
		# txbMonthPre
		# 
		self._txbMonthPre.Location = System.Drawing.Point(264, 150)
		self._txbMonthPre.Name = "txbMonthPre"
		self._txbMonthPre.Size = System.Drawing.Size(58, 20)
		self._txbMonthPre.TabIndex = 41
		# 
		# txbMonthNumberPost
		# 
		self._txbMonthNumberPost.Location = System.Drawing.Point(410, 124)
		self._txbMonthNumberPost.Name = "txbMonthNumberPost"
		self._txbMonthNumberPost.Size = System.Drawing.Size(58, 20)
		self._txbMonthNumberPost.TabIndex = 39
		# 
		# txbMonthNumberPre
		# 
		self._txbMonthNumberPre.Location = System.Drawing.Point(264, 124)
		self._txbMonthNumberPre.Name = "txbMonthNumberPre"
		self._txbMonthNumberPre.Size = System.Drawing.Size(58, 20)
		self._txbMonthNumberPre.TabIndex = 37
		# 
		# txbCountPost
		# 
		self._txbCountPost.Location = System.Drawing.Point(410, 96)
		self._txbCountPost.Name = "txbCountPost"
		self._txbCountPost.Size = System.Drawing.Size(58, 20)
		self._txbCountPost.TabIndex = 35
		# 
		# txbCountPre
		# 
		self._txbCountPre.Location = System.Drawing.Point(264, 96)
		self._txbCountPre.Name = "txbCountPre"
		self._txbCountPre.Size = System.Drawing.Size(58, 20)
		self._txbCountPre.TabIndex = 33
		# 
		# txbNumberPost
		# 
		self._txbNumberPost.Location = System.Drawing.Point(410, 68)
		self._txbNumberPost.Name = "txbNumberPost"
		self._txbNumberPost.Size = System.Drawing.Size(58, 20)
		self._txbNumberPost.TabIndex = 31
		# 
		# txbNumberPre
		# 
		self._txbNumberPre.Location = System.Drawing.Point(264, 68)
		self._txbNumberPre.Name = "txbNumberPre"
		self._txbNumberPre.Size = System.Drawing.Size(58, 20)
		self._txbNumberPre.TabIndex = 29
		# 
		# label1
		# 
		self._label1.AutoSize = True
		self._label1.Location = System.Drawing.Point(21, 47)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(33, 13)
		self._label1.TabIndex = 2
		self._label1.Text = "Prefix"
		# 
		# label2
		# 
		self._label2.AutoSize = True
		self._label2.Location = System.Drawing.Point(165, 45)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(38, 13)
		self._label2.TabIndex = 3
		self._label2.Text = "Postfix"
		# 
		# label3
		# 
		self._label3.AutoSize = True
		self._label3.Location = System.Drawing.Point(420, 43)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(38, 13)
		self._label3.TabIndex = 27
		self._label3.Text = "Postfix"
		# 
		# label4
		# 
		self._label4.AutoSize = True
		self._label4.Location = System.Drawing.Point(277, 45)
		self._label4.Name = "label4"
		self._label4.Size = System.Drawing.Size(33, 13)
		self._label4.TabIndex = 26
		self._label4.Text = "Prefix"
		# 
		# label5
		# 
		self._label5.AutoSize = True
		self._label5.Location = System.Drawing.Point(222, 45)
		self._label5.Name = "label5"
		self._label5.Size = System.Drawing.Size(26, 13)
		self._label5.TabIndex = 4
		self._label5.Text = "Pad"
		# 
		# label6
		# 
		self._label6.AutoSize = True
		self._label6.Location = System.Drawing.Point(478, 43)
		self._label6.Name = "label6"
		self._label6.Size = System.Drawing.Size(26, 13)
		self._label6.TabIndex = 28
		self._label6.Text = "Pad"
		# 
		# nudVolume
		# 
		self._nudVolume.Location = System.Drawing.Point(218, 208)
		self._nudVolume.Name = "nudVolume"
		self._nudVolume.Size = System.Drawing.Size(34, 20)
		self._nudVolume.TabIndex = 1
		# 
		# nudNumber
		# 
		self._nudNumber.Location = System.Drawing.Point(474, 68)
		self._nudNumber.Name = "nudNumber"
		self._nudNumber.Size = System.Drawing.Size(34, 20)
		self._nudNumber.TabIndex = 32
		# 
		# nudCount
		# 
		self._nudCount.Location = System.Drawing.Point(474, 96)
		self._nudCount.Name = "nudCount"
		self._nudCount.Size = System.Drawing.Size(34, 20)
		self._nudCount.TabIndex = 36
		# 
		# nudMonth
		# 
		self._nudMonth.Location = System.Drawing.Point(474, 124)
		self._nudMonth.Name = "nudMonth"
		self._nudMonth.Size = System.Drawing.Size(34, 20)
		self._nudMonth.TabIndex = 40
		# 
		# nudAltCount
		# 
		self._nudAltCount.Location = System.Drawing.Point(474, 208)
		self._nudAltCount.Name = "nudAltCount"
		self._nudAltCount.Size = System.Drawing.Size(34, 20)
		self._nudAltCount.TabIndex = 50
		# 
		# nudAltNumber
		# 
		self._nudAltNumber.Location = System.Drawing.Point(474, 236)
		self._nudAltNumber.Name = "nudAltNumber"
		self._nudAltNumber.Size = System.Drawing.Size(34, 20)
		self._nudAltNumber.TabIndex = 54
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
		self._sampleTextDir.Size = System.Drawing.Size(424, 35)
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
		self._ExPanel.Location = System.Drawing.Point(3, 49)
		self._ExPanel.Name = "ExPanel"
		self._ExPanel.Size = System.Drawing.Size(504, 204)
		self._ExPanel.TabIndex = 4
		# 
		# cmbExMatchType
		# 
		self._cmbExMatchType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cmbExMatchType.FormattingEnabled = True
		self._cmbExMatchType.Items.AddRange(System.Array[System.Object](
			["Any",
			"All"]))
		self._cmbExMatchType.Location = System.Drawing.Point(46, 22)
		self._cmbExMatchType.Name = "cmbExMatchType"
		self._cmbExMatchType.Size = System.Drawing.Size(60, 21)
		self._cmbExMatchType.TabIndex = 0
		# 
		# label16
		# 
		self._label16.Location = System.Drawing.Point(6, 25)
		self._label16.Name = "label16"
		self._label16.Size = System.Drawing.Size(100, 23)
		self._label16.TabIndex = 1
		self._label16.Text = "Match"
		# 
		# label17
		# 
		self._label17.Location = System.Drawing.Point(112, 25)
		self._label17.Name = "label17"
		self._label17.Size = System.Drawing.Size(214, 23)
		self._label17.TabIndex = 2
		self._label17.Text = "of the following rules."
		# 
		# btnExMetaAdd
		# 
		self._btnExMetaAdd.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._btnExMetaAdd.Location = System.Drawing.Point(388, 20)
		self._btnExMetaAdd.Name = "btnExMetaAdd"
		self._btnExMetaAdd.Size = System.Drawing.Size(72, 23)
		self._btnExMetaAdd.TabIndex = 3
		self._btnExMetaAdd.Text = "Add rule"
		self._btnExMetaAdd.UseVisualStyleBackColor = True
		self._btnExMetaAdd.Click += self.CreateRuleSet
		# 
		# flpExcludes
		# 
		self._flpExcludes.AutoSize = True
		self._flpExcludes.Location = System.Drawing.Point(3, 3)
		self._flpExcludes.MaximumSize = System.Drawing.Size(482, 0)
		self._flpExcludes.Name = "flpExcludes"
		self._flpExcludes.Size = System.Drawing.Size(482, 40)
		self._flpExcludes.TabIndex = 8
		# 
		# groupBox2
		# 
		self._groupBox2.Controls.Add(self._btnExMetaAdd)
		self._groupBox2.Controls.Add(self._ExPanel)
		self._groupBox2.Controls.Add(self._label17)
		self._groupBox2.Controls.Add(self._cmbExMatchType)
		self._groupBox2.Controls.Add(self._label16)
		self._groupBox2.Location = System.Drawing.Point(8, 167)
		self._groupBox2.Name = "groupBox2"
		self._groupBox2.Size = System.Drawing.Size(507, 253)
		self._groupBox2.TabIndex = 6
		self._groupBox2.TabStop = False
		self._groupBox2.Text = "Do not move eComics if their metadata is as follows"
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
		# txbStartYearPost
		# 
		self._txbStartYearPost.Location = System.Drawing.Point(155, 262)
		self._txbStartYearPost.Name = "txbStartYearPost"
		self._txbStartYearPost.Size = System.Drawing.Size(58, 20)
		self._txbStartYearPost.TabIndex = 55
		# 
		# txbStartYearPre
		# 
		self._txbStartYearPre.Location = System.Drawing.Point(8, 262)
		self._txbStartYearPre.Name = "txbStartYearPre"
		self._txbStartYearPre.Size = System.Drawing.Size(58, 20)
		self._txbStartYearPre.TabIndex = 56
		# 
		# btnStartYear
		# 
		self._btnStartYear.Location = System.Drawing.Point(73, 261)
		self._btnStartYear.Name = "btnStartYear"
		self._btnStartYear.Size = System.Drawing.Size(75, 23)
		self._btnStartYear.TabIndex = 57
		self._btnStartYear.Text = "Start Year*"
		self._btnStartYear.UseVisualStyleBackColor = True
		self._btnStartYear.Click += self.BtnStartYearClick
		# 
		# label15
		# 
		self._label15.Font = System.Drawing.Font("Microsoft Sans Serif", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
		self._label15.Location = System.Drawing.Point(8, 285)
		self._label15.Name = "label15"
		self._label15.Size = System.Drawing.Size(205, 17)
		self._label15.TabIndex = 58
		self._label15.Text = "* Calculated from available data in library"
		# 
		# tabPage1
		# 
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
		self._txbEmptyData.Location = System.Drawing.Point(275, 112)
		self._txbEmptyData.Name = "txbEmptyData"
		self._txbEmptyData.Size = System.Drawing.Size(235, 20)
		self._txbEmptyData.TabIndex = 22
		self._txbEmptyData.Leave += self.TxbEmptyDataLeave
		# 
		# label11
		# 
		self._label11.AutoSize = True
		self._label11.Location = System.Drawing.Point(195, 115)
		self._label11.Name = "label11"
		self._label11.Size = System.Drawing.Size(74, 13)
		self._label11.TabIndex = 21
		self._label11.Text = "Default Value:"
		# 
		# label10
		# 
		self._label10.AutoSize = True
		self._label10.Location = System.Drawing.Point(8, 115)
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
		self._cmbEmptyData.Location = System.Drawing.Point(68, 112)
		self._cmbEmptyData.Name = "cmbEmptyData"
		self._cmbEmptyData.Size = System.Drawing.Size(121, 21)
		self._cmbEmptyData.Sorted = True
		self._cmbEmptyData.TabIndex = 19
		self._cmbEmptyData.SelectedIndexChanged += self.CmbEmptyDataSelectedIndexChanged
		# 
		# label9
		# 
		self._label9.AutoSize = True
		self._label9.Location = System.Drawing.Point(8, 88)
		self._label9.Name = "label9"
		self._label9.Size = System.Drawing.Size(441, 13)
		self._label9.TabIndex = 15
		self._label9.Text = "To create a default value to be entered in the case of a blank data field, fill in the box below."
		# 
		# label8
		# 
		self._label8.AutoSize = True
		self._label8.Location = System.Drawing.Point(8, 30)
		self._label8.Name = "label8"
		self._label8.Size = System.Drawing.Size(204, 13)
		self._label8.TabIndex = 18
		self._label8.Text = "(Leave empty to remove blank directories)"
		# 
		# label7
		# 
		self._label7.AutoSize = True
		self._label7.Location = System.Drawing.Point(8, 10)
		self._label7.Name = "label7"
		self._label7.Size = System.Drawing.Size(181, 13)
		self._label7.TabIndex = 17
		self._label7.Text = "Replace blank directory names with: "
		self._label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
		# 
		# txbEmptyDir
		# 
		self._txbEmptyDir.Location = System.Drawing.Point(226, 15)
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
		self._ckbRemoveEmptyDir.Size = System.Drawing.Size(154, 24)
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
		self._nudVolume.EndInit()
		self._nudNumber.EndInit()
		self._nudCount.EndInit()
		self._nudMonth.EndInit()
		self._nudAltCount.EndInit()
		self._nudAltNumber.EndInit()
		self._tpExcludes.ResumeLayout(False)
		self._ExPanel.ResumeLayout(False)
		self._ExPanel.PerformLayout()
		self._groupBox2.ResumeLayout(False)
		self._tabPage1.ResumeLayout(False)
		self._tabPage1.PerformLayout()
		self.ResumeLayout(False)


	def BtnSepClick(self, sender, e):
		self.InsertText("\\", self._txbDirStruct)

	def ChkFileNamingCheckedChanged(self, sender, e):
		self._txbFileStruct.Enabled = self._ckbFileNaming.Checked
		self._lblFileStruct.Enabled = self._ckbFileNaming.Checked
		#self._label14.Enabled = self._ckbFileNaming.Checked
		#self._sampleTextFile.Enabled = self._ckbFileNaming.Checked
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
		#self._label12.Enabled = self._ckbDirectory.Checked
		#self._sampleTextDir.Enabled = self._ckbDirectory.Checked
		self.UpdateSampleText()

	def InsertItem(self, pre, post, item):
		#TODO: Replace with formatted string
		if self._ckbSpace.Checked:
			post += " "
		string = "{" + pre + "<" + item + ">" + post + "}"

		if self._tabs.SelectedIndex == 0:
			self.InsertText(string, self._txbDirStruct)

		elif self._tabs.SelectedIndex == 1:
			self.InsertText(string, self._txbFileStruct)

	def InsertItemPad(self, pre, post, item, pad):
		if self._ckbSpace.Checked:
			post += " "
		string = "{" + pre + "<" + item + pad + ">" + post + "}"

		if self._tabs.SelectedIndex == 0:
			self.InsertText(string, self._txbDirStruct)
			
		elif self._tabs.SelectedIndex == 1:
			self.InsertText(string, self._txbFileStruct)
	
	def InsertText(self, string, textbox):
		if textbox.Tag[1] > 0:
			s = textbox.Text
			s = s.Remove(textbox.Tag[0], textbox.Tag[1])
			s = s.Insert(textbox.Tag[0], string)
			textbox.Text = s
			textbox.Tag[1] = len(string)
			textbox.Focus()
		else:
			textbox.Text = textbox.Text.Insert(textbox.Tag[0], string)
			textbox.Tag[0] += len(string)
			textbox.Focus()
			
	def UpdateSampleText(self):
		if self._tabs.SelectedIndex == 0:
			if not ExcludePath(self.samplebook, list(self._lbExFolder.Items)) and not self.ExcludeMeta(self.samplebook, self.Excludes, self._cmbExMatchType.SelectedItem) and self._ckbDirectory.Checked:
				self._sampleTextDir.Text = self.PathCreator.CreateDirectoryPath(self.samplebook, self._txbDirStruct.Text, self._txbBaseDir.Text, self._txbEmptyDir.Text, self.settings.EmptyData)
			else:
				self._sampleTextDir.Text = self.samplebook.FileDirectory
		elif self._tabs.SelectedIndex == 1:
			if not ExcludePath(self.samplebook, list(self._lbExFolder.Items)) and not self.ExcludeMeta(self.samplebook, self.Excludes, self._cmbExMatchType.SelectedItem) and self._ckbFileNaming.Checked:
				self._sampleTextFile.Text = self.PathCreator.CreateFileName(self.samplebook, self._txbFileStruct.Text, self.settings.EmptyData, self._cmbImageFormat.SelectedItem)
			else:
				self._sampleTextFile.Text = self.samplebook.FileNameWithExtension

	def BtnBrowseClick(self, sender, e):
		self._FolderBrowser.ShowDialog()
		self._txbBaseDir.Text = self._FolderBrowser.SelectedPath
		self.UpdateSampleText()
		
	def BtnInsertPubClick(self, sender, e):
		self.InsertItem(self._txbPubPre.Text, self._txbPubPost.Text, "publisher")

	def BtnImprintClick(self, sender, e):
		self.InsertItem(self._txbImprintPre.Text, self._txbImprintPost.Text, "imprint")

	def BtnSeriesClick(self, sender, e):
		self.InsertItem(self._txbSeriesPre.Text, self._txbSeriesPost.Text, "series")
	
	def BtnTitleClick(self, sender, e):
		self.InsertItem(self._txbTitlePre.Text, self._txbTitlePost.Text, "title")

	def BtnFormatClick(self, sender, e):
		self.InsertItem(self._txbFormatPre.Text, self._txbFormatPost.Text, "format")

	def BtnVolumeClick(self, sender, e):
		self.InsertItemPad(self._txbVolumePre.Text, self._txbVolumePost.Text, "volume", str(self._nudVolume.Value))

	def BtnAltSeriesClick(self, sender, e):
		self.InsertItem(self._txbAltSeriesPre.Text, self._txbAltSeriesPost.Text, "altseries")

	def BtnNumberClick(self, sender, e):
		self.InsertItemPad(self._txbNumberPre.Text, self._txbNumberPost.Text, "number", str(self._nudNumber.Value))

	def BtnCountClick(self, sender, e):
		self.InsertItemPad(self._txbCountPre.Text, self._txbCountPost.Text, "count", str(self._nudCount.Value))
	
	def BtnMonthNumberClick(self, sender, e):
		self.InsertItemPad(self._txbMonthNumberPre.Text, self._txbMonthNumberPost.Text, "month#", str(self._nudMonth.Value))

	def BtnMonthTextClick(self, sender, e):
		self.InsertItem(self._txbMonthPre.Text, self._txbMonthPost.Text, "month")

	def BtnYearClick(self, sender, e):
		self.InsertItem(self._txbYearPre.Text, self._txbYearPost.Text, "year")

	def BtnAltCountClick(self, sender, e):
		self.InsertItemPad(self._txbAltCountPre.Text, self._txbAltCountPost.Text, "altcount", str(self._nudAltCount.Value))

	def BtnAltNumberClick(self, sender, e):
		self.InsertItemPad(self._txbAltNumberPre.Text, self._txbAltNumberPost.Text, "altnumber", str(self._nudAltNumber.Value))
		
	def BtnStartYearClick(self, sender, e):
		self.InsertItem(self._txbStartYearPre.Text, self._txbStartYearPost.Text, "startyear")

	def TxbFocusLeave(self, sender, e):
		sender.Tag[0] = sender.SelectionStart
		sender.Tag[1] = sender.SelectionLength

	def TxbFocusEnter(self, sender, e):
		if sender.Tag:
			sender.SelectionStart = sender.Tag[0]
			sender.SelectionLength = sender.Tag[1]
			
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
		
		#Reload the selected index of the emptydata cmb
		self.CmbEmptyDataSelectedIndexChanged(self._cmbEmptyData, None)
		
		#Excludes
		
		#TODO: Temporaraly disbabled
		"""
		self._lbExFolder.Items.Clear()
		for i in self.settings.ExcludeFolders:
			self._lbExFolder.Items.Add(i)
		
		#If changing settings to a different one. Delete the existing rule sets.

		self._flpExcludes.Controls.Clear()
		self.Excludes = []

		count = 0
		for i in self.settings.ExcludeMetaData:
			self.CreateRuleSet(None, None)
			self.Excludes[count][1].SelectedItem = i[0]
			self.Excludes[count][2].SelectedItem = i[1]
			self.Excludes[count][3].Text = i[2]
			count += 1
			
		self._cmbExMatchType.SelectedItem = self.settings.ExcludeOperator
		"""
		self._ckbFileless.Checked = self.settings.MoveFileless
		self._cmbImageFormat.SelectedItem = self.settings.FilelessFormat

		self._ckbRemoveEmptyDir.Checked = self.settings.RemoveEmptyDir
		self._lbRemoveEmptyDir.Items.Clear()
		self._lbRemoveEmptyDir.Items.AddRange(System.Array[System.String](self.settings.ExcludedEmptyDir))

		#post &  pre
		self._txbPubPre.Text = self.settings.Pre["Publisher"]
		self._txbPubPost.Text = self.settings.Post["Publisher"]
		self._txbImprintPre.Text = self.settings.Pre["Imprint"]
		self._txbImprintPost.Text = self.settings.Post["Imprint"]
		self._txbSeriesPost.Text = self.settings.Post["Series"]
		self._txbSeriesPre.Text = self.settings.Pre["Series"]
		self._txbTitlePost.Text = self.settings.Post["Title"]
		self._txbTitlePre.Text = self.settings.Pre["Title"]
		self._txbFormatPost.Text = self.settings.Post["Format"]
		self._txbFormatPre.Text = self.settings.Pre["Format"]
		self._txbVolumePost.Text = self.settings.Post["Volume"]
		self._txbVolumePre.Text = self.settings.Pre["Volume"]
		self._txbAltSeriesPost.Text = self.settings.Post["AltSeries"]
		self._txbAltSeriesPre.Text = self.settings.Pre["AltSeries"]
		self._txbAltNumberPost.Text = self.settings.Post["AltNumber"]
		self._txbAltNumberPre.Text = self.settings.Pre["AltNumber"]
		self._txbAltCountPost.Text = self.settings.Post["AltCount"]
		self._txbAltCountPre.Text = self.settings.Pre["AltCount"]
		self._txbYearPost.Text = self.settings.Post["Year"]
		self._txbYearPre.Text = self.settings.Pre["Year"]
		self._txbMonthPost.Text = self.settings.Post["Month"]
		self._txbMonthPre.Text = self.settings.Pre["Month"]
		self._txbMonthNumberPost.Text = self.settings.Post["Month#"]
		self._txbMonthNumberPre.Text = self.settings.Pre["Month#"]
		self._txbCountPost.Text = self.settings.Post["Count"]
		self._txbCountPre.Text = self.settings.Pre["Count"]
		self._txbNumberPost.Text = self.settings.Post["Number"]
		self._txbNumberPre.Text = self.settings.Pre["Number"]
		self._txbStartYearPost.Text = self.settings.Post["StartYear"]
		self._txbStartYearPre.Text = self.settings.Pre["StartYear"]
		
		#Other stuff has to be loaded before text boxes, otherwise they may have the wrong calculated value.
		#Base Textboxes
		self._txbBaseDir.Text = self.settings.BaseDir
		self._txbDirStruct.Text = self.settings.DirTemplate
		self._txbFileStruct.Text = self.settings.FileTemplate
		self._txbEmptyDir.Text = self.settings.EmptyDir
		
		#Set up textbox tags, this is to keep the selected position when clicking the add buttons
		self._txbDirStruct.Tag = [len(self._txbDirStruct.Text),0]
		self._txbFileStruct.Tag = [len(self._txbFileStruct.Text),0]
		
	def SaveSettings(self):
		#Base Textboxes
		self.settings.BaseDir = self._txbBaseDir.Text
		self.settings.DirTemplate = self._txbDirStruct.Text
		self.settings.FileTemplate = self._txbFileStruct.Text
		self.settings.EmptyDir = self._txbEmptyDir.Text
		
		#post &  pre
		self.settings.Pre["Publisher"] = self._txbPubPre.Text
		self.settings.Post["Publisher"] = self._txbPubPost.Text
		self.settings.Pre["Imprint"] = self._txbImprintPre.Text
		self.settings.Post["Imprint"] = self._txbImprintPost.Text
		self.settings.Post["Series"] = self._txbSeriesPost.Text
		self.settings.Pre["Series"] = self._txbSeriesPre.Text
		self.settings.Post["Title"] = self._txbTitlePost.Text
		self.settings.Pre["Title"] = self._txbTitlePre.Text
		self.settings.Post["Format"] = self._txbFormatPost.Text
		self.settings.Pre["Format"] = self._txbFormatPre.Text
		self.settings.Post["Volume"] = self._txbVolumePost.Text
		self.settings.Pre["Volume"] = self._txbVolumePre.Text
		self.settings.Post["AltSeries"] = self._txbAltSeriesPost.Text
		self.settings.Pre["AltSeries"] = self._txbAltSeriesPre.Text
		self.settings.Post["AltNumber"] = self._txbAltNumberPost.Text
		self.settings.Pre["AltNumber"] = self._txbAltNumberPre.Text
		self.settings.Post["AltCount"] = self._txbAltCountPost.Text
		self.settings.Pre["AltCount"] = self._txbAltCountPre.Text
		self.settings.Post["Year"] = self._txbYearPost.Text
		self.settings.Pre["Year"] = self._txbYearPre.Text
		self.settings.Post["Month"] = self._txbMonthPost.Text
		self.settings.Pre["Month"] = self._txbMonthPre.Text
		self.settings.Post["Month#"] = self._txbMonthNumberPost.Text
		self.settings.Pre["Month#"] = self._txbMonthNumberPre.Text
		self.settings.Post["Count"] = self._txbCountPost.Text
		self.settings.Pre["Count"] = self._txbCountPre.Text
		self.settings.Post["Number"] = self._txbNumberPost.Text
		self.settings.Pre["Number"] = self._txbNumberPre.Text
		self.settings.Post["StartYear"] = self._txbStartYearPost.Text
		self.settings.Pre["StartYear"] = self._txbStartYearPre.Text
		
		#Checkboxes
		
		self.settings.UseDirectory = self._ckbDirectory.Checked
		self.settings.UseFileName = self._ckbFileNaming.Checked
		
		#Excludes
		#TODO: Temporaraly disabled
		"""
		self.settings.ExcludeFolders = list(self._lbExFolder.Items)
		
		self.settings.ExcludeMetaData = []
		for i in self.Excludes:
			if i:
				a = []
				a.append(i[1].SelectedItem)
				a.append(i[2].SelectedItem)
				a.append(i[3].Text)
				self.settings.ExcludeMetaData.append(a)
		
		self.settings.ExcludeOperator = self._cmbExMatchType.SelectedItem
		"""
		self.settings.MoveFileless = self._ckbFileless.Checked
		self.settings.FilelessFormat = self._cmbImageFormat.SelectedItem

		self.settings.RemoveEmptyDir = self._ckbRemoveEmptyDir.Checked
		self.settings.ExcludedEmptyDir = list(self._lbRemoveEmptyDir.Items)

	def OkClick(self, sender, e):
		if not self.CheckFields():
			self.DialogResult = DialogResult.None
	
	def CheckFields(self):
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
	
	def CreateRuleSet(self, sender, e):
		#This will be the index of the new set of controls.
		#Don't have to add to the number because len() returns the number of items and a list is a zero based index
		#so the number returned will be the index of the newly created item.
		index = len(self.Excludes)
		
		r = ExcludeRule(self, index)
		
		"""
		
		
		controls = []
		controls.append(System.Windows.Forms.FlowLayoutPanel())
		controls[0].Size = System.Drawing.Size(451, 30)
		controls.append(System.Windows.Forms.ComboBox())
		controls[1].Items.AddRange(System.Array[System.String](
			["Alternate Count",
			"Alternate Number",
			"Alternate Series",
			"Count",			
			"Format",
			"Imprint",
			"Month",
			"Number",
			"Publisher",
			"Rating",
			"Tag",
			"Title",
			"Series",			
			"Year"]))
		controls[1].DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		controls[1].SelectedIndex = 0
		controls[1].Size = System.Drawing.Size(121, 21)
		controls.append(System.Windows.Forms.ComboBox())
		controls[2].Items.AddRange(System.Array[System.String](
			["contains",
			"does not contain",
			"is",
			"is not"]))
		controls[2].DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		controls[2].Size = System.Drawing.Size(110, 21)
		controls[2].SelectedIndex = 0
		controls.append(System.Windows.Forms.TextBox())
		controls[3].Size = System.Drawing.Size(175, 20)
		controls.append(System.Windows.Forms.Button())
		controls[4].Size = System.Drawing.Size(18, 23)
		controls[4].Click += self.RemoveRuleSet
		controls[4].Text = "-"
		controls[4].Tag = index
		controls[0].Controls.Add(controls[1])
		controls[0].Controls.Add(controls[2])
		controls[0].Controls.Add(controls[3])
		controls[0].Controls.Add(controls[4])"""
		self.Excludes.append(r)
		self._flpExcludes.Controls.Add(self.Excludes[-1].Panel)
		self._ExPanel.ScrollControlIntoView(r.Panel)

	def RemoveRuleSet(self, sender, e):
		index = sender.Tag
		
		self._flpExcludes.Controls.Remove(self.Excludes[index].Panel)
		#Don't delete the index of the list as that will screw up other deletion methods
		#Instead make it a null index so that it is skipped when putting it into settings
		self.Excludes[index] = None
		
		
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
		
		
	def ExcludeMeta(self, book, MetaSets, opperator):
		"""The reason the configform has it's own version of this fuction as opposed to using the one in libraryorganizer.py
		is because the pass MetaSets is a list of controls, not a list of strings.
		"""
		count = 0
		for set in MetaSets:
			if set == None:
				continue
			#0 is the field, 1 is the opeartor, 2 is the text to match
			if set[2].SelectedItem == "is":
				#Convert to string just in case
				#Replace the space in the altnerate fields so the it can get the attribute without erroring.
				if str(getattr(book, set[1].SelectedItem.replace(" ", ""))) == set[3].Text:
					count += 1
			elif set[2].SelectedItem == "does not contain":
				if set[3].Text not in str(getattr(book, set[1].SelectedItem.replace(" ", ""))):
					count += 1
			elif set[2].SelectedItem == "contains":
				if set[3].Text in str(getattr(book, set[1].SelectedItem.replace(" ", ""))):
					count += 1
			elif set[2].SelectedItem == "is not":
				if set[3].Text != str(getattr(book, set[1].SelectedItem.replace(" ", ""))):
					count += 1
		
		if opperator == "Any":
			if count > 0:
				return True
			else:
				return False
		
		elif opperator == "All":
			if count == len(MetaSets):
				return True
			else:
				return False
		
		
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



