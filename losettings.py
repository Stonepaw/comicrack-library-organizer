"""
losettings.py

Contains a class for settings

Author: Stonepaw

Version 1.4


Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.
"""

"""
Changes in 1.1:
	empty data changes:
		alternate Number is now : AlternateNumber
		alternate series is now : AlternateSeries
		alternate count is now:  AlternateCount
	added Exclude arrays and variables
	directory cleaning option (was automatic before)
	directories to exclude from cleaning
	fileless export, fileless export format
	
	
	changed from saving with cPickle to Xml, Sava and load functions created.
"""
import System
from System import Convert

import locommon
from locommon import ExcludeRule, ExcludeGroup, Mode

class settings:
	"""
	This class contains all the variables for saving any settings. It should be saved with XML formaly cPickle
	
	Settings are loaded into the config form in the form class
	"""
	def __init__(self):
		
		self.DirTemplate = ""
		self.BaseDir = ""
		self.FileTemplate = ""
		self.Name = ""
		self.EmptyDir = ""

		self.EmptyData = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", 
				"AlternateCount" : "", "StartYear" : "", "Manga" : "", "Characters" : "", "Genre" : "", "Tags" : "", 
				"Teams" : "", "Writer" : "", "SeriesComplete" : ""}
		
		self.Postfix = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", "AlternateCount" : "", 
				"MonthNumber" : "", "StartYear" : "", "Manga" : "", "Characters" : "", "Genre" : "", 
				"Tags" : "", "Teams" : "", "Writer" : "", "SeriesComplete" : ""}

		self.Prefix = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", 
				"AlternateCount" : "", "MonthNumber" : "", "StartYear" : "", "Manga" : "", 
				"Characters" : "", "Genre" : "", "Tags" : "", "Teams" : "", "Writer" : "", "SeriesComplete" : ""}

		self.Seperator = {"Characters" : "", "Genre" : "", "Tags" : "", "Teams" : "", "Writer" : ""}

		self.TextBox = {"Manga" : "", "SeriesComplete" : ""}
				
		self.UseDirectory = True
		
		self.UseFileName = True
		
		self.ExcludeFolders = []
	
		
		self.ExcludeRules = []
	
		self.ExcludeOperator = "Any"
		
		self.RemoveEmptyDir = True
		self.ExcludedEmptyDir = []
		
		self.MoveFileless = False		
		self.FilelessFormat = ".jpg"
		
		self.ExcludeMode = "Do not"
		
		self.Mode = Mode.Move

		self.CopyMode = True

	def Update(self):
		"""
		This is to update old settings from version 1.3 of this script
		"""
		if "AltSeries" in self.Prefix.keys():
			self.Prefix["AlternateSeries"] = self.Prefix["AltSeries"]
			del(self.Prefix["AltSeries"])

		if "AltSeries" in self.Postfix.keys():
			self.Postfix["AlternateSeries"] = self.Postfix["AltSeries"]
			del(self.Postfix["AltSeries"])

		if "AltCount" in self.Prefix.keys():
			self.Prefix["AlternateCount"] = self.Prefix["AltCount"]
			del(self.Prefix["AltCount"])

		if "AltCount" in self.Postfix.keys():
			self.Postfix["AlternateCount"] = self.Postfix["AltCount"]
			del(self.Postfix["AltCount"])

		if "AltNumber" in self.Prefix.keys():
			self.Prefix["AlternateNumber"] = self.Prefix["AltNumber"]
			del(self.Prefix["AltNumber"])

		if "AltNumber" in self.Postfix.keys():
			self.Postfix["AlternateNumber"] = self.Postfix["AltNumber"]
			del(self.Postfix["AltNumber"])

		if "Month#" in self.Prefix.keys():
			self.Prefix["MonthNumber"] = self.Prefix["Month#"]
			del(self.Prefix["Month#"])

		if "Month#" in self.Postfix.keys():
			self.Postfix["MonthNumber"] = self.Postfix["Month#"]
			del(self.Postfix["Month#"])
		
	def Save(self, xwriter):
		"""
		To save this single settings intance to the provided xml file.
		xwriter should be a XmlWriter instance.
		"""
		xwriter.WriteStartElement("Setting")
		xwriter.WriteAttributeString("Name", self.Name)
		xwriter.WriteElementString("DirTemplate", self.DirTemplate)
		xwriter.WriteElementString("BaseDir", self.BaseDir)
		xwriter.WriteElementString("FileTemplate", self.FileTemplate)
		xwriter.WriteElementString("EmptyDir", self.EmptyDir)
		xwriter.WriteElementString("Mode", self.Mode)
		
		xwriter.WriteStartElement("UseFileName")
		xwriter.WriteValue(self.UseFileName)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("UseDirectory")
		xwriter.WriteValue(self.UseDirectory)
		xwriter.WriteEndElement()

		xwriter.WriteStartElement("CopyMode")
		xwriter.WriteValue(self.CopyMode)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("Postfix")
		for i in self.Postfix:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.Postfix[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("Prefix")
		for i in self.Prefix:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.Prefix[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()

		xwriter.WriteStartElement("Seperator")
		for i in self.Seperator:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.Seperator[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()

		xwriter.WriteStartElement("TextBox")
		for i in self.TextBox:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.TextBox[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("EmptyData")
		for i in self.EmptyData:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.EmptyData[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("ExcludeFolders")
		for i in self.ExcludeFolders:
			xwriter.WriteElementString("Item", i)
		xwriter.WriteEndElement()
		print self.ExcludeRules.Count
		xwriter.WriteStartElement("ExcludeRules")
		xwriter.WriteAttributeString("Operator", self.ExcludeOperator)
		xwriter.WriteAttributeString("ExcludeMode", self.ExcludeMode)
		for i in self.ExcludeRules:
			if i:
				i.SaveXml(xwriter)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("MoveFileless")
		xwriter.WriteValue(self.MoveFileless)
		xwriter.WriteEndElement()
		
		xwriter.WriteElementString("FilelessFormat", self.FilelessFormat)
		
		xwriter.WriteStartElement("RemoveEmptyDir")
		xwriter.WriteValue(self.RemoveEmptyDir)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("ExcludedEmptyDir")
		for i in self.ExcludedEmptyDir:
			xwriter.WriteElementString("Item", i)
		xwriter.WriteEndElement()

		xwriter.WriteEndElement()
	
	def Load(self, Xml):
		"""
		Loads the settings instance from the Xml
		Xml should be a XmlNode containing the all the nodes in the Setting node
		"""
		#Text vars
		self.Name = Xml.Attributes["Name"].Value
		self.DirTemplate = Xml.SelectSingleNode("DirTemplate").InnerText
		self.BaseDir = Xml.SelectSingleNode("BaseDir").InnerText
		self.FileTemplate = Xml.SelectSingleNode("FileTemplate").InnerText
		self.EmptyDir = Xml.SelectSingleNode("EmptyDir").InnerText
		
		
		try:
			self.Mode = Xml.SelectSingleNode("Mode").InnerText
		except AttributeError:
			self.Mode = Mode.Move
		
		
		#Legacy, will not be needed in later versions
		op = Xml.SelectSingleNode("ExcludeOperator")
		if op:
			self.ExcludeOperator = op.InnerText
		
		
		
		self.FilelessFormat = Xml.SelectSingleNode("FilelessFormat").InnerText
		
		#Bools
		
		self.UseFileName = Convert.ToBoolean(Xml.SelectSingleNode("UseFileName").InnerText)
		self.UseDirectory = Convert.ToBoolean(Xml.SelectSingleNode("UseDirectory").InnerText)

		#If upgrading it will error
		try:
			self.CopyMode = Convert.ToBoolean(Xml.SelectSingleNode("CopyMode").InnerText)
		except:
			self.CopyMode = True
		
		self.MoveFileless = Convert.ToBoolean(Xml.SelectSingleNode("MoveFileless").InnerText)
		self.RemoveEmptyDir = Convert.ToBoolean(Xml.SelectSingleNode("RemoveEmptyDir").InnerText)
		
		#Dicts
		
		
		iter = Xml.SelectNodes("Prefix/Item")

		#Old loading
		if iter.Count ==  0:
			iter = Xml.SelectNodes("Pre/Item")
		for i in iter:
			self.Prefix[i.Attributes["Name"].Value] = i.Attributes["Value"].Value
			

		iter = Xml.SelectNodes("Postfix/Item")
		if iter.Count == 0:
			iter = Xml.SelectNodes("Post/Item")
		for i in iter:
			self.Postfix[i.Attributes["Name"].Value] = i.Attributes["Value"].Value		


		iter = Xml.SelectNodes("Seperator/Item")
		for i in iter:
			self.Seperator[i.Attributes["Name"].Value] = i.Attributes["Value"].Value

		iter = Xml.SelectNodes("TextBox/Item")
		for i in iter:
			self.TextBox[i.Attributes["Name"].Value] = i.Attributes["Value"].Value
			
		iter = Xml.SelectNodes("EmptyData/Item")
		for i in iter:
			self.EmptyData[i.Attributes["Name"].Value] = i.Attributes["Value"].Value	

		#Arrays
		
		iter = Xml.SelectNodes("ExcludeFolders/Item")
		if iter.Count > 0:
			for i in iter:
				self.ExcludeFolders.append(i.InnerText)
		else:
			self.ExcludeFolders = []
		
		#This is a legacy parser for the old Exclude rules setup. Will remove in later versions.
		iter = Xml.SelectNodes("ExcludeMetaData/Data")
		if iter.Count > 0:
			for i in iter:
				temp = []
				iter2 = i.SelectNodes("Item")
				for i2 in iter2:
					temp.append(i2.InnerText)
				
				t = ExcludeRule()
				t.SetFields(temp[0], temp[1], temp[2])
				self.ExcludeRules.append(t)

		
		#Exclude Rules
		node = Xml.SelectSingleNode("ExcludeRules")
		if node:
			self.ExcludeOperator = node.Attributes["Operator"].Value

			self.ExcludeMode = node.Attributes["ExcludeMode"].Value

			iter = node.ChildNodes
			for i in iter:
				if i.Name == "ExcludeRule":
					r = ExcludeRule()
					r.SetFields(i.Attributes["Field"].Value, i.Attributes["Operator"].Value, i.Attributes["Text"].Value)
					self.ExcludeRules.append(r)
	
				if i.Name == "ExcludeGroup":
					g = ExcludeGroup()
					g.SetOperator(i.Attributes["Operator"].Value)
					self.LoadExcludeRuleGroup(g, i)
					self.ExcludeRules.append(g)


		#Exclued empty dirs
		iter = Xml.SelectNodes("ExcludedEmptyDir/Item")
		if iter.Count > 0:
			for i in iter:
				self.ExcludedEmptyDir.append(i.InnerText)

		self.Update()
		
				
	def LoadExcludeRuleGroup(self, group, GroupNode):
		for i in GroupNode.ChildNodes:
			if i.Name == "ExcludeRule":
				group.CreateRule(None, None, i.Attributes["Field"].Value, i.Attributes["Operator"].Value, i.Attributes["Text"].Value)
			
			if i.Name == "ExcludeGroup":
				g = group.CreateGroup(None, None, i.Attributes["Operator"].Value)
				self.LoadExcludeRuleGroup(g, i)				
	
