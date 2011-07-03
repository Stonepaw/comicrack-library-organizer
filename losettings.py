"""
losettings.py

Contains a class for settings

Author: Stonepaw

Version 1.7


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
		
		self.FolderTemplate = ""
		self.BaseFolder = ""
		self.FileTemplate = ""
		self.Name = ""
		self.EmptyFolder = ""

		self.EmptyData = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", 
				"AlternateCount" : "", "StartYear" : "", "Manga" : "", "Characters" : "", "Genre" : "", "Tags" : "", 
				"Teams" : "", "Writer" : "", "SeriesComplete" : "", "AgeRating" : "", "ScanInformation" : "", "Language" : ""}
		
		self.Postfix = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", "AlternateCount" : "", 
				"MonthNumber" : "", "StartYear" : "", "Manga" : "", "Characters" : "", "Genre" : "", 
				"Tags" : "", "Teams" : "", "Writer" : "", "SeriesComplete" : "", "AgeRating" : "", "ScanInformation" : "", "Language" : ""}

		self.Prefix = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", 
				"AlternateCount" : "", "MonthNumber" : "", "StartYear" : "", "Manga" : "", 
				"Characters" : "", "Genre" : "", "Tags" : "", "Teams" : "", "Writer" : "", "SeriesComplete" : "", "AgeRating" : "", "ScanInformation" : "", "Language" : ""}

		self.Seperator = {"Characters" : "", "Genre" : "", "Tags" : "", "Teams" : "", "Writer" : "", "ScanInformation" : ""}

		self.IllegalCharacters = {"?" : "", "/" : "", "\\" : "", "*" : "", ":" : " - ", "<" : "[", ">" : "]", "|" : "!", "\"" : "'"}

		self.Months = {1 : "January", 2 : "February", 3 : "March", 4 : "April", 5 : "May", 6 : "June", 7 : "July", 8 :"August", 9 : "September", 10 : "October",
						11 : "November", 12 : "December"}

		self.TextBox = {"Manga" : "", "SeriesComplete" : ""}
				
		self.UseFolder = True
		
		self.UseFileName = True
		
		self.ExcludeFolders = []
	
		self.DontAskWhenMultiOne = False
		
		self.ExcludeRules = []
	
		self.ExcludeOperator = "Any"
		
		self.RemoveEmptyFolder = True
		self.ExcludedEmptyFolder = []
		
		self.MoveFileless = False		
		self.FilelessFormat = ".jpg"
		
		self.ExcludeMode = "Do not"
		
		self.Mode = Mode.Move

		self.CopyMode = True

	def Update(self):
		pass
		
	def Save(self, xwriter):
		"""
		To save this single settings intance to the provided xml file.
		xwriter should be a XmlWriter instance.
		"""
		xwriter.WriteStartElement("Setting")
		xwriter.WriteAttributeString("Name", self.Name)
		xwriter.WriteElementString("FolderTemplate", self.FolderTemplate)
		xwriter.WriteElementString("BaseFolder", self.BaseFolder)
		xwriter.WriteElementString("FileTemplate", self.FileTemplate)
		xwriter.WriteElementString("EmptyFolder", self.EmptyFolder)
		xwriter.WriteElementString("Mode", self.Mode)
		
		xwriter.WriteStartElement("UseFileName")
		xwriter.WriteValue(self.UseFileName)
		xwriter.WriteEndElement()

		xwriter.WriteStartElement("DontAskWhenMultiOne")
		xwriter.WriteValue(self.DontAskWhenMultiOne)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("UseFolder")
		xwriter.WriteValue(self.UseFolder)
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

		xwriter.WriteStartElement("Months")
		for i in self.Months:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", str(i))
			xwriter.WriteAttributeString("Value", self.Months[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()

		xwriter.WriteStartElement("IllegalCharacters")
		for i in self.IllegalCharacters:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.IllegalCharacters[i])
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
		
		xwriter.WriteStartElement("RemoveEmptyFolder")
		xwriter.WriteValue(self.RemoveEmptyFolder)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("ExcludedEmptyFolder")
		for i in self.ExcludedEmptyFolder:
			xwriter.WriteElementString("Item", i)
		xwriter.WriteEndElement()

		xwriter.WriteEndElement()
	
	def Load(self, Xml):
		"""
		Loads the settings instance from the Xml
		Xml should be a XmlNode containing a setting node
		"""
		try:
			#Text vars
			self.Name = Xml.Attributes["Name"].Value
		

			#From changes from 1.6 to 1.7
			try:
				self.FolderTemplate = Xml.SelectSingleNode("FolderTemplate").InnerText			
			except AttributeError:
				self.FolderTemplate = Xml.SelectSingleNode("DirTemplate").InnerText

			try:
				self.BaseFolder = Xml.SelectSingleNode("BaseFolder").InnerText
			except AttributeError:
				self.BaseFolder = Xml.SelectSingleNode("BaseDir").InnerText

			try:
				self.EmptyFolder = Xml.SelectSingleNode("EmptyFolder").InnerText
			except AttributeError:
				self.EmptyFolder = Xml.SelectSingleNode("EmptyDir").InnerText
			self.FileTemplate = Xml.SelectSingleNode("FileTemplate").InnerText
		
		
		
			try:
				self.Mode = Xml.SelectSingleNode("Mode").InnerText
			except AttributeError:
				self.Mode = Mode.Move
		

			self.FilelessFormat = Xml.SelectSingleNode("FilelessFormat").InnerText
		
			#Bools
		
			self.UseFileName = Convert.ToBoolean(Xml.SelectSingleNode("UseFileName").InnerText)

			#From 1.6 to 1.7 of the script
			try:
				self.UseFolder = Convert.ToBoolean(Xml.SelectSingleNode("UseFolder").InnerText)
			except AttributeError:
				self.UseFolder = Convert.ToBoolean(Xml.SelectSingleNode("UseDirectory").InnerText)


			self.CopyMode = Convert.ToBoolean(Xml.SelectSingleNode("CopyMode").InnerText)


			self.DontAskWhenMultiOne = Convert.ToBoolean(Xml.SelectSingleNode("DontAskWhenMultiOne").InnerText)
		
			self.MoveFileless = Convert.ToBoolean(Xml.SelectSingleNode("MoveFileless").InnerText)

			try:
				self.RemoveEmptyFolder = Convert.ToBoolean(Xml.SelectSingleNode("RemoveEmptyFolder").InnerText)
			except AttributeError:
				self.RemoveEmptyFolder = Convert.ToBoolean(Xml.SelectSingleNode("RemoveEmptyDir").InnerText)
		
			#Dicts
		
		
			iter = Xml.SelectNodes("Prefix/Item")

			for i in iter:
				self.Prefix[i.Attributes["Name"].Value] = i.Attributes["Value"].Value
			

			iter = Xml.SelectNodes("Postfix/Item")

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

			#added in 1.7
			iter = Xml.SelectNodes("Months/Item")
			if iter.Count > 0:
				for i in iter:
					self.Months[int(i.Attributes["Name"].Value)] = i.Attributes["Value"].Value

			iter = Xml.SelectNodes("IllegalCharacters/Item")
			if iter.Count > 0:
				for i in iter:
					self.IllegalCharacters[i.Attributes["Name"].Value] = i.Attributes["Value"].Value

			#Arrays
		
			iter = Xml.SelectNodes("ExcludeFolders/Item")
			if iter.Count > 0:
				for i in iter:
					self.ExcludeFolders.append(i.InnerText)
			else:
				self.ExcludeFolders = []
		
	
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
			try:
				iter = Xml.SelectNodes("ExcludedEmptyFolder/Item")
			except AttributeError:
				iter = Xml.SelectNodes("ExcludedEmptyDir/Item")
			if iter.Count > 0:
				for i in iter:
					self.ExcludedEmptyFolder.append(i.InnerText)

			self.Update()

		except Exception, ex:
			print ex
			return False
		
				
	def LoadExcludeRuleGroup(self, group, GroupNode):
		for i in GroupNode.ChildNodes:
			if i.Name == "ExcludeRule":
				group.CreateRule(None, None, i.Attributes["Field"].Value, i.Attributes["Operator"].Value, i.Attributes["Text"].Value)
			
			if i.Name == "ExcludeGroup":
				g = group.CreateGroup(None, None, i.Attributes["Operator"].Value)
				self.LoadExcludeRuleGroup(g, i)				
	
