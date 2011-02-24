"""
losettings.py

Contains a class for settings

Author: Stonepaw

Version 1.1

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

class settings:
	"""
	This class contains all the variables for saving any settings. It should be saved with XML formaly cPickle
	
	Settings are loaded into the config form in the form class
	"""
	def __init__(self):
		self.ExcludeMetaData = []
		
		self.DirTemplate = ""
		self.BaseDir = ""
		self.FileTemplate = ""
		self.Name = ""
		self.EmptyDir = ""

		self.EmptyData = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AlternateSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AlternateNumber" : "", "Count" : "", "Month" : "", "Year" : "", "AlternateCount" : "", "StartYear" : ""}
		
		self.Post = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AltSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AltNumber" : "", "Count" : "", "Month" : "", "Year" : "", "AltCount" : "", "Month#" : "", "StartYear" : ""}

		self.Pre = {"Publisher" : "", "Imprint" : "", "Series" : "", "Title" : "", 
				"AltSeries" : "", "Format" : "", "Volume" : "", "Number" : "", 
				"AltNumber" : "", "Count" : "", "Month" : "", "Year" : "", "AltCount" : "", "Month#" : "", "StartYear" : ""}
				
		self.UseDirectory = True
		
		self.UseFileName = True
		
		self.ExcludeFolders = []
	
		self.ExcludeMetaData = []
	
		self.ExcludeOperator = "Any"
		
		self.RemoveEmptyDir = True
		self.ExcludedEmptyDir = []
		
		self.MoveFileless = False		
		self.FilelessFormat = ".jpg"

	def Update(self):
		"""
		This is to update old settings from version 1.0 of this script
		"""
		if "Alternate Series" in self.EmptyData:
			self.EmptyData["AlternateSeries"] = self.EmptyData["Alternate Series"]
			del(self.EmptyData["Alternate Series"])
		
		if "Alternate Count" in self.EmptyData:
			self.EmptyData["AlternateCount"] = self.EmptyData["Alternate Count"]
			del(self.EmptyData["Alternate Count"])
		
		if "Alternate Number" in self.EmptyData:
			self.EmptyData["AlternateNumber"] = self.EmptyData["Alternate Number"]
			del(self.EmptyData["Alternate Number"])
		
		self.EmptyData["StartYear"] = ""		
		self.Post["StartYear"] = ""
		self.Pre["StartYear"] = ""
		
		self.ExcludeMetaData = []
		
		self.ExcludeOperator = "Any"
		self.ExcludeFolders = []
		self.Name = "Default"
		
		self.RemoveEmptyDir = True
		self.ExcludedEmptyDir = []
		
		self.MoveFileless = False
		self.FilelessFormat = ".jpg"
		
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
		xwriter.WriteElementString("ExcludeOperator", self.ExcludeOperator)

		xwriter.WriteStartElement("UseFileName")
		xwriter.WriteValue(self.UseFileName)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("UseDirectory")
		xwriter.WriteValue(self.UseDirectory)
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("Post")
		for i in self.Post:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.Post[i])
			xwriter.WriteEndElement()
		xwriter.WriteEndElement()
		
		xwriter.WriteStartElement("Pre")
		for i in self.Pre:
			xwriter.WriteStartElement("Item")
			xwriter.WriteAttributeString("Name", i)
			xwriter.WriteAttributeString("Value", self.Pre[i])
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
		print self.ExcludeMetaData.Count
		xwriter.WriteStartElement("ExcludeMetaData")
		for i in self.ExcludeMetaData:
			xwriter.WriteStartElement("Data")
			for ii in i:
				xwriter.WriteElementString("Item", ii)
			xwriter.WriteEndElement()
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
		self.ExcludeOperator = Xml.SelectSingleNode("ExcludeOperator").InnerText
		self.FilelessFormat = Xml.SelectSingleNode("FilelessFormat").InnerText
		
		#Bools
		
		self.UseFileName = Convert.ToBoolean(Xml.SelectSingleNode("UseFileName").InnerText)
		self.UseDirectory = Convert.ToBoolean(Xml.SelectSingleNode("UseDirectory").InnerText)
		
		self.MoveFileless = Convert.ToBoolean(Xml.SelectSingleNode("MoveFileless").InnerText)
		self.RemoveEmptyDir = Convert.ToBoolean(Xml.SelectSingleNode("RemoveEmptyDir").InnerText)
		
		#Dicts
		
		iter = Xml.SelectNodes("Pre/Item")
		for i in iter:
			self.Pre[i.Attributes["Name"].Value] = i.Attributes["Value"].Value
			
		iter = Xml.SelectNodes("Post/Item")
		for i in iter:
			self.Post[i.Attributes["Name"].Value] = i.Attributes["Value"].Value		

		iter = Xml.SelectNodes("Post/Item")
		for i in iter:
			self.Post[i.Attributes["Name"].Value] = i.Attributes["Value"].Value
			
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
			
		iter = Xml.SelectNodes("ExcludeMetaData/Data")
		if iter.Count > 0:
			for i in iter:
				temp = []
				iter2 = i.SelectNodes("Item")
				for i2 in iter2:
					temp.append(i2.InnerText)
				
				self.ExcludeMetaData.append(temp)
				
		iter = Xml.SelectNodes("ExcludedEmptyDir/Item")
		if iter.Count > 0:
			for i in iter:
				self.ExcludedEmptyDir.append(i.InnerText)
		
				
		