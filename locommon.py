"""
locommon.py

Author: Stonepaw

Version: 1.7.5
			-Added ReadPercentage to Rules

Contains several classes and functions. All are used in several files


Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.

"""




import clr

import System

clr.AddReference("System.Drawing")
from System.Drawing import Size, Point

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import TextBox, Button, ComboBox, FlowLayoutPanel, Panel, Label, ComboBoxStyle


from System.IO import Path, FileInfo

SCRIPTDIRECTORY = FileInfo(__file__).DirectoryName

SETTINGSFILE = Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat")

OLDSETTINGSFILE = Path.Combine(SCRIPTDIRECTORY, "losettings.dat")

ICON = Path.Combine(SCRIPTDIRECTORY, "libraryorganizer.ico")

UNDOFILE = Path.Combine(SCRIPTDIRECTORY, "undo.dat")

clr.AddReferenceByPartialName('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

startbooks = {}

class Mode:
	Move = "Move"
	Copy = "Copy"
	Test = "Test"

class CopyMode:
	AddToLibrary = True
	DoNotAdd = False
				
class ExcludeGroup(object):
	"""This class is the object that contains a rule group"
	
	"""
	
	def __init__(self):
		#rules is the rules this rule group contains. it can contain either rules or more rule groups
		self.rules = []
		
		self.Panel = Panel()
		self.Panel.Size = Size(451, 70)
		self.Panel.MinimumSize = Size(451, 70)
		self.Panel.AutoSize = True
		
		self.Text1 = Label()
		self.Text1.Text = "Match"
		self.Text1.Size = Size(40, 23)
		self.Text1.Location = Point(3, 2)
		self.Text1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		
		self.Operator = ComboBox()
		self.Operator.DropDownStyle = ComboBoxStyle.DropDownList
		self.Operator.Location = Point(44, 2)
		self.Operator.Items.AddRange(System.Array[System.Object](
			["Any",
			"All"]))
		self.Operator.Size = Size(46, 23)
		self.Operator.SelectedIndex = 0
		
		self.Text2 = Label()
		self.Text2.Location = Point(96, 4)
		self.Text2.Text = "of the following rules"
		self.Text2.Size = Size(120, 20)
		self.Text2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		
		
		self.AddRule = Button()
		self.AddRule.Size = Size(75, 23)
		self.AddRule.Location = Point(346, 2)
		self.AddRule.Text = "Add Rule"
		self.AddRule.Click += self.CreateRule

		self.AddGroup = Button()
		self.AddGroup.Location = Point(265, 2)
		self.AddGroup.Size = Size(75, 23)
		self.AddGroup.Text = "Add Group"
		self.AddGroup.Click += self.CreateGroup
		
		self.Remove = Button()
		self.Remove.Text = "-"
		self.Remove.Location = Point(427, 2)
		self.Remove.Size = Size(18, 23)
		self.Remove.Tag = self
	
		self.RulesContainer = FlowLayoutPanel()
		self.RulesContainer.Size = Size(458, 30)
		self.RulesContainer.MinimumSize = Size(458, 30)
		self.RulesContainer.AutoSize = True
		self.RulesContainer.Location = Point(15, 32)
		self.RulesContainer.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
		
		self.Panel.Controls.Add(self.Text1)
		self.Panel.Controls.Add(self.Text2)
		self.Panel.Controls.Add(self.Operator)
		self.Panel.Controls.Add(self.AddGroup)
		self.Panel.Controls.Add(self.AddRule)
		self.Panel.Controls.Add(self.Remove)
		self.Panel.Controls.Add(self.RulesContainer)
		
	def CreateRule(self, sender, e, Field = None, Operator = None, Text = None):
		r = ExcludeRule()
		self.rules.append(r)
		self.RulesContainer.Controls.Add(r.Panel)
		r.Remove.Click += self.RemoveRule
		
		if Field:
			r.SetField(Field)
		if Operator:
			r.SetOperator(Operator)
		if Text:
			r.SetText(Text)
			
		return r
		
	def RemoveRule(self, sender, e):
		
		index = self.rules.index(sender.Tag)
		
		self.RulesContainer.Controls.Remove(self.rules[index].Panel)
		del(self.rules[index])


	def CreateGroup(self, sender, e, Operator = None):
		g = ExcludeGroup()
		self.rules.append(g)
		g.Remove.Click += self.RemoveRule
		self.RulesContainer.Controls.Add(g.Panel)
		
		if Operator:
			g.SetOperator(Operator)
		return g
		
	def SetOperator(self, operator):
		if operator in ("Any", "All"):
			self.Operator.SelectedItem = operator
		
	def Calculate(self, book):
		
		#Keeps track of the amout of rules the book fell under
		count = 0
		
		#Keep track of the total amount of rules
		total = 0
		
		for i in self.rules:
			
			#Empty rule. Do not count towards total
			if i == None:
				continue
			
			r = i.Calculate(book)
			
			#Something went wrong, possible empty group. Thus we don't count that rule
			if r == None:
				continue
		
			count += r
			total += 1

		if total == 0:
			return None
		
		if self.Operator.SelectedItem == "Any":
			if count > 0:
				return 1
			else:
				return 0
		else:
			if count == total:
				return 1
			else:
				return 0

	def SaveXml(self, xmlwriter):
		xmlwriter.WriteStartElement("ExcludeGroup")
		xmlwriter.WriteAttributeString("Operator", self.Operator.SelectedItem)
		for i in self.rules:
			if i:
				i.SaveXml(xmlwriter)
		xmlwriter.WriteEndElement()
			
class ExcludeRule(object):
	"""This class is the object of a rule. It can either be in a rule group or by itself"""
	
	def __init__(self):
		
		#Flow Layout Panel
		self.Panel = FlowLayoutPanel()
		self.Panel.Size = Size(451, 30)
		
		#Field selector
		self.Field = ComboBox()
		self.Field.Items.AddRange(System.Array[System.String](
			["Age Rating",
			"Alternate Count",
			"Alternate Number",
			"Alternate Series",
			"Black And White",
			"Characters",
			"Count",
			"File Name",
			"File Path",
			"File Format",
			"Format",
			"Genre",	
			"Imprint",
			"Language",
			"Locations",
			"Manga",
			"Month",
			"Number",
			"Notes",
			"Publisher",
			"Rating",
			"Read Percentage",
			"Series Complete",
			"Tags",
			"Teams",
			"Title",
			"Scan Information",
			"Series",
			"Start Month",
			"Start Year",
			"Volume",
			"Web",		
			"Year"]))
			
		self.Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self.Field.Size = Size(121, 21)
		self.Field.MaxDropDownItems = 15
		self.Field.IntegralHeight = False
		self.Field.Sorted = True
		self.Field.SelectedIndex= 0
		
		#Operator selector
		self.Operator = ComboBox()
		self.Operator.Items.AddRange(System.Array[System.String](
			["contains",
			"does not contain",
			"greater than",
			"less than",
			"is",
			"is not"]))
		self.Operator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self.Operator.Size = Size(110, 21)
		self.Operator.SelectedIndex = 0
		
		#Textbox
		self.TextBox = TextBox()
		self.TextBox.Size = Size(175, 20)
		
		#Remove Button
		self.Remove = Button()
		self.Remove.Size = Size(18, 23)
		self.Remove.Text = "-"
		self.Remove.Tag = self

		self.Selection = ComboBox()
		self.Selection.Size = Size(175, 21)
		self.Selection.Enabled = False
		self.Selection.Visible = False
		self.Selection.DropDownStyle = ComboBoxStyle.DropDownList

		self.Field.SelectedIndexChanged += self.FieldSelectionIndexChanged
		
		#Add controls
		self.Panel.Controls.Add(self.Field)
		self.Panel.Controls.Add(self.Operator)
		self.Panel.Controls.Add(self.TextBox)
		self.Panel.Controls.Add(self.Selection)
		self.Panel.Controls.Add(self.Remove)
		
	def GetField(self):
		f = self.Field.SelectedItem

		if f in ["Series", "Count", "Format", "Number", "Title", "Volume", "Year"]:
			return "Shadow" + f

		if f == "Language":
			return "LanguageAsText"
		return f
	
	def GetText(self):
		return self.TextBox.Text

	def GetYesNo(self):
		if self.Field.SelectedItem == "Manga":
			if self.Selection.SelectedItem == "Yes (Right to Left)":
				return MangaYesNo.YesAndRightToLeft
			else:
				return getattr(MangaYesNo, self.Selection.SelectedItem)

		return getattr(YesNo, self.Selection.SelectedItem)
	
	def GetOperator(self):
		return self.Operator.SelectedItem
	
	def SetFields(self, Field, Operator, Text):
		self.Field.SelectedItem = Field
		self.Operator.SelectedItem = Operator
		if Field in ["Manga", "Series Complete", "Black And White"]:
			self.Selection.SelectedItem = Text
		else:
			self.TextBox.Text = Text
		
	def SetField(self, Field):
		self.Field.SelectedItem = Field
		
	def SetOperator(self, Operator):
		self.Operator.SelectedItem = Operator

	def FieldSelectionIndexChanged(self, sender, e):

		def ShowSelection():
			self.Operator.Items.Clear()
			self.Operator.Items.AddRange(System.Array[System.String](["is", "is not"]))
			self.Operator.SelectedIndex = 0
			self.TextBox.Enabled = False
			self.TextBox.Visible = False
			self.Selection.Enabled = True
			self.Selection.Visible = True

		if sender.SelectedItem in ["Series Complete", "Black And White"]:
			self.Selection.Items.Clear()
			self.Selection.Items.AddRange(System.Array[System.String](["Yes", "No", "Unknown"]))
			self.Selection.SelectedIndex = 0
			ShowSelection()

		elif sender.SelectedItem == "Manga":
			self.Selection.Items.Clear()
			self.Selection.Items.AddRange(System.Array[System.String](["Yes", "Yes (Right to Left)", "No", "Unknown"]))
			self.Selection.SelectedIndex = 0
			ShowSelection()

		else:
			if self.TextBox.Enabled == False:
				if self.Operator.Items.Count != 6:
					self.Operator.Items.Clear()
					self.Operator.Items.AddRange(System.Array[System.String](["contains", "does not contain", "greater than", "less than", "is", "is not"]))
					self.Operator.SelectedIndex = 0
				self.TextBox.Enabled = True
				self.TextBox.Visible = True
				self.Selection.Enabled = False
				self.Selection.Visible = False
	
	def SetText(self, Text):
		self.TextBox.Text = Text
		
	def Calculate(self, book):
		if self.Selection.Enabled:
			return self.CalculateYesNo(book)
			
		field = self.GetField()
		field = field.replace(" ", "")

		if field in ["StartYear", "StartMonth"]:
			return self.CalculateStart(book, field)

		text = self.GetText()
		operator = self.GetOperator()

		if operator == "is":
			#Convert to string just in case
			if unicode(getattr(book, field)) == text:
				return 1
			else:
				return 0
		elif operator == "does not contain":
			if text not in unicode(getattr(book, field)):
				return 1
			else:
				return 0
		elif operator == "contains":
			if text in unicode(getattr(book, field)):
				return 1
			else:
				return 0
		elif operator == "is not":
			if text != unicode(getattr(book, field)):
				return 1
			else:
				return 0
		elif operator == "greater than":
			#Try to use the int value to compare if possible
			try:
				number = int(text)
				if number < int(getattr(book, field)):
					return 1
				else:
					return 0
			except ValueError:
				if text < unicode(getattr(book, field)):
					return 1
				else:
					return 0
		elif operator == "less than":
			try:
				number = int(text)
				if number > int(getattr(book, field)):
					return 1
				else:
					return 0
			except ValueError:
				if text > unicode(getattr(book, field)):
					return 1
				else:
					return 0
		
	def CalculateYesNo(self, book):
		operator = self.GetOperator()
		yesno = self.GetYesNo()
		field = self.GetField()
		field = field.replace(" ", "")
		
		if operator == "is":
			if (getattr(book, field)) == yesno:
				return 1
			else:
				return 0
		elif operator == "is not":
			if yesno != (getattr(book, field)):
				return 1
			else:
				return 0

	def CalculateStart(self, book, field):

		text = self.GetText()
		operator = self.GetOperator()

		startbook = GetEarliestBook(book)
		
		if field == "StartMonth":
			fieldtext = unicode(startbook.Month)

		else:
			fieldtext = unicode(startbook.ShadowYear)

		if operator == "is":
			#Convert to string just in case
			if fieldtext == text:
				return 1
			else:
				return 0
		elif operator == "does not contain":
			if text not in fieldtext:
				return 1
			else:
				return 0
		elif operator == "contains":
			if text in fieldtext:
				return 1
			else:
				return 0
		elif operator == "is not":
			if text != fieldtext:
				return 1
			else:
				return 0
		elif operator == "greater than":
			#Try to use the int value to compare if possible
			try:
				number = int(text)
				if number < int(fieldtext):
					return 1
				else:
					return 0
			except ValueError:
				if text < fieldtext:
					return 1
				else:
					return 0
		elif operator == "less than":
			try:
				number = int(text)
				if number > int(fieldtext):
					return 1
				else:
					return 0
			except ValueError:
				if text > fieldtext:
					return 1
				else:
					return 0

	def SaveXml(self, xmlwriter):
		xmlwriter.WriteStartElement("ExcludeRule")
		xmlwriter.WriteAttributeString("Field", self.Field.SelectedItem)
		xmlwriter.WriteAttributeString("Operator", self.GetOperator())
		if self.TextBox.Enabled == True:
			xmlwriter.WriteAttributeString("Text", self.GetText())
		else:
			xmlwriter.WriteAttributeString("Text", self.Selection.SelectedItem)
		xmlwriter.WriteEndElement()

def SaveDict(dict, file):
	"""
	Saves a dict of strings to a file
	"""
	try:

		w = open(file, 'w')
		for i in dict:
			w.write(i.encode("utf8") + "|" + dict[i].encode("utf8") + "\n")
		w.close()
	except IOError, err:
		print "Somthing went wrong saving the undo list"
		print err

def LoadDict(file):
	dict = {}
	try:
		r = open(file, 'r')
		
		for i in r:
			i = i.decode('utf8')
			parts = i.split("|")
			dict[parts[0]] = parts[1].strip()
		
		r.close()

	except IOError, err:
		dict = None
		print "Error loading dict from file " + file
		print err
	return dict

def GetEarliestBook(book):
	"""
	Finds the first published issue of a series in the library
	Returns a ComicBook object
	"""
	#Find the Earliest by going through the whole list of comics in the library find the earliest year field and month field of the same series and volume
		
	index = book.Publisher+book.ShadowSeries+str(book.ShadowVolume)
		
	if startbooks.has_key(index):
		startbook = startbooks[index]
	else:
		startbook = book
			
		for b in ComicRack.App.GetLibraryBooks():
			if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume and b.Publisher == book.Publisher:
					
				#Notes:
				#Year can be empty (-1)
				#Month can be empty (-1)

				#In case the initial value is bad
				if startbook.ShadowYear == -1 and b.ShadowYear != 1:
					startbook = b
					
				#Check if the current book's year 
				if b.ShadowYear != -1 and b.ShadowYear < startbook.ShadowYear:
					startbook = b

				#Check if year the same and a valid month
				if b.ShadowYear == startbook.ShadowYear and b.Month != -1:

					#Current book has empty month
					if startbook.Month == -1:
						startbook = b
						
					#Month is earlier
					elif b.Month < startbook.Month:
						startbook = b
			
		#Store this final result in the dict so no calculation require for others of the series.
		startbooks[index] = startbook

	return startbook