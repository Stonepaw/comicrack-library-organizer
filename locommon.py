"""
locommon.py

Author: Stonepaw

Version: 1.6
			

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
			"Count",
			"File Name",
			"File Path",
			"File Format",
			"Format",
			"Imprint",
			"Month",
			"Number",
			"Notes",
			"Publisher",
			"Rating",
			"Tags",
			"Title",
			"Scan Information",
			"Series",
			"Volume",			
			"Year"]))
			
		self.Field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self.Field.SelectedIndex = 0
		self.Field.Size = Size(121, 21)
		
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
		
		#Add controls
		self.Panel.Controls.Add(self.Field)
		self.Panel.Controls.Add(self.Operator)
		self.Panel.Controls.Add(self.TextBox)
		self.Panel.Controls.Add(self.Remove)
		
	def GetField(self):
		return self.Field.SelectedItem
	
	def GetText(self):
		return self.TextBox.Text
	
	def GetOperator(self):
		return self.Operator.SelectedItem
	
	def SetFields(self, Field, Operator, Text):
		self.Field.SelectedItem = Field
		self.Operator.SelectedItem = Operator
		self.TextBox.Text = Text
		
	def SetField(self, Field):
		self.Field.SelectedItem = Field
		
	def SetOperator(self, Operator):
		self.Operator.SelectedItem = Operator
		
	def SetText(self, Text):
		self.TextBox.Text = Text
		
	def Calculate(self, book):
		operator = self.GetOperator()
		text = self.GetText()
		field = self.GetField()
		
		if operator == "is":
			#Convert to string just in case
			#Replace a space with nothing in the case of Alternate fields
			if unicode(getattr(book, field.replace(" ", ""))) == text:
				return 1
			else:
				return 0
		elif operator == "does not contain":
			if text not in unicode(getattr(book, field.replace(" ", ""))):
				return 1
			else:
				return 0
		elif operator == "contains":
			if text in unicode(getattr(book, field.replace(" ", ""))):
				return 1
			else:
				return 0
		elif operator == "is not":
			if text != unicode(getattr(book, field.replace(" ", ""))):
				return 1
			else:
				return 0
		elif operator == "greater than":
			#Try to use the int value to compare if possible
			try:
				number = int(text)
				if number < int(getattr(book, field.replace(" ", ""))):
					return 1
				else:
					return 0
			except ValueError:
				if text < unicode(getattr(book, field.replace(" ", ""))):
					return 1
				else:
					return 0
		elif operator == "less than":
			try:
				number = int(text)
				if number > int(getattr(book, field.replace(" ", ""))):
					return 1
				else:
					return 0
			except ValueError:
				if text > unicode(getattr(book, field.replace(" ", ""))):
					return 1
				else:
					return 0
		

	def SaveXml(self, xmlwriter):
		xmlwriter.WriteStartElement("ExcludeRule")
		xmlwriter.WriteAttributeString("Field", self.GetField())
		xmlwriter.WriteAttributeString("Operator", self.GetOperator())
		xmlwriter.WriteAttributeString("Text", self.GetText())
		xmlwriter.WriteEndElement()

def SaveDict(dict, file):
	"""
	Saves a dict of strings to a file
	"""
	try:
		w = open(file, 'w')
		for i in dict:
			w.write(i + "|" + dict[i] + "\n")
		w.close()
	except IOError, err:
		print "Somthing went wrong saving the undo list"
		print err

def LoadDict(file):
	dict = {}
	try:
		r = open(file, "r")
		
		for i in r.readlines():
			parts = i.split("|")
			dict[parts[0]] = parts[1].strip()
		
		r.close()

	except IOError, err:
		dict = None
		print "Error loading dict from file " + file
		print err
	return dict