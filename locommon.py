"""
locommon.py

Author: Stonepaw

Version 1.1
    Support for unicode characters

Copyright 2010-2012 Stonepaw

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

"""




import clr

import System

clr.AddReference("System.Drawing")
from System.Drawing import Size, Point

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import TextBox, Button, ComboBox, FlowLayoutPanel, Panel, Label, ComboBoxStyle


from System.IO import Path, FileInfo

SCRIPTDIRECTORY = FileInfo(__file__).DirectoryName

PROFILEFILE = Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat")

ICON = Path.Combine(SCRIPTDIRECTORY, "libraryorganizer.ico")

UNDOFILE = Path.Combine(SCRIPTDIRECTORY, "undo.dat")

VERSION = 2.1

clr.AddReferenceByPartialName('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

startbooks = {}
endbooks = {}

name_to_field = {"Added Date" : "AddedDate",
                 "Age Rating" : "AgeRating", 
                 "Alternate Count" : "AlternateCount", 
                 "Alternate Number" : "AlternateNumber",
                 "Alternate Series" : "AlternateSeries", 
                 "Black And White" : "BlackAndWhite", 
                 "Characters" : "Characters", 
                 "Colorist" : "Colorist",
                 "Counter" : "Counter", 
                 "Count" : "ShadowCount", 
                 "Cover Artist" : "CoverArtist", 
                 "Day" : "Day",
                 "Editor" : "Editor", 
                 "End Year" : "EndYear",
                 "End Month" : "EndMonth",
                 "File Format" : "FileFormat", 
                 "File Name" : "FileName", 
                 "File Path" : "FilePath", 
                 "First Letter" : "FirstLetter",
                 "Format" : "ShadowFormat", 
                 "Genre" : "Genre", 
                 "Imprint" : "Imprint", 
                 "Inker" : "Inker", 
                 "Language" : "LanguageISO", 
                 "Letterer" : "Letterer", 
                 "Locations" : "Locations",
                 "Main Character Or Team": "MainCharacterOrTeam", 
                 "Manga" : "Manga", 
                 "Month" : "Month", 
                 "Notes" : "Notes", 
                 "Number" : "ShadowNumber", 
                 "Penciller" : "Penciller", 
                 "Publisher" : "Publisher", 
                 "Rating" : "Rating", 
                 "Read Percentage" : "ReadPercentage",
                 "Released Date" : "Released Date", 
                 "Review" : "Review", 
                 "Scan Information" : "ScanInformation", 
                 "Series" : "ShadowSeries",
                 "Series Complete" : "SeriesComplete", 
                 "Series Group" : "SeriesGroup", 
                 "Start Month" : "StartMonth", 
                 "Start Year" : "StartYear", 
                 "Story Arc": "StoryArc", 
                 "Tags" : "Tags",
                 "Teams" : "Teams", 
                 "Title" : "ShadowTitle", 
                 "Volume" : "ShadowVolume", 
                 "Web" : "Web", 
                 "Writer" : "Writer", 
                 "Year" : "ShadowYear"}

field_to_name = {"AddedDate" : "Added Date",
                 "AgeRating" : "Age Rating", 
                 "AlternateCount" : "Alternate Count", 
                 "AlternateNumber" : "Alternate Number", 
                 "AlternateSeries" : "Alternate Series", 
                 "BlackAndWhite" : "Black And White", 
                 "Characters" : "Characters", 
                 "Colorist" : "Colorist",
                 "Counter" : "Counter", 
                 "CoverArtist" : "Cover Artist",
                 "Day" : "Day",
                 "Editor" : "Editor", 
                 "EndMonth" : "End Month",
                 "EndYear" : "End Year",
                 "FileFormat" : "File Format", 
                 "FileName" : "File Name", 
                 "FilePath" : "File Path", 
                 "FirstLetter" : "First Letter", 
                 "Genre" : "Genre", 
                 "Imprint" : "Imprint", 
                 "Inker" : "Inker",
                 "LanguageISO" : "Language", 
                 "Letterer" : "Letterer", 
                 "Locations" : "Locations", 
                 "MainCharacterOrTeam": "Main Character Or Team", 
                 "Manga" : "Manga", 
                 "Month" : "Month", 
                 "Notes" : "Notes", 
                 "Penciller" : "Penciller",
                 "Publisher" : "Publisher", 
                 "Rating" : "Rating", 
                 "ReadPercentage" : "Read Percentage", 
                 "Read" : "Read Percentage", 
                 "ReleasedDate" : "Released Date",
                 "Review" : "Review", 
                 "ScanInformation" : "Scan Information", 
                 "SeriesComplete" : "Series Complete", 
                 "ShadowCount" : "Count", 
                 "ShadowFormat" : "Format", 
                 "ShadowNumber" : "Number", 
                 "ShadowSeries" : "Series", 
                 "SeriesGroup" : "Series Group", 
                 "StoryArc": "Story Arc", 
                 "ShadowTitle" : "Title", 
                 "ShadowVolume" : "Volume", 
                 "ShadowYear" : "Year", 
                 "StartMonth" : "Start Month", 
                 "StartYear" : "Start Year", 
                 "Tags" : "Tags", 
                 "Teams" : "Teams", 
                 "Web" : "Web", 
                 "Writer" : "Writer"}


                                                               

class Mode(object):
    Move = "Move"
    Copy = "Copy"
    Simulate = "Simulate"



class CopyMode(object):
    AddToLibrary = True
    DoNotAdd = False


                
class ExcludeGroup(object):
    """
    Contains a list of rules or rule groups and can calculate if a book should be moved under it's rules.
    """
    
    def __init__(self, operator, rules = None):

        if rules is None:
            self.rules = []
        else:
            self.rules = rules
        
        self.operator = operator
        

    def add_rule(self, rule):
        """Adds a single rule to this group's list of rules."""
        self.rules.append(rule)


    def add_rules(self, rules):
        """Adds a list of rules to this group's list of rules."""
        self.rules.extend(rules)

        
    def book_should_be_moved(self, book):
        """
        Checks if the book should be moved under the rules in the rule group.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """
        
        #Keeps track of the amount of rules the book fell under
        count = 0
        
        #Keep track of the total amount of rules
        total = 0
        
        for rule in self.rules:
            

            result = rule.book_should_be_moved(book)
            
            #Something went wrong, possible empty group. Thus we don't count that rule
            if result is None:
                continue
        
            count += result
            total += 1

        if total == 0:
            return None
        
        if self.operator == "Any":
            if count > 0:
                return 1
            else:
                return 0
        else:
            if count == total:
                return 1
            else:
                return 0


    def save_xml(self, xmlwriter):
        """Saves this rule group and its containing rules to an xml file using the specified xmlwriter."""
        xmlwriter.WriteStartElement("ExcludeGroup")
        xmlwriter.WriteAttributeString("Operator", self.operator)
        for rule in self.rules:
            rule.save_xml(xmlwriter)
        xmlwriter.WriteEndElement()


    def load_from_xml(self, xml_node):
        """Loads the rules and groups from the xml_node."""
        for node in xml_node.ChildNodes:
            if node.Name == "ExcludeRule":
                #Changes from 1.7.17 to 2.0
                try:
                    self.add_rule(ExcludeRule(node.Attributes["Field"].Value, node.Attributes["Operator"].Value, node.Attributes["Value"].Value))
                except AttributeError:
                    self.add_rule(ExcludeRule(node.Attributes["Field"].Value, node.Attributes["Operator"].Value, node.Attributes["Text"].Value))
            
            elif node.Name == "ExcludeGroup":
                group = ExcludeGroup(node.Attributes["Operator"].Value)
                group.load_from_xml(node)
                self.add_rule(group)


            
class ExcludeRule(object):
    """Contains the data of an exlude rule. It can calculate if a book should be moved using it's rules."""
    
    def __init__(self, field, operator, value):
        
        self.field = field
        
        self.operator = operator
        
        self.value = value
        

    def get_yes_no_value(self):
        """Returns the correct YesNo value."""
        if self.field == "Manga":
            if self.value == "Yes (Right to Left)":
                return MangaYesNo.YesAndRightToLeft
            else:
                return getattr(MangaYesNo, self.value)

        return getattr(YesNo, self.value)

        
    def book_should_be_moved(self, book):
        """
        Finds if the book should be moved using this rule.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """

        field = name_to_field[self.field]

        if field in ("Manga", "SeriesComplete", "BlackAndWhite"):
            return self.calculate_book_should_be_moved(book, getattr(book, field), self.get_yes_no_value())

        elif field in ("StartYear", "StartMonth"):
            return self.calculate_book_should_be_moved(book, self.get_start_field_data(book, field), self.value)

        else:
            return self.calculate_book_should_be_moved(book, unicode(getattr(book, field)), self.value)
    

    def calculate_book_should_be_moved(self, book, field_data, value):
        """
        Checks if the book should be moved based on this single rule.
        field_data -> the contents of the field.
        value -> the value to check the contents of the field against.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """

        if self.operator == "is":
        #Convert to string just in case
            if field_data == value:
                return 1
            else:
                return 0
        elif self.operator == "does not contain":
            if value not in field_data:
                return 1
            else:
                return 0
        elif self.operator == "contains":
            if value in field_data:
                return 1
            else:
                return 0
        elif self.operator == "is not":
            if value != field_data:
                return 1
            else:
                return 0
        elif self.operator == "greater than":
            #Try to use the int value to compare if possible
            try:
                if int(value) < int(field_data):
                    return 1
                else:
                    return 0
            except ValueError:
                if value < field_data:
                    return 1
                else:
                    return 0
        elif self.operator == "less than":
            try:
                if int(value) > int(field_data):
                    return 1
                else:
                    return 0
            except ValueError:
                if value > field_data:
                    return 1
                else:
                    return 0
        

    def get_start_field_data(self, book, field):
        """
        Finds the field contents for the earlies book of the same series in the ComicRack library.
        book -> the book of the series to search for.
        field -> The string of the field to retrieve.
        
        returns -> Unicode string of the field.
        """

        startbook = get_earliest_book(book)
        
        if field == "StartMonth":
            return unicode(startbook.Month)

        else:
            return unicode(startbook.ShadowYear)


    def save_xml(self, xmlwriter):
        """Save this rule to an xml file using the specified xmlwriter."""
        xmlwriter.WriteStartElement("ExcludeRule")
        xmlwriter.WriteAttributeString("Field", self.field)
        xmlwriter.WriteAttributeString("Operator", self.operator)
        xmlwriter.WriteAttributeString("Value", self.value)
        xmlwriter.WriteEndElement()



class UndoCollection(object):

    def __init__(self):
        self._undo_paths = []
        self._current_paths = []
        self._profiles = []


    def __len__(self):
        return len(self._undo_paths)


    def append(self, undo_path, new_path, profile_name):
        #Book has been moved more than once. In this case keep only a record of the original location and it's current location.
        if undo_path in self._current_paths:
            self._current_paths[self._current_paths.index(undo_path)] = new_path
        else:
            self._undo_paths.append(undo_path)
            self._current_paths.append(new_path)
            self._profiles.append(profile_name)


    def undo_path(self, path):
        """Gets an undo path from a current path."""
        return self._undo_paths[self._current_paths.index(path)]


    def profile(self, path):
        """Gets the name of the profile used to move a book.
        path->The current path of the book.
        Returns the string name of the profile.
        """
        return self._profiles[self._current_paths.index(path)]


    def save(self, file_path):
        """Saves the undo collection to a file.
        file_path->The complete file path of the file to save to.
        """
        try:
            with open(file_path, 'w') as f:
                for index in range(0, len(self._current_paths)):
                    f.write(self._profiles[index].encode("utf8") + "|" + self._current_paths[index].encode("utf8") + "|" + self._undo_paths[index].encode("utf8") + "\n")

        except IOError, ex:
            print "Somthing went wrong saving the undo list"
            print ex


    def load(self, file_path):
        """Loads an undo collection from a file into the undo collection instance.
        file_path->The complete file path of the file to load the undo collection from.
        """
        try:
            with open(file_path, 'r') as f:
                for line in f:
                    i = line.decode('utf8')
                    parts = i.strip().split("|")
                    self.append(parts[2], parts[1], parts[0])

            #Reverse them because otherwise if multiple profiles have been used a book could have been moved multiple times.
            #This way it ends back in the original spot eventually.
            self._current_paths.reverse()
            self._profiles.reverse()
            self._undo_paths.reverse()
        except IOError, ex:
            print "Error loading dict from file " + file_path
            print ex



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



def get_earliest_book(book):
    """
    Finds the first published issue of a series in the library.
    Returns a ComicBook object.
    """
    #Find the Earliest by going through the whole list of comics in the library find the earliest year field and month field of the same series and volume
        
    index = book.Publisher+book.ShadowSeries+str(book.ShadowVolume)
        
    if index in startbooks:
        return startbooks[index]

    startbook = book
            
    for b in ComicRack.App.GetLibraryBooks():
        if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume and b.Publisher == book.Publisher:
                    
            #Notes:
            #Year can be empty (-1)
            #Month can be empty (-1)

            #In case the initial value is bad
            if startbook.ShadowYear == -1 and b.ShadowYear != 1:
                startbook = b
                    
            #Check if the current book's year is earlier
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

                #Month is the same so check for later issue numbers:
                elif b.Month == startbook.Month:
                    if b.ShadowNumber < startbook.ShadowNumber:
                        startbook = b
            
    #Store this final result in the dict so no calculation require for others of the series.
    startbooks[index] = startbook

    return startbook



def get_last_book(book):
    """
    Finds the last published issue of a series in the library.
    Returns a ComicBook object.
    """
    #Find the last by going through the whole list of comics in the library find the earliest issue of the series
        
    index = book.Publisher+book.ShadowSeries+str(book.ShadowVolume)
        
    #If we have already found this series:
    if index in endbooks:
        return endbooks[index]

    endbook = book
            
    for b in ComicRack.App.GetLibraryBooks():
        if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume and b.Publisher == book.Publisher:
            if not endbook.ShadowNumber.isdigit():
                endbook = b
                continue
            if (endbook.ShadowNumber.isdigit() and b.ShadowNumber.isdigit()) and int(endbook.ShadowNumber) < int(b.ShadowNumber):
                endbook = b
          
    #Store this final result in the dict so no calculation require for others of the series.
    endbooks[index] = endbook

    return endbook


def check_metadata_rules(book, profile):
    """Goes through the metadata rules and sees if the book should be moved or not.
    
    Returns True if the book should be moved, returns false if the book should not be moved.
    
    """
    count = 0
    total = 0
    
    qualifies = False
    
    for rule in profile.ExcludeRules:
      
        result = rule.book_should_be_moved(book)
        
        if result == None:
            continue
        
        count += result
        total += 1
    
    if total == 0:
        #No rules so the book should be moved regardless
        return True
    
    if profile.ExcludeOperator == "Any":
        if count > 0:
            qualifies = True
    
    elif profile.ExcludeOperator == "All":
        if count == total:
            qualifies = True

    if profile.ExcludeMode == "Only":
        #book falls under rules and should be moved if it does so return True to have it moved
        if qualifies == True:
            return True
        
        #book doesn't fall under rules and should not be moved. So return False to have it skipped
        else:
            return False
    
    elif profile.ExcludeMode == "Do not":
        #book falls under rules and should not be moved so return False to have it skipped
        if qualifies == True:
            return False
        #book does not fall under rules and should be moved, so return True to have it moved
        else:
            return True



def check_excluded_folders(book_path, profile):
    """Checks the excluded paths of a profile.
    Returns False if the book is located in an excluded path. Returns True otherwise.
    """
    for path in profile.ExcludeFolders:
        if path in book_path:
            return False
    return True


def get_custom_value_keys():
    """Retrieves a list of all the custom value keys in the library"""
    keys = []
    for book in ComicRack.App.GetLibraryBooks():
        for pair in book.GetCustomValues():
            if pair.Key not in keys:
                keys.append(pair.Key)
    return keys