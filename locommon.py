"""
locommon.py

Author: Stonepaw


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
import re
import System
from System.Collections.Generic import Dictionary, SortedDictionary
clr.AddReference("System.Drawing")
from System.Drawing import Size, Point

from System.IO import Path, FileInfo


SCRIPTDIRECTORY = FileInfo(__file__).DirectoryName
PROFILEFILE = Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat")
ICON = Path.Combine(SCRIPTDIRECTORY, "libraryorganizer.ico")
UNDOFILE = Path.Combine(SCRIPTDIRECTORY, "undo.dat")
GLOBAL_SETTINGS_FILE = Path.Combine(SCRIPTDIRECTORY, "globalsettings.dat")
VERSION = "2.2"
MAJORVERSION = 2
MINORVERSION = 2

SUBSTITUTION_REGEX = re.compile("{(?P<prefix>[^{}<]*)<"
                               "(?P<name>[^\d\s(>]+)"
                               "(?P<args>\d*|(?:\([^)]*\))*)>"
                               "(?P<suffix>[^{}]*)}")

#clr.AddReferenceByPartialName('ComicRack.Engine')
#from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

REQUIRED_ILLEGAL_CHARS = ['?', '>', '<', ':', '|', '\\', '"', '/', '*']

startbooks = {}
endbooks = {}

#The usable date formated for any date fields
date_formats = ["dd-MM-yy", "dd-MM-yyyy", "MM-dd-yy", "MM-dd-yyyy", 
                "yyyy-MM-dd", "M", "Y", "MMMM d, yyyy", "MMMM dd, yyyy", 
                "MMM d, yyyy", "MMM dd, yyyy", "MMM yyyy", "MMMM yyyy", 
                "MMM, yyyy"]

#Don't actually need these anymore because all names are now localized. These do have to be kept for updating from previous versions though.
name_to_field = {"Age Rating" : "AgeRating", "Alternate Count" : "AlternateCount", "Alternate Number" : "AlternateNumber",
                 "Alternate Series" : "AlternateSeries", "Black And White" : "BlackAndWhite", "Cover Artist" : "CoverArtist", "File Format" : "FileFormat", "File Name" : "FileName", 
                 "File Path" : "FilePath", "First Letter" : "FirstLetter", "Language" : "LanguageAsText",
                 "Main Character Or Team": "MainCharacterOrTeam", "Read Percentage" : "ReadPercentage", "Scan Information" : "ScanInformation", "Series Complete" : "SeriesComplete", "Series Group" : "SeriesGroup", "Start Month" : "StartMonth", "Start Year" : "StartYear", "Story Arc": "StoryArc"}

field_to_name = {"AgeRating" : "Age Rating", "AlternateCount" : "Alternate Count", "AlternateNumber" : "Alternate Number", 
                 "AlternateSeries" : "Alternate Series", "BlackAndWhite" : "Black And White", "Characters" : "Characters", "Colorist" : "Colorist",
                 "Counter" : "Counter", "CoverArtist" : "Cover Artist", "Editor" : "Editor", "FileFormat" : "File Format", "FileName" : "File Name", "FilePath" : "File Path", 
                 "FirstLetter" : "First Letter", "Genre" : "Genre", "Imprint" : "Imprint", "Inker" : "Inker",
                 "LanguageISO" : "Language", "Letterer" : "Letterer", "Locations" : "Locations", "MainCharacterOrTeam": "Main Character Or Team", "Manga" : "Manga", "Month" : "Month", "Notes" : "Notes", "Penciller" : "Penciller",
                 "Publisher" : "Publisher", "Rating" : "Rating", "ReadPercentage" : "Read Percentage", "Read" : "Read Percentage", "Review" : "Review", "ScanInformation" : "Scan Information", 
                 "SeriesComplete" : "Series Complete", "ShadowCount" : "Count", "ShadowFormat" : "Format", "ShadowNumber" : "Number", 
                 "ShadowSeries" : "Series", "SeriesGroup" : "Series Group", "StoryArc": "Story Arc", "ShadowTitle" : "Title", "ShadowVolume" : "Volume", "ShadowYear" : "Year", "StartMonth" : "Start Month", 
                 "StartYear" : "Start Year", "Tags" : "Tags", "Teams" : "Teams", "Web" : "Web", "Writer" : "Writer"}

#These next few lists are just to make certain functions easier while providing one place needed to add new fields in buy simply add the property to the relavent list.

#Contains all the used fields in the comic. This is used to build the translations
comic_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 'AlternateSeries',
                'BlackAndWhite', 'BookAge', 'BookCollectionStatus', 'BookCondition',
                'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 'BookStore', 
                'Characters', 'Checked', 'Colorist', 'CommunityRating', 'ShadowCount', 
                'CoverArtist', 'Editor', 'FileDirectory', 'FileFormat', 'FileName', 
                'FileIsMissing', 'FileNameWithExtension', 'FilePath', 'FileSize', 
                'ShadowFormat', 'Genre', 'HasBeenOpened', 'HasBeenRead', 'ISBN', 
                'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 
                'MainCharacterOrTeam', 'Manga', 'Month', 'Notes', 'ShadowNumber', 
                'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 
                'ScanInformation', 'ShadowSeries', 'SeriesComplete', 'SeriesGroup', 
                'StoryArc', 'Summary', 'Tags', 'Teams', 'ShadowTitle', 'ShadowVolume', 
                'Web', 'Writer', 'ShadowYear']

#This contains the fields that are available to add into the template. Used for building the correct list of things later.
template_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 'AlternateSeries', 'BlackAndWhite', 'BookAge', 'BookCollectionStatus', 'BookCondition', 'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 'BookStore', 'Characters', 'Colorist', 'CommunityRating', 'Conditional', 'Count', 'Counter', 'CoverArtist', 'Editor', 'FirstIssueNumber', 'FirstLetter', 'Format', 'Genre', 'ISBN', 'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 'MainCharacterOrTeam', 'Manga', 'Month', 'Number', 'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 'ScanInformation', 'Series', 'SeriesComplete', 'SeriesGroup', 'StartMonth', 'StartYear', 'StoryArc', 'Summary', 'Tags', 'Teams', 'Title', 'Volume', 'Writer', 'Year']

#These are special fields that are created in library organizer. Needs to be seperate for ease of getting translations
library_organizer_fields = ["Counter", "FirstLetter", "Conditional", "StartMonth", "StartYear", "FirstIssueNumber"]

#These are the fields useable in the exclude rules.
exclude_rule_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 'AlternateSeries', 'BlackAndWhite', 'BookAge', 'BookCollectionStatus', 'BookCondition', 'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 'BookStore', 'Characters', 'Checked', 'Colorist', 'CommunityRating', 'Count', 'CoverArtist', 'Editor', 'FileDirectory', 'FileFormat', 'FileName', 'FileIsMissing', 'FileNameWithExtension', 'FilePath', 'FileSize', 'Format', 'Genre', 'HasBeenOpened', 'HasBeenRead', 'ISBN', 'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 'MainCharacterOrTeam', 'Manga', 'Month', 'Notes', 'Number', 'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 'ScanInformation', 'Series', 'SeriesComplete', 'SeriesGroup', 'StoryArc', 'Summary', 'Tags', 'Teams', 'Title', 'Volume', 'Web', 'Writer', 'Year']

#These are the fields that require the multiple value treatment.
multiple_value_fields = ["AlternateSeries", "Characters", "Colorist", "CoverArtist", "Editor", "Genre", "Inker", "Letterer", "Locations", "Penciller", "ScanInformation", "Tags", "Teams", "Writer"]

#These are fields that are usable in the first letter selector. Typically string fields.
first_letter_fields = ['AlternateSeries', 'Imprint', 'Publisher', 'Series' ]

#This contains the fields usable in the selectors for condtitional. Should basically contain everything but Conditional with the additon of Text\
conditional_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 'AlternateSeries', 'BlackAndWhite', 'BookAge', 'BookCollectionStatus', 'BookCondition', 'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 'BookStore', 'Characters', 'Colorist', 'CommunityRating', 'Count', 'CoverArtist', 'Editor', 'FirstIssueNumber', 'FirstLetter', 'Format', 'Genre', 'ISBN', 'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 'MainCharacterOrTeam', 'Manga', 'Month', 'Number', 'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 'ScanInformation', 'Series', 'SeriesComplete', 'SeriesGroup', 'StartMonth', 'StartYear', 'StoryArc', 'Summary', 'Tags', 'Text', 'Teams', 'Title', 'Volume', 'Writer', 'Year']
conditional_then_else_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 'AlternateSeries', 'BlackAndWhite', 'BookAge', 'BookCollectionStatus', 'BookCondition', 'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 'BookStore', 'Characters', 'Colorist', 'CommunityRating', 'Count', 'Counter', 'CoverArtist', 'Editor', 'FirstIssueNumber', 'FirstLetter', 'Format', 'Genre', 'ISBN', 'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 'MainCharacterOrTeam', 'Manga', 'Month', 'Number', 'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 'ScanInformation', 'Series', 'SeriesComplete', 'SeriesGroup', 'StartMonth', 'StartYear', 'StoryArc', 'Summary', 'Tags', 'Text', 'Teams', 'Title', 'Volume', 'Writer', 'Year']



#Although we could assume that any of the shadow fields are filled it is better to get they value from the shadow fields when possible.
#Can use book.GetPropertyValue[type](Name, bool:check shadow)


class Translations(object):
    def __init__(self):
        self.rules_operators = SortedDictionary[str, str]({"contains" : "contains", "does not contain" : "does not contain", "greater than" : "greater than", "less than" : "less than",  "is" : "istrans", "is not" : "is not trans", })
        self.rules_operators_yes_no = SortedDictionary[str, str]({"is" : "istrans", "is not" : "is not trans"})
    rules_fields = SortedDictionary[str, str](field_to_name)
    rules_values_manga = SortedDictionary[str, str]({"Yes" : "Yes", "Yes (Right to Left)" : "Yes (Right to Left)", "No" : "No", "Unknown" : "Unknown"})
    rules_values = SortedDictionary[str, str]({"Yes" : "Yes", "No" : "No", "Unknown" : "Unknown"})
    rules_mode = SortedDictionary[str, str]({"Only" : "Only", "Do not" : "Do not"})
    rules_group_operators = SortedDictionary[str, str]({"All" : "All", "Any" : "Any"})
        

exclude_string_operators = ["is", 
                            "contains", 
                            "contains any of", 
                            "contains all of", 
                            "starts with", 
                            "ends with", 
                            "list contains", 
                            "regular expression"]

                                                  
class Mode(object):
    Move = "Move"
    Copy = "Copy"
    Simulate = "Simulate"


class CopyMode(object):
    AddToLibrary = True
    DoNotAdd = False


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




