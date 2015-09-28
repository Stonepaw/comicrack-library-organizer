"""
This module provides several classes that mock several methods provided by
ComicRack so that the script can be run in Visual Studio.
"""
import clr
clr.AddReference("System.Drawing")
import System
from System.Collections.Generic import Dictionary
from System.Drawing import SystemIcons, Bitmap
from System.IO import FileStream, FileMode, FileInfo, Path
from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter

class ComicRack(object):
    
    def __init__(self):
        self.App = App()

    def Localize(self, resource, key, text):
        """The Localize function would normally lookup a translated string
        from a resource via a key. The text argument would be returned if
        the key is not found in the resource. This mock function will
        always return the text.
        """
        return text


class App(object):
    """This is a replacement for methods provided by ComicRack. 
    This allows developing the program in Visual Studio without running CR."""
    
    def GetComicFields(self):
        return Dictionary[str, str]({
            "Tags " : "Tags","File Path " : "FilePath","Book Age " : "BookAge",
            "Book Condition " : "BookCondition","Book Store " : "BookStore",
            "Book Owner " : "BookOwner",
            "Book Collection Status " : "BookCollectionStatus",
            "Book Notes " : "BookNotes","Book Location " : "BookLocation",
            "ISBN " : "ISBN","Title " : "Title","Series " : "Series",
            "Number " : "Number","Alternate Series " : "AlternateSeries",
            "Alternate Number " : "AlternateNumber","Story Arc " : "StoryArc",
            "Series Group " : "SeriesGroup","Summary " : "Summary",
            "Notes" : "Notes","Review" : "Review","Writer" : "Writer",
            "Penciller " : "Penciller","Inker" : "Inker",
            "Colorist" : "Colorist","Letterer" : "Letterer",
            "Cover Artist" : "CoverArtist","Editor" : "Editor",
            "Publisher" : "Publisher","Imprint " : "Imprint","Genre" : "Genre",
            "Web " : "Web","Format " : "Format","Age Rating " : "AgeRating",
            "Characters " : "Characters","Teams " : "Teams",
            "Main Character Or Team " : "MainCharacterOrTeam",
            "Locations" : "Locations","Scan Information" : "ScanInformation",})

    def GetComicThumbnail(self, book, index):
        """ Mocks returning a bitmap page image by using SystemIcons.Error. """
        path = Path.Combine(Path.GetDirectoryName(__file__), "437px-Mystery_Men_Comics_L.jpg")
        return Bitmap(path)

    def GetLibraryBooks(self):
        """ Mocks calling ComicRack's built in function to retrieve all
        the library books. It uses a file sample_data.dat which
        is dumped from ComicRack using a BinaryFormatter"""
        with FileStream(Path.Combine(FileInfo(__file__).DirectoryName,
                                     "sample data.dat"), FileMode.Open) as f:
            b = BinaryFormatter()

            return b.Deserialize(f)

    def RemoveBook(self, book):
        """Mocks the RemoveBook Method"""
        pass

class ComicBook(object):
    """Mocks a ComicBook object"""
    def __init__(self):
        self.FilePath = ""
        self.Caption = ""
        self.FileDirectory = ""


