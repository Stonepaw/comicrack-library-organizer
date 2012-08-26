import System
from System.Collections.Generic import Dictionary

class ComicRack(object):
    
    def __init__(self):
        self.App = App()

    def Localize(self, resource, key, text):
        return text

class App(object):
    """This is a replacement for methods provided by comicrack. This allows developing the program in Visual Studio without running CR."""
    

    def GetComicFields(self):
        return Dictionary[str, str]({"Tags " : "Tags","File Path " : "FilePath","Book Age " : "BookAge","Book Condition " : "BookCondition","Book Store " : "BookStore","Book Owner " : "BookOwner","Book Collection Status " : "BookCollectionStatus","Book Notes " : "BookNotes","Book Location " : "BookLocation","ISBN " : "ISBN","Title " : "Title","Series " : "Series","Number " : "Number","Alternate Series " : "AlternateSeries","Alternate Number " : "AlternateNumber","Story Arc " : "StoryArc","Series Group " : "SeriesGroup","Summary " : "Summary","Notes " : "Notes","Review " : "Review","Writer " : "Writer","Penciller " : "Penciller","Inker " : "Inker","Colorist " : "Colorist","Letterer " : "Letterer","Cover Artist " : "CoverArtist","Editor " : "Editor","Publisher " : "Publisher","Imprint " : "Imprint","Genre " : "Genre","Web " : "Web","Format " : "Format","Age Rating " : "AgeRating","Characters " : "Characters","Teams " : "Teams","Main Character Or Team " : "MainCharacterOrTeam","Locations " : "Locations","Scan Information " : "ScanInformation",})



