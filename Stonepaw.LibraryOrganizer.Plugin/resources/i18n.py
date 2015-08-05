""" This a modified version of cbanack's localizer utility in 
ComicVine Scraper. I have just tweaked it to remove functionality
I don't need, add some functionality I do need and changed the zip utility 
to one ComicRack already ships with.

All credit goes to cbanack.
"""

import clr
clr.AddReference("System.Xml.Linq")
from System.IO import Path, StreamReader
from System.Xml.Linq import XElement

clr.AddReference("ICSharpCode.SharpZipLib")
from ICSharpCode.SharpZipLib.Zip import ZipFile, FastZip

from locommon import SCRIPTDIRECTORY

__i18n = None

#So that different folders can be used when launching from visual studio
resourcespath = SCRIPTDIRECTORY
default_resource = "en.zip"

def setup(comicrack):
    """Creates the localizer utility with a comicrack object and extracts the
    default English keys and translations from the en.zip.

    Args:
        comicrack: The ComicRack object passed from ComicRack to the script.
    """
    global __i18n
    __i18n = __i18n(comicrack)
    __i18n._get_defaults()

def get(key):
    """Retrieves a translation with a key or the default string if no
    translation is given. Throws an error if the localizer is not setup.

    Args:
        key: The string key of the translation.

    Returns:
        The translated string or the English string if no translation is found.
    """
    if __i18n is not None:
        return __i18n.get(key)
    else:
        raise "Localizer not initialized"

#TODO: use comicrack built in method GetComicFields to retrive some comic field names.
#Dictionary[str, str]({'Tags' : 'Tags', 'File Path' : 'FilePath', 'Book Age' : 'BookAge', 'Book Condition' : 'BookCondition', 'Book Store' : 'BookStore', 'Book Owner' : 'BookOwner', 'Book Collection Status' : 'BookCollectionStatus', 'Book Notes' : 'BookNotes', 'Book Location' : 'BookLocation', 'ISBN' : 'ISBN', 'Custom Values Store' : 'CustomValuesStore', 'Title' : 'Title', 'Series' : 'Series', 'Number' : 'Number', 'Alternate Series' : 'AlternateSeries', 'Alternate Number' : 'AlternateNumber', 'Story Arc' : 'StoryArc', 'Series Group' : 'SeriesGroup', 'Summary' : 'Summary', 'Notes' : 'Notes', 'Review' : 'Review', 'Writer' : 'Writer', 'Penciller' : 'Penciller', 'Inker' : 'Inker', 'Colorist' : 'Colorist', 'Letterer' : 'Letterer', 'Cover Artist' : 'CoverArtist', 'Editor' : 'Editor', 'Publisher' : 'Publisher', 'Imprint' : 'Imprint', 'Genre' : 'Genre', 'Web' : 'Web', 'Format' : 'Format', 'Age Rating' : 'AgeRating', 'Characters' : 'Characters', 'Teams' : 'Teams', 'Main Character Or Team' : 'MainCharacterOrTeam', 'Locations' : 'Locations', 'Scan Information' : 'ScanInformation'})

class __i18n(object):
    """ A hidden class that implements the localize functions. Extremely based
    on cbanack's i18n utility in ComicVine Scraper.
    """

    def __init__(self, comicrack):
        """ Initializes the i18n class and loads the default keys and strings
        from the en.zip file.

        Args:
            comicrack: The ComicRack object passed from ComicRack that contains
            the Localize method. This must not be none.
        """
        self._comicrack = comicrack
        defaults = self._get_defaults()
        self.defaults = defaults

    def _get_defaults(self):
        """Retrieves the default English translations and keys from the en.zip.

        parses in a file that looks like this:
        <?xml version="1.0" encoding="UTF-8"?>
         <TR xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                Name="Script.LibraryOrganizer" CultureName="en">
            <Texts>
               <Text Key="key" Text="Text" Comment="text"/>
               ...
            </Texts>
         </TR>

        Returns:
            A dictionary of {key, default text}.
        """
        defaults = {}
        z = ZipFile(Path.Combine(resourcespath,default_resource))
        xelement = None
        with StreamReader(z.GetInputStream(z.GetEntry("Script.LibraryOrganizer.xml"))) as f:
            xelement = XElement.Load(f)
        if xelement is None:
            raise "Could not load default translations"
        entries = xelement.Element("Texts").Elements("Text")
        return {entry.Attribute("Key").Value : entry.Attribute("Text").Value 
                for entry in entries}

    def get(self, key, resource="Script.LibraryOranizer"):
        """ Retrieves a translated string using ComicRack's localize function. 

        Args:
            key: The key to retrieve from.
            resource: The resource to retrieve from, defaults to LibraryOrganizer.

        Returns:
            The translated string or the default string if no translation is
            found.
        """
        return self._comicrack.Localize(resource, key, self.defaults[key])