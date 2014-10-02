import wpf
from System.Windows import Application, Window
import sys
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Xml")
from System.IO import Path, File, FileMode, DirectoryInfo, FileInfo

#VERY IMPORTANT
# ComicRack doesn't support folders in it's script installer
from ComicRack import ComicRack
c = ComicRack()
import localizer
localizer.ComicRack = c


from System.Windows.Forms import MessageBox
from System.Xml import XmlDocument, XmlWriter, XmlWriterSettings, XmlConvert

import configureform
from configureform import ConfigureForm

import configure_form_new
import locommon
from locommon import SCRIPTDIRECTORY, VERSION

#import excluderules
from losettings import Profile
import xmlserializer
from xmlserializer import XmlDeserializer, XmlSerializer
import xml2py
from global_settings import GlobalSettings
#from exclude_rule_test import ExcludeRuleTest


def load_profiles(file_path):
    """
    Load profiles from a xml file. If no profiles are found it creates a blank profile.
    file_path->The absolute path to the profile file

    Returns a dict of the found profiles and a list of the lastused profile(s)
    """
    xmldeserializer = XmlDeserializer()
    profiles, lastused = xmldeserializer.deserialize_profiles(file_path)
    #profiles, lastused = load_profiles_from_file(file_path)

    if len(profiles) == 0:
        #Just in case
        profiles["Default"] = Profile()
        profiles["Default"].Name = "Default"
        #Some default templates
        profiles["Default"].FileTemplate = "{<series>}{ Vol.<volume>}{ #<number2>}{ (of <count2>)}{ ({<month>, }<year>)}"
        profiles["Default"].FolderTemplate = "{<publisher>}\{<imprint>}\{<series>}{ (<startyear>{ <format>})}"
        
    if not lastused:
        lastused = [profiles.keys()[0]]
       
    return profiles, lastused


if __name__ == '__main__':
    ComicRack = ComicRack()
    books = ComicRack.App.GetLibraryBooks()
    #localizer.ComicRack = ComicRack
    configureform.ComicRack = ComicRack
    locommon.ComicRack = ComicRack
    configure_form_new.RUNNER = True

    settings, lastused = load_profiles(Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"))
    global_settings = GlobalSettings()
    global_settings.load(Path.Combine(SCRIPTDIRECTORY,'globalsettings.dat'))
    #d = ConfigureForm(settings, lastused[0])
    #e = ExcludeRuleTest()
    d = configure_form_new.ConfigureForm(settings, global_settings)

    Application().Run(d)
    serializer = xmlserializer.XmlSerializer()
    serializer.serialize_profiles(settings, lastused, Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"))
    global_settings.save(Path.Combine(SCRIPTDIRECTORY,'globalsettings.dat'))
    #save_profiles(Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"), settings, lastused)


