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

import excluderules
from losettings import Profile
import xmlserializer
from xmlserializer import XmlDeserializer, XmlSerializer
import xml2py

class GlobalSettings(object):

    def __init__(self):
        self.version = VERSION
        self.last_used_profiles = []
        self.generate_localized_months = True
        self.remove_empty_folders = True
        self.excluded_empty_folders = []

        self.month_names = {"1" : "January", 
                           "2" : "February", 
                           "3" : "March", 
                           "4" : "April",
                           "5" : "May", 
                           "6" : "June", 
                           "7" : "July", 
                           "8" :"August", 
                           "9" : "September",
                           "10" : "October", 
                           "11" : "November", 
                           "12" : "December"}

        self.illegal_character_replacements = {"?" : "", 
                                               "/" : "", 
                                               "\\" : "", 
                                               "*" : "", 
                                               ":" : " - ",
                                               "<" : "[", 
                                               ">" : "]", 
                                               "|" : "!", 
                                               "\"" : "'"}
    def save(self, file_path):
        serializer = XmlSerializer()
        xSettings = XmlWriterSettings()
        xSettings.Indent = True
        with XmlWriter.Create(file_path, xSettings) as writer:
            writer.WriteStartElement("GlobalSettings")
            for i in self.__dict__:
                value = getattr(self, i)
                if type(value) is bool:
                    serializer.serialize_bool(value, i, writer)
                elif type(value) is list:
                    serializer.serialize_list(value, i, writer)
                elif type(value) is dict:
                    serializer.serialize_dict(value, i, writer, True)
                elif type(value) is str:
                    writer.WriteElementString(i, value)
            writer.WriteEndElement()

    def load(self, file_path):
        if not File.Exists(file_path):
            return
        with File.Open(file_path, FileMode.Open) as f:
            xml = xml2py.parse(f)
            xmlserializer.deserialize_object_from_xml2py(self, xml)


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
    d = configure_form_new.ConfigureForm(settings, global_settings)

    Application().Run(d)
    serializer = xmlserializer.XmlSerializer()
    serializer.serialize_profiles(settings, lastused, Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"))
    global_settings.save(Path.Combine(SCRIPTDIRECTORY,'globalsettings.dat'))
    #save_profiles(Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"), settings, lastused)


