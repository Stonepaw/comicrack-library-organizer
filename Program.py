import wpf
from System.Windows import Application, Window

import clr
clr.AddReference("System.Windows.Forms")

from ComicRack import ComicRack
c = ComicRack()
import localizer
localizer.ComicRack = c

from System.Windows.Forms import MessageBox

from System.IO import Path
import configureform
from configureform import ConfigureForm
from configure_form_new import ConfigureFormNew
import locommon
from locommon import SCRIPTDIRECTORY

import excluderules
from losettings import Profile
import xmlserializer
from xmlserializer import XmlDeserializer

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
    #localizer.ComicRack = ComicRack
    configureform.ComicRack = ComicRack
    locommon.ComicRack = ComicRack
    settings, lastused = load_profiles(Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"))
    d = ConfigureForm(settings, lastused[0])
    #d = ConfigureFormNew()
    Application().Run(d)
    serializer = xmlserializer.XmlSerializer()
    serializer.serialize_profiles(settings, lastused, Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"))
    #save_profiles(Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat"), settings, lastused)


