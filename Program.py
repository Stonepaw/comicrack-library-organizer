"""
This file loads some sample comics from a file and shows the config form.
Simply for testing purposes.
"""


import clr
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference("System.Xml")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\ComicRack.Engine.dll")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\cYo.Common.dll")

from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter

from System.Xml import XmlDocument

from System.IO import StreamReader, File, FileStream, FileMode

from System.Windows.Forms import Application, DialogResult
import configureform

import losettings

from loworkerform import ProfileSelector

SETTINGSFILE = "losettingsx.dat"

try:
    print "Loading test data"
    f = FileStream("Sample.dat", FileMode.Open)
    bf = BinaryFormatter()
    books = bf.Deserialize(f)
    print "Done loading sample data"
    print "Starting to load profiles"
    profiles, last_used_profiles = losettings.load_profiles(SETTINGSFILE)
    print "Done loading profiles"
    Application.EnableVisualStyles()

    selector = ProfileSelector(profiles.keys(), last_used_profiles)

    selector.ShowDialog()

    print selector.get_profiles_to_use()

    form = configureform.ConfigureForm(profiles, last_used_profiles[0], books)
    r = form.ShowDialog()

    if r != DialogResult.Cancel:
        form.save_profile()
        last_used_profile = form._profile_selector.SelectedItem
        losettings.save_profiles(SETTINGSFILE, profiles, last_used_profile)
except Exception, ex:
    print ex
    System.Console.Read()

