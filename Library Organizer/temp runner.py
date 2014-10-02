import clr
import wpf
import globals
globals.RUNNER = True
import ComicRack
globals.ComicRack = ComicRack.ComicRack()
import System


from System.IO import Path
from System.Windows import Application, Window

import locommon
import configure_form
from losettings import Profile
from global_settings import GlobalSettings
from common import SCRIPTDIRECTORY

profiles = {}
profiles["Default"] = Profile()
profiles["Default"].Name = "Default"
#Some default templates
profiles["Default"].FileTemplate = "{<series>}{ Vol.<volume>}{ #<number2>}{ (of <count2>)}{ ({<month>, }<year>)}"
profiles["Default"].FolderTemplate = "{<publisher>}\{<imprint>}\{<series>}{ (<startyear>{ <format>})}"

if __name__ == '__main__':
    try:
        ComicRack = ComicRack.ComicRack()
        books = ComicRack.App.GetLibraryBooks()
        configure_form.ComicRack = ComicRack
        locommon.ComicRack = ComicRack

    
        global_settings = GlobalSettings()
        global_settings.load(Path.Combine(SCRIPTDIRECTORY,'globalsettings.dat'))
        #d = ConfigureForm(settings, lastused[0])
        #e = ExcludeRuleTest()
        d = configure_form.ConfigurationWindow(profiles, global_settings)

        Application().Run(d)
    except Exception, ex:
        print ex
        System.Console.ReadLine()