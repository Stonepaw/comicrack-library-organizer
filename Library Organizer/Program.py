import clr
clr.AddReference("PresentationFramework")

import globals
globals.RUNNER = True
import ComicRack
globals.ComicRack = ComicRack.ComicRack()

import System

from System.Windows import Application


import settings
from configuration_window import ConfigurationWindow
from global_settings import GlobalSettings

def main():

    profiles = settings.get_profiles()
    print profiles
    global_settings = GlobalSettings()
    d = ConfigurationWindow(profiles, global_settings)
    Application().Run(d)
    settings.save_profiles(profiles)
    System.Console.ReadLine()

if __name__ == "__main__":
    main()