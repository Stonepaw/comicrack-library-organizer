"""
libraryorganizer.py

The entry area for ComicRack

Version 2.0

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
import System

import System.IO
from System.IO import File, StreamReader, StreamWriter

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon

clr.AddReference("System.Xml")
import System.Xml
from System.Xml import XmlWriter, Formatting, XmlTextWriter, XmlWriterSettings, XmlDocument

import configureform
from configureform import ConfigureForm

import losettings
from losettings import load_profiles, save_profiles, save_last_used

import loworkerform
from loworkerform import ProfileSelector, WorkerForm, WorkerFormUndo

import locommon
from locommon import PROFILEFILE, UNDOFILE, UndoCollection

import bookmover
import pathmaker

import stdoutnlogtarget
from System.IO import Directory

Directory.SetCurrentDirectory(locommon.SCRIPTDIRECTORY)

#@Name Library Organizer
#@Hook Books
#@Key library-organizer-main
#@Image libraryorganizer.png
def LibraryOrganizer(books):
    if books:
        try:
            profiles, lastused = load_profiles(PROFILEFILE)

            loworkerform.ComicRack = ComicRack
            locommon.ComicRack = ComicRack
            bookmover.ComicRack = ComicRack
            pathmaker.ComicRack = ComicRack
            #Create the config form
            print "Creating config form"
            if show_config_form(profiles, lastused, books):
                show_worker_form(profiles, lastused, books)

        except Exception, ex:
            print "The following error occured"
            print Exception
            MessageBox.Show(str(ex))


#@Name Configure Library Organizer
#@Hook Library
#@Image libraryorganizer.png
def ConfigureLibraryOrganizer(books):
    if books is None:
        books = ComicRack.App.GetLibraryBooks()
    try:
        locommon.ComicRack = ComicRack
        bookmover.ComicRack = ComicRack
        pathmaker.ComicRack = ComicRack
        profiles, lastused = load_profiles(PROFILEFILE)

        show_config_form(profiles, lastused, books)
        
    except Exception, ex:
        print "The Following error occured"
        print Exception
        MessageBox.Show(str(ex))


#@Key library-organizer-main
#@Hook ConfigScript
def ConfigLibraryOrganizer():
    ConfigureLibraryOrganizer(None)


#@Name Library Organizer (Quick)
#@Hook Books
#@Key library-organizer-quick
#@Image libraryorganizerquick.png
def LibraryOrganizerQuick(books):
    if books:
        try:
            loworkerform.ComicRack = ComicRack
            locommon.ComicRack = ComicRack
            bookmover.ComicRack = ComicRack
            pathmaker.ComicRack = ComicRack
            profiles, lastused = load_profiles(PROFILEFILE)

            if len(profiles) == 1 and profiles[profiles.keys()[0]].BaseFolder == "":
                MessageBox.Show("Library Organizer will not work as expected when the BaseFolder is empty. Please run the normal Library Organizer script or the Configure Library Organizer script before running Library Organizer Quick", "BaseFolder empty", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            
            show_worker_form(profiles, lastused, books)

        except Exception, ex:
            print "The following error occured"
            print Exception
            print ex
            MessageBox.Show(str(ex))
      
              
#@Name Library Organizer - Undo last move
#@Hook Library
#@Image libraryorganizer.png
def LibraryOrganizerUndo(books):
    try:
        if File.Exists(UNDOFILE):
            loworkerform.ComicRack = ComicRack
            locommon.ComicRack = ComicRack
            bookmover.ComicRack = ComicRack
            pathmaker.ComicRack = ComicRack
            profiles, lastused = load_profiles(PROFILEFILE)
                
            undo_collection = UndoCollection()

            undo_collection.load(UNDOFILE)

            if len(undo_collection) > 0:
                undo_form = WorkerFormUndo(undo_collection, profiles)
                undo_form.ShowDialog()
                undo_form.Dispose()
                File.Delete(UNDOFILE)
            else:
                MessageBox.Show("Error loading Undo file", "Library Organizer - Undo")
        else:
            MessageBox.Show("Nothing to Undo", "Library Organizer - Undo")
    except Exception, ex:
        print "The following error occured"
        print Exception
        MessageBox.Show(str(ex))


#@Name Library Organizer - Startup
#@Enabled false
#@Hook Startup
#@Image libraryorganizer.png
def LibraryOrganizerStartup():
    books = ComicRack.App.GetLibraryBooks()
    LibraryOrganizerQuick(books)


def show_config_form(profiles, lastused, books):
    """Shows the configure form and saves the changes if the user press okay.
    Returns True if the user press Okay.
    Returns False if the user pressed cancel."""
    configform = ConfigureForm(profiles, lastused[0], books)
    result = configform.ShowDialog()
    configform.save_profile()
    configform.Dispose()
    if result != DialogResult.Cancel:
        save_profiles(PROFILEFILE, profiles, lastused)
        return True
    return False


def show_worker_form(profiles, lastused, books):
    """Gets the profile(s) to use and shows the worker form."""
    if len(profiles) > 1:
        profile_selector = ProfileSelector(profiles.keys(), lastused)
        result = profile_selector.ShowDialog()
        if result == DialogResult.Cancel:
            profile_selector.Dispose()
            return
        lastused = profile_selector.get_profiles_to_use()
        save_last_used(PROFILEFILE, lastused)
        profile_selector.Dispose()

    profiles_to_use = [profiles[name] for name in lastused]

    worker_form = WorkerForm(books, profiles_to_use)
    worker_form.ShowDialog()