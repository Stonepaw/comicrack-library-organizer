"""
duplicate_form.py

Author: Stonepaw


Description: Contains a Form for displaying two comic images and various information


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

clr.AddReference("IronPython.Wpf")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Xml")

import wpf
import System.IO
from System import Uri, UriKind
from System.Drawing.Imaging import ImageFormat
from System.IO import Path, MemoryStream, FileInfo, FileStream, FileMode
from System.Windows import Visibility, Window
from System.Windows.Forms import DialogResult
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System.Xml import XmlReader

from locommon import SCRIPTDIRECTORY, ICON, Mode

import i18n

ICONDIRECTORY = SCRIPTDIRECTORY

class DuplicateForm(Window):
    """ A wpf window that mimics the windows 7 copy dialog that shows when
    a duplicate file exists.

    """
    
    def __init__(self):
        """Initializes the DuplicateForm"""

        self.Resources.Add("Arrow", BitmapImage(
                    Uri(Path.Combine(ICONDIRECTORY, 'arrow.png'))))
        
        self._action = None
        self._load_text()


        #The the correct text based on what mode we are in
        #Mode is set by default so only change if in Copy or Simulation mode
        #if mode == Mode.Copy:
        #    self.win.FindName("MoveHeader").Content = "Copy and Replace"
        #    self.win.FindName("MoveText").Content = "Replace the file in the destination folder with the file you are copying:"

        #    self.win.FindName("DontMoveHeader").Content = "Don't Copy"
            
        #    self.win.FindName("RenameHeader").Content = "Copy, but keep both files"
        #    self._rename_text = "The file you are copying will be renamed: "
        #    self.win.FindName("RenameText").Text = "The file you are copying will be renamed: "

        #if mode == Mode.Simulate:
        #    self.win.FindName("Subtitle").Content = "Click the file you want to keep (simulated, no files will be deleted or moved)"

        wpf.LoadComponent(self, Path.Combine(FileInfo(__file__).DirectoryName, 
                                                 'DuplicateForm.xaml'))
        self.DataContext = self
        icon = BitmapImage()
        icon.BeginInit()
        icon.UriSource = Uri(ICON, UriKind.Absolute)
        icon.EndInit()
        self.Icon = icon
    def _reset_text(self, mode, diff_ext):
        if diff_ext:
            pass

        elif mode == Mode.Move:
            self._rename_text = "The file you are moving will be renamed to {0}"
        elif mode == Mode.Copy:
            self._rename_text = "The file you are copying will be renamed to {0}"

    def ReplaceClick(self, sender, e):
        self._action = DuplicateAction.Overwrite
        #Note: set to hide so that the dialog can be reopened.
        self.win.Hide()

    def CancelClick(self, sender, e):
        self._action = DuplicateAction.Cancel
        self.win.Hide()

    def RenameClick(self, sender, e):
        self._action = DuplicateAction.Rename
        self.win.Hide()

    def ShowForm(self, newbook, oldbook, renamefile, count):
        self._action = DuplicateAction.Cancel
        self.SetupFields(newbook, oldbook, renamefile, count)

        self.win.ShowDialog()
        return DuplicateResult(self._action, self.win.FindName("DoAll").IsChecked)


    def FormClosing(self, sender, e):
        e.Cancel = True
        self._action = DuplicateAction.Cancel
        self.Hide()

    def _load_text(self):
        keys = ("DupHeader",
                "DupSubHeading",
                "DupMoveReplace",
                "DupCopyReplace",
                "DupReplaceDescription",
                "DupCancelMove",
                "DupCancelCopy",
                "DupCancelDescription",
                "DupRenameMove",
                "DupRenameCopy",
                "DupRenameDescriptionMove",
                "DupRenameDescriptionCopy",
                "DupAlwaysDo")

        self._texts = {key:i18n.get(key) for key in keys}

        #self._header_text = i18n.get("DupHeader")
        #self._subheader_text = i18n.get("DupSubHeading")
        #self._move_replace_text = i18n.get("DupMoveReplace")
        #self._copy_replace_text = i18n.get("DupCopyReplace")
        #self._replace_text = self._move_replace_text
        #self._replace_decription = i18n.get("DupReplaceDescription")
        #self._cancel_move_text = i18n.get("DupCancelMove")
        #self._cancel_copy_text = i18n.get("DupCancelCopy")
        #self._cancel_description = i18n.get("DupCancelDescription")
        #self._move_rename_text = i18n.get("DupRenameMove")
        #self._copy_rename_text = i18n.get("DupRenameCopy")
        #self._rename_description_move = i18n.get("DupRenameDescriptionMove")
        #self._rename_description_copy = i18n.get("DupRenameDescriptionCopy")
        #self._always_do_text = i18n.get("DupAlwaysDo")

        

    def SetupFields(self, newbook, oldbook, renamefile, count):

        self.win.FindName("RenameText").Text = self._rename_text + renamefile

        if type(newbook) != FileInfo:
            #oldbook can be either a ComicBook object or a FileInfo object
            self.win.FindName("NewSeries").Content = newbook.ShadowSeries
            self.win.FindName("NewVolume").Content = newbook.ShadowVolume
            self.win.FindName("NewNumber").Content = newbook.ShadowNumber
            self.win.FindName("NewPages").Content = newbook.PageCount
            self.win.FindName("NewFileSize").Content = newbook.FileSizeAsText
            self.win.FindName("NewPublishedDate").Content = newbook.MonthAsText + ", " + newbook.YearAsText
            self.win.FindName("NewAddedDate").Content = newbook.AddedTime
            self.win.FindName("NewPath").Text = newbook.FilePath
            self.win.FindName("NewScanInfo").Content = newbook.ScanInformation

            #Load the new book cover
            ncs = MemoryStream();
            ComicRack.App.GetComicThumbnail(newbook, newbook.PreferredFrontCover).Save(ncs, ImageFormat.Png);
            ncs.Position = 0;
            newcover = BitmapImage();
            newcover.BeginInit();
            newcover.StreamSource = ncs;
            newcover.EndInit();
            self.win.FindName("NewCover").Source = newcover
        else:
            self.win.FindName("NewSeries").Content = "Comic not in Library"
            self.win.FindName("NewVolume").Content = ""
            self.win.FindName("NewNumber").Content = ""
            self.win.FindName("NewPages").Content = ""
            self.win.FindName("NewFileSize").Content = '%.2f MB' %  (newbook.Length/1048576.0) #Calculate bytes to MB
            self.win.FindName("NewPublishedDate").Content = ""
            self.win.FindName("NewAddedDate").Content = ""
            self.win.FindName("NewScanInfo").Content = ""
            self.win.FindName("NewPath").Text = newbook.FullName
            self.win.FindName("NewCover").Source = None

        #old book, possible for it to be a FileInfo object
        if type(oldbook) != FileInfo:
            self.win.FindName("OldSeries").Content = oldbook.ShadowSeries
            self.win.FindName("OldVolume").Content = oldbook.ShadowVolume
            self.win.FindName("OldNumber").Content = oldbook.ShadowNumber
            self.win.FindName("OldPages").Content = oldbook.PageCount
            self.win.FindName("OldFileSize").Content = oldbook.FileSizeAsText
            self.win.FindName("OldPublishedDate").Content = oldbook.MonthAsText + ", " + oldbook.YearAsText
            self.win.FindName("OldAddedDate").Content = oldbook.AddedTime
            self.win.FindName("OldScanInfo").Content = oldbook.ScanInformation
            self.win.FindName("OldPath").Text = oldbook.FilePath

            #cover
            ocs = MemoryStream();
            ComicRack.App.GetComicThumbnail(oldbook, oldbook.PreferredFrontCover).Save(ocs, ImageFormat.Png);
            ncs.Position = 0;
            oldcover = BitmapImage();
            oldcover.BeginInit();
            oldcover.StreamSource = ocs;
            oldcover.EndInit();
            self.win.FindName("OldCover").Source = oldcover
        
        else:
            self.win.FindName("OldSeries").Content = "Comic not in Library"
            self.win.FindName("OldVolume").Content = ""
            self.win.FindName("OldNumber").Content = ""
            self.win.FindName("OldPages").Content = ""
            self.win.FindName("OldFileSize").Content = '%.2f MB' %  (oldbook.Length/1048576.0) #Calculate bytes to MB
            self.win.FindName("OldPublishedDate").Content = ""
            self.win.FindName("OldAddedDate").Content = ""
            self.win.FindName("OldScanInfo").Content = ""
            self.win.FindName("OldPath").Text = oldbook.FullName
            self.win.FindName("OldCover").Source = None

        if count > 1:
            self.win.FindName("DoAll").Visibility = Visibility.Visible
            self.win.FindName("DoAll").Content = "Do this for all conficts (" + str(count) + ")"
        else:
            self.win.FindName("DoAll").Visibility = Visibility.Hidden


class DuplicateResult(object):

    def __init__(self, action, always_do_action):
        self.action = action
        self.always_do_action = always_do_action


class DuplicateAction(object):
    Overwrite = 1
    Cancel = 2
    Rename = 3