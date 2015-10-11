"""
duplicatewindow.py

Author: Stonepaw


Description: Contains a Window for displaying two comic images and various 
information so that the user can pick to overwrite, rename or cancel.


Copyright 2010-2015 Stonepaw

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
clr.AddReferenceToFile("ExtractLargeIconFromFile.dll")

import wpf
from System import Uri, UriKind, ArgumentException  # @UnresolvedImport
from System.IO import Path, FileInfo
from System.Windows import Window  # @UnresolvedImport
from System.Windows.Controls import DataTemplateSelector  # @UnresolvedImport
from System.Windows.Data import IValueConverter  # @UnresolvedImport
from System.Windows.Media.Imaging import BitmapImage  # @UnresolvedImport

from ExtractLargeIconFromFile import ShellEx

import i18n
from locommon import SCRIPTDIRECTORY, ICON, Mode
from wpfutils import NotifyPropertyChangedBase, notify_property, bitmap_to_bitmapimage

ICONDIRECTORY = SCRIPTDIRECTORY
ComicRack = None

class DuplicateWindow(Window):
    """ A wpf window that mimics the windows 7 copy dialog that shows when
    a duplicate file exists.

    Usage: Due to performance cost of recreating the window, ShowDialog is 
        overridden to keep the window and reset the ui based on new values
        passed to that method.

        ShowDialog Returns a DuplicateResult
    """
    
    def __init__(self):
        """Initializes the DuplicateWindow"""

        # Load resources since I can't seem to get relative paths to work in xaml
        self.Resources.Add("Arrow", BitmapImage(
                    Uri(Path.Combine(ICONDIRECTORY, 'arrow.png'))))
        self.Resources.Add("ItemTypeSelector", DuplicateFormDataTemplateSelector())
        self.Resources.Add("ByteToMBConverter", BytesToMBConverter())
        
        self._action = None
        self._view_model = DuplicateWindowViewModel()
        wpf.LoadComponent(self, Path.Combine(FileInfo(__file__).DirectoryName, 
                                                 'DuplicateWindow.xaml'))
        self.DataContext = self._view_model

        #Set the Icon
        icon = BitmapImage()
        icon.BeginInit()
        icon.UriSource = Uri(ICON, UriKind.Absolute)
        icon.EndInit()
        self.Icon = icon

    def _replace_click(self, sender, e):
        self.DialogResult = True
        self._action = DuplicateAction.Overwrite

    def _cancel_click(self, sender, e):
        self.DialogResult = True
        self._action = DuplicateAction.Cancel

    def _rename_click(self, sender, e):
        self.DialogResult = True
        self._action = DuplicateAction.Rename

    def ShowDialog(self, processing_book, existing_book, rename_path, mode, count, diff_ext=False):
        """Sets up and shows the DuplicateWindow

        Args:
            processing_book: The ComicBook object that is being moved.
            existing_book: The ComicBook or FileInfo object that already exists 
                at the intended path.
            rename_path: The string containing the path if the ComicBook is
                moved/copied and then renamed.
            mode: The Mode to be used.
            count: The number of duplicates left to process.

        """
        self._action = DuplicateAction.Cancel # Default
        self._view_model.setup(processing_book, existing_book, rename_path, mode, count)
        super(DuplicateWindow, self).ShowDialog()
        return DuplicateResult(self._action, self._view_model.always_do_action)

    def _form_closing(self, sender, e):
        """ In order for the form to be reused after ShowDialog() the closing 
        has to be canceled and then the window has to be hidden.
        """
        e.Cancel = True
        self.Hide()


class DuplicateWindowViewModel(NotifyPropertyChangedBase):
    """The ViewModel for the DuplicateWindow.

    This controls mostly the different text that can be displayed depending
    on the mode being used.

    It also allows the DuplicateWindow to be reused easily by easily updating
    the bindings for the different text and covers.
    
    Use setup() to change the ComicBook/FileInfo objects.
    """
    def __init__(self):
        super(DuplicateWindowViewModel, self).__init__()
        self._load_text()
        self._rename_text = ""
        self._cancel_text = ""
        self._replace_header = ""
        self._processing_book = None
        self._existing_book = None
        self._do_all_text = ""
        self._show_do_all = False
        self._existing_cover = None
        self._processing_cover = None
        self._rename_header = ""
        self.always_do_action = False
    
    #region properties
    @notify_property
    def rename_text(self):
        return self._rename_text

    @rename_text.setter
    def rename_text(self, value):
        self._rename_text = value

    @notify_property
    def rename_header(self):
        return self._rename_header

    @rename_header.setter
    def rename_header(self, value):
        self._rename_header = value

    @notify_property
    def cancel_text(self):
        return self._cancel_text

    @cancel_text.setter
    def cancel_text(self, value):
        self._cancel_text = value

    @notify_property
    def replace_header(self):
        return self._replace_header

    @replace_header.setter
    def replace_header(self, value):
        self._replace_header = value

    @notify_property
    def processing_book(self):
        return self._processing_book

    @processing_book.setter
    def processing_book(self, value):
        self._processing_book = value

    @notify_property
    def existing_book(self):
        return self._existing_book

    @existing_book.setter
    def existing_book(self, value):
        self._existing_book = value

    @notify_property
    def do_all_text(self):
        return self._do_all_text

    @do_all_text.setter
    def do_all_text(self, value):
        self._do_all_text = value

    @notify_property
    def show_do_all(self):
        return self._show_do_all

    @show_do_all.setter
    def show_do_all(self, value):
        self._show_do_all = value

    @notify_property
    def existing_cover(self):
        return self._existing_cover

    @existing_cover.setter
    def existing_cover(self, value):
        self._existing_cover = value

    @notify_property
    def processing_cover(self):
        return self._processing_cover

    @processing_cover.setter
    def processing_cover(self, value):
        self._processing_cover = value

    #endregion

    def setup(self, processing_book, existing_book, rename_path, mode, count):
        """ Setups up the various properites of the view model depending on what
        mode is being used.

        Args:
            processing_book: The ComicBook object being moved.
            existing_book: The ComicBook or FileInfo object of the existing book
            rename_path: String of the path that the file can be renamed to.
            mode: The LibraryOrganizer Mode being used.
            count: The number of duplicates left to process
        """
        self.processing_book = processing_book
        self.existing_book = existing_book
        self._reset_text(mode, rename_path, count)
        self._reset_images()

    def _reset_text(self, mode, rename_path, count):
        """ Resets the various texts on the window depending on the mode.

        Args:
            mode: The LibraryOrganized mode to use.
            rename_path: The string containing the new rename path to use.
            count: The number of duplicates left to process.
        """
        if count > 1:
            self.show_do_all = True
            self.do_all_text = " ".join((self._texts["DupAlwaysDo"], count))
        else:
            self.show_do_all = False
        if mode == Mode.Copy:
            self.rename_text = self._texts["DupRenameDescriptionCopy"].format(rename_path)
            self.cancel_text = self._texts["DupCancelCopy"]
            self.replace_header = self._texts["DupCopyReplace"]
            self.rename_header = self._texts["DupRenameCopy"]
        else:
            self.rename_text = self._texts["DupRenameDescriptionMove"].format(rename_path)
            self.cancel_text = self._texts["DupCancelMove"]
            self.replace_header = self._texts["DupMoveReplace"]
            self.rename_header = self._texts["DupRenameMove"]
        
    def _reset_images(self):
        """Loads the new cover images or icon thumbnails depending on the 
        filetype.
        """
        if type(self._existing_book) == FileInfo:
            try:
                icon = ShellEx.GetBitmapFromFilePath(
                    self._existing_book.FullName,ShellEx.IconSizeEnum.LargeIcon48)
            except ArgumentException:
                self.existing_cover = None
            else:
                self.existing_cover = bitmap_to_bitmapimage(icon)
        else:
            self.existing_cover = bitmap_to_bitmapimage(
                ComicRack.App.GetComicThumbnail(
                    self._existing_book, self._existing_book.PreferredFrontCover))

        self.processing_cover = bitmap_to_bitmapimage(
            ComicRack.App.GetComicThumbnail(
                self._processing_book, 
                self._processing_book.PreferredFrontCover))

    def _load_text(self):
        """ Loads all the translations for the text in the Window """
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


class DuplicateResult(object):

    def __init__(self, action, always_do_action):
        self.action = action
        self.always_do_action = always_do_action


class DuplicateAction(object):
    Overwrite = 1
    Cancel = 2
    Rename = 3


class DuplicateFormDataTemplateSelector(DataTemplateSelector):
    """ This class is used to select the correct display in the wpf Window 
    depending if the book type is a ComicBook or FileInfo.
    """
    def SelectTemplate(self, item, container):
        if type(item) == FileInfo:
            return container.FindResource("FileInfoCommandButton")
        else:
            return container.FindResource("ComicBookCommandButton")


class BytesToMBConverter(IValueConverter):
    """ Returns a pretty string from a size in bytes"""
    def Convert(self, value, targetType, parameter, culture):
        if value >= 1048576:
            return "{0:.2f} MB".format(float(value)/1024/1024.0)
        elif value >=1024:
            return "{0:.2f} KB".format(float(value)/1024.0)
        else:
            return "{0:.2f} bytes".format(float(value))

    def ConvertBack(self, value, targetType, parameter, culture):
        return ""