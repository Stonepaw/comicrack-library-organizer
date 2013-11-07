import clr

clr.AddReference("IronPython.Modules")
clr.AddReferenceToFile("Microsoft.WindowsAPICodePack.dll")
clr.AddReferenceToFile("Microsoft.WindowsAPICodePack.Shell.dll")

from IronPython.Modules import PythonLocale
from Microsoft.WindowsAPICodePack.Dialogs import CommonOpenFileDialog, CommonFileDialogResult
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Specialized import NotifyCollectionChangedAction

from fieldmappings import FIELDS, library_organizer_fields, template_fields, TemplateItem
from insert_view_models import ConditionalInsertViewModel, NumberInsertViewModel, SelectFieldTemplateFromType
from locommon import REQUIRED_ILLEGAL_CHARS
from losettings import Profile
from wpfutils import Command, notify_property, ViewModelBase


PythonLocale.setlocale(PythonLocale.LC_ALL, "")

class ConfigureFormViewModel(ViewModelBase):

    def __init__(self, profiles, global_settings):
        super(ConfigureFormViewModel, self).__init__()
        self.FileFolderViewModel = ConfigureFormFileFolderViewModel()
        self.OptionsViewModel = ConfigureFormOptionsViewModel(global_settings)
        self.Profiles = ObservableCollection[Profile](profiles.values())
        self._profile = None
        self._profile_names = profiles.keys()
        self.Profile = self.Profiles[0]
        self._input_is_visible = False

        #Commands
        self.SelectBaseFolderCommand = Command(self.select_base_folder)
        self.NewProfileCommand = Command(self.add_new_profile, 
                lambda x: x and x not in self._profile_names, True)

    #Profile
    @notify_property
    def Profile(self):
        return self._profile

    @Profile.setter
    def Profile(self, value):
        self._profile = value
        self.FileFolderViewModel.Profile = value
        self.OptionsViewModel.Profile = value
        self.OnPropertyChanged("BaseFolder")
    
    #BaseFolder
    @notify_property
    def BaseFolder(self):
        return self._profile.BaseFolder

    @BaseFolder.setter
    def BaseFolder(self, value):
        self._profile.BaseFolder = value

    #InputVisible
    @notify_property
    def IsInputVisible(self):
        return self._input_is_visible

    @IsInputVisible.setter
    def IsInputVisible(self, value):
        self._input_is_visible = value

    def select_base_folder(self):
        """
        Shows a folder browser dialog and sets the BaseFolder to the 
        selected folder
        """
        c = CommonOpenFileDialog()
        c.IsFolderPicker = True
        if c.ShowDialog() == CommonFileDialogResult.Ok:
            self.BaseFolder = c.FileName

    def add_new_profile(self, name):
        new_profile = Profile()
        new_profile.Name = name
        self.Profiles.Add(new_profile)
        self._profile_names.append(name)
        self.Profile = new_profile
        self.IsInputVisible = False
       

class ConfigureFormFileFolderViewModel(ViewModelBase):
    """
    Controls the Files/Folder template builder page of the configureform
    """
    def __init__(self):
        super(ConfigureFormFileFolderViewModel, self).__init__()
        self._field_options = NumberInsertViewModel("Number")
        self.template_field_selectors = sorted(
                [FIELDS.get_item_by_field(field) for field in template_fields], 
                key=lambda x: x.name, cmp=PythonLocale.strcoll)
        self._selectedField = TemplateItem("", "", "", "")
        self.ConditionalViewModel = ConditionalInsertViewModel()

        #Commands
        self.InsertFieldCommand = Command(self.insert_field)
        self.InsertFolderCommand = Command(self.insert_folder)

        #File
        self._file_selection_start = 0
        self.FileSelectionLength = 0
        self._file_template = ""
        self.FileVisible = False

        #Folder
        self._folder_selection_start = 0
        self.FolderSelectionLength = 0
        self._folder_template = ""
        
        #profile
        self._profile = None

        

    #Profile
    @notify_property
    def Profile(self):
        return self._profile

    @Profile.setter
    def Profile(self, value):
        self._profile = value
        self.OnPropertyChanged("FileTemplate")
        self.OnPropertyChanged("FolderTemplate")
        self.FileSelectionStart = len(self.FileTemplate)
        self.FolderSelectionStart = len(self.FolderTemplate)


    def insert_field(self):
        """
        Generates the template from the currently selected field and options and
        inserts it into the file or folder template. It keeps the current
        selection in the text boxes.
        """
        template = ""
        original_template = ""
        start = 0
        length = 0
        if self.SelectedField.type == "Conditional":
            template = self.ConditionalViewModel.make_template(self.Profile.AutoSpaceFields)
        else:
            template = self.FieldOptions.make_template(self.Profile.AutoSpaceFields)

        if self.FileVisible:
            start = self._file_selection_start
            length = self.FileSelectionLength
            self.FileTemplate = self.insert_item_into_template(start,
                                                               length,
                                                               self.FileTemplate,
                                                               template)

            self.FileSelectionStart = start + len(template)
        else:
            start = self._folder_selection_start
            length = self.FolderSelectionStart
            self.FolderTemplate = self.insert_item_into_template(start,
                                                                 length,
                                                                 self.FolderTemplate,
                                                                 template)

            self.FolderSelectionStart = start + len(template)

    def insert_item_into_template(self, start, length, template, new_item):
        """
        Inserts a new item into a string while replacing a length of the
        original string.

        Args:
            start: The (int)start position to insert
            length: The (int)length of the text to remove when inserting
            template: The original template string
            new_item: The string to insert into the template

        Returns:
            The template with the new item inserted into the existing template.
        """
        return "".join((template[:start], new_item, template[start + length:]))

    def insert_folder(self):
        """
        Inserts a folder separator character "\" into the folder template
        """
        start = self.FolderSelectionStart
        length = self.FolderSelectionLength
        self.FolderTemplate = self.insert_item_into_template(start, 
                                                             length, 
                                                             self.FolderTemplate, 
                                                             "\\")
        self.FolderSelectionStart = start + 1


    #FieldOptions
    @notify_property
    def FieldOptions(self):
        return self._field_options

    @FieldOptions.setter
    def FieldOptions(self, value):
        self._field_options = value
    
    #SelectedField
    @property
    def SelectedField(self):
        return self._selectedField

    @SelectedField.setter
    def SelectedField(self, value):
        self._selectedField = value
        if type(value) == TemplateItem:
            self.FieldOptions = SelectFieldTemplateFromType(
                    value.type, value.template, value.name)

    #FileTemplate
    @notify_property
    def FileTemplate(self):
        return self._profile.FileTemplate

    @FileTemplate.setter
    def FileTemplate(self, value):
        self._profile.FileTemplate = value

    #FolderTemplate
    @notify_property
    def FolderTemplate(self):
        return self._profile.FolderTemplate

    @FolderTemplate.setter
    def FolderTemplate(self, value):
        self._profile.FolderTemplate = value

    #FileSelectionStart
    @notify_property
    def FileSelectionStart(self):
        return self._file_selection_start

    @FileSelectionStart.setter
    def FileSelectionStart(self, value):
        self._file_selection_start = value

    #FolderSelectionStart
    @notify_property
    def FolderSelectionStart(self):
        return self._folder_selection_start

    @FolderSelectionStart.setter
    def FolderSelectionStart(self, value):
        self._folder_selection_start = value


class ConfigureFormOptionsViewModel(ViewModelBase):
    """
    Coordinates between fields in the profile/global options and the view.
    """
    
    empty_fields = None

    def __init__(self, global_settings):
        super(ConfigureFormOptionsViewModel, self).__init__()
        self._profile = None
        self.global_settings = global_settings

        self._add_illegal_char_checked = False

        if self.empty_fields is None:
            self.empty_fields = [FIELDS.get_item_by_field(field) 
                for field in template_fields 
                if field not in library_organizer_fields]

            #Sorting with PythonLocale to fix sorting localized strings
            self.empty_fields.sort(key=lambda x: x.name, 
                                   cmp=PythonLocale.strcoll)

        self._selected_empty_field = self.empty_fields[0].field
        self._selected_failed_fields = ObservableCollection[TemplateItem]()
        self._selected_failed_fields.CollectionChanged += self.failed_fields_changed

        self.excluded_empty_folders = (ObservableCollection[str](global_settings.excluded_empty_folders))
        self.excluded_empty_folders.CollectionChanged += self.excluded_empty_folders_changed

        self.illegal_characters = (ObservableCollection[str](global_settings.illegal_character_replacements.keys()))
        self.illegal_characters.CollectionChanged += self.illegal_characters_changed
        self._selected_illegal_character = self.illegal_characters[0]

        self.months = sorted(global_settings.month_names.keys(), key=int)
        self._selected_month = self.months[0]

        #Commands
        self.AddExcludedFolderCommand = Command(self.add_excluded_folder)
        self.RemoveExcludedFolderCommand = Command(lambda x: self.excluded_empty_folders.Remove(x),
                                                   lambda x: x is not None,
                                                   True)
        self.RemoveIllegalCharacterCommand = Command(lambda x: self.illegal_characters.Remove(x),
                                                     lambda x: x in self.illegal_characters and x not in REQUIRED_ILLEGAL_CHARS,
                                                     True)
        self.AddIllegalCharacterCommand = Command(self.add_illegal_character,
                                                  lambda x: x and x not in self.illegal_characters,
                                                  True)
    def add_illegal_character(self, char):
        if char:
            self.illegal_characters.Add(char)
            self.AddIllegalCharacterIsChecked = False

    def add_excluded_folder(self):
        """Shows a folder browser and if the folder is not already
        in the list, added it to the excluded_empty_folders list"""

        c = CommonOpenFileDialog()
        c.IsFolderPicker = True
        if c.ShowDialog() == CommonFileDialogResult.Ok:
            if c.FileName not in self.excluded_empty_folders:
                self.excluded_empty_folders.Add(c.FileName)

    """ Profile needs to be a notify property so the view knows when
    the profile changes """
    @notify_property
    def Profile(self):
        return self._profile

    @Profile.setter
    def Profile(self, value):
        self._profile = value
        self.OnPropertyChanged("EmptyFieldReplacement")
        self.SelectedFailedFields = (ObservableCollection[TemplateItem]([FIELDS.get_item_by_field(field) 
                                       for field in self._profile.FailedFields]))

    # SelectedIllegalCharacter
    # Needs to be a notify property so the view can update the
    # replacement textbox.
    @notify_property
    def SelectedIllegalCharacter(self):
        return self._selected_illegal_character

    @SelectedIllegalCharacter.setter
    def SelectedIllegalCharacter(self, value):
        self._selected_illegal_character = value
        self.OnPropertyChanged("IllegalCharacterReplacement")

    # IllegalCharacterReplacement
    # Needs to be a notify property so the view can update the
    # replacement textbox.
    @notify_property
    def IllegalCharacterReplacement(self):
        return (self.global_settings.illegal_character_replacements[self.SelectedIllegalCharacter])

    @IllegalCharacterReplacement.setter
    def IllegalCharacterReplacement(self, value):
        self.global_settings.illegal_character_replacements[self.SelectedIllegalCharacter] = value

    @notify_property
    def AddIllegalCharacterIsChecked(self):
        return self._add_illegal_char_checked

    @AddIllegalCharacterIsChecked.setter
    def AddIllegalCharacterIsChecked(self, value):
        self._add_illegal_char_checked = value

    # SelectedMonth
    # Needs to be a notify property so the view can update the
    # replacement textbox.
    @notify_property
    def SelectedMonth(self):
        return self._selected_month

    @SelectedMonth.setter
    def SelectedMonth(self, value):
        self._selected_month = value
        self.OnPropertyChanged("MonthReplacement")

    # MonthReplacement
    # Needs to be a notify property so the view can update the
    # replacement textbox.
    @notify_property
    def MonthReplacement(self):
        return self.global_settings.month_names[self.SelectedMonth]

    @MonthReplacement.setter
    def MonthReplacement(self, value):
        self.global_settings.month_names[self.SelectedMonth] = value

    # SelectedEmptyField
    # Needs to be a notify property so the view can update the
    # replacement textbox.
    @notify_property
    def SelectedEmptyField(self):
        return self._selected_empty_field

    @SelectedEmptyField.setter
    def SelectedEmptyField(self, value):
        self._selected_empty_field = value
        self.OnPropertyChanged("EmptyFieldReplacement")
    
    # EmptyFieldReplacement
    # Needs to be a notify property in order to update the view when profile or
    # selected empty field changes.
    @notify_property
    def EmptyFieldReplacement(self):
        if self._profile.EmptyData.has_key(self._selected_empty_field):
            return self._profile.EmptyData[self._selected_empty_field]
        else:
            return ""
    
    @EmptyFieldReplacement.setter
    def EmptyFieldReplacement(self, value):
        self._profile.EmptyData[self._selected_empty_field] = value

    #SelectedFailedFields
    @notify_property
    def SelectedFailedFields(self):
        return self._selected_failed_fields
    
    @SelectedFailedFields.setter
    def SelectedFailedFields(self, value):
        self._selected_failed_fields.CollectionChanged -= self.failed_fields_changed
        self._selected_failed_fields = value
        self._selected_failed_fields.CollectionChanged += self.failed_fields_changed

    def failed_fields_changed(self, sender, e):
        """ Handles adding and deleting the failed fields for the current
        profile's failed fields
        """
        if e.Action == NotifyCollectionChangedAction.Add:
            for i in e.NewItems:
                self._profile.FailedFields.append(i.field)
        elif e.Action == NotifyCollectionChangedAction.Remove:
            for i in e.OldItems:
                self._profile.FailedFields.remove(i.field)

    def illegal_characters_changed(self, sender, e):
        """ Handles adding and deleting illegal characters"""
        if e.Action == NotifyCollectionChangedAction.Add:
            for i in e.NewItems:
                self.global_settings.illegal_character_replacements[i] = ""
                self.SelectedIllegalCharacter = i
        elif e.Action == NotifyCollectionChangedAction.Remove:
            for i in e.OldItems:
                index = e.OldStartingIndex
                if index > len(self.illegal_characters) - 1:
                    index -= 1
                del(self.global_settings.illegal_character_replacements[i])
                self.SelectedIllegalCharacter = self.illegal_characters[index]
                

    def excluded_empty_folders_changed(self, sender, e):
        """ Updates the global settings excluded empty folders when a
        folder is removed or added
        """
        if e.Action == NotifyCollectionChangedAction.Add:
            for i in e.NewItems:
                self.global_settings.excluded_empty_folders.append(i)
        elif e.Action == NotifyCollectionChangedAction.Remove:
            for i in e.OldItems:
                self.global_settings.excluded_empty_folders.remove(i)