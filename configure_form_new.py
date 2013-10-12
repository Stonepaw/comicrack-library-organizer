import clr
clr.AddReference("IronPython.Wpf")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("IronPython.Modules")
from IronPython.Modules import PythonLocale
PythonLocale.setlocale(PythonLocale.LC_ALL, "")
from IronPython.Modules import Wpf
from System.Collections.Generic import SortedDictionary, Dictionary
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, Visibility
from System.Windows.Controls import Grid
from System.Windows.Media import SolidColorBrush, Colors
from System import DateTime
from wpfutils import ViewModelBase, notify_property, Command
from System.Windows.Data import IValueConverter, Binding
clr.AddReference("System.Windows.Forms")
from System.Windows.Controls import DataTemplateSelector
from fieldmappings import FIELDS, TemplateItem, template_fields, first_letter_fields, conditional_fields, conditional_then_else_fields
from locommon import SCRIPTDIRECTORY, date_formats
from System.IO import Path
from losettings import Profile
import localizer
#clr.AddReferenceToFile("CodeBoxControl.dll")
#from CodeBoxControl.Decorations import MultiStringDecoration, RegexGroupDecoration
clr.AddReferenceToFile("Microsoft.WindowsAPICodePack.dll")
clr.AddReferenceToFile("Microsoft.WindowsAPICodePack.Shell.dll")
from Microsoft.WindowsAPICodePack.Dialogs import CommonOpenFileDialog, CommonFileDialogResult
clr.AddReferenceToFile("Ookii.Dialogs.dll")
clr.AddReferenceToFile("Ookii.Dialogs.Wpf.dll")

from codeboxdecorations import (LibraryOrganizerNameDecoration, 
                                LibraryOrganizerArgsDecoration,
                                LibraryOrganizerPrefixSuffixDecoration)


class ConfigureForm(Window):
    def __init__(self, profiles, last_used_profiles):
        self.ViewModel = ConfigureFormViewModel(profiles, last_used_profiles)
        self.DataContext = self.ViewModel
        self.Resources.Add("InsertFieldTemplateSelector", 
                           InsertFieldTemplateSelector())
        self.Resources.Add("ComparisonConverter", ComparisonConverter())
        Wpf.LoadComponent(self, Path.Combine(SCRIPTDIRECTORY, 
                                             'ConfigureFormNew.xaml'))
        self.setup_text_highlighting();

    def setup_text_highlighting(self):
        """
        Setups up the text highlighting for the template text boxes
        """
        names = LibraryOrganizerNameDecoration()
        names.Brush = SolidColorBrush(Colors.Blue)
        self.FileTemplateTextBox.Decorations.Add(names);
        self.FolderTemplateBox.Decorations.Add(names);

        prefix = LibraryOrganizerPrefixSuffixDecoration()
        prefix.Brush = SolidColorBrush(Colors.Teal)
        self.FileTemplateTextBox.Decorations.Add(prefix);
        self.FolderTemplateBox.Decorations.Add(prefix);

        args = LibraryOrganizerArgsDecoration();
        args.Brush = SolidColorBrush(Colors.Red);
        self.FileTemplateTextBox.Decorations.Add(args);
        self.FolderTemplateBox.Decorations.Add(args);

    def new_profile_clicked(self, *args):
        self.ProfileNameInputBox.Text = ""
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Visible)
        self.ProfileNameInputBox.Focus()

    def close_inputbox(self, *args):
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Collapsed)

    


class ConfigureFormViewModel(ViewModelBase):
    def __init__(self, profiles, last_used_profiles):
        super(ConfigureFormViewModel, self).__init__()
        self.FileFolderViewModel = ConfigureFormFileFolderViewModel();
        self.Profiles = ObservableCollection[Profile](profiles.values())
        self._profile = None
        self._profile_names = profiles.keys()
        self.Profile = self.Profiles[0]
        self._input_is_visible = False;
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
        self.template_field_selectors = sorted([FIELDS.get_item_by_field(field) 
                                                for field in template_fields], 
                                                key=lambda x: x.name,
                                                cmp=PythonLocale.strcoll)
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
            self.FieldOptions = SelectFieldTemplateFromType(value.type, 
                                                            value.template, 
                                                            value.name)

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

           
class InsertViewModel(object):
    """
    The insert view model for a string. Also the base insert view model
    that other insert view models inherit from.
    """
    def __init__(self, template):
        self.Prefix = ""
        self.Suffix = ""
        self._template = template

    def make_template(self, autospace, args=""):
        """
        Creates the template string with the provided args.

        Args:
            autospace: Inserts a space before the prefix
            args: (Optional)The args to insert into the template

        Returns:
            The string of the constructed template
        """
        if autospace:
            return "{ %s<%s%s>%s}" % (self.Prefix, self._template, args, 
                                      self.Suffix)
        return "{%s<%s%s>%s}" % (self.Prefix, self._template, args, 
                                 self.Suffix)


class NumberInsertViewModel(InsertViewModel):
    """
    The insert view model for a number field.
    """
    def __init__(self, field):
        self.Padding = 0
        super(NumberInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a number field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)" % (self.Padding)
        return super(NumberInsertViewModel, self).make_template(autospace, args)


class MultipleValueInsertViewModel(InsertViewModel):
    """
    The insert view model for a multiple value field.
    """
    def __init__(self, template, name):
        self.SelectWhichMultipleValue = False
        self.SelectOnceForEachSeries = False
        self.SelectWhichText = "Select which %s to insert" % (name.lower())
        self.Seperator = ""
        super(MultipleValueInsertViewModel, self).__init__(template)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a 
        multiple value field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = ""
        if self.SelectWhichMultipleValue:
            if self.SelectOnceForEachSeries:
                args = "(%s)(series)" % (self.Seperator)
            else:
                args = "(%s)(issue)" % (self.Seperator)
        return super(MultipleValueInsertViewModel, self).make_template(autospace, args)


class MonthValueInsertViewModel(InsertViewModel):
    """
    The insert view model for a month field.
    """
    def __init__(self, template):
        self.UseMonthNames = True
        self.Padding = 0
        super(MonthValueInsertViewModel, self).__init__(template)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a month field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = ""
        if not self.UseMonthNames:
            args = "(%s)" % (self.Padding)
        return super(MonthValueInsertViewModel, self).make_template(autospace, args)


class CounterInsertViewModel(InsertViewModel):
    """
    The insert view model for a counter field.
    """
    def __init__(self, field):
        self.Padding = 0
        self.Start = 1
        self.Increment = 1
        super(CounterInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a counter field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)(%s)(%s)" % (self.Start, self.Increment, self.Padding)
        return super(CounterInsertViewModel, self).make_template(autospace, args)


class DateInsertViewModel(InsertViewModel):
    """
    The insert view model for a date field.
    """
    DateFormats = SortedDictionary[str, str]()
    
    def __init__(self, field):
        if self.DateFormats.Count == 0:
            self.DateFormats["Custom"] = "Custom"
            date = DateTime.Now
            for d in date_formats:
                self.DateFormats[date.ToString(d)] = d
        self.SelectedDateFormat = "MMMM dd, yyyy"
        self.CustomDateFormat = ""
        super(DateInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a date field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        if self.SelectedDateFormat == "Custom":
            args = "(%s)" % (self.CustomDateFormat)
        else:
            args = "(%s)" % (self.SelectedDateFormat)
        return super(DateInsertViewModel, self).make_template(autospace, args)


class YesNoInsertViewModel(InsertViewModel):
    """
    The insert view model for a YesNo field.
    """
    YesNoOperators = SortedDictionary[str, str](localizer.get_yes_no_operators())

    def __init__(self, field):
        self.SelectedYesNo = "Yes"
        self.Invert = False
        self.TextToInsert = ""
        super(YesNoInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a Yes No field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)(%s)" % (self.TextToInsert, self.SelectedYesNo)
        if self.Invert:
            args += "(!)"
        return super(YesNoInsertViewModel, self).make_template(autospace, args)


class MangaYesNoInsertViewModel(YesNoInsertViewModel):
    """
    The insert view model for a MangaYesNo field.
    """

    YesNoOperators = SortedDictionary[str, str](localizer.get_manga_yes_no_operators())


class FirstLetterInsertViewModel(InsertViewModel):
    """
    The insert view model for a first letter field.
    """
    FirstLetterFields = []

    def __init__(self, field):
        if not self.FirstLetterFields:
            self.FirstLetterFields = [FIELDS.get_item_by_field(f) for f in first_letter_fields]
            self.FirstLetterFields.sort(key=lambda x: x.name, cmp=PythonLocale.strcoll)
        self.SelectedField = "ShadowSeries"
        super(FirstLetterInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a first letter field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)" % (self.SelectedField)
        return super(FirstLetterInsertViewModel, self).make_template(autospace, args)


class ReadPercentageInsertViewModel(InsertViewModel):
    """
    The insert view model for the read percentage field.
    """
    ReadPercentageOperators = ["=", "<", ">"]

    def __init__(self, field):
        self.TextToInsert = ""
        self.Operator = "="
        self.Percent = 90
        super(ReadPercentageInsertViewModel, self).__init__(field)

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a read percentage field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)(%s)(%s)" % (self.TextToInsert, self.Operator, self.Percent)
        return super(ReadPercentageInsertViewModel, self).make_template(autospace, args)


class ConditionalInsertViewModel(ViewModelBase):
    """
    The view model for a conditional field.
    """

    ConditionalStringOperators = Dictionary[str, str](localizer.get_conditional_string_operators())
    ConditionalNumberOperators = Dictionary[str, str]({"=" : "=", "<": "<", ">" :  ">"})
    ConditionalYesNoOperators = Dictionary[str, str](localizer.get_yes_no_operators())
    ConditionalMangaYesNoOperators = Dictionary[str, str](localizer.get_manga_yes_no_operators())
    ConditionalDateOperators = Dictionary[str, str](localizer.get_date_operators())

    def __init__(self):
        super(ConditionalInsertViewModel, self).__init__()
        self.ConditionalFields = sorted([FIELDS.get_item_by_field(f) for f in conditional_fields],
                                        PythonLocale.strcoll,
                                        lambda x: x.name)
        self.ConditionalThenElseFields = sorted([FIELDS.get_item_by_field(f) for f in conditional_then_else_fields],
                                        PythonLocale.strcoll,
                                        lambda x: x.name)
        self._selected_conditional_field = None
        self._selected_then_field = None
        self._selected_else_field = None
        self._conditional_operators = None
        self._selected_conditional_operator = self.ConditionalStringOperators["is"]
        self._then_field_options = InsertViewModel("")
        self._else_field_options = InsertViewModel("")
        self.SelectedConditionalField = self.ConditionalFields[0]
        self.SelectedElseField = self.ConditionalThenElseFields[0]
        self.SelectedThenField = self.ConditionalThenElseFields[0]
        self.UseElse = False

        #
        self.StringValue = ""
        self.NumberValue = 0
        self.DateValue = DateTime.Now
        self.Invert = False
        
    #SelectedConditionalField
    @notify_property
    def SelectedConditionalField(self):
        return self._selected_conditional_field

    @SelectedConditionalField.setter
    def SelectedConditionalField(self, value):
        if value is None:
            return
        if self._selected_conditional_field is None:
            old_type = None
        else:
            old_type = self._selected_conditional_field.type
        if old_type != value.type:
            if value.type in ("Number", "ReadPercentage", "Year", "Month"):
                self.ConditionalOperators = self.ConditionalNumberOperators
                self.SelectedConditionalOperator = "="
            if value.type == "YesNo":
                self.ConditionalOperators = self.ConditionalYesNoOperators
                self.SelectedConditionalOperator = "Yes"
            if value.type == "DateTime":
                self.ConditionalOperators = self.ConditionalDateOperators
                self.SelectedConditionalOperator = "is"
            if value.type == "MangaYesNo":
                self.ConditionalOperators = self.ConditionalMangaYesNoOperators
                self.SelectedConditionalOperator = "Yes"
            if value.type == "String" or value.type == "MultipleValue":
                self.ConditionalOperators = self.ConditionalStringOperators
                self.SelectedConditionalOperator = "is"
        self._selected_conditional_field = value
    
    #SelectedConditionalOperator
    @notify_property
    def SelectedConditionalOperator(self):
        return self._selected_conditional_operator

    @SelectedConditionalOperator.setter
    def SelectedConditionalOperator(self, value):
        self._selected_conditional_operator = value

    #ConditionalOperators
    @notify_property
    def ConditionalOperators(self):
        return self._conditional_operators

    @ConditionalOperators.setter
    def ConditionalOperators(self, value):
        self._conditional_operators = value

    #SelectedThenField
    @notify_property
    def SelectedThenField(self):
        return self._selected_then_field

    @SelectedThenField.setter
    def SelectedThenField(self, value):
        if value is None:
            return
        self._selected_then_field = value
        self.ThenFieldOptions = SelectFieldTemplateFromType(value.type, 
                                                            value.template, 
                                                            value.name)

    #ThenFieldOptions
    @notify_property
    def ThenFieldOptions(self):
        return self._then_field_options

    @ThenFieldOptions.setter
    def ThenFieldOptions(self, value):
        self._then_field_options = value

    #SelectedElseField
    @notify_property
    def SelectedElseField(self):
        return self._selected_else_field

    @SelectedElseField.setter
    def SelectedElseField(self, value):
        if value is not None:
            self._selected_else_field = value
            self.ElseFieldOptions = SelectFieldTemplateFromType(value.type, 
                                                                value.template, 
                                                                value.name)
        else:
            self.SelectedElseField = self.ConditionalThenElseFields[0]

    #ElseFieldOptions
    @notify_property
    def ElseFieldOptions(self):
        return self._else_field_options

    @ElseFieldOptions.setter
    def ElseFieldOptions(self, value):
        self._else_field_options = value

    def make_template(self, autospace):
        """
        Creates the template for the conditional field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The created conditional template
        """
        template = ""
        if self.Invert:
            template = "!%s(%s)" % (self.SelectedConditionalField.template, 
                                    self.SelectedConditionalOperator)
        else:
            template = "?%s(%s)" % (self.SelectedConditionalField.template, 
                                    self.SelectedConditionalOperator)

        if self.SelectedConditionalField.type in ("Number", "ReadPercentage", 
                                                  "Year", "Month"):
            template += "(%s)" % (self.NumberValue)
        elif self.SelectedConditionalField.type in ("String", "MultipleValue"):
            template += "(%s)" % (self.StringValue)
        elif self.SelectedConditionalField.type == "DateTime":
            template += "(%s)" % (self.DateValue.ToShortDateString())

        if self.UseElse:
            return "{%s<%s>%s}" % (self.ElseFieldOptions.make_template(autospace), 
                                   template,
                                   self.ThenFieldOptions.make_template(autospace))
        else:
            return "{<%s>%s}" % (template, self.ThenFieldOptions.make_template(autospace))


def SelectFieldTemplateFromType(type, template, name):
    if type == "Number":
        return NumberInsertViewModel(template)
    elif type == "String" or type == "Year":
        return InsertViewModel(template)
    elif type == "Month":
        return MonthValueInsertViewModel(template)
    elif type == "Counter":
        return CounterInsertViewModel(template)
    elif type == "DateTime":
        return DateInsertViewModel(template)
    elif type == "YesNo":
        return YesNoInsertViewModel(template)
    elif type == "MangaYesNo":
        return MangaYesNoInsertViewModel(template)
    elif type == "FirstLetter":
        return FirstLetterInsertViewModel(template)
    elif type == "ReadPercentage":
        return ReadPercentageInsertViewModel(template)
    elif type == "MultipleValue":
        return MultipleValueInsertViewModel(template, name)
    elif type == "Conditional":
        return None


class InsertFieldTemplateSelector(DataTemplateSelector):
    """Selects the correct datatemplate for the insert rule view templates"""
    def SelectTemplate(self, item, container):
        if type(item) is InsertViewModel:
            return container.FindResource("StringInsertFieldTemplate")
        elif type(item) is NumberInsertViewModel:
            return container.FindResource("NumberInsertFieldTemplate")
        elif type(item) is MultipleValueInsertViewModel:
            return container.FindResource("MultipleValueInsertFieldTemplate")
        elif type(item) is MonthValueInsertViewModel:
            return container.FindResource("MonthInsertFieldTemplate")
        elif type(item) is CounterInsertViewModel:
            return container.FindResource("CounterInsertFieldTemplate")
        elif type(item) is YesNoInsertViewModel or type(item) is MangaYesNoInsertViewModel:
            return container.FindResource("YesNoInsertFieldTemplate")
        elif type(item) is DateInsertViewModel:
            return container.FindResource("DateInsertFieldTemplate")
        elif type(item) is FirstLetterInsertViewModel:
            return container.FindResource("FirstLetterInsertFieldTemplate")
        elif type(item) is ReadPercentageInsertViewModel:
            return container.FindResource("ReadPercentageInsertFieldTemplate")


class ComparisonConverter(IValueConverter):
    """A converter to databind wpf radio buttons to an enum"""
    def Convert(self, value, targetType, parameter, culture):
        return value == parameter

    def ConvertBack(self, value, targetType, parameter, culture):
        if value:
            return parameter
        else:
            return Binding.DoNothing