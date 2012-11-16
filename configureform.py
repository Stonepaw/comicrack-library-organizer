import wpf
import clr

clr.AddReferenceByPartialName("WPFToolkit.Extended")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
clr.AddReference("Ookii.Dialogs.Wpf")
clr.AddReference("Ookii.Dialogs")
clr.AddReferenceToFile("ComicRack.Engine.dll")
clr.AddReferenceToFile("cYo.Common.dll")

import System
import localizer
from System import Single, Double, Int64
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import SortedDictionary
from System.Windows import Window, PropertyPath, Visibility
from System.Windows.Controls import ValidationRule, ValidationResult, TextBox, DataTemplateSelector, ComboBox, BooleanToVisibilityConverter
from System.Windows.Controls.Primitives import Popup
from System.Windows.Data import Binding, IValueConverter

from Ookii.Dialogs.Wpf import VistaFolderBrowserDialog
from locommon import Translations, Mode, template_fields, exclude_rule_fields, multiple_value_fields, library_organizer_fields, first_letter_fields, conditional_fields, conditional_then_else_fields
from excluderules import ExcludeRule, ExcludeGroup
from cYo.Projects.ComicRack.Engine import YesNo, ComicBook, MangaYesNo
from wpfutils import notify_property, NotifyPropertyChangedBase

#Use global so it isn't contantly created in the converter class
COMICBOOK = ComicBook()

class ConfigureForm(Window):

    def __init__(self, profiles, lastused):
        #TODO: Clean this up a bit
        #Setup profiles
        self.profiles = profiles
        self.profilenames = ObservableCollection[str](profiles.keys())
        self.profile = profiles[lastused]

        self.comicbook = ComicBook()
        
        #For the exclude rules data template
        rulestemplateselector = RulesTemplateSelector()
        self.Resources.Add("rulestemplateselector", rulestemplateselector)
        self.field_name_to_type_name_converter = FieldNameToTypeNameConverter()
        self.Resources.Add("FieldNameToTypeNameConverter", self.field_name_to_type_name_converter)
        fieldnametomultivaluedescription = FieldNameToMultiValueDescription()
        self.Resources.Add("FieldNameToMultiValueDescription", fieldnametomultivaluedescription)
        fieldnametovisiblityconverter = FieldNameToVisiblityConverter()
        self.Resources.Add("FieldNameToVisibilityConverter", fieldnametovisiblityconverter)
        self.Resources.Add("InverseFieldNameToVisibilityConverter", InverseFieldNameToVisiblityConverter())
        inverse_converter = InverseBooleanToVisibilityConverter()
        self.Resources.Add("InverseBooleanToVisibilityConverter", inverse_converter)
        self.translations = Translations()
        
        self.filelessformats = [".bmp", ".jpg", ".png"]

        #Setup some things need for the Illegal characters
        self.illegal_characters = ObservableCollection[str](self.profile.IllegalCharacters.keys())
        self.default_illegal_characters = ["?", "/", "\\", "*", ":", "<", ">", "|", "\""]
        self.months_selectors = sorted(self.profile.Months.keys(), key=int)
        self.localize()
        #Load the xaml file
        wpf.LoadComponent(self, 'ConfigureForm.xaml')
        self.ProfileSelector.SetValue(ComboBox.SelectedItemProperty, lastused)
        self.illegal_characters_selector.SelectedIndex = 0
        self.MonthSelector.SelectedIndex = 0

        self.folder_dialog = VistaFolderBrowserDialog()

    def localize(self):
        translated_fields = localizer.get_comic_fields()
        self.exclude_rule_field_names = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() 
                                                                    if property in exclude_rule_fields})
        self.exclude_rule_string_operators = SortedDictionary[str, str](localizer.get_exclude_rule_string_operators())
        self.exclude_rule_numeric_operators = SortedDictionary[str, str](localizer.get_exclude_rule_numeric_operators())
        self.exclude_rule_yes_no_operators = SortedDictionary[str, str](localizer.get_exclude_rule_yes_no_operators())
        self.exclude_rule_manga_yes_no_operators = SortedDictionary[str, str](localizer.get_exclude_rule_manga_yes_no_operators())
        self.exclude_rule_bool_operators = SortedDictionary[str, str](localizer.get_exclude_rule_bool_operators())
        self.template_field_selectors = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() 
                                                                    if property in template_fields})
        self.first_letter_fields_names = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() if property in first_letter_fields})
        self.conditional_field_names = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() if property in conditional_fields})
        self.conditional_then_else_field_names = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() if property in conditional_then_else_fields})

    def Button_Browse_Click(self, sender, e):
        """Shows a folder browser dialog and sets the base folder to the selected path."""        
        if self.folder_dialog.ShowDialog():
            self.BaseFolder.SetValue(TextBox.TextProperty, self.folder_dialog.SelectedPath)

    def Page_Button_Checked(self, sender, e):
        """Handler to switch pages in the form"""
        self.PagesContainer.SelectedIndex = int(sender.Tag)

    #These two methods load and save the mode radio buttons
    def mode_check_changed(self, sender, e):
        """Saves the mode to the profile when a different mode is selected"""
        #There has to be a better way to do the radio buttons but I couldn't find a way
        if sender == self.Move:
            self.profile.Mode = "Move"
        elif sender == self.Copy:
            self.profile.Mode = "Copy"
        elif sender == self.Simulate:
            self.profile.Mode = "Simulate"

    def load_mode(self):
        """Sets the correct mode radio box check from the values in the profile."""
        #There has to be a better way to do the radio buttons but I couldn't find a way
        if self.profile.Mode == Mode.Move:
            self.Move.IsChecked = True
        elif self.profile.Mode == Mode.Copy:
            self.Copy.IsChecked = True
        elif self.profile.Mode == Mode.Simulate:
            self.Simulate.IsChecked = True

    def drop_down_button_clicked(self, sender, e):
        """This method fakes a dropdown button using the contextmenu."""
        sender.ContextMenu.IsEnabled = True
        sender.ContextMenu.PlacementTarget = sender
        sender.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
        sender.ContextMenu.IsOpen = True

    def Profile_SelectionChanged(self, sender, e):
        self.profile = self.profiles[sender.SelectedItem]

        #Reload the illegal characters selector. Have to do this because of custom characters.
        index = self.illegal_characters_selector.SelectedIndex
        self.illegal_characters.Clear()
        for i in self.profile.IllegalCharacters:
            self.illegal_characters.Add(i)
        if self.illegal_characters.Count < index - 1:
            index -= 1

        self.load_mode()
        self.PagesContainer.DataContext = self.profiles[sender.SelectedItem]
        self.illegal_characters_selector.SelectedIndex = index
        self.FolderStructure.CaretIndex = len(self.profile.FolderTemplate)
        self.FileStructure.CaretIndex = len(self.profile.FileTemplate)

    def add_profile(self, sender, e):
        pass

    def insert_field_clicked(self, sender, e):
        if self.TemplateFieldSelector.SelectedValue == "Conditional":
            if_args = self.get_field_args(self.ConditionalIfField.SelectedValue, "ConditionalIf")
            if self.ConditionalIfInvert.IsChecked:
                field = "!" + self.ConditionalIfField.SelectedValue
            else:
                field = "?" + self.ConditionalIfField.SelectedValue
            if self.ConditionalIfRegex.IsChecked:
                field += "(!%s)%s" % (self.ConditionalIfMatchText.Text, if_args)
            else:
                field += "(!%s)%s" % (self.ConditionalIfMatchText.Text, if_args)
            if self.TemplateBuilderAutoSpaceFields.IsChecked:
                then_prefix = " " + self.ConditionalThenPrefix.Text
                else_prefix = " " + self.ConditionalElsePrefix.Text
            else:
                then_prefix = self.ConditionalThenPrefix.Text
                else_prefix = self.ConditionalElsePrefix.Text
            then_args = self.get_field_args(self.ConditionalThenField.SelectedValue, "ConditionalThen")
            then_field = "{%s<%s%s>%s}" % (then_prefix, self.ConditionalThenField.SelectedValue, then_args, self.ConditionalThenSuffix.Text)
            if self.ConditionalElseCheckBox.IsChecked:
                else_args = self.get_field_args(self.ConditionalElseField.SelectedValue, "ConditionalElse")
                else_field = "{%s<%s%s>%s}" % (else_prefix, self.ConditionalElseField.SelectedValue, else_args, self.ConditionalElseSuffix.Text)
            else:
                else_field = ""
            insert_text = "{%s<%s>%s}" % (else_field, field, then_field)
            self.insert_text_into_template(insert_text)

        else:
            args = self.get_field_args(self.TemplateFieldSelector.SelectedValue, "TemplateBuilder")
        #field_type = self.field_name_to_type_name_converter.Convert(self.TemplateFieldSelector.SelectedValue, None, None, None)
        #if field_type == "Numeric":
        #    args = str(self.TemplateBuilderPadding.Value)
        #elif field_type == "YesNo" or field_type == "MangaYesNo":
        #    args = "(%s)" % (self.TemplateBuilderYesNoText.Text)
        #    if self.TemplateBuilderYesNoInvert.IsChecked:
        #        args += "(!)"
        #elif field_type == "Month":
        #    pass
        #elif field_type == "FirstLetter":
        #    args = "(%s)" % (self.FirstLetterSeriesSelector.SelectedValue)
        #elif field_type == "MultipleValue":
        #    if self.TemplateBuilderSelectMultipleValue.IsChecked:                
        #        if self.TemplateBuilderMultipleValueSelectOnce.IsChecked:
        #            args = "(%s)(series)" % (self.TemplateBuilderMultipleValueSeperator.Text)
        #        else:
        #            args = "(%s)(issue)" % (self.TemplateBuilderMultipleValueSeperator.Text)
        #elif field_type == "Counter":
        #    args = "(%s)(%s)(%s)" % (self.TemplateBuilderCounterStart.Value, 
        #                             self.TemplateBuilderCounterIncrement.Value, 
        #                             self.TemplateBuilderPadding.Value)
        #elif field_type == "ReadPercentage":
        #    args = "(%s)(%s)(%s)" % (self.TemplateBuilderReadPercentageText.Text,
        #                             self.TemplateBuilderReadPercentageOperator.Text,
        #                             self.TemplateBuilderReadPercentageNumber.Value)
            if self.TemplateBuilderAutoSpaceFields.IsChecked:
                prefix = " " + self.TemplateBuilderPrefix.Text
            else:
                prefix = self.TemplateBuilderPrefix.Text

            insert_text = "{%s<%s%s>%s}" % (prefix, self.TemplateFieldSelector.SelectedItem.Value, args, self.TemplateBuilderSuffix.Text)
            self.insert_text_into_template(insert_text)

    def create_field_text(self, field_type):
        pass

    def get_field_args(self, field, arg_type):
        field_type = self.field_name_to_type_name_converter.Convert(field, None, None, None)

        if field_type == "Numeric":
            return str(self.FindName(arg_type + "Padding").Value)

        elif field_type == "YesNo" or field_type == "MangaYesNo":
            args = "(%s)" % (self.FindName(arg_type + "YesNoText").Text)
            if self.FindName(arg_type + "YesNoInvert").IsChecked:
                args += "(!)"
            return args

        elif field_type == "Month":
            return ""

        elif field_type == "FirstLetter":
            return "(%s)" % (self.FindName(arg_type + "FirstLetterSeriesSelector").SelectedValue)

        elif field_type == "MultipleValue":
            if self.FindName(arg_type + "SelectMultipleValue").IsChecked:                
                if self.FindName(arg_type + "MultipleValueSelectOnce").IsChecked:
                    return "(%s)(series)" % (self.FindName(arg_type + "MultipleValueSeperator").Text)
                else:
                    return "(%s)(issue)" % (self.FindName(arg_type + "MultipleValueSeperator").Text)

        elif field_type == "Counter":
            return "(%s)(%s)(%s)" % (self.FindName(arg_type + "CounterStart").Value, 
                                     self.FindName(arg_type + "CounterIncrement").Value, 
                                     self.FindName(arg_type + "Padding").Value)

        elif field_type == "ReadPercentage":
            args = "(%s)(%s)(%s)" % (self.TemplateBuilderReadPercentageText.Text,
                                        self.TemplateBuilderReadPercentageOperator.Text,
                                        self.TemplateBuilderReadPercentageNumber.Value)
        return ""

    def insert_folder_clicked(self, sender, e):
        self.insert_text_into_template("\\")


    def insert_text_into_template(self, insert_text):
        """Inserts text into the correct template textbox. 
        
        Replaces the selected text or inserts at the caret location."""
        if self.FolderButton.IsChecked:
            self.FolderStructure.SelectedText = insert_text
            self.FolderStructure.CaretIndex += len(insert_text)
            self.FolderStructure.SelectionLength = 0
        elif self.FileButton.IsChecked:
            self.FileStructure.SelectedText = insert_text
            self.FileStructure.CaretIndex += len(insert_text)
            self.FileStructure.SelectionLength = 0


    #These are for the various combobox/textbox options
    def months_selection_changed(self, sender, e = None):
        if sender.SelectedIndex == -1:
            sender.SelectedIndex = 0
            return
        binding = Binding()
        binding.Path = PropertyPath("Months[" + str(sender.SelectedItem) + "]")
        self.MonthName.SetBinding(TextBox.TextProperty, binding)

    def illegal_characters_selector_selection_changed(self, sender, e):
        if sender.SelectedIndex == -1:
            sender.SelectedIndex = 0
            return
        binding = Binding()
        binding.Path = PropertyPath("IllegalCharacters[" + str(sender.SelectedItem) + "]")
        self.illegal_character_textbox.SetBinding(TextBox.TextProperty, binding)

    def add_illegal_character(self, sender, e):
        if self.add_illegal_character_textbox.Text and self.add_illegal_character_textbox.Text not in self.profile.IllegalCharacters:
            self.profile.IllegalCharacters[self.add_illegal_character_textbox.Text] = ""
            self.illegal_characters.Add(self.add_illegal_character_textbox.Text)
            self.add_illegal_character_popup.SetValue(Popup.IsOpenProperty, False)
            self.illegal_characters_selector.SetValue(ComboBox.SelectedItemProperty, self.add_illegal_character_textbox.Text)
            self.add_illegal_character_textbox.Text = ""

    def remove_illegal_character(self, sender, e):
        if self.illegal_characters_selector.SelectedItem not in self.default_illegal_characters:
            del(self.profile.IllegalCharacters[self.illegal_characters_selector.SelectedItem])
            index = self.illegal_characters_selector.SelectedIndex
            self.illegal_characters.Remove(self.illegal_characters_selector.SelectedItem)
            self.illegal_characters_selector.SelectedIndex = index - 1

    def illegal_character_textbox_preview_text_input(self, sender, e):
        if e.Text in self.profile.IllegalCharacters:
            e.Handled = True

    def add_excluded_folder(self, sender, e):
        if self.folder_dialog.ShowDialog():
            self.profile.ExcludedEmptyFolder.Add(self.folder_dialog.SelectedPath)


class RulesTemplateSelector(DataTemplateSelector):
    """Selects the correct datatemplete for ExcludeRules or ExcludeGroups."""
    def SelectTemplate(self, item, container):
        if type(item) is ExcludeRule:
            return container.FindResource("ExcludeRule")
        elif type(item) is ExcludeGroup:
            return container.FindResource("ExcludeGroup")


class EmptyTextBoxValidationRule(ValidationRule):
    
    def Validate(self, value, cultureInfo):
        if not value:
            return ValidationResult(False, "Error")
        else:
            return ValidationResult.ValidResult


class FieldNameToTypeNameConverter(IValueConverter):

    def Convert(self, value, targetType, parameter, culture):
        global COMICBOOK
        if value is None:
            return ""

        if value in ("Counter", "FirstLetter", "Conditional", "Year", "StartYear", "ReadPercentage"):
            return value
        elif value in ("StartMonth","Month"):
            return "Month"
        elif value in multiple_value_fields:
            return "MultipleValue"
        if value not in library_organizer_fields and value not in ("Number", "AlternateNumber"):
            value = type(getattr(COMICBOOK, value))
        if value is YesNo:
            return "YesNo"
        elif value is MangaYesNo:
            return "MangaYesNo"
        elif value in ("Number", "AlternateNumber", "FirstIssueNumber", "LastIssueNumber", int, Double, Int64, float, Single):
            return "Numeric"
        elif value is bool:
            return "Bool"
        elif value is str:
           return "String"
        print self._template_selector_type


class FieldNameToVisiblityConverter(IValueConverter):

    fieldnameconverter = FieldNameToTypeNameConverter()

    def Convert(self, value, targetType, parameter, culture):
        """Converts the name of a field to a type and then compares it to the parameter(s).

        Args:
            parameter: Can either be a single type or a | seperated list of parameters.
        
        Returns:
            System.Windows.Visiblity.Visible or System.Windows.Visiblity.Collapsed
        """
        fieldtype = self.fieldnameconverter.Convert(value, targetType, parameter, culture)
        parameters = parameter.split("|")
        for p in parameters:
            if fieldtype == p:
                return Visibility.Visible
        return Visibility.Collapsed


class InverseFieldNameToVisiblityConverter(FieldNameToVisiblityConverter):

    def Convert(self, value, targetType, parameter, culture):
        visiblity = super(InverseFieldNameToVisiblityConverter, self).Convert(value, targetType, parameter, culture)
        if visiblity is Visibility.Visible:
            return Visibility.Collapsed
        else:
            return Visibility.Visible


class FieldNameToMultiValueDescription(IValueConverter):

    def Convert(self, value, targetType, paramater, culture):
        if value is None:
            return ""
        return "Select which %s to use" % (value.lower())


class InverseBooleanToVisibilityConverter(IValueConverter):
    """Converts a boolean value to a System.Windows.Visibility value."""
    _boolean_converter = BooleanToVisibilityConverter()

    def Convert(self, value, targetType, parameter, culture):
        """Converts a boolean value to a System.Windows.Visibility value.

        Args:
            value: The Bool to convert.
            other args are not used.

        Returns:
            System.Windows.Visibility.Collapsed when the value is True; 
            otherwise Visiblity.Collapsed is returned.
        """
        return self._boolean_converter.Convert(not value, targetType, parameter, culture)

    def ConvertBack(self, value, targetType, parameter, culture):
        """Converts a System.Windows.Visibility value to a boolean.

        Args:
            value: The System.Windows.Visibility to convert.
            other args are not used.

        Returns:
            True when the value is System.Windows.Visibility.Collapsed; 
            otherwise False is returned.
        """
        return self._boolean_converter.ConvertBack(not value, targetType, parameter, culture)
