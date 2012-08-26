import wpf
import clr

clr.AddReferenceByPartialName("WPFToolkit.Extended")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
clr.AddReference("Ookii.Dialogs.Wpf")
clr.AddReference("Ookii.Dialogs")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\ComicRack.Engine.dll")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\cYo.Common.dll")

import System
import pyevent
import localizer
import locommon
import Ookii.Dialogs.Wpf

from System import Single, Double, Int64
from System.IO import Path
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import Dictionary, SortedDictionary
from System.Windows import Window, PropertyPath, Visibility
from System.Windows.Controls import ValidationRule, ValidationResult, TextBox, DataTemplateSelector, ComboBox, Grid
from System.Windows.Controls.Primitives import Popup
from System.Windows.Data import Binding, BindingMode, BindingOperations, UpdateSourceTrigger, IValueConverter
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Forms import MessageBox, FolderBrowserDialog, DialogResult
from System.Windows.Media import VisualTreeHelper


from Ookii.Dialogs.Wpf import VistaFolderBrowserDialog
from Ookii.Dialogs import InputDialog

from Xceed.Wpf.Toolkit import DropDownButton

from locommon import SCRIPTDIRECTORY, Translations, ExcludeGroup, ExcludeRule, NotifyPropertyChangedBase, notify_property, Mode, template_fields, exclude_rule_fields
from cYo.Projects.ComicRack.Engine import YesNo, ComicBook, MangaYesNo

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


        self._tab_index = 0
        
        self.translations = Translations()
        
        self.filelessformats = [".bmp", ".jpg", ".png"]

        #Setup some things need for the Illegal characters
        self.illegal_characters = ObservableCollection[str](self.profile.IllegalCharacters.keys())
        self.default_illegal_characters = ["?", "/", "\\", "*", ":", "<", ">", "|", "\""]
        translated_fields = localizer.get_comic_fields()
        self.exclude_rule_field_names = SortedDictionary[str, str]({name : property for name, property in translated_fields.iteritems() if property in exclude_rule_fields})
        self.exclude_rule_string_operators = SortedDictionary[str, str](localizer.get_exclude_rule_string_operators())
        self.exclude_rule_numeric_operators = SortedDictionary[str, str](localizer.get_exclude_rule_numeric_operators())
        self.exclude_rule_yes_no_operators = SortedDictionary[str, str](localizer.get_exclude_rule_yes_no_operators())
        self.exclude_rule_manga_yes_no_operators = SortedDictionary[str, str](localizer.get_exclude_rule_manga_yes_no_operators())
        self.exclude_rule_bool_operators = SortedDictionary[str, str](localizer.get_exclude_rule_bool_operators())
        #Load the xaml file
        wpf.LoadComponent(self, 'ConfigureForm.xaml')
        #Set the current profile
        #self.template_field_names = []
        #for name, property in translated_fields.iteritems():
        #    if property in template_fields:
                #self.template_field_names.append(name)

        self.TemplateFieldSelector.ItemsSource = sorted([name for name, property in translated_fields.iteritems() if property in template_fields])


        print self.exclude_rule_field_names
        self.ProfileSelector.SetValue(ComboBox.SelectedItemProperty, lastused)
        self.illegal_characters_selector.SelectedIndex = 0
        self.MonthSelector.SelectedIndex = 0

        self.folder_dialog = VistaFolderBrowserDialog()


    def Button_Browse_Click(self, sender, e):
        """
        Shows a folderbrowser dialog and sets the basefolder to the selected path.
        """
        
        self.folder_dialog.ShowDialog()
        self.BaseFolder.SetValue(TextBox.TextProperty, folder_dialog.SelectedPath)


    def Page_Button_Checked(self, sender, e):
        """
        Handler to switch pages in the form
        """
        self.PagesContainer.SelectedIndex = int(sender.Tag)

    
    #These two methods load and save the mode radio buttons
    def mode_check_changed(self, sender, e):
        """
        Saves the mode to the profile when a different mode is selected
        """
        #There has to be a better way to do the radio buttons but I couldn't find a way
        if sender == self.Move:
            self.profile.Mode = "Move"
        elif sender == self.Copy:
            self.profile.Mode = "Copy"
        elif sender == self.Simulate:
            self.profile.Mode = "Simulate"


    def load_mode(self):
        """
        Sets the correct mode radio box check from the values in the profile.
        """
        #There has to be a better way to do the radio buttons but I couldn't find a way
        if self.profile.Mode == Mode.Move:
            self.Move.IsChecked = True
        elif self.profile.Mode == Mode.Copy:
            self.Copy.IsChecked = True
        elif self.profile.Mode == Mode.Simulate:
            self.Simulate.IsChecked = True


    def drop_down_button_clicked(self, sender, e):
        """
        This method fakes a dropdown button using the contextmenu.
        """
        sender.ContextMenu.IsEnabled = True
        sender.ContextMenu.PlacementTarget = sender
        sender.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
        sender.ContextMenu.IsOpen = True

    
    def Profile_SelectionChanged(self, sender, e):
        self.profile = self.profiles[sender.SelectedItem]
        index = self.illegal_characters_selector.SelectedIndex
        self.illegal_characters.Clear()
        for i in self.profile.IllegalCharacters:
            self.illegal_characters.Add(i)
        if self.illegal_characters.Count < index - 1:
            index -= 1
        

        self.load_mode()
        self.months_selection_changed(self.MonthSelector, None)
        self.PagesContainer.DataContext = self.profiles[sender.SelectedItem]
        self.illegal_characters_selector.SelectedIndex = index


    def add_profile(self, sender, e):
        pass


    # These 6 methods controls the addition, removal and management of exclude rules/groups
    def add_rule(self, sender, e):
        sender.DataContext.add_rule(ExcludeRule())


    def rule_add_rule(self, sender, e):
        sender.DataContext.parent.insert_rule_after_rule(ExcludeRule(), sender.DataContext)


    def add_rule_group(self, sender, e):
        sender.DataContext.add_group()


    def rule_add_rule_group(self, sender, e):
        group = ExcludeGroup()
        group.add_rule(ExcludeRule())
        sender.DataContext.parent.insert_rule_after_rule(group, sender.DataContext)


    def remove_rule(self, sender, e):
        sender.DataContext.parent.remove_rule(sender.DataContext)


    def exclude_rules_field_selection_changed(self, sender, e):
        
        field_type = type(getattr(self.comicbook, sender.SelectedValue))

        if field_type is str:
            sender.DataContext.Type = "String"
            if sender.DataContext.operator not in self.exclude_rule_string_operators.Values:
                sender.DataContext.Operator = "is"

        elif field_type in [int, float, Single, Double, Int64]:
            sender.DataContext.Type = "Numeric"
            if sender.DataContext.operator not in self.exclude_rule_numeric_operators.Values:
                sender.DataContext.Operator = "is"

        elif field_type is bool:
            sender.DataContext.Type = "Bool"
            if sender.DataContext.operator not in self.exclude_rule_bool_operators.Values:
                sender.DataContext.Operator = "is True"

        elif field_type is YesNo:
            sender.DataContext.Type = "YesNo"
            if sender.DataContext.operator not in self.exclude_rule_yes_no_operators.Values:
                sender.DataContext.Operator = "is Yes"

        elif field_type is MangaYesNo:
            sender.DataContext.Type = "MangaYesNo"
            if sender.DataContext.operator not in self.exclude_rule_manga_yes_no_operators.Values:
                sender.DataContext.Operator = "is Yes"



    #These are for the various combobox/textbox options
    def months_selection_changed(self, sender, e = None):
        if sender.SelectedIndex == -1:
            sender.SelectedIndex = 0
            return
        #Note:  I tried to use a binding here in a similar function as the illegal characters
        #       but for some reason the integer index doesn't work with python dicts. I may fix this later by changing it to string indexes
        #       but for now it works.
        self.MonthName.Text = self.profile.Months[int(self.MonthSelector.SelectedItem)]


    def months_text_changed(self, sender, e):
        self.profile.Months[int(self.MonthSelector.SelectedItem)] = sender.Text


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
    """
    This is for selecting the right datatemplate for the exclude rules
    """
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


class ConfigureFormViewModel(NotifyPropertyChangedBase):
    

    def __init__(self, profiles):
        super(ConfigureFormViewModel, self).__init__()



    


class ConfigureFormViewModel2(NotifyPropertyChangedBase):

    def __init__(self, months, illegals, empty):
        super(ConfigureFormViewModel, self).__init__()
        self._months = months
        self._month_index = 1
        self._illegals = illegals
        self.IllegalCharacters = ObservableCollection[str](illegals.keys())
        self._empty = empty
        self._illegal_index = 1


    @notify_property
    def MonthName(self):
        return self._months[self._month_index]

    @MonthName.setter
    def MonthName(self, value):
        self._months[self._month_index] = value

    @notify_property
    def MonthIndex(self):
        return self._month_index

    @MonthIndex.setter
    def MonthIndex(self, value):
        self._month_index = value
        self.OnPropertyChanged("MonthName")


    @notify_property
    def IllegalIndex(self):
        return self._illegal_index

    @IllegalIndex.setter
    def IllegalIndex(self, value):
        pass

    def add_illegal_character(self, character):
        self._illegals[character] = ""
        self.IllegalCharacters.Add(character)



    def remove_character(self, character):
        pass
    def change_months(self, months):
        self._months = months
        self.OnPropertyChanged("MonthName") 