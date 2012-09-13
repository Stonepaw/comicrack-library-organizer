import wpf
import clr
import System
from System.Collections.ObjectModel import ObservableCollection
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
from System.Windows import Window
from System.Windows.Forms import MessageBox
from System.Windows.Controls import DataTemplateSelector
from locommon import ExcludeRule, ExcludeGroup
from locommon import Translations
from System.Windows.Media import VisualTreeHelper

class rulestest(Window):
    def __init__(self, rules):
        #Note: exclude rules type programatically
        self.rules = ObservableCollection[object](rules)
        self.rulestemplateselector = RulesTemplateSelector()
        self.translations = Translations()
        self.Resources.Add("rulestemplateselector", self.rulestemplateselector)
        wpf.LoadComponent(self, 'rulestest.xaml')
    
    def add_rule(self, sender, e):
        sender.DataContext.Add(ExcludeRule("Series", "is", ""))

    def add_rule_group(self, sender, e):
        sender.DataContext.Add(ExcludeGroup("All", ObservableCollection[object]([ExcludeRule("Series", "is", "")])))

    def remove_rule(self, sender, e):
        #Probably a better way to do this but this works.
        sender.Tag.Remove(sender.DataContext)

    def rules_field_selection_changed(self, sender, e):
        """Changes the items in the rules operator and value comboboxes when the selected field is changed.
        Hiding and showing the value combobox or textbox is also done here
        """
        operator = sender.DataContext.operator
        operator_combobox = VisualTreeHelper.GetChild(sender.Parent, 1)
        value_textbox = VisualTreeHelper.GetChild(sender.Parent, 2)
        value_combobox = VisualTreeHelper.GetChild(sender.Parent, 3)
        if sender.DataContext.field in ["Manga", "BlackAndWhite", "SeriesComplete"]:
            value_textbox.Visibility = System.Windows.Visibility.Collapsed
            value_combobox.Visibility = System.Windows.Visibility.Visible
            if operator_combobox.ItemsSource != self.translations.rules_operators_yes_no:
                operator_combobox.ItemsSource = self.translations.rules_operators_yes_no
                if operator not in ["is", "is not"]:
                    operator = "is"
                sender.DataContext.Operator = operator

            if sender.DataContext.field == "Manga":
                if value_combobox.ItemsSource != self.translations.rules_values_manga:
                    value_combobox.ItemsSource = self.translations.rules_values_manga
                if sender.DataContext.value is None or sender.DataContext.value not in self.translations.rules_values_manga.Keys:
                    sender.DataContext.Value = "Yes"
            else:
                if value_combobox.ItemsSource != self.translations.rules_values:
                    value_combobox.ItemsSource = self.translations.rules_values
                if sender.DataContext.value is None or sender.DataContext.value not in self.translations.rules_values.Keys:
                    sender.DataContext.Value = "Yes"

        else:
            if operator_combobox.ItemsSource != self.translations.rules_operators:
                operator_combobox.ItemsSource = self.translations.rules_operators
                sender.DataContext.Operator = operator

                value_textbox.Visibility = System.Windows.Visibility.Visible
                value_combobox.Visibility = System.Windows.Visibility.Collapsed
                sender.DataContext.Value = ""

    
    def ComboBox_SourceUpdated(self, sender, e):
        pass

    
    def ComboBox_DataContextChanged(self, sender, e):
        pass
    
    def ComboBox_SelectionChanged(self, sender, e):
        pass

class RulesTemplateSelector(DataTemplateSelector):

    def SelectTemplate(self, item, container):
        if type(item) is ExcludeRule:
            return container.FindResource("ExcludeRule")
        elif type(item) is ExcludeGroup:
            return container.FindResource("ExcludeGroup")