import clr
import System

clr.AddReference("System.Core")
clr.AddReference("IronPython.Modules")
clr.AddReference("PresentationFramework")


clr.ImportExtensions(System.Linq)

from IronPython.Modules import PythonLocale
from System.Collections.Generic import SortedDictionary
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import DataTemplateSelector

from exclude_rules import ExcludeRuleBase, ExcludeRuleGroup
from fieldmappings import exclude_rule_fields, FIELDS, FieldType
from localizer import get_exclude_rule_bool_operators, get_exclude_rule_numeric_operators, get_exclude_rule_string_operators, get_exclude_rule_yes_no_operators, get_manga_yes_no_operators, Localizer
from locommon import get_custom_value_keys
from wpfutils import Command, notify_property, ViewModelBase

LOCALIZER = Localizer()

class ExcludeRuleCollectionViewModel(ViewModelBase):
    
    operators = SortedDictionary[str, str](LOCALIZER.all_any_operators)

    def __init__(self, exclude_rule_collection):
        self._exclude_rule_collection = exclude_rule_collection

        self.rule_view_models = ObservableCollection[object]()

        for rule in exclude_rule_collection:
            if isinstance(rule, ExcludeRuleBase):
                self.rule_view_models.Add(ExcludeRuleViewModel(rule, self))
        self.create_commands()
        return super(ExcludeRuleCollectionViewModel, self).__init__()

    def create_commands(self):
        self.AddRuleCommand = Command(self.add_rule, uses_parameter=True)
        self.AddRuleGroupCommand = Command(self.add_rule_group, uses_parameter=True)
        self.RemoveRuleCommand = Command(self.remove_rule, uses_parameter=True)
        self.RemoveRuleGroupCommand = Command(self.remove_rule_group, uses_parameter=True)

    def add_rule(self, rule_view_model=None):
        """ Creates and adds a new ExcludeRule and inserts it after a
       certain rule or at the end if no rule is specified """
        if rule_view_model is not None:
            index = self.rule_view_models.IndexOf(rule_view_model) + 1
        else:
            index = len(self.rule_view_models)
        r = ExcludeRuleBase()
        self._exclude_rule_collection.insert(index, r)
        self.rule_view_models.Insert(index, ExcludeRuleViewModel(r, self))

    def add_rule_group(self, rule_view_model=None):
        """Creates and adds a new ExcludeRuleGroup with a blank 
        ExlcudeRule and inserts it after a certain rule or at the end 
        if no rule is specified"""
        if rule_view_model is not None:
            index = self.rule_view_models.IndexOf(rule_view_model) + 1
        else:
            index = len(self.rule_view_models)
        rg = ExcludeRuleGroup()
        self._exclude_rule_collection.insert(index, rg)
        rgvm = ExcludeRuleGroupViewModel(rg, self)
        rgvm.add_rule()
        self.rule_view_models.Insert(index, rgvm)

    def remove_rule(self, rule_view_model):
        self._exclude_rule_collection.remove(rule_view_model._exclude_rule)
        self.rule_view_models.Remove(rule_view_model)

    def remove_rule_group(self, group_view_model):
        self._exclude_rule_collection.remove(group_view_model._exclude_rule_collection)
        self.rule_view_models.Remove(group_view_model)
        
    @notify_property
    def Operator(self):
        return self._exclude_rule_collection.operator

    @Operator.setter
    def Operator(self, value):
        self._exclude_rule_collection.operator = value

    @notify_property
    def Invert(self):
        return self._exclude_rule_collection.invert

    @Invert.setter
    def Invert(self, value):
        self._exclude_rule_collection.invert = value




class ExcludeRuleViewModel(ViewModelBase):
    """The viewmodel for an exclude_rules.ExludeRule object"""

    fields = []
    number_operators = SortedDictionary[int, str](LOCALIZER.number_operators)
    string_operators = SortedDictionary[int, str](LOCALIZER.string_operators)
    manga_yes_no_operators = SortedDictionary[int, str](LOCALIZER.manga_yes_no_operators)
    yes_no_operators = SortedDictionary[int, str](LOCALIZER.yes_no_operators)
    bool_operators = SortedDictionary[int, str](LOCALIZER.bool_operators)
    date_operators = SortedDictionary[int, str](LOCALIZER.date_operators)
    custom_values = get_custom_value_keys()

    def __init__(self, exclude_rule, parent):
        self.parent = parent
        if not self.fields:
            self.fields = FIELDS.exclude_rule_fields
            self.fields.sort(cmp=PythonLocale.strcoll, key=lambda f: f.name)
        self._exclude_rule = exclude_rule
        self._operators = self.string_operators
        
        if not self._exclude_rule.field:
            self.SelectedField = self.fields[0]

        return super(ExcludeRuleViewModel, self).__init__()

    #SelectedField
    @notify_property
    def SelectedField(self):
        try:
            return FIELDS.get_by_field(self._exclude_rule.field)
        except KeyError:
            return None

    @SelectedField.setter
    def SelectedField(self, value):
        if value:
            self._exclude_rule.field = value.field
            self.Type = value.type

    #Operator
    @notify_property
    def Operator(self):
        return self._exclude_rule.operator

    @Operator.setter
    def Operator(self, value):
        if value != self._exclude_rule.operator:
            if (self.Type == FieldType.DateTime and 
                self._exclude_rule.operator == ExcludeDateTimeOperators.Range):
                self._exclude_rule.value2 = ""
            self._exclude_rule.operator = value
            

    #Value
    @notify_property
    def Value(self):
        return self._exclude_rule.value

    @Value.setter
    def Value(self, value):
        self._exclude_rule.value = value

    #Value2
    @notify_property
    def Value2(self):
        return self._exclude_rule.value2

    @Value2.setter
    def Value2(self, value):
        self._exclude_rule.value2 = value

    #Invert
    @notify_property
    def Invert(self):
        return self._exclude_rule.invert

    @Invert.setter
    def Invert(self, value):
        self._exclude_rule.invert = value

    #Type
    @notify_property
    def Type(self):
        return self._exclude_rule.type

    @Type.setter
    def Type(self, value):
        if value == self._exclude_rule.type:
            return
        self._exclude_rule.type = value
        if value == "MangaYesNo":
            self.Operators = self.manga_yes_no_operators
            #for whatever reason the binding doesn't update correct when the operators change
            self.Operator = -1
            self.Operator = 0
        elif value == "String" or value == "MultipleValue":
            self.Operators = self.string_operators
            self.Operator = self.Operators.Keys.First()
        elif value == "YesNo":
            self.Operators = self.yes_no_operators
            self.Operator = 0
        elif value == "Number":
            self.Operators = self.number_operators
            self.Operator = self.Operators.Keys.First()
        elif value == FieldType.DateTime:
            self.Operators = self.date_operators
            self.Operator = self.Operators.Keys.First()
        try:
            if not self.Operators.ContainsKey(self.Operator):
                self.Operator = self.Operators.Keys.First()
        except Exception, ex:
            print ex

        print self.Operator

    #Operators
    @notify_property
    def Operators(self):
        return self._operators

    @Operators.setter
    def Operators(self, value):
        self._operators = value

class ExcludeRuleGroupViewModel(ExcludeRuleCollectionViewModel):

    def __init__(self, rule_collection, parent):
        self.parent = parent
        super(ExcludeRuleGroupViewModel, self).__init__(rule_collection)

    def create_commands(self):
        self.AddRuleCommand = Command(self.add_rule, uses_parameter=True)
        self.AddRuleGroupCommand = Command(self.add_rule_group, uses_parameter=True)
        self.RemoveRuleCommand = Command(self.remove_rule, lambda: len(self.rule_view_models) > 1, uses_parameter=True)
        self.RemoveRuleGroupCommand = Command(self.remove_rule_group, lambda: len(self.rule_view_models) > 1, uses_parameter=True)


class RulesTemplateSelector(DataTemplateSelector):
    """Selects the correct datatemplete for ExcludeRules or ExcludeGroups."""
    def SelectTemplate(self, item, container):
        if type(item) is ExcludeRuleViewModel:
            return container.FindResource("ExcludeRuleTemplate")
        elif type(item) is ExcludeRuleGroupViewModel:
            return container.FindResource("ExcludeGroupTemplate")

