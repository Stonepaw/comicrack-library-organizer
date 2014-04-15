"""
Copyright 2013 Stonepaw

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

clr.AddReference("System.Core")
clr.AddReference("IronPython.Modules")
clr.AddReference("PresentationFramework")

from IronPython.Modules import PythonLocale
from System.Collections.Generic import SortedDictionary
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import DataTemplateSelector

from exclude_rules import _ExcludeRuleBase, ExcludeRuleGroup, BooleanExcludeRule, CustomValueExcludeRule, DateExcludeRule, MangeYesNoExcludeRule, NumberExcludeRule, StringExcludeRule, YesNoExcludeRule
from fieldmappings import FIELDS, FieldType
from localizer import Localizer
from locommon import get_custom_value_keys
from wpfutils import Command, notify_property, ViewModelBase

LOCALIZER = Localizer()
DEBUG = True

class ExcludeRuleCollectionViewModel(ViewModelBase):
    """Provides an view model for an exclude rule collection.

    Provides Notify properties for the Operator and Invert variable.

    Provides several useful commands that are bindable in the view:
        AddRuleCommand: Adds a new default rule.
        AddRuleGroupCommand: Adds a new rule group with a rule
        RemoveRuleCommand: Removes a passed in rule
        RemoveRuleGroupCommand: Removes a passed in rule group
        RemoveAllRulesCommands: Removes all the rules and view models
    """
        
    operators = SortedDictionary[str, str](LOCALIZER.all_any_operators)

    def __init__(self, exclude_rule_collection):
        self._exclude_rule_collection = exclude_rule_collection

        self.rule_view_models = ObservableCollection[object]()

        for rule in exclude_rule_collection:
            if isinstance(rule, _ExcludeRuleBase):
                self.rule_view_models.Add(ExcludeRuleViewModel(rule, self))
        self.create_commands()
        return super(ExcludeRuleCollectionViewModel, self).__init__()

    def create_commands(self):
        """Creates commands that can be bound in the WPF view.
        """
        self.AddRuleCommand = Command(self.add_rule, uses_parameter=True)
        self.AddRuleGroupCommand = Command(self.add_rule_group, uses_parameter=True)
        self.RemoveRuleCommand = Command(self.remove_rule, uses_parameter=True)
        self.RemoveRuleGroupCommand = Command(self.remove_rule_group, uses_parameter=True)
        self.RemoveAllRulesCommand = Command(self.remove_all_rules, lambda: len(self.rule_view_models) > 0)

    def add_rule(self, rule_view_model=None):
        """ Creates and adds a new ExcludeRule and inserts it after a
       certain rule or at the end if no rule is specified """
        if rule_view_model is not None:
            index = self.rule_view_models.IndexOf(rule_view_model) + 1
        else:
            index = len(self.rule_view_models)
        r = DateExcludeRule()
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
        """Removes a rule and rule view model from the collection"""
        self._exclude_rule_collection.remove(rule_view_model._exclude_rule)
        self.rule_view_models.Remove(rule_view_model)

    def remove_rule_group(self, group_view_model):
        """Removes a rule group and rule group view model from the 
        collection"""
        self._exclude_rule_collection.remove(group_view_model._exclude_rule_collection)
        self.rule_view_models.Remove(group_view_model)

    def remove_all_rules(self):
        """Removes all the rules in this rule collection."""
        self.rule_view_models.Clear()
        self._exclude_rule_collection.Clear()
        
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
    """The view model for an exclude rule object
    
    Exposes properties via the notify property interface to the wpf view.
    Exposes:
        Rule: the rule itself.
        SelectedField: The TemplateItem object of the selected field
        Operator: The rule operator
        Value: The rule value
        Value2: The rule's secondary value.    
    """

    #Some global things for the view to use
    fields = []
    number_operators = SortedDictionary[str, str](LOCALIZER.number_operators)
    string_operators = SortedDictionary[str, str](LOCALIZER.string_operators)
    manga_yes_no_operators = SortedDictionary[str, str](LOCALIZER.manga_yes_no_operators)
    yes_no_operators = SortedDictionary[str, str](LOCALIZER.yes_no_operators)
    bool_operators = SortedDictionary[str, str](LOCALIZER.bool_operators)
    date_operators = SortedDictionary[str, str](LOCALIZER.date_operators)
    custom_values = get_custom_value_keys()

    def __init__(self, exclude_rule, parent):
        """Creates a new ExcludeRuleViewModel.

        Args:
            exclude_rule: The ExcludeRuleBase object or derivatives 
                that this view model represents.
            parent: The parent view model of the group or collection that
                this rule is part of.
        """
        #The parent group has to be kept in order to call the parent's
        #delete function in the view.
        self.parent = parent
        if not self.fields:
            fields = LOCALIZER.translated_field_list.exclude_rule_fields
            fields.sort(cmp=PythonLocale.strcoll, key=lambda f: f.name)
            ExcludeRuleViewModel.fields = fields
        self._exclude_rule = exclude_rule
        
        if not self._exclude_rule.field:
            self.SelectedField = self.fields[0]

        super(ExcludeRuleViewModel, self).__init__()
        self.OnPropertyChanged("SelectedField")

    @notify_property
    def Rule(self):
        return self._exclude_rule

    @Rule.setter
    def Rule(self, value):
        self._exclude_rule = value


    #SelectedField
    @notify_property
    def SelectedField(self):
        try:
            return LOCALIZER.translated_field_list.get_by_field(self.Rule.field)
        except KeyError:
            return None

    @SelectedField.setter
    def SelectedField(self, value):
        if value:
            self.Rule.field = value.field
            if value.type != self._exclude_rule.type:
                #If the type is different then the base rule type has to change
                #Except for some cases where the fieldtype has the same exclude
                #rule type.
                if (self._exclude_rule.type == FieldType.String and 
                    value.type == FieldType.MultipleValue):
                    return
                elif (self._exclude_rule.type == FieldType.Number and
                      value.type in (FieldType.Month, FieldType.Year)):
                      return
                self._switch_rule_type(value.type)

    def _switch_rule_type(self, new_type):
        """Switches the _exclude_rule to a new type.

        Args:
            new_type: A FieldType with the new type.
        """
        if new_type == FieldType.DateTime:
            if DEBUG: print "DateTime"
            self.Rule = DateExcludeRule.from_exclude_rule(self._exclude_rule)

        elif new_type == FieldType.String or new_type == FieldType.MultipleValue:
            if DEBUG: print "String"
            self.Rule = StringExcludeRule.from_exclude_rule(self._exclude_rule)
            
        elif new_type in (FieldType.Number, FieldType.Month, FieldType.Year):
            if DEBUG: print "Number"
            self.Rule = NumberExcludeRule.from_exclude_rule(self._exclude_rule)

        elif new_type == FieldType.Boolean:
            if DEBUG: print "Bool"
            self.Rule = BooleanExcludeRule.from_exclude_rule(self._exclude_rule)

        elif new_type == FieldType.YesNo:
            if DEBUG: print "YesNo"
            self.Rule = YesNoExcludeRule.from_exclude_rule(self._exclude_rule)

        elif new_type == FieldType.MangaYesNo:
            if DEBUG: print "Manga Yes No"
            self.Rule = MangeYesNoExcludeRule.from_exclude_rule(self._exclude_rule)

        elif new_type == FieldType.CustomValue:
            if DEBUG: print "Custom Value"
            self.Rule = CustomValueExcludeRule.from_exclude_rule(self._exclude_rule)

        self.OnPropertyChanged("Type")

    #Operator
    @notify_property
    def Operator(self):
        return self.Rule.operator

    @Operator.setter
    def Operator(self, value):
        if value != self.Rule.operator:
            if (self.Type == FieldType.DateTime and 
                self.Rule.operator == "Range"):
                self.Rule.value2 = ""
            self.Rule.operator = value

    #Value
    @notify_property
    def Value(self):
        return self._exclude_rule.value

    @Value.setter
    def Value(self, value):
        self._exclude_rule.value = value
        print self._exclude_rule.value

    #Value2
    @notify_property
    def Value2(self):
        return self._exclude_rule.value2

    @Value2.setter
    def Value2(self, value):
        self._exclude_rule.value2 = value

    #Type
    @notify_property
    def Type(self):
        return self.Rule.type


class ExcludeRuleGroupViewModel(ExcludeRuleCollectionViewModel):
    """Provides an view model for an exclude rule group.

    Provides Notify properties for the Operator and Invert variable.

    Provides several useful commands that are bindable in the view:
        AddRuleCommand: Adds a new default rule.
        AddRuleGroupCommand: Adds a new rule group with a rule
        RemoveRuleCommand: Removes a passed in rule
        RemoveRuleGroupCommand: Removes a passed in rule group
    """

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