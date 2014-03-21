# exclude_rules.py
# Copyright 2014 Andrew Feltham (Stonepaw)
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""This modules contains the classes for the exclude rule implementation.

Usage:
    The following exclude rules are defined for use:
    StringExcludeRule
    YesNoExcludeRule
    BooleanExcludeRule
    CustomValueExcludeRule
    DateExcludeRule
    MangeYesNoExcludeRule
    NumberExcludeRule

    ExcludeRuleCollection and ExcludeRuleGroup are list based classes
    for contain multiple exclude rules.

    Each class has a function evaluate which requires a ComicBook passed
    as a parameter. This function returns True if the book should 
    be moved or False if it should not.
"""

import clr
from System import DateTime, FormatException
from System.Text.RegularExpressions import Regex

clr.AddReference("ComicRack.Engine")
from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

from fieldmappings import FieldType, FIELDS
from common import ExcludeBoolOperators, ExcludeDateOperators, ExcludeMangaYesNoOperators, ExcludeNumberOperators, ExcludeStringOperators, ExcludeYesNoOperators

class ExcludeRuleCollectionMode(object):
    Only = "Only"
    DoNot = "Do not"


class ExcludeRuleGroupOperator(object):
    Any = "Any"
    All = "All"


class _ExcludeRuleBase(object):
    """A Base for the more specific forms of exclude rules.

    Do not use this class directly as it is a common base for 
    more specific types of exclude rules and contains no logic.

    To implement this class you must define the class variables:
    default_operator, default_field and type as well as implement the 
    function: evaluate(book).

    In evaluate function call get_field_value to get the book field.
    Override get_field_value if special cases are required.
    """
    default_field = FIELDS[0].field
    default_operator = None
    type = FieldType.String

    def __init__(self, field=default_field, operator=default_operator, value='', value2='', 
                 invert=False):
        self.field = field
        self.operator = operator
        self.value = value
        self.value2 = value2
        self.invert = invert

    @classmethod
    def from_exclude_rule(cls, rule):
        """Creates a new exclude rule from the contents of another rule.

        Copies the operator (if valid), the selected field and the invert
        to a new rule and returns that rule.

        If the operator is not valid then it uses the default_operator.

        Args:
            rule: An ExcludeRuleBase object that the fields will be retrieved 
                from.

        Returns:
            The created ExcludeRuleBase
        """
        operator = cls.default_operator
        return cls(rule.field, operator=operator, invert=rule.invert)

    def get_field_value(self, book):
        """Retrieves the field value from the book"""
        return getattr(book, self.field)

    def result_with_invert(self, result):
        if self.invert:
            return not result
        return result

    def evaluate(self, book):
        raise NotImplementedError


class StringExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a string field.

    Contains the logic required for matching a string field.
    """
    type = FieldType.String
    default_operator = ExcludeStringOperators.ContainsAll

    def evaluate(self, book):
        """Evaluates this rule against a book.

        Returns: True if the book matches this rule or False if it does
                not.
        """
        field_value = self.get_field_value(book)
        result = False
        if self.operator == ExcludeStringOperators.Is:
            result = field_value == self.value

        elif self.operator == ExcludeStringOperators.Contains:
            result = self.value in field_value
        
        elif self.operator == ExcludeStringOperators.StartsWith:
            result = field_value.startswith(self.value)

        elif self.operator == ExcludeStringOperators.EndsWith:
            result = field_value.endswith(self.value)

        elif self.operator == ExcludeStringOperators.ContainsAll:
            operators = self.value.split()
            result = all(x in field_value for x in operators)

        elif self.operator == ExcludeStringOperators.ContainsAny:
            operators = self.value.split()
            result = any(x in field_value for x in operators)

        elif self.operator == ExcludeStringOperators.ListContains:
            field_items = [i.strip() for i in field_value.split(",")]
            match_items = [i.strip() for i in self.value.split(",")]
            result = all(x in field_items for x in match_items)

        elif self.operator == ExcludeStringOperators.RegEx:
            regex = Regex(self.value)
            result = regex.Match(field_value).Success

        return self.result_with_invert(result)


class DateExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a date field.

    Contains the logic required for a matching a date field.
    """
    default_operator = ExcludeDateOperators.Is
    type = FieldType.DateTime

    def evaluate(self, book):
        field_value = self.get_field_value(book)
        result = False

        #Try to parse the value string into a DateTime object.
        if self.operator != ExcludeDateOperators.InTheLast:
            try:
                date = DateTime.Parse(self.value)
            except FormatException:
                print "%s is not a valid DateTime. Rule is marked as false" % (self.value)
                return False
        
        if self.operator == ExcludeDateOperators.Is:
            result = date.Date == field_value.Date

        elif self.operator == ExcludeDateOperators.After:
            result = date.Date < field_value.Date

        elif self.operator == ExcludeDateOperators.Before:
            result = date.Date > field_value.Date

        elif self.operator == ExcludeDateOperators.InTheLast:
            result = field_value.Date >= DateTime.Now.AddDays(-int(self.value)).Date

        elif self.operator == ExcludeDateOperators.Range:
            try:
                date2 = DateTime.Parse(self.value2)
            except FormatException:
                print "%s is not a valid DateTime. Rule is marked as false" % (self.value2)
                return False

            result = date.Date <= field_value.Date <= date2.Date

        return self.result_with_invert(result)


class NumberExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a number field.

    Contains the logic required for matching a number field.
    """
    default_operator = ExcludeNumberOperators.Is
    type = FieldType.Number

    def evaluate(self, book):
        field_value = self.get_field_value(book)
        value = None
        result = False

        if self.operator == ExcludeNumberOperators.Is:
            # Certain fields can possibly be string values so on the is
            # comparison it's best to compare strings to strings
            result = unicode(self.value) == unicode(field_value)
        else:
            try:
                field_value = float(field_value)
                value = float(self.value)
            except ValueError:
                return False
            if self.operator == ExcludeNumberOperators.Greater:
                result = field_value > value
            elif self.operator == ExcludeNumberOperators.Smaller:
                result = field_value < value
            elif self.operator == ExcludeNumberOperators.Range:
                try:
                    result = value <= field_value <= float(self.value2)
                except ValueError:
                    return False
        return self.result_with_invert(result)

        
        
    

class BooleanExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a boolean field.

    Contains the logic required for matching a boolean field.
    """
    type = FieldType.Boolean
    default_operator = ExcludeBoolOperators.True

    def evaluate(self, book):
        field_value = self.get_field_value(book)
        result = False

        if self.operator == ExcludeBoolOperators.True:
            result = field_value == True
            
        elif self.operator == ExcludeBoolOperators.False:
            result = field_value == False    

        return self.result_with_invert(result)
    
    def get_field_value(self, book):
        """Returns the field value from the book"""
        if self.field == "HasCustomValues":
            return any(book.GetCustomValues())
        else:
            return super(BooleanExcludeRule, self).get_field_value(book)


class YesNoExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a YesNo field.

    Contains the logic required for matching a YesNo field.
    """
    type = FieldType.YesNo
    default_operator = ExcludeYesNoOperators.Yes

    def evaluate(self, book):
        field_value = self.get_field_value(book)
        print field_value
        result = False
        if self.operator == ExcludeYesNoOperators.Yes:
            result = field_value == YesNo.Yes
            print result

        elif self.operator == ExcludeYesNoOperators.No:
            result = field_value == YesNo.No
            print result

        elif self.operator == ExcludeYesNoOperators.Unknown:
            result = field_value == YesNo.Unknown
            print result

        return self.result_with_invert(result)


class MangeYesNoExcludeRule(_ExcludeRuleBase):
    """The exclude rule class for a MangaYesNo field.

    Contains the logic required for matching a MangaYesNo field.
    """
    type = FieldType.MangaYesNo
    default_operator = ExcludeMangaYesNoOperators.Yes

    def evaluate(self, book):
        field_value = self.get_field_value(book)
        result = False
        if self.operator == ExcludeMangaYesNoOperators.Yes:
            result = field_value == MangaYesNo.Yes

        elif self.operator == ExcludeMangaYesNoOperators.No:
            result = field_value == MangaYesNo.No

        elif self.operator == ExcludeMangaYesNoOperators.Unknown:
            result = field_value == MangaYesNo.Unknown

        elif self.operator == ExcludeMangaYesNoOperators.YesRightToLeft:
            result = field_value == MangaYesNo.YesRightToLeft
            
        return self.result_with_invert(result)


class CustomValueExcludeRule(StringExcludeRule):
    type = FieldType.CustomValue
    default_operator = ExcludeStringOperators.Is

    def get_field_value(self, book):
        c = book.GetCustomValue(self.value2)
        print c
        return c


class ExcludeRuleCollection(list):
    
    def __init__(self, invert=False, operator=ExcludeRuleGroupOperator.All):
        self.invert = invert
        self.operator = operator
        return super(ExcludeRuleCollection, self).__init__()

    def evaluate(self, book):
        result = False
        results = []
        for rule in self:
            results.append(rule.evaluate(book))

        if operator == ExcludeRuleGroupOperator.All:
            result = all(results)
        elif operator == ExcludeRuleGroupOperator.Any:
            result = any(results)
        
        if self.invert:
            return not result
        return result


class ExcludeRuleGroup(ExcludeRuleCollection):
    """ Contains a collection of ExcludeRules and Groups"""