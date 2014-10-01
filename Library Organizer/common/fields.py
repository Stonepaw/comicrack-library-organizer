"""
fields.py

Copyright 2014 Stonepaw

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

"""This file contains the Field object that can be used to get the 
field name, template, and english name"""

import clr
clr.AddReference("NLog")

import os
from utils import str_to_bool
path = os.path.abspath(os.path.dirname(__file__))

from NLog import LogManager

logger = LogManager.GetLogger("Fields")

class FieldType(object):
    """ An enum for field types """
    String = 'String'
    Number = 'Number'
    MultipleValue = 'MultipleValue'
    YesNo = 'YesNo'
    MangaYesNo = 'MangaYesNo'
    DateTime = 'DateTime'
    Boolean = 'Boolean'
    Month = 'Month'
    Year = 'Year'
    CustomValue = 'CustomValue'
    Conditional = 'Conditional'
    ReadPercentage = 'ReadPercentage'
    Text = 'Text'
    FirstLetter = 'FirstLetter'
    Counter = 'Counter'


class Field(object):
    """Contains various important varibles for each field.

    Attributes:
        name: The English name of the field.
        field: The ComicRack field name.
        template: The string to use for the template builder. If 
            template is an empty string then the field should not 
            be used as a vaild template
        type: The field type. Uses a FieldType class.
        exclude: A bool of if this field should be used in exclude rules.
        conditional: A bool of if this field should be used in conditional

    """
    def __init__(self, backup_name, field, type, template=None, 
                 exclude=True, conditional=True):
        """Creates a new field.

        Args:
            backup_name: The name used in the translation if no
                translation is available.
            field: The string representation of the property on the
                ComicBook.
            type: The FieldType of this field.
            template: The string to use for the template builder. If 
                template is an empty string then the field should not 
                be used as a vaild template
            exclude: A boolean if this field should be usable in
                exclude rules.
            conditional: A boolean if this field should be usable in a
                conditional field.
        """
        self.default_name = self.name = backup_name
        self.field = field
        self.template = template
        self.type = type
        self.exclude = exclude
        self.conditional = conditional


class FieldList(list):

    def __init__(self, *args, **kwargs):
        self._exclude_rule_fields = None
        self._template_fields = None
        return super(FieldList, self).__init__(*args, **kwargs)

    def get_by_field(self, field):
        for item in self.__iter__():
            if item.field == field:
                return item
        raise KeyError

    @property
    def exclude_rule_fields(self):
        if self._exclude_rule_fields is None:
            l = [f for f in self.__iter__() if f.exclude]
            self._exclude_rule_fields = l
        return self._exclude_rule_fields

    def add(self, backup_name, field, type, template=None, exclude=True,
            conditional=True):
        """Adds a field to this field list.
        Args:
            backup_name: The name used in the translation if no
                translation is available.
            field: The string representation of the property on the
                ComicBook.
            type: The FieldType of this field.
            template:The string to use for the template builder. If 
                template is an empty string then the field should not 
                be used as a vaild template
            exclude: A boolean if this field should be usable in
                exclude rules. Defaults to True
            conditional: A boolean if this field should be usable in a
                conditional field. Defaults to True.
        """
        self.append(Field(backup_name, field, type, template, exclude,
                          conditional))

    @property
    def template_fields(self):
        return [i for i in self if i.template]

    @property
    def conditional_fields(self):
        return [i for i in self if i.conditional != False and i.field != "Text"]

    @property
    def conditional_then_else_fields(self):
        return [i for i in self if i.conditional != False]

    @property
    def exclude_fields(self):
        return [i for i in self if i.exclude == True]


def create_fields():
    fields = FieldList()
    logger.Info("Getting fields from csv")
    try:
        fields_source = []
        with open(os.path.join(path, 'fields.csv'), 'r') as f:
            fields_source = f.readlines()
        for i in fields_source[1:]:
            ii = i.strip().split(",")
            fields.add(ii[0], ii[1], ii[2], ii[3], str_to_bool(ii[4]), 
                       str_to_bool(ii[5]))
        logger.Info("Sucessfuly loaded fields info")
    except IOError, ex:
        logger.Error(ex)
        raise IOError("Could not locate fields.csv")
    return fields

FIELDS = create_fields()