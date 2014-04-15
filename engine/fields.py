

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

    def __init__(self, backup_name, field, type, template=None, 
                 exclude=True, conditional=True):
        """Creates a new field.

        Args:
            backup_name: The name used in the translation if no
                translation is available.
            field: The string representation of the property on the
                ComicBook.
            type: The FieldType of this field.
            template: The string to use for the template builder. Pass
                False if this field should not be used in the template
                builder. Pass None if the template name is the same as
                the field string.
            exclude: A boolean if this field should be usable in
                exclude rules.
            conditional: A boolean if this field should be usable in a
                conditional field.
        """
        self.backup_name = backup_name
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
            template: The string to use for the template builder. Pass
                False if this field should not be used in the template
                builder. Pass None if the template name is the same as
                the field string. Defaults to None.
            exclude: A boolean if this field should be usable in
                exclude rules. Defaults to True
            conditional: A boolean if this field should be usable in a
                conditional field. Defaults to True.
        """
        self.append(Field(backup_name, field, type, template, exclude,
                          conditional))

    def get_template_fields(self):
        return [i for i in self if i.template != False]

    def get_conditional_fields(self):
        return [i for i in self if i.conditional != False and i.field != "Text"]

    def get_conditional_then_else_fields(self):
        return [i for i in self if i.conditional != False]

    def get_exclude_fields(self):
        return [i for i in self if i.exclude == True]

class fields(object):
    """description of class"""



