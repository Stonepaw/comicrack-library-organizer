import clr
clr.AddReference("PresentationFramework")

from System import DateTime, FormatException
from System.Collections.Generic import Dictionary, SortedDictionary
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import BackgroundWorker
from IronPython.Modules import PythonLocale
from System.Windows.Controls import DataTemplateSelector

from fieldmappings import conditional_fields, conditional_then_else_fields, FIELDS, first_letter_fields
import localizer
from locommon import date_formats, get_custom_value_keys
from wpfutils import notify_property, NotifyPropertyChangedBase, ViewModelBase


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


class DateInsertViewModel(InsertViewModel, ViewModelBase):
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
        self._custom_date_format = ""
        self._custom_date_preview = ""
        super(DateInsertViewModel, self).__init__(field)
        super(ViewModelBase, self).__init__()

    @notify_property
    def CustomDateFormat(self):
        return self._custom_date_format

    @CustomDateFormat.setter
    def CustomDateFormat(self, value):
        try:
            self.CustomDateFormatPreview = DateTime.Now.ToString(value)
        except FormatException:
            self.CustomDateFormatPreview = "Error parsing Custom Date Format"
        self._custom_date_format = value


    @notify_property
    def CustomDateFormatPreview(self):
        return self._custom_date_preview

    @CustomDateFormatPreview.setter
    def CustomDateFormatPreview(self, value):
        self._custom_date_preview = value

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
            self.FirstLetterFields = [FIELDS.get_by_field(f) for f in first_letter_fields]
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
        self.ConditionalFields = sorted([FIELDS.get_by_field(f) for f in conditional_fields],
                                        PythonLocale.strcoll,
                                        lambda x: x.name)
        self.ConditionalThenElseFields = sorted([FIELDS.get_by_field(f) for f in conditional_then_else_fields],
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


class CustomFieldInsertViewModel(InsertViewModel, NotifyPropertyChangedBase):

    custom_keys = ObservableCollection[str]()

    def __init__(self, template):
        self._retrieving_custom_keys = False
        self._selected_item = ""

        if self.custom_keys.Count == 0:
            self.IsRetrievingCustomKeys = True
            b = BackgroundWorker()
            b.DoWork += self.retrieve_custom_values
            b.RunWorkerCompleted += self.retrieve_custom_values_completed
            b.RunWorkerAsync()
            self._retrieving_custom_keys = True
        
        else:
            self._selected_item = self.custom_keys[0]
            
        NotifyPropertyChangedBase.__init__(self)
        InsertViewModel.__init__(self, template)

    def retrieve_custom_values(self, sender, e):
        keys = get_custom_value_keys()
        e.Result = keys

    def retrieve_custom_values_completed(self, sender, e):
        for i in e.Result:
            self.custom_keys.Add(i)
        self.IsRetrievingCustomKeys = False
        self.SelectedItem = self.custom_keys[0]

    @notify_property
    def IsRetrievingCustomKeys(self):
        return self._retrieving_custom_keys
        
    @IsRetrievingCustomKeys.setter
    def IsRetrievingCustomKeys(self, value):
        self._retrieving_custom_keys = value

    @notify_property
    def SelectedItem(self):
        return self._selected_item

    @SelectedItem.setter
    def SelectedItem(self, value):
        self._selected_item = value

    def make_template(self, autospace):
        """
        Creates the insertable template with correct args for a read percentage field.

        Args:
            autospace: Inserts a space before the prefix

        Returns:
            The string of the constructed insertable template
        """
        args = "(%s)" % (self.SelectedItem)
        return super(CustomFieldInsertViewModel, self).make_template(autospace, args)


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
    elif type == "CustomValue":
        return CustomFieldInsertViewModel(template)


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
        elif type(item) is CustomFieldInsertViewModel:
            return container.FindResource("CustomValueInsertFieldTemplate")