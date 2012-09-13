import clr
import System
clr.AddReference("System.Xml")
from System import Single, Double, Int64, Int32
from System.Collections.ObjectModel import ObservableCollection
from System.Xml import XmlConvert
clr.AddReference("ComicRack.Engine")
clr.AddReference("cYo.Common")
from cYo.Projects.ComicRack.Engine import ComicBook, MangaYesNo, YesNo, ComicValueType

import localizer
from locommon import name_to_field
from wpfutils import Command, NotifyPropertyChangedBase, notify_property
#TODO: improve rule calculation. Currently broken.

class ExcludeRuleCollection(object):
    """Contains and manages a collection of ExcludeRules and/or ExcludeGroups.

    Can also calculate if a book should be moved or not moved under its rules
    """
    
    def __init__(self, operator="All", mode="Do not"):
        """Initiates a new ExcludeRuleCollection.

        Args:
            operator: The operator to use. Should be "All" or "Any".
            mode: The mode to use. Should be "Only" or "Do not".
        """
        self.rules = ObservableCollection[object]()
        if operator not in ("All", "Any"): operator = "All"
        self.operator = operator
        if mode not in ("Only", "Do not"): mode = "Do not"
        self.mode = mode

        #Create Commands
        self.AddNewRule = Command(self.add_new_rule)
        self.InsertNewRule = Command(self.insert_new_rule, uses_parameter=True)
        self.AddNewGroup = Command(self.add_new_group)
        self.InsertNewGroup = Command(self.insert_new_group, uses_parameter=True)
        self.RemoveRule = Command(self.remove_rule, uses_parameter=True)
        self.MoveRuleUp = Command(self.move_rule_up, can_execute=self.can_move_rule_up, uses_parameter=True)
        self.MoveRuleDown = Command(self.move_rule_down, can_execute=self.can_move_rule_down, uses_parameter=True)

    def __len__(self):
        return self.rules.Count

    def add_new_rule(self):
        """Creates and adds a new ExcludeRule to the end of the rule collection."""
        rule = ExcludeRule()
        rule.parent = self
        self.rules.Add(rule)
        
    def insert_new_rule(self, rule_to_add_after):
        """Creates a new ExcludeRule and inserts it after an existing rule or group.

        If the existing rule is not in the rule collection then the new rule is
        not created.

        Args:
            rule_to_add_after: The ExcludeRule or ExcludeGroup to insert the
                new rule after.
        """
        if rule_to_add_after in self.rules:
            rule = ExcludeRule()
            rule.parent = self
            self.rules.Insert(self.rules.IndexOf(rule_to_add_after) + 1, rule)

    def add_new_group(self):
        """Creates and adds a new ExcludeGroup to the end of the rule collection."""
        group = ExcludeGroup()
        group.parent = self
        group.add_new_rule()
        self.rules.Add(group)

    def insert_new_group(self, rule_to_add_after):
        """Creates a new ExcludeGroup and inserts it after an existing rule or group.

        If the existing rule is not in the rule collection then the new group is
        not created.

        Args:
            rule_to_add_after: The ExcludeRule or ExcludeGroup to insert the
                new group after.
        """
        if rule_to_add_after in self.rules:
            group = ExcludeGroup()
            group.parent = self
            group.add_new_rule()
            self.rules.Insert(self.rules.IndexOf(rule_to_add_after) + 1, group)

    def add_rule(self, rule):
        """Adds an ExcludeRule or ExcludeGroup into the end of the rule collection.

        Args:
            rule: The ExcludeRule or ExcludeGroup to add.
        """
        if rule is not None:
            rule.parent = self
            self.rules.Add(rule)

    def can_move_rule_up(self, rule):
        """Checks if the rule is not the first in the rule collection.

        Args:
            rule: The rule to check.

        Returns:
            True if the rule can be moved up, False if it cannot.
        """
        if rule in self.rules:
            if self.rules.IndexOf(rule) > 0:
                return True
        return False

    def can_move_rule_down(self, rule):
        """Checks if the rule is not the last in the rule collection.

        Args:
            rule: The rule to check.

        Returns:
            True if the rule can be moved down, False if it cannot.
        """
        if rule in self.rules:
            if self.rules.Count > 1 and self.rules.IndexOf(rule) < self.rules.Count - 1:
                return True
        return False

    def move_rule_up(self, rule):
        """Moves a rule in the rule collection up one space.
        
        Args:
            rule: The rule to move up.
        """
        if rule in self.rules:
           index = self.rules.IndexOf(rule)
           if index > 0:
               self.rules.Move(index, index - 1)

    def move_rule_down(self, rule):
        """Moves a rule in the rule collection up down space.
        
        Args:
            rule: The rule to move down.
        """
        if rule in self.rules:
           index = self.rules.IndexOf(rule)
           if index < self.rules.Count:
               self.rules.Move(index, index + 1)

    def remove_rule(self, rule):
        """Removes an ExcludeRule or ExcludeGroup from the rule collection.

        Args:
            rule: The ExcludeRule or ExcludeGroup to remove.
        """
        if rule in self.rules:
            self.rules.Remove(rule)
        
                
class ExcludeGroup(ExcludeRuleCollection):
    """Contains and manages a collection of ExcludeRules and/or ExlcudeGroups.
   
    Can also calculate if a book should be moved under its rules.
    """
    def __init__(self, operator="All", invert=False):
        """Initiates a new ExcludeGroup.

        Args:
            operator: The operator to use. Should be "All" or "Any".
            invert: A bool if the calculation should be inverted.
        """
        self.rules = ObservableCollection[object]()
        self.invert = invert
        if operator not in ("All", "Any"): operator = "All"
        self.operator = operator
        self.parent = None

        #ICommands
        self.InsertNewRule = Command(self.insert_new_rule, uses_parameter=True)
        self.InsertNewGroup = Command(self.insert_new_group, uses_parameter=True)
        self.RemoveRule = Command(self.remove_rule, can_execute=lambda: len(self) > 1, uses_parameter=True)
        self.MoveRuleUp = Command(self.move_rule_up, can_execute=self.can_move_rule_up, uses_parameter=True)
        self.MoveRuleDown = Command(self.move_rule_down, can_execute=self.can_move_rule_down, uses_parameter=True)

    def book_should_be_moved(self, book):
        """
        Checks if the book should be moved under the rules in the rule group.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """
        
        #Keeps track of the amount of rules the book fell under
        count = 0
        
        #Keep track of the total amount of rules
        total = 0
        
        for rule in self.rules:
            

            result = rule.book_should_be_moved(book)
            
            #Something went wrong, possible empty group. Thus we don't count that rule
            if result is None:
                continue
        
            count += result
            total += 1

        if total == 0:
            return None
        
        if self.operator == "Any":
            if count > 0:
                return 1
            else:
                return 0
        else:
            if count == total:
                return 1
            else:
                return 0


class ExcludeRule(NotifyPropertyChangedBase):
    """Contains the data of an exlude rule. It can calculate if a book should be moved using it's rules."""
    
    #Class variables for memory savings.
    _comicbook = ComicBook()
    _numeric_operators = localizer.get_exclude_rule_numeric_operators().values()
    _string_operators = localizer.get_exclude_rule_string_operators().values()
    _manga_yes_no_operators = localizer.get_exclude_rule_manga_yes_no_operators().values()
    _yes_no_operators = localizer.get_exclude_rule_yes_no_operators().values()
    _bool_operators = localizer.get_exclude_rule_bool_operators().values()

    def __init__(self, field="Series", operator="is", value="", invert=False):
        """Creates a new exclude rule. 

        If no args are passed, then a blank exclude rule is created with the defaults.
        
        Args:
            field: The ComicBook property name to use.
            operator: The operator to use. Should be appropriate for the field type.
            value: The value to check the field against.
            invert: A bool if the moving calculation should be inverted.
        """
        super(ExcludeRule, self).__init__()
        self._type = ""
        self.invert = invert
        #updates from 2.1 to 2.2 because the invert function was added
        if operator == "is not":
            operator = "is"
            self.invert = True
        elif operator == "does not contain":
            operator = "contains"
            self.invert = True
        self._operator = operator
        self._value = value
        self.parent = None
        self._field = ""
        #Because these were changed to use the field property name instead of the english name in 2.2,
        #have to check and use the property name
        if field in name_to_field:
            field = name_to_field[field]
        self.Field = field

    #Have to use notify property on Field, Type and Operator to make the changing comboboxes work correctly in the wpf form
    @notify_property
    def Field(self):
        return self._field

    @Field.setter
    def Field(self, value):
        self._field = value
        t = type(getattr(self._comicbook, value))
        print t
        if t is str:
            self.Type = "String"
            if self._operator not in self._string_operators:
                self.Operator = "is"
        elif t is MangaYesNo:
            self.Type = "MangaYesNo"
            if self._operator not in self._manga_yes_no_operators:
                self.Operator = "is Yes"
        elif t is YesNo:
            self.Type = "YesNo"
            if self._operator not in self._yes_no_operators:
                self.Operator = "is Yes"
        elif t is bool:
            print "is Bool"
            self.Type = "Bool"
            if self._operator not in self._bool_operators:
                self.Operator = "is True"
        elif t in (Int32, Int64, float, Double, Single, int):
            self.Type = "Numeric"
            if self._operator not in self._numeric_operators:
                self.Operator = "is"

    @notify_property
    def Type(self):
        return self._type

    @Type.setter
    def Type(self, value):
        self._type = value

    @notify_property
    def Operator(self):
        return self._operator

    @Operator.setter
    def Operator(self, value):
        self._operator = value

    @notify_property
    def Value(self):
        return self._value

    @Value.setter
    def Value(self, value):
        self._value = value

    #TODO improve this
    def get_yes_no_value(self):
        """Returns the correct YesNo value."""
        if self._field == "Manga":
            if self._value == "Yes (Right to Left)":
                return MangaYesNo.YesAndRightToLeft
            else:
                return getattr(MangaYesNo, self._value)

        return getattr(YesNo, self._value)

    def book_should_be_moved(self, book):
        """
        Finds if the book should be moved using this rule.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """

        if field in ("Manga", "SeriesComplete", "BlackAndWhite"):
            return self.calculate_book_should_be_moved(book, getattr(book, field), self.get_yes_no_value())

        elif field in ("StartYear", "StartMonth"):
            return self.calculate_book_should_be_moved(book, self.get_start_field_data(book, field), self._value)

        else:
            return self.calculate_book_should_be_moved(book, getattr(book, field), self._value)
    
    def calculate_book_should_be_moved(self, book, field_data, value):
        """
        Checks if the book should be moved based on this single rule.
        field_data -> the contents of the field.
        value -> the value to check the contents of the field against.
        Returns 1 if the book should be moved.
        Returns 0 if the book should not be moved.
        """
        if self._type == "String":
            return self.compare_string_to_value(book)
        

        if self._operator == "is":
        #Convert to string just in case
            if field_data == value:
                return 1
            else:
                return 0
        elif self._operator == "does not contain":
            if value not in field_data:
                return 1
            else:
                return 0
        elif self._operator == "contains":
            if value in field_data:
                return 1
            else:
                return 0
        elif self._operator == "is not":
            if value != field_data:
                return 1
            else:
                return 0
        elif self._operator == "greater than":
            #Try to use the int value to compare if possible
            try:
                if int(value) < int(field_data):
                    return 1
                else:
                    return 0
            except ValueError:
                if value < field_data:
                    return 1
                else:
                    return 0
        elif self._operator == "less than":
            try:
                if int(value) > int(field_data):
                    return 1
                else:
                    return 0
            except ValueError:
                if value > field_data:
                    return 1
                else:
                    return 0
        
    def get_start_field_data(self, book, field):
        """
        Finds the field contents for the earlies book of the same series in the ComicRack library.
        book -> the book of the series to search for.
        field -> The string of the field to retrieve.
        
        returns -> Unicode string of the field.
        """

        startbook = get_earliest_book(book)
        
        if field == "StartMonth":
            return unicode(startbook.Month)

        else:
            return unicode(startbook.ShadowYear)

    def compare_string_to_value(self, book):
        #print books[0].GetStringPropertyValue("Series", ComicValueType.Shadow) #Gets the value and if failing, getting the shadowvalues

        field_value = book.GetStringPropertyValue(self._field, ComicValueType.Shadow)

        if self._operator == "is":
            return field_value == self._value
        elif self._operator == "contains":
            return self._value in field_value
        elif self._operator == "starts with":
            return field_value.startswith(self._value)
        elif self._operator == "ends with":
            return field_value.endswith(self._value)
        elif self._operator == "contains any of":
            values = [s.strip() for s in self.values.split(" ")]
            for s in values:
                if s in field_value:
                    return True
        elif self._operator == "contains all of":
            values = [s.strip() for s in self.values.split(" ")]
            matches = 0
            for s in values:
                if s in field_value:
                    matches = matches + 1
            return len(matches) == len(values)
        elif self._operator == "regular expression":
            pass

    def compare_numeric_to_value(self, book):
        try:
            value = float(self._value)
            field_value = book.GetPropertyValue[float](self._field, ComicValueType.Shadow)
        except ValueError:
            #If the comparing value is not a number then try and compare with strings
            value = self._value
            field_value = book.GetStringPropertyValue(self._field, ComicValueType.Shadow)
        if self._operator == "is":
            return value == field_value
        elif self._operator == "is greater":
            return field_value > value
        elif self._operator == "is smaller":
            return field_value < value