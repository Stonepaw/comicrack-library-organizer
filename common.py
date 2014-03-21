"""
These enums are for ease of comparing operators. Strings have been 
considered but Numbers are used for now. 
"""

#ExcludeDateOperators = ['Is', 'After', 'Before', 'InTheLast', 'Range']
#ExcludeStringOperators = ['Is', 'Contains', 'ContainsAny', 'ContainsAll',
#                          'StartsWith', 'EndsWith', 'ListContains', 'RegEx']
#ExcludeNumberOperators = ['Is', 'Greater', 'Smaller', 'Range']
#ExcludeBoolOperators = ['True', 'False']
#ExcludeYesNoOperators = ['Yes', 'No', 'Unknown']
#ExcludeMangaYesNoOperators = ['Yes', 'YesRightToLeft', 'No', 'Unknown']

class ExcludeDateOperators(object):
    Is = "Is"
    After = "After"
    Before = "Before"
    InTheLast = "InTheLast"
    Range = "Range"

    @classmethod
    def get_list(cls):
        return (cls.Is, cls.After, cls.Before, cls.InTheLast, cls.Range)


class ExcludeStringOperators(object):
    Is = "Is"
    Contains = "Contains"
    ContainsAny = "ContainsAny"
    ContainsAll = "ContainsAll"
    StartsWith = "StartsWith"
    EndsWith = "EndsWith"
    ListContains = "ListContains"
    RegEx = "RegEx"

    @classmethod
    def get_list(cls):
        return (cls.Is, cls.Contains, cls.ContainsAny, cls.ContainsAll,
                cls.StartsWith, cls.EndsWith, cls.ListContains, cls.RegEx)


class ExcludeNumberOperators(object):
    Is = 0
    Greater = 1
    Smaller = 2
    Range = 3

    @classmethod
    def get_list(cls):
        return (cls.Is, cls.Greater, cls.Smaller, cls.Range)


class ExcludeBoolOperators(object):
    True = "True"
    False = "False"

    @classmethod
    def get_list(cls):
        return (cls.True, cls.False)


class ExcludeYesNoOperators(object):
    Yes = "Yes"
    No = "No"
    Unknown = "Unknown"

    @classmethod
    def get_list(cls):
        return (cls.Yes, cls.No, cls.Unknown)


class ExcludeMangaYesNoOperators(object):
    Yes = "Yes"
    YesRightToLeft = "YesRightToLeft"
    No = "No"
    Unknown = "Unknown"

    @classmethod
    def get_list(cls):
        return (cls.Yes, cls.YesRightToLeft, cls.No, cls.Unknown)