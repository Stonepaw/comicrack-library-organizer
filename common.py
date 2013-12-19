"""
These enums are for ease of comparing operators. Strings have been 
considered but Numbers are used for now. 
"""
class ExcludeDateOperators(object):
    Is = 0
    After = 1
    Before = 2
    InTheLast = 3
    Range = 4

    @staticmethod
    def get_list():
        return range(5)


class ExcludeStringOperators(object):
    Is = 0
    Contains = 1
    ContainsAny = 2
    ContainsAll = 3
    StartsWith = 4
    EndsWith = 5
    ListContains = 6
    RegEx = 7

    @staticmethod
    def get_list():
        return range(8)


class ExcludeNumberOperators(object):
    Is = 0
    Greater = 1
    Smaller = 2
    Range = 3

    @staticmethod
    def get_list():
        return range(4)


class ExcludeBoolOperators(object):
    True = 0
    False = 1

    @staticmethod
    def get_list():
        return range(2)


class ExcludeYesNoOperators(object):
    Yes = 0
    No = 1
    Unknown = 2

    @staticmethod
    def get_list():
        return range(3)


class ExcludeMangaYesNoOperators(object):
    Yes = 0
    No = 1
    Unknown = 2
    YesRightToLeft = 3

    @staticmethod
    def get_list():
        return range(4)