from fieldmappings import FieldType

class ExcludeRuleCollectionMode(object):
    Only = "Only"
    DoNot = "Do not"


class ExcludeRuleGroupOperator(object):
    Any = "Any"
    All = "All"

class DateTimeOperators(object):
    Is = 0
    After = 1
    Before = 2
    Last = 3
    Range = 4



class ExcludeRuleBase(object):
    
    def __init__(self):
        
        self.field = ""
        self.operator = ""
        self.value = ""
        self.value2 = ""
        self.invert = False
        self.type = ""


class StringExcludeRule(ExcludeRuleBase):
        
        def __init__(self):
            super(ExcludeRuleBase, self).__init__()


class NumberExcludeRule(ExcludeRuleBase):
    pass

class ExcludeRuleGroup(list):
    """ Contains a collection of ExcludeRules and Groups"""

    def __init__(self, invert=False, operator=ExcludeRuleGroupOperator.All):

        self.invert = invert
        self.operator = operator

        return super(ExcludeRuleGroup, self).__init__()


class ExcludeRuleCollection(list):
    
    def __init__(self, invert=False, operator=ExcludeRuleGroupOperator.All):
        self.invert = invert
        self.operator = operator
        return super(ExcludeRuleCollection, self).__init__()



