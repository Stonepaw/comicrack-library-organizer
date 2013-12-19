from fieldmappings import FieldType, FIELDS

from common import ExcludeBoolOperators, ExcludeDateOperators, ExcludeMangaYesNoOperators, ExcludeNumberOperators, ExcludeStringOperators, ExcludeYesNoOperators

class ExcludeRuleCollectionMode(object):
    Only = "Only"
    DoNot = "Do not"


class ExcludeRuleGroupOperator(object):
    Any = "Any"
    All = "All"


class ExcludeRuleBase(object):
    
    operators = ExcludeStringOperators

    def __init__(self, field='AddedTime', operator=0, value='', value2='', 
                 invert=False):
        
        self.field = field
        self.operator = operator
        self.value = value
        self.value2 = value2
        self.invert = invert
        self.type = FIELDS.get_by_field(field).type

    @classmethod
    def from_exclude_rule(cls, rule):
        return cls(rule.field, invert=rule.invert)


class StringExcludeRule(ExcludeRuleBase):

    operators = ExcludeStringOperators
    
    def __init__(self):
        super(ExcludeRuleBase, self).__init__()


class DateExcludeRule(ExcludeRuleBase):
    operators = ExcludeDateOperators



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



