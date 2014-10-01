import unittest

import clr
clr.AddReference("ComicRack.Engine")

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c

from exclude_rules import StringExcludeRule
from common import ExcludeStringOperators
from cYo.Projects.ComicRack.Engine import ComicBook
from fieldmappings import FIELDS, FieldType

class TestStringExcludeRule(unittest.TestCase):

    def setUp(self):
        self.book = ComicBook()
        self.book.Publisher = "Test"
        self.book.Tags = "Test, Test2, Test3"

    def test_is(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Is, "Test")
        self.assertTrue(s.evaluate(self.book))

    def test_is_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Is, "Test2")
        self.assertFalse(s.evaluate(self.book))

    def test_is_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Is, "Test", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_contains(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Contains, "es")
        self.assertTrue(s.evaluate(self.book))

    def test_contains_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Contains, "asdfasdf")
        self.assertFalse(s.evaluate(self.book))

    def test_contains_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.Contains, "es", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_starts_with(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.StartsWith, "Te")
        self.assertTrue(s.evaluate(self.book))

    def test_starts_with_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.StartsWith, "asdf")
        self.assertFalse(s.evaluate(self.book))

    def test_starts_with_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.StartsWith, "Te", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_ends_with(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.EndsWith, "st")
        self.assertTrue(s.evaluate(self.book))

    def test_ends_with_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.EndsWith, "asdfasdf")
        self.assertFalse(s.evaluate(self.book))
        
    def test_ends_with_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.EndsWith, "st", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_contains_all(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAll, "Te st")
        self.assertTrue(s.evaluate(self.book))

    def test_contains_all_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAll, "Te st asdf")
        self.assertFalse(s.evaluate(self.book))
        
    def test_contains_all_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAll, "Te st", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_contains_any(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAny, "Test alskdjfas")
        self.assertTrue(s.evaluate(self.book))

    def test_contains_any_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAny, "asdlfkas")
        self.assertFalse(s.evaluate(self.book))
        
    def test_contains_any_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.ContainsAny, "Test alskdjfas", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_list_contains(self):
        s = StringExcludeRule("Tags", ExcludeStringOperators.ListContains, "Test")
        self.assertTrue(s.evaluate(self.book))

    def test_list_contains_false(self):
        s = StringExcludeRule("Tags", ExcludeStringOperators.ListContains, "asdlfkas")
        self.assertFalse(s.evaluate(self.book))
        
    def test_list_contains_invert(self):
        s = StringExcludeRule("Tags", ExcludeStringOperators.ListContains, "Test2", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_regex(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.RegEx, "[tT]es")
        self.assertTrue(s.evaluate(self.book))

    def test_regex_false(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.RegEx, "[Tt]ass")
        self.assertFalse(s.evaluate(self.book))
        
    def test_regex_invert(self):
        s = StringExcludeRule("Publisher", ExcludeStringOperators.RegEx, "st", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_all_string_fields(self):
        """Tests all the fields that will go into a StringExcludeRule to
        make sure all fields can get a value
        """ 
        l = []
        for f in FIELDS:
            if f.type in (FieldType.String, FieldType.MultipleValue):
                s = StringExcludeRule(f.field, ExcludeStringOperators.Is, "")
                r = s.get_field_value(self.book)
                print "%s: %s" % (f.field, r)
                l.append(r)
        self.assertTrue(all((type(x) is str for x in l)))


if __name__ == '__main__':
    unittest.main()
