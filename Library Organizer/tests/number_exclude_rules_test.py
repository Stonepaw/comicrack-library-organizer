import unittest

import clr
clr.AddReference("ComicRack.Engine")

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c

from exclude_rules import NumberExcludeRule
from common import ExcludeNumberOperators
from cYo.Projects.ComicRack.Engine import ComicBook
from fieldmappings import FIELDS, FieldType

class TestNumberExcludeRule(unittest.TestCase):

    def setUp(self):
        self.book = ComicBook()
        self.book.Number = "1"

    def test_is_with_string_and_number(self):
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Is, 1)
        self.assertTrue(s.evaluate(self.book))

    def test_is_with_number_and_number(self):
        book = ComicBook()
        book.Count = 1
        s = NumberExcludeRule("Count", ExcludeNumberOperators.Is, 1)
        self.assertTrue(s.evaluate(book))

    def test_is_with_string_and_string(self):
        book = ComicBook()
        book.Number = "1"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Is, "1")
        self.assertTrue(s.evaluate(book))

    def test_greater_with_string_and_string(self):
        book = ComicBook()
        book.Number = "2"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Greater, "1")
        self.assertTrue(s.evaluate(book))

    def test_smaller_with_string_and_string(self):
        book = ComicBook()
        book.Number = "2"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Smaller, "3")
        self.assertTrue(s.evaluate(book))

    def test_smaller__false_with_string_and_string(self):
        book = ComicBook()
        book.Number = "2"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Smaller, "1")
        self.assertFalse(s.evaluate(book))

    def test_is_with_string_and_string_invert(self):
        book = ComicBook()
        book.Number = "1"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Is, "1", invert=True)
        self.assertFalse(s.evaluate(book))

    def test_range_with_string_and_string(self):
        book = ComicBook()
        book.Number = "1"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Range, "0", "5")
        self.assertTrue(s.evaluate(book))

    def test_range_with_string_and_string_false(self):
        book = ComicBook()
        book.Number = "20"
        s = NumberExcludeRule("ShadowNumber", ExcludeNumberOperators.Range, "0", "5")
        self.assertFalse(s.evaluate(book))

    def test_all_number_fields(self):
        """Tests all the fields that will go into a StringExcludeRule to
        make sure all fields can get a value
        """ 
        l = []
        for f in FIELDS.exclude_rule_fields:
            if f.type in (FieldType.Number):
                s = NumberExcludeRule(f.field, ExcludeNumberOperators.Is, "")
                r = s.get_field_value(self.book)
                print "%s: %s: %s" % (f.field, r, type(r))
                l.append(r)
        self.assertTrue(all((type(x) is int for x in l)))


if __name__ == '__main__':
    unittest.main()
