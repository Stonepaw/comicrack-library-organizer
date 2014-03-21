import unittest

import clr
clr.AddReference("ComicRack.Engine")
from System import DateTime

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c

from exclude_rules import BooleanExcludeRule
from common import ExcludeBoolOperators
from cYo.Projects.ComicRack.Engine import ComicBook
from fieldmappings import FIELDS, FieldType

class TestBooleanExcludeRule(unittest.TestCase):
    def setUp(self):
        self.book = ComicBook()
        self.book.Checked = True

    def test_true(self):
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.True)
        self.assertTrue(s.evaluate(self.book))

    def test_true_false(self):
        book = ComicBook()
        book.Checked = False
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.True)
        self.assertFalse(s.evaluate(book))

    def test_true_invert(self):
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.True, invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_false(self):
        book = ComicBook()
        book.Checked = False
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.False)
        self.assertTrue(s.evaluate(book))

    def test_false_false(self):
        book = ComicBook()
        book.Checked = True
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.False)
        self.assertFalse(s.evaluate(book))

    def test_false_invert(self):
        book = ComicBook()
        book.Checked = False
        s = BooleanExcludeRule("Checked", ExcludeBoolOperators.False, invert=True)
        self.assertFalse(s.evaluate(book))

    def test_all_fields_are_bool(self):
        l = []
        for f in FIELDS:
            if f.type == FieldType.Boolean:
                s = BooleanExcludeRule(f.field, ExcludeBoolOperators.True)
                r = s.get_field_value(self.book)
                print "%s: %s" % (f.field, r)
                l.append(r)
        self.assertTrue(all((type(x) == bool for x in l)))

if __name__ == '__main__':
    unittest.main()
