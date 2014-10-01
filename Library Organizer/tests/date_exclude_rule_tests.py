import unittest

import clr
clr.AddReference("ComicRack.Engine")
from System import DateTime

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c

from exclude_rules import DateExcludeRule
from common import ExcludeDateOperators
from cYo.Projects.ComicRack.Engine import ComicBook
from fieldmappings import FIELDS, FieldType

class TestDateExcludeRule(unittest.TestCase):

    def setUp(self):
        self.book = ComicBook()
        self.book.AddedTime = DateTime(2013, 12, 29)

    def test_is(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Is, "29/12/2013")
        self.assertTrue(s.evaluate(self.book))

    def test_is_false(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Is, "28/12/2013")
        self.assertFalse(s.evaluate(self.book))

    def test_is_invert(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Is, "29/12/2013", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_is_invalid_date_time_string(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Is, "asdf")
        self.assertFalse(s.evaluate(self.book))

    def test_after(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.After, "28/12/2013")
        self.assertTrue(s.evaluate(self.book))

    def test_after_false(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.After, "30/12/2013")
        self.assertFalse(s.evaluate(self.book))

    def test_after_invert(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.After, "28/12/2013", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_before(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Before, "30/12/2013")
        self.assertTrue(s.evaluate(self.book))

    def test_before_false(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Before, "28/12/2013")
        self.assertFalse(s.evaluate(self.book))

    def test_before_invert(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Before, "30/12/2013", invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_in_the_last(self):
        book = ComicBook()
        book.AddedTime = DateTime().Now.AddDays(-1)
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.InTheLast, "2")
        self.assertTrue(s.evaluate(book))

    def test_in_the_last_false(self):
        book = ComicBook()
        book.AddedTime = DateTime.Now.AddDays(-3)
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.InTheLast, "2")
        self.assertFalse(s.evaluate(book))

    def test_in_the_last_invert(self):
        book = ComicBook()
        book.AddedTime = DateTime.Now.AddDays(-3)
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.InTheLast, "2", invert=True)
        self.assertTrue(s.evaluate(book))

    def test_in_the_range(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Range, "28/12/2013", "30/12/2013")
        self.assertTrue(s.evaluate(self.book))

    def test_in_the_range_false(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Range, "26/12/2013", "28/12/2013")
        self.assertFalse(s.evaluate(self.book))

    def test_in_the_range_invert(self):
        s = DateExcludeRule("AddedTime", ExcludeDateOperators.Range, "26/11/2013", "2/12/2013", invert=True)
        self.assertTrue(s.evaluate(self.book))

    def test_all_fields_are_datetime(self):
        l = []
        for f in FIELDS:
            if f.type == FieldType.DateTime:
                s = DateExcludeRule(f.field, ExcludeDateOperators.Is, "")
                r = s.get_field_value(self.book)
                print "%s: %s" % (f.field, r)
                l.append(r)
        self.assertTrue(all((type(x) == DateTime for x in l)))


if __name__ == '__main__':
    unittest.main()
