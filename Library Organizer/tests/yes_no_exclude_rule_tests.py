import unittest

import clr
clr.AddReference("ComicRack.Engine")
from System import DateTime

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c

from exclude_rules import YesNoExcludeRule
from common import ExcludeYesNoOperators
from cYo.Projects.ComicRack.Engine import ComicBook, YesNo
from fieldmappings import FIELDS, FieldType

class TestYesNoExcludeRule(unittest.TestCase):
    def setUp(self):
        self.book = ComicBook()
        self.book.BlackAndWhite = YesNo.Yes

    def test_is_yes(self):
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Yes)
        self.assertTrue(s.evaluate(self.book))

    def test_is_yes_false(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.No
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Yes)
        self.assertFalse(s.evaluate(book))

    def test_is_yes_invert(self):
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Yes, invert=True)
        self.assertFalse(s.evaluate(self.book))

    def test_is_no(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.No
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.No)
        self.assertTrue(s.evaluate(book))

    def test_is_no_false(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.Yes
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.No)
        self.assertFalse(s.evaluate(book))

    def test_is_no_invert(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.No
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.No, invert=True)
        self.assertFalse(s.evaluate(book))

    def test_is_unknown(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.Unknown
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Unknown)
        self.assertTrue(s.evaluate(book))

    def test_is_unknown_false(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.Yes
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Unknown)
        self.assertFalse(s.evaluate(book))

    def test_is_unknown_invert(self):
        book = ComicBook()
        book.BlackAndWhite = YesNo.Unknown
        s = YesNoExcludeRule("BlackAndWhite", ExcludeYesNoOperators.Unknown, invert=True)
        self.assertFalse(s.evaluate(book))

    def test_all_fields_are_yes_no(self):
        l = []
        for f in FIELDS:
            if f.type == FieldType.YesNo:
                s = YesNoExcludeRule(f.field, ExcludeYesNoOperators.Yes)
                r = s.get_field_value(self.book)
                print "%s: %s" % (f.field, r)
                l.append(r)
        self.assertTrue(all((type(x) == YesNo for x in l)))

if __name__ == '__main__':
    unittest.main()
