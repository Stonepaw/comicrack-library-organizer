import unittest

import ComicRack
c = ComicRack.ComicRack()

import localizer
localizer.ComicRack = c
from exclude_rules import CustomValueExcludeRule
from common import ExcludeStringOperators

import clr
clr.AddReference("ComicRack.Engine")
from cYo.Projects.ComicRack.Engine import ComicBook

class Test_custom_value_exclude_rule(unittest.TestCase):
    def test_is(self):
        b = ComicBook()
        b.SetCustomValue("Test", "abcd")
        c = CustomValueExcludeRule("CustomValue", ExcludeStringOperators.Is, "abcd", "Test")
        self.assertTrue(c.evaluate(b))

    def test_nonexistant_custom_value(self):
        b = ComicBook()
        c = CustomValueExcludeRule("CustomValue", ExcludeStringOperators.Is, "abcd", "Test")
        self.assertFalse(c.evaluate(b))

if __name__ == '__main__':
    unittest.main()
