"""
This module contains all the unittests for the path maker
"""
import clr
from System import Console

from unittest import TestCase, TestSuite
import unittest
import clr
clr.AddReference("ComicRack.Engine")
clr.AddReference("Stonepaw.LibraryOrganizer")
from cYo.Projects.ComicRack.Engine import ComicBook
from pathmaker import PathMaker
from System.IO import File, Path
from Stonepaw.LibraryOrganizer import Profile


class TestTemplateFields(object):

    def setUp(self):
        self.pathmaker = PathMaker()
        self.book = ComicBook()
        self.book.Series = "Star Wars"
        tempFolder = Path.GetTempPath()
        self.path = Path.Combine(tempFolder, "test book2.cbz")
        self.book.FilePath = self.path
        File.Create(self.path)
        self.pathmaker.book = self.book
        p = Profile()
        self.pathmaker.profile = p


    def test_series(self):
        pass

    def test_endmonth(self):
        pass

    def test_firstLetter(self):
        template = "{<FirstLetter(Series)>}"
        s = self.pathmaker.insert_fields_into_template(template)
        print s
        self.assertEquals(s, 'S')

    #def tearDown(self):
    #    File.Delete(self.path)
t = TestTemplateFields()
t.setUp()
t.test_firstLetter()
Console.ReadLine()
