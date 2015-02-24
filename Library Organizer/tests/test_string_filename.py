import unittest
import clr
clr.AddReference("ComicRack.Engine")
from cYo.Projects.ComicRack.Engine import ComicBook
from pathmaker import PathMaker
from System.IO import File, Path

class Test_test_string_filename(unittest.TestCase):

    def setUp(self):
        self.pathmaker = PathMaker
        self.book = ComicBook()
        self.book.Series = "Star Wars"
        tempFolder = Path.GetTempPath()
        self.path = Path.Combine(tempFolder, "test book.cbz")
        self.book.FilePath = self.path
        File.Create(self.path)


    def test_Series(self):
        template = "{<Series>}"
        self.assertEquals(self.pathmaker.make_file_name(self.book,template, None),
                          "Star Wars")

    def tearDown(self):
        File.Delete(self.path)

if __name__ == '__main__':
    unittest.main()
