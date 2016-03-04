from losettings import Profile
from unittest import TestCase

import clr
clr.AddReference('ComicRack.Engine')
clr.AddReference('System.IO')
from cYo.Projects.ComicRack.Engine import ComicBook
from System.IO import Path, File, Directory


class BookAndProfileTestCase(TestCase):

    def setUp(self):
        self.book = ComicBook()
        self.profile = Profile()
        self.cleanup_folder_paths = []
        self.cleanup_file_paths = []

    def tearDown(self):
        for path in self.cleanup_file_paths:
            if File.Exists(path):
                File.Delete(path)
        for path in self.cleanup_folder_paths:
            if Directory.Exists(path):
                Directory.Delete(path)

    def create_temp_path(self, f):
        return Path.Combine(Path.GetTempPath(), f)
