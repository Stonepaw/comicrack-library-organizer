import clr
from unittest import TestCase

from System.IO import Path

from ComicRack import ComicRack
from bookandprofiletestcase import BookAndProfileTestCase
from bookmanager import BookManager
from common import Mode
from losettings import Profile
from movereporter import MoveReporter

clr.AddReference('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import ComicBook


class TestBookManager(BookAndProfileTestCase):
    def setUp(self):
        super(TestBookManager, self).setUp()
        self.manager = BookManager(ComicRack(), None)

    """
    Tests:

    fileless books:
        enabled option
        disabled option
        no thumbnail

    book:
        missing file

    features to test:
        failed fields
        multiple profiles
        missing info
        duplicates

    """

    def test_process_books(self):
        self.fail()

    def test__create_paths(self):
        self.manager._create_paths([self.book], [self.profile])

    def test__get_profile_for_book(self):
        self.fail()

    def test__create_path(self):
        result = self.manager._create_path(self.book, self.profile)
        self.assertEquals(result,"Marvel\\Spider-Man\\1.cbz")

        self.book.Series = ""
        self.profile.FailedFields.append("Series")
        result = self.manager._create_path(self.book, self.profile)

    def test__create_path_failed_fields_and_move(self):
        self.book.Publisher = ""
        self.profile.FolderTemplate = "{<publisher>}"
        self.profile.FileTemplate = "{<series>}"
        self.profile.FailedFields.append("Publisher")
        self.profile.MoveFailed = True
        self.profile.FailedFolder = self.create_temp_path("")
        result = self.manager._create_path(self.book, self.profile)
        self.assertEquals(result,self.book.FilePath)

    def test_create_path_failed_fields(self):
        self.book.Publisher = ""
        self.profile.FolderTemplate = "{<publisher>}"
        self.profile.FileTemplate = "{<series>}"
        self.profile.FailedFields.append("Publisher")
        self.profile.FailEmptyValues = True
        self.book.Series = "Test"
        self.profile.MoveFailed = False
        result = self.manager._create_path(self.book, self.profile)
        self.assertEqual(result, None)


    def test__book_should_be_moved(self):
        self.fail()
