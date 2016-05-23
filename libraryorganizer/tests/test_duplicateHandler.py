from System.IO import File, FileInfo

import i18n
from ComicRack import ComicRack
from bookandprofiletestcase import BookAndProfileTestCase
from common import MoveFailedException, BookToMove, MoveSkippedException
from duplicatehandler import DuplicateHandler
from movereporter import MoveReporter

i18n.setup(ComicRack())
from duplicatewindow import DuplicateResult, DuplicateAction


class TestDuplicateHandler(BookAndProfileTestCase):
    def setUp(self):
        super(TestDuplicateHandler, self).setUp()
        self.handler = DuplicateHandler(MoveReporter(), [], ComicRack())

    def test__create_rename_path(self):
        """ This tests the create rename path function that rename paths are
        created correctly when several exist."""
        f1 = self._create_temp_file("test (1).txt")
        f2 = self._create_temp_file("test (2).txt")
        path = self.create_temp_path("test.txt")
        new_path = self.handler._create_rename_path(path)
        self.assertEqual(new_path, self.create_temp_path("test (3).txt"))

    def test__create_rename_path_already_duplicate(self):
        """ This tests the create rename path function that rename paths are
        created correctly when several exist and the current path is already a duplicate."""
        f = self._create_temp_file("test.txt")
        f1 = self._create_temp_file("test (1).txt")
        f2 = self._create_temp_file("test (2).txt")
        path = self.create_temp_path("test (1).txt")
        new_path = self.handler._create_rename_path(path)
        self.assertEqual(new_path, self.create_temp_path("test (3).txt"))

    def test__delete_duplicate(self):
        """ This tests that the duplicate delete code works as expected in a
        normal usage."""
        dup_path = self._create_temp_file("test duplicate.txt")
        self.handler._delete_duplicate(dup_path)
        self.assertFalse(File.Exists(dup_path))

    def test__delete_duplicate_file_in_use(self):
        """ This tests that the duplicate delete code works as expected in a
        normal usage."""
        dup_path = self.create_temp_path("test duplicate.txt")
        f1 = File.Create(dup_path)
        self.cleanup_file_paths.append(dup_path)
        with self.assertRaises(MoveFailedException):
            self.handler._delete_duplicate(dup_path)
        self.assertTrue(File.Exists(dup_path))
        f1.Close()

    def test__find_book_in_library(self):
        """ Tests that finding a real book in the library returns the book if on exists
        and return None if it does not.

        Tests that if there is a slight capitialization difference, the existing books is
        still returned.

        """
        self.assertIsNotNone(self.handler._find_book_in_library(
            'G:\\Comics\\DC Comics\\Adventure Comics (2009 Series)\\Adventure Comics Vol.2009 #01 (October, 2009).cbz'))
        self.assertIsNotNone(self.handler._find_book_in_library(
            'G:\\Comics\\DC Comics\\Adventure Comics (2009 Series)\\adventure comics Vol.2009 #01 (October, 2009).cbz'))
        self.assertIsNone(self.handler._find_book_in_library(
            'G:\\Comics\\DC Comics\\Adventure Comics (2009 Series)\\Adventure Comics Vol.2009 #19 (October, 2009).cbz'))

    def test_handle_duplicate_book(self):
        self.fail()

    def test__process_different_extension_books_all_overwrite(self):
        def return_overwrite(*args):
            return DuplicateResult(DuplicateAction.Overwrite, False)

        self.handler._ask_user = return_overwrite
        f1 = self.create_temp_path("lotest.cbz")
        f2 = self._create_temp_file("lotest.cbr")
        f3 = self._create_temp_file("lotest.cbt")
        file2 = FileInfo(f2)
        file3 = FileInfo(f3)
        self.cleanup_file_paths.append(f1)
        b = BookToMove(self.book, f1, 1, None)
        b.duplicate_different_extension = True
        b.duplicate_ext_files = [file2, file3]
        self.assertEqual(self.handler._process_different_extension_books(b, self.profile), DuplicateAction.Overwrite)
        self.assertFalse(File.Exists(f1))
        self.assertFalse(File.Exists(f2))

    def test__do_action(self):
        self.fail()

    def test__overwrite(self):
        self.fail()

    def test__overwrite_simulated(self):
        self.fail()

    def test__ask_user(self):
        self.fail()

    def test_copy_read_percentage_on_delete(self):
        self.fail()

    def test_remove_book_from_library(self):
        self.fail()

    def test__cancelled(self):
        with (self.assertRaises(MoveSkippedException)):
            self.handler._cancelled(BookToMove(self.book, "a path", 1, []))
