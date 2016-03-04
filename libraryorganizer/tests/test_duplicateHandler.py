from System.IO import File, FileInfo
import comicracknlogtarget
from ComicRack import ComicRack
from bookandprofiletestcase import BookAndProfileTestCase
from bookmanager import DuplicateHandler
from common import MoveFailedException, BookToMove
from movereporter import MoveReporter
import i18n
i18n.setup(ComicRack())
from duplicatewindow import DuplicateResult, DuplicateAction


class TestDuplicateHandler(BookAndProfileTestCase):
    def setUp(self):
        super(TestDuplicateHandler, self).setUp()
        self.handler = DuplicateHandler(MoveReporter(), [], ComicRack())

    def test__create_rename_path(self):
        """ This tests the create rename path function that rename paths are
        created correctly when several exist."""
        f1 = self.create_temp_path("test (1).txt")
        f2 = self.create_temp_path("test (2).txt")
        File.Create(f1).Close()
        File.Create(f2).Close()
        self.cleanup_file_paths.append(f1)
        self.cleanup_file_paths.append(f2)
        path = self.create_temp_path("test.txt")

        new_path = self.handler._create_rename_path(path)
        self.assertEqual(new_path, self.create_temp_path("test (3).txt"))

    def test__create_rename_path_already_duplicate(self):
        """ This tests the create rename path function that rename paths are
        created correctly when several exist."""
        f = self.create_temp_path("test.txt")
        f1 = self.create_temp_path("test (1).txt")
        f2 = self.create_temp_path("test (2).txt")
        File.Create(f).Close()
        File.Create(f1).Close()
        File.Create(f2).Close()
        self.cleanup_file_paths.append(f1)
        self.cleanup_file_paths.append(f2)
        self.cleanup_file_paths.append(f)
        path = self.create_temp_path("test (1).txt")

        new_path = self.handler._create_rename_path(path)
        self.assertEqual(new_path, self.create_temp_path("test (3).txt"))

    def test__delete_duplicate(self):
        """ This tests that the duplicate delete code works as expected in a
        normal usage."""
        dup_path = self.create_temp_path("test duplicate.txt")
        File.Create(dup_path).Close()
        self.cleanup_file_paths.append(dup_path)
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
        self.fail()

    def test_handle_duplicate_book(self):
        self.fail()

    def test__process_different_extension_books_all_overwrite(self):
        def return_overwrite(*args):
            return DuplicateResult(DuplicateAction.Overwrite, False)
        self.handler._ask_user = return_overwrite
        f1 = self.create_temp_path("lotest.cbz")
        f2 = self.create_temp_path("lotest.cbr")
        f3 = self.create_temp_path("lotest.cbt")
        File.Create(f2).Close()
        File.Create(f3).Close()
        file2 = FileInfo(f2)
        file3 = FileInfo(f3)
        self.cleanup_file_paths.extend([f1,f2,f3])
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
