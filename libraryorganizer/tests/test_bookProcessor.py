import clr

from ComicRack import ComicRack
from movereporter import MoveReporter

clr.AddReference("System.IO")
clr.AddReference("System.Windows.Forms")
import System
from System.IO import File, Path, Directory, FileInfo, FileMode
from System.Windows.Forms import DialogResult
from bookandprofiletestcase import BookAndProfileTestCase
from bookprocessor import MoveBookProcessor, FilelessBookProcessor, BaseBookProcessor, SimulatedBookProcessor
from common import BookToMove, MoveFailedException, MoveSkippedException, DuplicateExistsException, Mode
from loforms import PathTooLongForm
clr.AddReference("NLog.dll")
from NLog import LogManager

class TestBaseBookProcessor(BookAndProfileTestCase):
    def setUp(self):
        super(TestBaseBookProcessor, self).setUp()
        self.processor = BaseBookProcessor(ComicRack())

    # _check_destination
    def test_check_file_invalid(self):
        """ Ensures that if an invalid file path is used, a MovedFailedException occurs.
        """
        print "Testing file_invalid with colons"
        with self.assertRaises(MoveFailedException):
            # Colon in middle
            self.processor._check_file("C:\\oew&:", False)

    def test_check_source_already_at_location(self):
        """ Checked that MoveSkipped is raised when file is at destination
        """
        print "Testing file already at location"
        with self.assertRaises(MoveSkippedException):
            self.book.FilePath = self._create_temp_file("LOalreadyexists.cbz")
            self.processor._check_source(self.book, self.book.FilePath, True)

    def test_source_book_already_at_destination_different_capitalization(self):
        """ Tests that if a book is already at the correct location with different capitalization.
        MoveSkipped should be raised.
        """
        file_location = self._create_temp_file("lodifferentcaptializationrename.cbz")
        capitalized_path = file_location.replace("lo", "LO")
        self.book.FilePath = file_location
        with self.assertRaises(MoveSkippedException):
            self.processor._check_source(self.book, capitalized_path, True)
        self.assertEqual(self.book.FilePath, capitalized_path)
        self.assertEquals(Directory.GetFiles(Path.GetDirectoryName(capitalized_path),
                                                 Path.GetFileName(capitalized_path))[0], capitalized_path)

    def test_source_book_already_at_destination_different_capitalization_no_rename(self):
        """ Tests that if a book is already at the correct location with different capitalization and is not renamed when
        false is passed in.
        MoveSkipped should be raised.
        """
        file_location = self._create_temp_file("lodifferentcaptializationrename2.cbz")
        capitalized_path = file_location.replace("lo", "LO")
        self.book.FilePath = file_location
        with self.assertRaises(MoveSkippedException):
            self.processor._check_source(self.book, capitalized_path, False)
        self.assertEqual(self.book.FilePath, file_location)
        self.assertEquals(Directory.GetFiles(Path.GetDirectoryName(capitalized_path),
                                                 Path.GetFileName(capitalized_path))[0], file_location)

    def test_check_source_not_exist(self):
        """  Checks that MoveFailed is raised when source file doesn't exits
        """
        with self.assertRaises(MoveFailedException) as e:
            self.book.FilePath = self.create_temp_path("LOdoesnotexist.cbz")
            self.processor._check_source(self.book, self.create_temp_path("lonull.cbz"), True)

    def test_add_custom_value_on_success(self):
        self.processor._add_custom_values_to_book(self.book, {"1" : "2", "3" : "4"})
        self.assertEquals(self.book.GetCustomValue("1"), "2")
        self.assertEquals(self.book.GetCustomValue("3"), "4")

    def test_add_tag_on_success(self):
        """ Checks that adding tags works correctly and sorts the tag list
        """
        self.book.Tags = "NEW TAG"
        self.processor._add_tag_on_success(self.book, ["BOOK", "ADDED"])
        self.assertEquals(self.book.Tags, "ADDED, BOOK, NEW TAG")



class TestBookProcessor(BookAndProfileTestCase):
    def setUp(self):
        super(TestBookProcessor, self).setUp()
        self.processor = MoveBookProcessor(ComicRack())
        self.book.FilePath = self.book_path = self.create_temp_path("book.cbz")
        self.cleanup_file_paths = [self.book_path]
        File.Create(self.book_path).Close()

    def test_process_book_simple_move(self):
        """ Tests that a simple move operation from one file name to another works.
        """
        new_path = self.create_temp_path("book2.cbz")
        self.cleanup_file_paths.extend([new_path])
        b = BookToMove(self.book, new_path, 1, [])
        self.processor.process_book(b, self.profile)
        self.assertTrue(File.Exists(new_path))
        self.assertTrue(not File.Exists(self.book_path))
        self.assertTrue(self.book.FilePath == new_path)

    def test_process_book_simple_move_create_directory(self):
        """ Tests that a simple move operation into a new directory works.
        """
        new_path = self.create_temp_path("lotest\\book2.cbz")
        self.cleanup_file_paths.extend([new_path])
        self.cleanup_folder_paths.append(self.create_temp_path("lotest"))
        b = BookToMove(self.book, new_path, 1, [])
        self.processor.process_book(b, self.profile)
        self.assertTrue(File.Exists(new_path))
        self.assertTrue(not File.Exists(self.book_path))

    def test_process_book_move_path_too_long_user_fixes(self):
        """ Tests that a simple move operation into a new directory works.
        """
        new_path = self.create_temp_path(
            "book2;lkjaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa aasldkfj;laskjdf;laskjf;kajsdf;lkjasl;fkjas"
            ";lkjfdl;askjfl;akjs ;lakj ;alskjdf ajsldf ;laskjf ;askjf ;laskjf ;laskj fd;laskj f;lakjs f;lkjasf;lj "
            "sa;lfkj  ;alskjdf sj;aslfdj fl;kjasfdlk.cbz")
        new_path2 = self.create_temp_path("book2.cbz")

        def fake_show_dialog(self, *args, **kwargs):
            return DialogResult.OK

        def get_path(self, *arg, **kwargs):
            return self.new_path

        PathTooLongForm.ShowDialog = fake_show_dialog
        PathTooLongForm.get_path = get_path
        PathTooLongForm.new_path = new_path2
        self.cleanup_file_paths.append(new_path)
        self.cleanup_file_paths.append(new_path2)
        b = BookToMove(self.book, new_path, 1, [])
        self.processor.process_book(b, self.profile)
        self.assertTrue(self.book.FilePath != new_path)
        self.assertTrue(File.Exists(self.book.FilePath))
        self.assertTrue(not File.Exists(self.book_path))

    def test_process_book_move_path_too_long_user_declines(self):
        """ Tests that a simple move operation into a new directory works.
        """
        new_path = self.create_temp_path(
            "book2;lkjaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa aasldkfj;laskjdf;laskjf;kajsdf;lkjasl;fkjas"
            ";lkjfdl;askjfl;akjs ;lakj ;alskjdf ajsldf ;laskjf ;askjf ;laskjf ;laskj fd;laskj f;lakjs f;lkjasf;lj "
            "sa;lfkj  ;alskjdf sj;aslfdj fl;kjasfdlk.cbz")

        def fake_show_dialog(self, *args, **kwargs):
            return DialogResult.Cancel

        PathTooLongForm.ShowDialog = fake_show_dialog
        self.cleanup_file_paths.append(new_path)
        with self.assertRaises(MoveSkippedException):
            b = BookToMove(self.book, new_path, 1, [])
            self.processor.process_book(b, self.profile)

    def test_source_book_does_not_exist(self):
        new_path = self.create_temp_path("testlo\\book2.cbz")
        File.Delete(self.book.FilePath)
        with self.assertRaises(MoveFailedException):
            self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)

    def test_source_book_already_at_destination(self):
        new_path = self.book_path
        with self.assertRaises(MoveSkippedException):
            self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)



    def test_destination_book_already_exists(self):
        new_path = self.create_temp_path("book2.cbz")
        File.Create(new_path).Close()
        self.cleanup_file_paths.append(new_path)
        with self.assertRaises(DuplicateExistsException):
            self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)

    def test_destination_book_already_exists_different_capitalization(self):
        new_path = self.create_temp_path("book2.cbz")
        new_path2 = self.create_temp_path("BOOK2.cbz")
        File.Create(new_path2).Close()
        self.cleanup_file_paths.append(new_path)
        self.cleanup_file_paths.append(new_path2)
        with self.assertRaises(DuplicateExistsException):
            self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)

    def test_destination_book_with_different_extension_exists(self):
        new_path = self.create_temp_path("book2.cbz")
        new_path2 = self.create_temp_path("book2.cbr")
        self.cleanup_file_paths.append(new_path)
        self.cleanup_file_paths.append(new_path2)
        self.profile.DifferentExtensionsAreDuplicates = True
        File.Create(new_path2).Close()
        with self.assertRaises(DuplicateExistsException):
            self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)

    def test_undo_custom_value_created(self):
        self.test_process_book_simple_move()
        self.assertEquals(self.book.GetCustomValue("lo_previous_path"), self.book_path)

    def test_simple_copy(self):
        new_path = self.create_temp_path("book2.cbz")
        self.cleanup_file_paths.append(new_path)
        self.profile.Mode = Mode.Copy
        self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)
        self.assertTrue(File.Exists(new_path))
        self.assertTrue(File.Exists(self.book_path))

    def test_cleanup_source_folders(self):
        source_folder = self.create_temp_path("testLOshouldbedeleted")
        Directory.CreateDirectory(source_folder).Refresh()
        self.cleanup_folder_paths.append(source_folder)

        book_path = Path.Combine(source_folder, "book2.cbz")
        self.book.FilePath = book_path
        self.cleanup_file_paths.append(book_path)
        File.Create(book_path).Close()

        new_path = self.create_temp_path("book_source_folders.cbz")
        self.cleanup_file_paths.append(new_path)
        self.profile.RemoveEmptyFolder = True

        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertFalse(File.Exists(book_path))
        self.assertTrue(File.Exists(new_path))
        self.assertFalse(Directory.Exists(source_folder))

    def test_do_not_cleanup_source_folders(self):
        source_folder = self.create_temp_path("testLOshouldbedeleted")
        Directory.CreateDirectory(source_folder).Refresh()
        self.cleanup_folder_paths.append(source_folder)

        book_path = Path.Combine(source_folder, "book2.cbz")
        self.book.FilePath = book_path
        self.cleanup_file_paths.append(book_path)
        File.Create(book_path).Close()

        new_path = self.create_temp_path("book_source_folders.cbz")
        self.cleanup_file_paths.append(new_path)
        self.profile.RemoveEmptyFolder = False

        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertFalse(File.Exists(book_path))
        self.assertTrue(File.Exists(new_path))
        self.assertTrue(Directory.Exists(source_folder))

    def test_cleanup_source_folders_excluded(self):
        source_folder = self.create_temp_path("LOtest should not be deleted")
        Directory.CreateDirectory(source_folder).Refresh()
        self.cleanup_folder_paths.append(source_folder)
        self.profile.ExcludedEmptyFolder.append(source_folder)

        book_path = Path.Combine(source_folder, "book2.cbz")
        self.book.FilePath = book_path
        self.cleanup_file_paths.append(book_path)
        File.Create(book_path).Close()

        new_path = self.create_temp_path("book_source_folders.cbz")
        self.cleanup_file_paths.append(new_path)
        self.profile.RemoveEmptyFolder = True

        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertFalse(File.Exists(book_path))
        self.assertTrue(File.Exists(new_path))
        self.assertTrue(Directory.Exists(source_folder))

    def test_cleanup_created_folder_when_failed(self):
        def return_failed(*args):
            raise MoveFailedException("")

        self.processor._move_book = return_failed

        new_source_path = self.create_temp_path("LOTestShouldBeDeleted")
        self.cleanup_folder_paths.append(new_source_path)
        new_path = Path.Combine(new_source_path, "book.cbz")
        self.cleanup_file_paths.append(new_path)

        with self.assertRaises(MoveFailedException):
            self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertFalse(Directory.Exists(new_source_path))

    def test_move_book(self):
        """ Tests move book"""
        new_path = self.create_temp_path("book2.cbz")
        self.cleanup_file_paths.append(new_path)
        self.processor._move_book(self.book, FileInfo(self.book_path), new_path)
        self.assertTrue(File.Exists(new_path))
        self.assertFalse(File.Exists(self.book_path))
        self.assertEqual(self.book.FilePath, new_path)

    def test_move_book_add_tag_on_success(self):
        """ Tests move book"""
        self.profile.SuccessTags.append("Success")
        new_path = self.create_temp_path("book2tagtest.cbz")
        self.cleanup_file_paths.append(new_path)
        self.processor.process_book(BookToMove(self.book, new_path, 0, []),self.profile)
        self.assertTrue(File.Exists(new_path))
        self.assertFalse(File.Exists(self.book_path))
        self.assertEqual(self.book.FilePath, new_path)
        self.assertEqual("Success", self.book.Tags)

    def test_move_book_add_custom_on_success(self):
        """ Tests move book"""
        self.profile.SuccessCustomValues["Success"] = "Success"
        new_path = self.create_temp_path("book2tagtest.cbz")
        self.cleanup_file_paths.append(new_path)
        self.processor.process_book(BookToMove(self.book, new_path, 0, []),self.profile)
        self.assertTrue(File.Exists(new_path))
        self.assertFalse(File.Exists(self.book_path))
        self.assertEqual(self.book.FilePath, new_path)
        self.assertEqual(self.book.GetCustomValue("Success"), "Success")

    def test_move_book_errors(self):
        """ Tests move book"""
        new_path = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles),
                                "book.cbz")
        with self.assertRaises(MoveFailedException):
            self.processor._move_book(self.book, FileInfo(self.book_path), new_path)
        self.assertNotEqual(self.book.FilePath, new_path)
        self.assertFalse(File.Exists(new_path))
        self.assertTrue(File.Exists(self.book_path))

    def test_move_book_file_in_use(self):
        """ Tests move book"""
        new_path = self.create_temp_path("book_in_use_dest.cbr")
        new_path2 = self.create_temp_path("book_in_use.cbr")
        f = File.Create(new_path2)
        self.book.FilePath = new_path2
        self.cleanup_file_paths.append(new_path)
        self.cleanup_file_paths.append(new_path2)
        with self.assertRaises(MoveFailedException):
            try:
                self.processor.process_book(BookToMove(self.book, new_path, 1, []), self.profile)
            except MoveFailedException as e:
                print e.message
                raise
        self.assertNotEqual(self.book.FilePath, new_path)
        self.assertFalse(File.Exists(new_path))
        self.assertTrue(File.Exists(self.book_path))
        f.Close()

    def test_destination_book_with_different_extension_exists_and_exact_file_exists(self):
        new_path = self.create_temp_path("book2.cbz")
        new_path2 = self.create_temp_path("book2.cbr")
        new_path3 = self.create_temp_path("book2.cbt")
        self.cleanup_file_paths.append(new_path)
        self.cleanup_file_paths.append(new_path2)
        self.cleanup_file_paths.append(new_path3)
        self.profile.DifferentExtensionsAreDuplicates = True
        File.Create(new_path3).Close()
        File.Create(new_path2).Close()
        File.Create(new_path).Close()
        b = BookToMove(self.book, new_path, 1, [])
        with self.assertRaises(DuplicateExistsException):
            self.processor.process_book(b, self.profile)
        self.assertEqual(len(b.duplicate_ext_files), 3)

    def test_destination_book_with_exact_file_exists_and_different_extensions_enabled(self):
        """ An error occurred where duplicate extensions were always reported even if only an exact duplicate existed
           when search for different extensions.
        """
        new_path = self.create_temp_path("book2.cbz")
        self.cleanup_file_paths.append(new_path)
        self.profile.DifferentExtensionsAreDuplicates = True
        File.Create(new_path).Close()
        b = BookToMove(self.book, new_path, 1, [])
        with self.assertRaises(DuplicateExistsException):
            self.processor.process_book(b, self.profile)
        self.assertEqual(len(b.duplicate_ext_files), 0)
        self.assertFalse(b.duplicate_different_extension)

    def test_failed_fields_move_to_folder_fails(self):
        new_path = self.create_temp_path("lo_destination.cbz")
        self.cleanup_file_paths.append(new_path)
        self.cleanup_folder_paths.append(self.profile.FailedFolder)

        with self.assertRaises(MoveFailedException) as ex:
            try:
                self.processor.process_book(BookToMove(self.book, new_path, 0, ["Series", "Publisher"]), self.profile)
            except MoveFailedException as e:
                print e.message
                raise
        self.assertNotEqual(self.book.FilePath, self.book_path)
        self.assertEqual(self.book.FilePath, new_path)
        self.assertTrue(File.Exists(new_path))
        self.cleanup_file_paths.append(self.book.FilePath)

class TestMoveBookProcessor(BookAndProfileTestCase):
    pass

class TestFilelessBookProcessor(BookAndProfileTestCase):
    def setUp(self):
        super(TestFilelessBookProcessor, self).setUp()
        self.processor = FilelessBookProcessor(ComicRack())
        self.b = BookToMove(self.book, "", 0, [])

    def test_simple_create_fileless_not_enabled(self):
        """ Tests that MoveSkippedException is raised when create fileless is not enabled
        """
        self.book.FilePath = ""
        new_path = self.create_temp_path("lo_test_image.jpg")
        self.cleanup_file_paths.append(new_path)
        self.b.path = new_path
        self.profile.MoveFileless = False
        with self.assertRaises(MoveSkippedException):
            self.processor.process_book(self.b, self.profile)
        self.assertFalse(File.Exists(new_path))

    def test_simple_create_fileless(self):
        """ Tests that the fileless image is created sucessfully in a simple operation"""
        self.book.FilePath = ""
        new_path = self.create_temp_path("lo_test_image.jpg")
        self.cleanup_file_paths.append(new_path)
        self.b.path = new_path
        self.profile.MoveFileless = True
        self.processor.process_book(self.b, self.profile)
        self.assertTrue(File.Exists(new_path))

    def test_create_fileless_not_enabled_ensure_folders_not_created(self):
        """ Tests that folders are not created if fileless images are not being created

        Fixes a bug discovered in development.
        """
        self.book.FilePath = ""
        new_path_folder = self.create_temp_path("lo_test")
        new_path = Path.Combine(new_path_folder, "lo_test_image.jpg")
        self.cleanup_file_paths.append(new_path)
        self.cleanup_folder_paths.append(new_path_folder)
        self.b.path = new_path
        self.profile.MoveFileless = False
        with self.assertRaises(MoveSkippedException):
            self.processor.process_book(self.b, self.profile)
        self.assertFalse(File.Exists(new_path))
        self.assertFalse(Directory.Exists(new_path_folder))

    def test_create_fileless_fails_ensure_folders_cleaned_up(self):
        """ Tests that if a fileless operation fails, the created folders are cleaned up"""
        self.book.FilePath = ""
        new_path_folder = self.create_temp_path("lo_test")
        new_path = Path.Combine(new_path_folder, "lo_test_image.jpg")
        self.cleanup_file_paths.append(new_path)
        self.cleanup_folder_paths.append(new_path_folder)
        self.b.path = new_path
        self.profile.MoveFileless = True

        def make_fail(*args):
            raise MoveFailedException("")

        self.processor._do_operation = make_fail
        with self.assertRaises(MoveFailedException):
            self.processor.process_book(self.b, self.profile)
        self.assertFalse(File.Exists(new_path))
        self.assertFalse(Directory.Exists(new_path_folder))


class TestSimulatedBookProcessor(BookAndProfileTestCase):

    def setUp(self):
        super(TestSimulatedBookProcessor, self).setUp()
        self.profile.Mode = Mode.Simulate
        self.processor = SimulatedBookProcessor(ComicRack(), MoveReporter())

    def test_simple_move(self):
        file_path = self._create_temp_file("Testsimulatesimplemove.cbz")
        new_path = self.create_temp_path("shouldnotbecreated.cbz")
        self.book.FilePath = file_path
        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertEquals(self.book.FilePath, file_path)
        self.assertFalse(File.Exists(new_path))

    def test_tags_are_not_added(self):
        file_path = self._create_temp_file("Testsimulatesimplemove2.cbz")
        new_path = self.create_temp_path("shouldnotbecreated2.cbz")
        self.book.FilePath = file_path
        self.book.Tags = "Tags"
        self.profile.SuccessTags = ["Sucess"]
        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertEquals(self.book.FilePath, file_path)
        self.assertFalse(File.Exists(new_path))
        self.assertEqual(self.book.Tags, "Tags")

    def test_custom_values_not_added(self):
        file_path = self._create_temp_file("Testsimulatesimplemove2.cbz")
        new_path = self.create_temp_path("shouldnotbecreated2.cbz")
        self.book.FilePath = file_path
        self.profile.SuccessCustomValues = {"Success": "Custom"}
        self.processor.process_book(BookToMove(self.book, new_path, 0, []), self.profile)
        self.assertEquals(self.book.FilePath, file_path)
        self.assertFalse(File.Exists(new_path))
        self.assertEquals(self.book.GetCustomValue("Success"), None)


