import clr
clr.AddReference("System.IO")
from System.IO import DirectoryInfo, Path, File, FileInfo

import unittest
import bookmover
from bookmover import BookMover, MoveResult, BookToMove
from ComicRack import ComicRack, ComicBook
from locommon import Mode
from lologger import Logger
from losettings import Profile
import i18n

#TODO: Tests to make:
#  process_books
#_create_book_path
#_process_book
# _process_duplicate_book
# _move_book
# _save_fileless_image
# _book_should_be_moved_with_rules
# _ check_bookpath_same_as_newpath
# _get_fixed_path
# _get_samller_path

class TestBookmover(unittest.TestCase):
    def setUp(self):
        """Sets up a bookmover object and profile the subclasses can use"""
        i18n.setup(ComicRack())
        bookmover.ComicRack = ComicRack()
        self.profile = Profile()
        self.profile.Mode = Mode.Move
        self.mover = BookMover(None,None, Logger())
        self.mover.profile = self.profile

class TestBookMoverCreateDeleteFolders(TestBookmover):

    def test_create_folder_new(self):
        """Tests that the folder creation works"""
        d = DirectoryInfo(create_path("LOTEST"))
        try:
            r = self.mover._create_folder(d)
            self.assertEqual(r, MoveResult.Success)
            d.Refresh()
            self.assertTrue(d.Exists)
        finally:
            if d.Exists:
                d.Delete()

    def test_create_folder_simulate(self):
        """Tests that the folder creation in simulate mode doesn't create a folder"""
        d = DirectoryInfo(create_path("LOTEST"))
        self.profile.Mode = Mode.Simulate
        r = self.mover._create_folder(d)
        self.assertEqual(r, MoveResult.Success)
        d.Refresh()
        self.assertFalse(d.Exists)
        print self.mover.report.ToArray()
        if d.Exists:
            d.Delete()
            
    def test_delete_folder(self):
        """Tests that the delete folder functions with a single empty folder"""
        f1 = create_path("LOTEST")
        self.profile.Mode = Mode.Move
        d = DirectoryInfo(f1)
        d.Create()
        try:
            self.mover._delete_empty_folders(d)
            d.Refresh()
            self.assertFalse(d.Exists)
        finally:
            if d.Exists:
                d.Delete()
                
    def test_delete_nested_folders(self):
        """Tests deleting recursively empty folders """
        f1 = create_path("LOTEST")
        f2 = create_path("LOTEST\LOTEST2")
        self.profile.Mode = Mode.Move
        d = DirectoryInfo(f1)
        d2 = DirectoryInfo(f2)
        d2.Create()
        try:
            self.mover._delete_empty_folders(d2)
            d2.Refresh()
            d.Refresh()
            self.assertFalse(d2.Exists)
            self.assertFalse(d.Exists)

        finally:
            if d2.Exists:
                d2.Delete()
            if d.Exists:
                d.Delete()
                
    def test_delete_exlcuded_folder(self):
        """Tests that an excluded folder will not be deleted"""
        f1 = create_path("LOTEST")
        self.profile.ExcludedEmptyFolder.append(f1)
        d = DirectoryInfo(f1)
        d.Create()
        try:
            self.mover._delete_empty_folders(d)
            d.Refresh()
            self.assertTrue(d.Exists)
        finally:
            if d.Exists:
                d.Delete()
    

class TestBookMoverDuplicates(TestBookmover):

    def test_duplicate_move_overwrite(self):
        """ Mocks passing a duplicate that needs to be overwritten to the 
        main duplicate handling method. 
        """
        def returnoverwrite(*args, **kwargs):
            return self.DuplicateResult(1,False)
        self.mover._duplicate_window.ShowDialog = returnoverwrite

        dup_path = create_path("test overwrite.txt")
        File.Create(dup_path).Close()

        to_move_path = create_path("test to move.txt")
        File.Create(to_move_path).Close()
        
        c = ComicBook()
        c.FilePath = to_move_path
        c.FileDirectory = DirectoryInfo(to_move_path).FullName
        try:
            b = BookToMove(c, dup_path,0,None)
            assert self.mover._process_duplicate_book(b) == MoveResult.Success
            assert File.Exists(dup_path) == True
            assert File.Exists(to_move_path) == False
        finally:
            File.Delete(dup_path)
            if File.Exists(to_move_path):
                File.Delete(to_move_path)
                
    def test_duplicate_move_overwrite_simulate(self):
        """ Tests that the duplicate handling works correctly when overwrite is
        chosen with simulate mode. 
        """
        def returnoverwrite(*args, **kwargs):
            return self.DuplicateResult(1,False)
        self.mover._duplicate_window.ShowDialog = returnoverwrite
        self.profile.Mode = Mode.Simulate
        dup_path = create_path("test overwrite.txt")
        File.Create(dup_path).Close()

        to_move_path = create_path("test to move.txt")
        File.Create(to_move_path).Close()
        
        c = ComicBook()
        c.FilePath = to_move_path
        c.FileDirectory = DirectoryInfo(to_move_path).FullName
        try:
            b = BookToMove(c, dup_path,0,None)
            assert self.mover._process_duplicate_book(b) == MoveResult.Success
            assert File.Exists(dup_path) == True
            assert File.Exists(to_move_path) == True
        finally:
            File.Delete(dup_path)
            if File.Exists(to_move_path):
                File.Delete(to_move_path)
                
    def test_duplicate_overwrite_with_different_extension(self):
        """ Tests that the duplicate handling works correctly when overwrite is
        chosen with simulate mode. 
        """
        def returnoverwrite(*args, **kwargs):
            return self.DuplicateResult(1,False)
        self.mover._duplicate_window.ShowDialog = returnoverwrite
        dup_path = create_path("test overwrite.tmp")
        File.Create(dup_path).Close()
        dup_file = FileInfo(dup_path)
        to_move_path = create_path("test to move.txt")
        File.Create(to_move_path).Close()
        dest_path = create_path("test.txt")
        c = ComicBook()
        c.FilePath = to_move_path
        c.FileDirectory = DirectoryInfo(to_move_path).FullName
        try:
            b = BookToMove(c, dest_path,0,None)
            b.duplicate_different_extension = True
            b.duplicate_ext_files.append(dup_file)
            assert self.mover._process_duplicate_book(b) == MoveResult.Success
            assert File.Exists(dup_path) == False
            assert File.Exists(to_move_path) == False
            assert File.Exists(dest_path) == True
        finally:
            File.Delete(dup_path)
            if File.Exists(to_move_path):
                File.Delete(to_move_path)
            if File.Exists(dest_path):
                File.Delete(dest_path)
                
    def test_duplicate_overwrite_with_multiple_different_extension(self):
        """ Tests that the duplicate handling works correctly when overwrite is
        chosen with simulate mode. 
        """
        def returnoverwrite(*args, **kwargs):
            return self.DuplicateResult(1,False)
        self.mover._duplicate_window.ShowDialog = returnoverwrite
        dup_path = create_path("testbook.cbt")
        File.Create(dup_path).Close()
        dup_path2 = create_path("testbook.cbr")
        File.Create(dup_path2).Close()
        self.mover.profile.DifferentExtensionsAreDuplicates = True
        to_move_path = create_path("test to move.txt")
        File.Create(to_move_path).Close()
        dest_path = create_path("testbook.cbz")
        c = ComicBook()
        c.FilePath = to_move_path
        c.FileDirectory = DirectoryInfo(to_move_path).FullName
        try:
            b = BookToMove(c, dest_path,0,None)
            b.duplicate_different_extension = True
            b.duplicate_ext_files.append(FileInfo(dup_path))
            b.duplicate_ext_files.append(FileInfo(dup_path2))
            assert self.mover._process_duplicate_book(b) == MoveResult.Success
            self.assertFalse(File.Exists(dup_path))
            self.assertFalse(File.Exists(dup_path2))
            self.assertFalse(File.Exists(to_move_path))
            self.assertTrue(File.Exists(dest_path))
        finally:
            if File.Exists(dup_path): File.Delete(dup_path)
            if File.Exists(dup_path2): File.Delete(dup_path)
            if File.Exists(to_move_path): File.Delete(to_move_path)
            if File.Exists(dest_path): File.Delete(dest_path)
                
    def test_delete_duplicate(self):
        """ This tests that the duplicate delete code works as expected"""
        try:
            dup_path = Path.Combine(Path.GetTempPath(),"test duplicate.txt")
            f1 = File.Create(dup_path)
            f1.Close()
            self.mover._delete_duplicate(dup_path)
            assert not File.Exists(dup_path)
        finally:
            File.Delete(dup_path)

    def test_delete_duplicate_in_use(self):
        """ This tests that the duplicate delete code fails gracefully when
        the duplicate file is in use."""
        try:
            dup_path = Path.Combine(Path.GetTempPath(),"test duplicate.txt")
            f1 = File.Create(dup_path)
            assert self.mover._delete_duplicate(dup_path) == MoveResult.Failed
            assert File.Exists(dup_path)
        finally:
            f1.Close()
            File.Delete(dup_path)

    def test_create_rename_path(self):
        """ This tests the create rename path function that rename paths are
        created correctly when several exist."""
        try:
            f1 = create_path("test (1).txt")
            f2 = create_path("test (2).txt")
            File.Create(f1).Close()
            File.Create(f2).Close()
            self.assertEqual(
                self.mover._create_rename_path(create_path("test.txt")), 
                create_path("test (3).txt"))
        finally:
            File.Delete(f1)
            File.Delete(f2)

    def test_get_files_with_different_ext(self):
        """ Test that the function picks up files with different ext."""
        try:
            f1 = create_path("test.cbt")
            f2 = create_path("test.cbr")
            File.Create(f1).Close()
            File.Create(f2).Close()
            a = self.mover._get_files_with_different_ext(create_path("test.cbz"))

        finally:
            File.Delete(f1)
            File.Delete(f2)
        self.assertTrue(len(a) == 2)

    class DuplicateResult(object):
        def __init__(self, action, always_do_action):
            self.action = action
            self.always_do_action = always_do_action
            
    def test_get_files_with_different_ext_empty(self):
        """ Tests that the correct input is returns if there are no duplicates
        with different extensions.
        """
        a = self.mover._get_files_with_different_ext(create_path("test.cbz"))
        self.assertTrue(len(a) == 0)

def create_path(f):
    return Path.Combine(Path.GetTempPath(),f)
    
if __name__ == '__main__':
    try:
        unittest.main()
    except Exception as ex:
        print ex

