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

class TestBookmover(unittest.TestCase):
    def setUp(self):
        """Sets up a bookmover object and profile the subclasses can use"""
        i18n.setup(ComicRack())
        bookmover.ComicRack = ComicRack()
        self.profile = Profile()
        self.profile.Mode = Mode.Move
        self.mover = BookMover(None,None, Logger())
        self.mover.profile = self.profile


class TestBookMoverDeleteDuplicate(TestBookmover):
    def test_duplicate_in_use(self):
        """ This tests that the exceptions are dealt with correctly when
        the duplicate can't be deleted after overwrite is chosen"""
        dup_path = Path.Combine(Path.GetTempPath(),"test duplicate.txt")
        f1 = File.Create(dup_path)
        self.assertFalse(self.mover._delete_duplicate(dup_path))
        f1.Close()
        File.Delete(dup_path)


class TestBookMoverCreateFolders(TestBookmover):

    def test_create_folder_new(self):
        d = DirectoryInfo(Path.Combine(Path.GetTempPath(),"LOTEST"))
        r = self.mover._create_folder(d)
        self.assertEqual(r, MoveResult.Success)
        d.Refresh()
        self.assertTrue(d.Exists)
        if d.Exists:
            d.Delete()

    def test_create_folder_simulate(self):
        d = DirectoryInfo(Path.Combine(Path.GetTempPath(),"LOTEST"))
        self.profile.Mode = Mode.Simulate
        r = self.mover._create_folder(d)
        self.assertEqual(r, MoveResult.Success)
        d.Refresh()
        self.assertFalse(d.Exists)
    
    
class TestBookMoverGetRenamePath(TestBookmover):
    def test_create_rename_path(self):
        f1 = File.Create(Path.Combine(Path.GetTempPath(),"test (1).txt"))
        f2 = File.Create(Path.Combine(Path.GetTempPath(),"test (2).txt"))
        f1.Close()
        f2.Close()
        self.assertEqual(
            self.mover._create_rename_path(Path.Combine(Path.GetTempPath(),"test.txt")), 
            Path.Combine(Path.GetTempPath(),"test (3).txt"))
        File.Delete(Path.Combine(Path.GetTempPath(),"test (1).txt"))
        File.Delete(Path.Combine(Path.GetTempPath(),"test (2).txt"))

class TestBookMoverDuplicates(TestBookmover):

    def test_duplicate_move_overwrite(self):
        """ Mocks passing a duplicate that needs to be overwritten to the 
        main duplicate handeling method. 
        """
        def returnoverwrite(*args, **kwargs):
            return self.DuplicateResult(1,False)
        self.mover._duplicate_window.ShowDialog = returnoverwrite

        dup_path = Path.Combine(Path.GetTempPath(),"test overwrite.txt")
        dup_file = File.Create(dup_path)
        dup_file.Close()

        to_move_path = Path.Combine(Path.GetTempPath(),"test to move.txt")
        to_move_file = File.Create(to_move_path)
        to_move_file.Close()
        c = ComicBook()
        c.FilePath = to_move_path
        c.FileDirectory = DirectoryInfo(to_move_path).FullName
        b = BookToMove(c, dup_path,0,None)
        assert self.mover._process_duplicate_book(b) == MoveResult.Success
        assert File.Exists(dup_path) == True
        assert File.Exists(to_move_path) == False
        File.Delete(dup_path)


    class DuplicateResult(object):
        def __init__(self, action, always_do_action):
            self.action = action
            self.always_do_action = always_do_action

    
if __name__ == '__main__':
    try:
        unittest.main()
    except Exception as ex:
        print ex

