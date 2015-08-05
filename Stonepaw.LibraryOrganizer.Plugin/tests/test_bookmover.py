import unittest
from bookmover import BookMover, MoveResult
from losettings import Profile
from locommon import Mode
from lologger import Logger
import clr
clr.AddReference("System.IO")
from System.IO import DirectoryInfo, Path

class TestBookmover(unittest.TestCase):
    def setUp(self):
        self.p = Profile()
        self.p.Mode = Mode.Move
        self.b = BookMover(None,None, Logger())
        self.b.profile = p

    def test_create_folder_new(self):
        d = DirectoryInfo(Path.Combine(Path.GetTempPath(),"LOTEST"))
        r = self.b._create_folder(d)
        self.assertEqual(r, MoveResult.Success)
        d.Refresh()
        self.assertTrue(d.Exists)
        if d.Exists:
            d.Delete()

    def test_create_folder_simulate(self):
        d = DirectoryInfo(Path.Combine(Path.GetTempPath(),"LOTEST"))
        self.p.Mode = Mode.Simulate
        r = self.b._create_folder(d)
        self.assertEqual(r, MoveResult.Success)
        d.Refresh()
        self.assertFalse(d.Exists)

if __name__ == '__main__':
    unittest.main()
