import unittest
from System.IO import File, Path, DirectoryInfo
from fileutils import get_files_with_different_ext, delete_empty_folders


class TestFileUtils(unittest.TestCase):
    def test_get_files_with_different_ext(self):
        """ Test that the function picks up files with different ext."""
        f1 = create_temp_path("test.cbt")
        f2 = create_temp_path("test.cbr")
        try:
            File.Create(f1).Close()
            File.Create(f2).Close()
            a = get_files_with_different_ext(
                create_temp_path("test.cbz"))
        finally:
            File.Delete(f1)
            File.Delete(f2)
        self.assertTrue(len(a) == 2)

    def test_get_files_with_different_ext_empty(self):
        """ Tests that the correct input is returns if there are no duplicates
        with different extensions.
        """
        a = get_files_with_different_ext(create_temp_path("test.cbz"))
        self.assertTrue(len(a) == 0)

    def test_delete_folder(self):
        """Tests that the delete folder functions with a single empty folder"""
        f1 = create_temp_path("LOTEST")
        d = DirectoryInfo(f1)
        d.Create()
        try:
            delete_empty_folders(d)
            d.Refresh()
            self.assertFalse(d.Exists)
        finally:
            if d.Exists:
                d.Delete()

    def test_delete_nested_folders(self):
        """Tests deleting recursively empty folders """
        f1 = create_temp_path("LOTEST")
        f2 = create_temp_path("LOTEST\LOTEST2")
        d = DirectoryInfo(f1)
        d2 = DirectoryInfo(f2)
        d2.Create()
        try:
            delete_empty_folders(d2)
            d2.Refresh()
            d.Refresh()
            self.assertFalse(d2.Exists)
            self.assertFalse(d.Exists)
        finally:
            if d2.Exists:
                d2.Delete()
            if d.Exists:
                d.Delete()

    def test_delete_excluded_folder(self):
        """Tests that an excluded folder will not be deleted"""
        f1 = create_temp_path("LOTEST")
        excluded_folders = [f1]
        d = DirectoryInfo(f1)
        d.Create()
        try:
            delete_empty_folders(d, excluded_folders)
            d.Refresh()
            self.assertTrue(d.Exists)
        finally:
            if d.Exists:
                d.Delete()


def create_temp_path(f):
    """Creates a path in the temp file location
    Args:
        f: The desired local path in the Temp folder
    """
    return Path.Combine(Path.GetTempPath(), f)


if __name__ == '__main__':
    unittest.main()
