import unittest
import i18n
import ComicRack


class Test_test_i18n(unittest.TestCase):
    def test_get_defaults(self):
        localizer = i18n.__i18n(ComicRack.ComicRack)
        self.assertTrue(len(localizer.defaults) > 0)

if __name__ == '__main__':
    unittest.main()
