import clr
from unittest import TestCase

clr.AddReference("System.IO")
import bookmover
from bookmover import BookMover
from ComicRack import ComicRack
clr.AddReference('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import ComicBook  # @UnresolvedImport
from locommon import Mode
from losettings import Profile
from movereporter import MoveReporter
import i18n
clr.AddReference("NLog.dll")
from NLog import LogManager

log = LogManager.GetLogger("BookMoverTests")


class TestBookMover(TestCase):

    def setUp(self):
        """Sets up a bookmover object and profile the subclasses can use"""
        i18n.setup(ComicRack())
        bookmover.ComicRack = ComicRack()
        self.profile = Profile()
        self.profile.Name = "Test"
        self.profile.Mode = Mode.Move
        self.book = ComicBook()
        self.mover = BookMover(None, None, MoveReporter())
        self.mover.profile = self.profile
        self.mover._set_book_and_profile(self.book, self.profile)
        self.mover._report.create_profile_reports([self.profile], 1)

    def tearDown(self):
        print self.mover._report._log

    def test__get_canceled_report(self):
        self.fail()

    def test__set_book_and_profile(self):
        self.fail()

    def test__increase_progress(self):
        self.fail()

    def test_process_books(self):
        self.fail()

    def test__create_book_paths(self):
        self.fail()

    def test__create_book_path(self):
        self.fail()

    def test__process_book(self):
        self.fail()

    def test__check_destination(self):
        self.fail()

    def test__process_duplicate_book(self):
        self.fail()

    def test__check_for_duplicates(self):
        self.fail()

    def test__do_mode_operation(self):
        self.fail()

    def test__move_book(self):
        self.fail()

    def test__copy_book(self):
        self.fail()

    def test__move_book_simulated(self):
        self.fail()

    def test__create_folder(self):
        self.fail()

    def test__create_folder_simulated(self):
        self.fail()

    def test__save_fileless_image(self):
        self.fail()

    def test__book_should_be_moved(self):
        self.fail()

    def test__check_book_already_at_dest(self):
        self.fail()

    def test__get_shorter_path(self):
        self.fail()
