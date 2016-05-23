import clr

from duplicatehandler import DuplicateHandler
from movereporter import MoveReporter

clr.AddReference("Microsoft.VisualBasic")
clr.AddReference("System.IO")
from System.IO import Path

# 3rd-party
clr.AddReference("NLog.dll")
from NLog import LogManager
# Local
from bookprocessor import MoveBookProcessor, FilelessBookProcessor, SimulatedBookProcessor, CopyBookProcessor
from common import BookToMove, Mode, MoveFailedException, MoveSkippedException, DuplicateExistsException
from locommon import check_excluded_folders, check_metadata_rules
from pathmaker import PathMaker

_log = LogManager.GetLogger("BookManager")


class BookManager(object):
    def __init__(self, comicrack, worker):
        """

        Args:
            comicrack:
            worker:

        Returns:

        """
        self._comic_rack = comicrack
        self._path_maker = PathMaker(None, None)
        self._report = MoveReporter()

        # Different Processors
        self._move_book_processor = MoveBookProcessor(comicrack)
        self._copy_book_processor = CopyBookProcessor(comicrack)
        self._fileless_book_processor = FilelessBookProcessor(comicrack)
        self._simulate_book_processor = SimulatedBookProcessor(comicrack, self._report)

        # Progress Reporting
        self._progress = 0.0
        self._progress_increment = 0
        self._worker = worker

        # Duplicates
        self._held_duplicate_books = []

    def process_books(self, books, profiles):
        """ The main entry point to process a collection of books

        * Generates paths for each book
        * Does the operation for that book

        Args:
            books: A list containing the books to move
            profiles: A list containing the profiles to use to move
                these books. They should be order so that the last possible
                profile will be used to move the book

        Returns:

        """
        _log.Info("Starting to process {0} books", len(books))

        self._report.create_profile_reports(profiles, len(books))
        books_to_move = self._create_paths(books, profiles)

        # Setup progress
        self._progress_increment = 1.0 / len(books_to_move) * 100

        # Initial processing
        for book in books_to_move:
            if self._handle_cancellation():
                return
            profile = profiles[book.profile_index]
            self._process_book(book, profile)

        # Handle Duplicate books
        if self._held_duplicate_books:
            self._process_duplicate_books(profiles)

        return self._report

    def _process_book(self, book_to_move, profile):

        self._set_book_and_profile(book_to_move, profile)
        try:
            self._get_processor(book_to_move.book, profile.Mode).process_book(book_to_move, profile)
        except MoveFailedException as e:
            self._report.fail(e.message)
            _log.Error("Failed: {0}", e.message)
        except MoveSkippedException as e:
            self._report.skip(e.message)
            _log.Info("Skipped {0}", e.message)
        except DuplicateExistsException:
            self._held_duplicate_books.append(book_to_move)
            return
        else:  # Succeeds!
            self._report.success()

        self._increase_progress()

    def _set_book_and_profile(self, book_to_move, profile):
        self._report.set_profile(profile)
        self._report.current_book = book_to_move.book

    def _process_duplicate_books(self, profiles):
        _log.Info(
            "\n {0} Beginning to process {1} duplicates books {0}".format("*" * 10, len(self._held_duplicate_books)))
        count = len(self._held_duplicate_books)
        duplicate_handler = DuplicateHandler(self._report, self._simulate_book_processor.moved_books, self._comic_rack)
        for book_to_move in self._held_duplicate_books:
            if self._handle_cancellation():
                return
            profile = profiles[book_to_move.profile_index]
            self._set_book_and_profile(book_to_move, profile)
            self._process_duplicate_book(book_to_move, profile, duplicate_handler, count)
            count -= 1

    def _process_duplicate_book(self, book_to_move, profile, duplicate_handler, count):
        try:
            duplicate_handler.handle_duplicate_book(book_to_move, profile, count)
            processor = self._get_processor(book_to_move.book, profile.Mode)
            processor.process_book(book_to_move, profile, True)
        except MoveFailedException as e:
            self._report.fail(e.message)
            _log.Error("Failed: {0}", e.message)
        except MoveSkippedException as e:
            self._report.skip(e.message)
            _log.Info("Skipped {0}", e.message)
        except DuplicateExistsException:
            return self._process_duplicate_book(book_to_move, profile, duplicate_handler, count)
        else:  # Succeeds!
            self._report.success()
        self._increase_progress()

    #region path making

    def _create_paths(self, books, profiles):
        """ Creates a path for each book

        Goes through each profile for each book and creates the possible path for the book.

        Only the last possible path is used.

        Returns: An array of BookToMove Objects
        """
        _log.Info("Finding paths for all the books using {0} profiles", len(profiles))
        books_to_move = []

        for book in books:
            self._report.current_book = book
            _log.Info("\n" + "*" * 40 + "\n")
            _log.Info("Finding profile for {0}", self._report.current_book)

            # Fetch the profile
            index = self._get_profile_index_for_book(book, profiles)

            if index == -1:  # No profile found so skip this book
                _log.Info("Couldn't find a profile. Skipping.")
                continue

            _log.Info("Moving with profile {0}", profiles[index].Name)

            # Start Generating the path
            self._report.set_profile(profiles[index])
            result = self._create_path(book, profiles[index])

            # If failed
            if result is None:
                self._report.fail("Field(s) {0} empty.".format(self._path_maker.failed_fields))
                _log.Info("Field(s) {0} empty.", self._path_maker.failed_fields)
            else:
                books_to_move.append(BookToMove(book, result, index, self._path_maker.failed_fields))
        return books_to_move

    def _create_path(self, book, profile):
        """ Creates a path for a book with a profile

        Returns: The created file path

        """
        _log.Trace("Beginning create_book_path for {0}", self._report.current_book)
        self._path_maker.profile = profile  # TODO: fix this on pathmaker
        folder_path, file_name, failed = self._path_maker.make_path(book, profile.FolderTemplate,
                                                                    profile.FileTemplate)
        full_path = Path.Combine(folder_path, file_name)
        _log.Info("Created the path: {0}", full_path)

        if failed and not profile.MoveFailed:
            _log.Warn("Failed fields and not moving to a location")
            return None
        return full_path

    def _get_profile_index_for_book(self, book, profiles):
        current_profile_index = -1

        for index, profile in enumerate(profiles):
            self._report.set_profile(profile)

            if not self._book_should_be_moved(book, profile):
                continue

            # If there is a previous valid profile, mark it as skipped in the report
            if current_profile_index != -1:
                self._report.skip("The book is moved by a later profile", profiles[current_profile_index].Name)

            # Set the current valid index
            current_profile_index = index

        return current_profile_index

    def _book_should_be_moved(self, book, profile):
        """ Checks the excluded folders and metadata rules to see if the book
        should be moved.

        Handles reporting the skip to the report if the book should be skipped

        Args:
            book: The ComicBook object to check with.
        Returns:
            True if the book should be moved. False otherwise
        """
        if not check_excluded_folders(book.FilePath, profile):
            _log.Info("Skipped because the book is located in an excluded path")
            self._report.skip("The book is located in an excluded path")
            return False

        if not check_metadata_rules(book, profile):
            _log.Info("Skipped because the book was excluded under the exclude rules")
            self._report.skip("The book qualified under the exclude rules")
            return False
        return True

    #endregion

    def _get_processor(self, book, mode):
        if mode == Mode.Simulate:
            return self._simulate_book_processor
        elif not book.FilePath:
            return self._fileless_book_processor
        elif mode == Mode.Move:
            return self._move_book_processor
        elif mode == Mode.Copy:
            return self._copy_book_processor
        else:
            raise MoveFailedException("Unknown Mode")

    def _increase_progress(self, book_to_move, profile_mode):
        """Increases the progress percentage and reports to the worker form,
        if any.
        """
        if self._worker is None:
            return

        self._progress += self._progress_increment
        book = book_to_move.book
        progress_text = "{0} {1} to {2}"
        if profile_mode == Mode.Move:
            progress_text.format("Move", book.FilePath, book_to_move.path)
        elif profile_mode == Mode.Copy:
            progress_text.format("Copying", book.FilePath, book_to_move.path)
        else:
            progress_text.format("Moving (Simulated)", book.FilePath, book_to_move.path)
        # self._worker.ReportProgress(int(round(self._progress)))
        self._worker.ReportProgress(int(round(self._progress)), (progress_text, self._comic_rack.App.GetComicThumbnail(book, book.PreferredFrontCover)))

    def _handle_cancellation(self):
        if self._worker is not None and self._worker.CancellationPending:
            remaining = (100.0 - self._progress) * self._progress_increment
            _log.Debug(remaining)
            self._report.cancelled = True
            self._report.log("Cancelled", "{0} operations".format(remaining), "User cancelled the script")
            return True
        return False