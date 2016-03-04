import clr
import re

from duplicatewindow import DuplicateWindow, DuplicateAction
from movereporter import MoveReporter

clr.AddReference("Microsoft.VisualBasic")
clr.AddReference("System.IO")
from Microsoft.VisualBasic import FileIO
from System import OperationCanceledException, UnauthorizedAccessException
from System.Security import SecurityException
from System.IO import Path, File, IOException, FileInfo

# 3rd-party
clr.AddReference("NLog.dll")
from NLog import LogManager
# Local
from bookprocessor import BookProcessor, FilelessBookProcessor, SimulatedBookProcessor
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
        self._book_processor = BookProcessor(comicrack)
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
        processor = self._get_processor(book_to_move.book, profile.Mode)
        try:
            processor.process_book(book_to_move, profile)
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
        else:
            return self._book_processor

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


class DuplicateHandler(object):
    """Handles functions for processing a duplicate book."""

    def __init__(self, reporter, moved_books, comicrack):
        self._reporter = reporter
        self._comicrack = comicrack
        self._simulated_moved_books = moved_books
        self._log = LogManager.GetLogger("DuplicateHandler")
        self._duplicate_window = None
        self.duplicate_action = None
        self.always_do_duplicate_action = None

    def _create_rename_path(self, path):
        """Creates a Windows formatted duplicate path in the manner of
        Filename (#).abc. It searches until it finds the first unused number.

        Originally written by pescuma and modified slightly.

        Args:
            path (str): The fully qualified path to find the duplicate for.

        Returns:
            The fully qualified path that was created.
        """
        self._log.Trace("Trying to find a rename path")
        extension = Path.GetExtension(path)
        base = path[:-len(extension)]
        base = re.sub(" \([0-9]\)$", "", base)
        i = 0
        while True:
            i += 1
            new_path = "{0} ({1}){2}".format(base, i + 1, extension)
            # For simulated mode we check _simulated_moved_books
            if new_path in self._simulated_moved_books or File.Exists(new_path):
                continue
            else:
                self._log.Info("Found rename path: {0}", new_path)
                return new_path

    def _delete_duplicate(self, path):
        """Deletes a duplicate file by sending it to the Recycle bin
        and handles the errors.

        Args:
            path: The string file path to try and move to the recycle bin.

        Returns: MoveResult.Success if it succeeds or MoveResult.Failed if it
            fails

        Raises:
            MoveFailedException if the duplicate cannot be deleted.
        """
        self._log.Trace("Trying to delete {0}", path)
        try:
            FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            self._log.Info("Deleted duplicate at {0}", path)
        except (IOException, OperationCanceledException, SecurityException, UnauthorizedAccessException) as e:
            raise MoveFailedException("Could not delete the existing duplicate {0} : {1}".format(path, e.Message))

    def _find_book_in_library(self, path):
        """Tries to find a ComicBook in the ComicRack library by checking the
        path against all the books in the library.

        It is possible to have different cases in the stored path and
        created paths, therefore both paths are converted to lowercase despite
        the possible performance hit this may have.

        Args:
            path: The string path to search in the library for.

        Returns:
            A ComicBook object if located in the library, None otherwise.
        """
        self._log.Info("Trying to find the duplicate book in the library")
        path = path.lower()
        for book in self._comicrack.App.GetLibraryBooks():
            # Fix Issue #5: different capitalization in path names.
            if book.FilePath.lower() == path:
                self._log.Info("Found a duplicate in the library")
                return book
        self._log.Info("Could not find a duplicate in the library")
        return None

    def handle_duplicate_book(self, book_to_move, profile, count):
        """ Process a duplicate book by asking the user what to do.

        If the user chooses to overwrite then the duplicate file if sent to the Recycle Bin and the script returns
        DuplicateAction.Overwrite.

        If the user chooses to rename then the script sets the rename_path to the book_to_move.path and returns
        DuplicateAction.Rename

        If the user chooses to Cancel then the MoveSkippedException is raised to notify the reporter that the book
        is skipped.

        Args:
            book_to_move (BookToMove): The BookToMove object with a duplicate to deal with.
            profile (Profile): The profile being used.
            count (int): The number of held duplicate books to display in the
                duplicate form.

        Returns:
                DuplicateAction.Rename if the user chooses to rename the book.
                DuplicateAction.Overwrite if the user chooses to overwrite the
                    book.

        Raises:
                MoveSkippedException: If the user skips the duplicates.
                MoveFailedExceptions: If the duplicate cannot be deleted.
        """
        if book_to_move.duplicate_different_extension:
            return self._process_different_extension_books(book_to_move, profile)

        move_path = book_to_move.path
        self._log.Info("Starting duplicate handling for {0} \n with a duplicate at {1}",
                       self._reporter.current_book,
                       move_path)

        duplicate_book = self._find_book_in_library(move_path)

        if duplicate_book is None:
            # No need to catch errors here since this path was validated in the BookProcessor
            duplicate_book = FileInfo(move_path)

        rename_path = self._create_rename_path(move_path)

        # Check if the user has an action saved for every book
        if self.always_do_duplicate_action is not None:
            action = self.always_do_duplicate_action
        else:
            result = self._ask_user(book_to_move.book, duplicate_book, rename_path, profile.Mode, count)
            action = result.action

            # Save action if required
            if result.always_do_action:
                self.always_do_duplicate_action = action

        return self._do_action(action, book_to_move, duplicate_book, rename_path, profile)

    def _process_different_extension_books(self, book_to_move, profile):
        """ Processes different extension files. These have to be dealt with differently since there is possible
        more
        than one.

        Each file is asked what to do separately.
                If renamed is chosen for any then no more are asked.
                If cancel is chosen then it continues asking and if all were skipped the MoveSkippedException is
                raised.
                If overwrite is chosen
        Args:
            book_to_move (BookToMove):
            profile (Profile):

        Returns:
            DuplicateAction.Rename if the file should be moved to the correct location
            DuplicateAction.Overwrite if any of the file are deleted

        Raises:
            MoveSkippedException if all files were skipped
        """
        self._log.Info("Starting duplicate handling for {0} \n with duplicates with different extensions",
                       self._reporter.current_book)
        count = len(book_to_move.duplicate_ext_files)
        skipped = []
        always_action = None

        for file_info in book_to_move.duplicate_ext_files:

            duplicate_book = self._find_book_in_library(file_info.FullName)

            if duplicate_book is None:
                duplicate_book = file_info

            if always_action is None:
                result = self._ask_user(book_to_move.book, duplicate_book, book_to_move.path, profile.Mode, count)
                action = result.action
                if result.always_do_action:
                    always_action = action
            else:
                action = always_action

            # If there is several duplicate file and the user chooses to cancel only one of them then we will
            # want to
            # ask about each one of them if we are not at the end
            if action == MoveSkippedException:
                skipped.append(file_info)
                continue

            # if renaming we can return right away as the path won't change
            elif action == DuplicateAction.Rename:
                return action

            # if overwriting then we may want to ask the user about other duplicates.
            elif action == DuplicateAction.Overwrite:
                self._overwrite(book_to_move.book, file_info.FullName, duplicate_book, profile.Mode,
                                profile.CopyReadPercentage)

        if skipped:
            text = "\n".join(
                ["A file already exists at: {0} and the user declined to overwrite or rename it".format(
                    file_info.path)
                 for file_info in skipped])
            # All were skipped so report that we can skip this file
            if len(skipped) == len(book_to_move.duplicate_ext_files):
                raise MoveSkippedException(text)

            # Not all were skipped, log the text and return Overwrite
            else:
                _log.Info(text)

        # If the code reaches this point then it will be overwrite as rename is returned already and skipped is
        # raised if all actions were skip
        return DuplicateAction.Overwrite

    def _do_action(self, action, book_to_move, duplicate_book, rename_path, profile):
        """ Performs the required action

        Args:
            action:
            book_to_move:
            duplicate_book:
            rename_path:
            mode:

        Returns:

        """
        if action == DuplicateAction.Cancel:
            # Raising MoveSkippedException will let the BookManager handle the skipped book.
            self._cancelled(book_to_move)
        elif action == DuplicateAction.Rename:
            # Simply set the book_to_move_path as the new path and let the BookManager move the book
            self._rename(book_to_move, rename_path)
        elif self.duplicate_action == DuplicateAction.Overwrite:
            self._overwrite(book_to_move.book, book_to_move.path, duplicate_book, profile.mode,
                            profile.CopyReadPercentage)

    def _cancelled(self, book_to_move):
        raise MoveSkippedException(
            "A file already exists at: {0} and the user declined to overwrite or rename it".format(
                book_to_move.path))

    def _rename(self, book_to_move, rename_path):
        self._log.Info("User chose to rename the processing ComicBook")
        book_to_move.path = rename_path
        return DuplicateAction.Rename

    def _overwrite(self, book, duplicate_path, duplicate_book, mode, copy_read_percent):
        """ The Overwrite Action will move the duplicate_path to the Recycle Bin
        and remove the duplicate book from the ComicRack library (if it exists)

        Args:
            book:
            duplicate_path:
            duplicate_book:
            mode:

        Returns:

        Raises:
            MoveFailedException if the file is unable to be deleted

        """
        self._log.Info("User choose to overwrite the duplicate")
        if mode == Mode.Simulate:
            self._overwrite_simulated(duplicate_path)
        else:
            self._delete_duplicate(duplicate_path)
            if book.FilePath and type(duplicate_book) is not FileInfo:
                if copy_read_percent:
                    book.LastPageRead = duplicate_book.LastPageRead
                    self._log.Info("Copied the read percentage from the duplicate book")
                self._comicrack.App.RemoveBook(duplicate_book)
                self._log.Info("Removed the duplicate book from the library")
        return DuplicateAction.Overwrite

    def _overwrite_simulated(self, move_path):
        _log.Info("Deleted (simulated) {0}", move_path)
        self._reporter.log("Deleted (simulated)", move_path)

    def _ask_user(self, source_book, duplicate_book, rename_path, mode, count):
        if self._duplicate_window is None:
            self._duplicate_window = DuplicateWindow(self._comicrack)

        return self._duplicate_window.ShowDialog(source_book, duplicate_book, rename_path, mode, count)
