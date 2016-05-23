import clr
import re

clr.AddReference("System.IO")
clr.AddReference("Microsoft.VisualBasic")

from Microsoft.VisualBasic.FileIO import FileSystem, UIOption, RecycleOption
from System import UnauthorizedAccessException, OperationCanceledException
from System.Security import SecurityException
from System.IO import File, Path, IOException, FileInfo

from common import MoveFailedException, MoveSkippedException, Mode
from duplicatewindow import DuplicateAction, DuplicateWindow

clr.AddReference("Nlog")

from NLog import LogManager


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
        Filename (#).abc by taking the file path and stripping the extension and any
        (#) at the end and then appending (#) starting at 1 and continuing until it finds a file
        path that doesn't already exist.

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
            new_path = "{0} ({1}){2}".format(base, i, extension)
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
            FileSystem.DeleteFile(path, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin)
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
                self._log.Info(text)

        # If the code reaches this point then it will be overwrite as rename is returned already and skipped is
        # raised if all actions were skip
        return DuplicateAction.Overwrite

    def _do_action(self, action, book_to_move, duplicate_book, rename_path, profile):
        """ Performs the required action

        If action is DuplicateAction.Cancel a MoveSkippedException is raise
        If action is DuplicateAction.Rename the book_to_move path is changed to the generated duplicate path
        If action is Overwrite the existing book is moved to the recyclebin

        Args:
            action: The DuplicateAction to perform
            book_to_move: The BookToMove object
            duplicate_book: A ComicBook object if the book exists in the CR library or a FileInfo
                            object if it does not
            rename_path: The string path of the calculated path to rename the book to
            profile: The profile this book is being moved with.

        Returns:
            The same action that was passed in.

        Raises:
            MoveSkippedException if the user chose to cancel.

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
        return action

    def _cancelled(self, book_to_move):
        """ Creates and raises a MoveSkippedException that the user skipped overwriting or renaming the book

        Args:
            book_to_move: The BookToMove object of the book being moved.

        Raises: MoveSkippedException.
        """
        raise MoveSkippedException(
            "A file already exists at: {0} and the user declined to overwrite or rename it".format(
                book_to_move.path))

    def _rename(self, book_to_move, rename_path):
        """ Sets the book_to_move destination path to the rename_path.

        Args:
            book_to_move: The BookToMove object of the book that is being moved.
            rename_path: The new path the book should be moved to.

        Returns: DuplicationAction.Rename

        """
        self._log.Info("User chose to rename the processing ComicBook")
        book_to_move.path = rename_path
        return DuplicateAction.Rename

    def _overwrite(self, book, duplicate_path, duplicate_book, mode, copy_read_percent):
        """ The Overwrite Action will move the duplicate_path to the Recycle Bin
        and remove the duplicate book from the ComicRack library (if it exists)

        If the mode is Simulate then the book isn't actually deleted.

        Args:
            book: The ComicBook object of the book being moved.
            duplicate_path: The string path the book is being moved to.
            duplicate_book: The ComicBook or FileInfo object of the book that already exists.
            mode: The profile mode for this operation
            copy_read_percent: Boolean if the read percentage should be copied from the existing book before
                it is removed from the library.

        Returns:
            DuplicateAction.Overwrite

        Raises:
            MoveFailedException if the file is unable to be deleted

        """
        self._log.Info("User choose to overwrite the duplicate")
        if mode == Mode.Simulate:
            self._overwrite_simulated(duplicate_path)
        else:
            self._delete_duplicate(duplicate_path)
            if book.FilePath and type(duplicate_book) is not FileInfo and copy_read_percent:
                if copy_read_percent:
                    book.LastPageRead = duplicate_book.LastPageRead
                    self._log.Info("Copied the read percentage from the duplicate book")
                self._comicrack.App.RemoveBook(duplicate_book)
                self._log.Info("Removed the duplicate book from the library")
        return DuplicateAction.Overwrite

    def _overwrite_simulated(self, move_path):
        """ Creates log messages that a book was deleted without actually deleting it.

        Args:
            move_path: The path that would have been deleted.
        """
        self._log.Info("Deleted (simulated) {0}", move_path)
        self._reporter.log("Deleted (simulated)", move_path)

    def _ask_user(self, source_book, duplicate_book, rename_path, mode, count):
        """ Creates and shows a dialog to the user if they want to rename, cancel, or overwrite a duplicate.

        Args duplicate window is created if required otherwise the previously created one is reused.

        Args:
            source_book: The ComicBook object being moved.
            duplicate_book: The ComicBook or File Info object that already exists at the destination
            rename_path: The generated path that the book can be renamed to.
            mode: The profile mode used for this book.
            count: The number of duplicates that need to be handled.

        Returns: The DuplicateResult the user chose.

        """
        if self._duplicate_window is None:
            self._duplicate_window = DuplicateWindow(self._comicrack)

        return self._duplicate_window.ShowDialog(source_book, duplicate_book, rename_path, mode, count)
