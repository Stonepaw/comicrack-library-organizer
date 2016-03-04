# System
import clr

clr.AddReference("System")
clr.AddReference("System.IO")
clr.AddReference("System.Windows.Forms")
from System import ArgumentException, ArgumentNullException, NotSupportedException, UnauthorizedAccessException, \
    DateTime
from System.IO import IOException, FileInfo, PathTooLongException
from System.Runtime.InteropServices import ExternalException
from System.Security import SecurityException
from System.Windows.Forms import DialogResult

# 3rd-party
clr.AddReference("NLog.dll")
from NLog import LogManager

# Local
from common import BookToMove, MoveSkippedException, MoveFailedException, Mode, DuplicateExistsException
from fileutils import create_folder, get_files_with_different_ext, delete_empty_folders
from loforms import PathTooLongForm
from comicbook_utils import convert_tags_to_list, copy_data_to_new_book

_log = LogManager.GetLogger("BookProcessor")


class BaseBookProcessor(object):
    process_text = "Beginning to {0} {1}.\nDestination: {2}"

    def __init__(self, comicrack):
        self._comic_rack = comicrack

    def process_book(self, book_to_move, profile, ignore_duplicates=False):
        """ The main entry point for the class to move or copy a book.

        Args:
            book_to_move (BookToMove): The BookToMove object
            profile (Profile): The Profile this book is being moved with.
            ignore_duplicates(bool): If set true, no duplicate checking will be done. This is used for cases where
                duplicates with different extensions are found and the user chooses to rename the file. It also is used
                when processing a simulated process when duplicates were found.

        Returns:

        Raises:
            MoveFailedException: The operations fails
            MoveSkippedException: The file doesn't need to be move or the user canceled fixing a path
            DuplicateExistsException: A duplicate exists at the destination
        """
        new_path = book_to_move.path
        book = book_to_move.book
        _log.Info("\n" + "*" * 40 + "\n")  # Pretty separator
        _log.Info(self.process_text, profile.Mode, self._get_book_text(book), new_path)

        # Check source and destinations
        destination_file = self._check_destination(new_path)
        source_file = self._check_source(book, new_path)

        # Check duplicates
        if not ignore_duplicates:
            self._check_for_duplicates(destination_file, book_to_move, profile.DifferentExtensionsAreDuplicates)

        # Create destination folder if required
        self._create_folder(destination_file.Directory)

        try:
            self._do_operation(book, source_file, destination_file, profile)
        except MoveFailedException:
            # Clean up created empty folders
            delete_empty_folders(destination_file.Directory, profile.ExcludedEmptyFolder)
            raise

        # On success add tags and custom values
        self._add_tag_on_success(book, profile)
        self._add_custom_value_on_success(book, profile)

        if book_to_move.failed_fields:
            raise MoveFailedException(
                "Field(s) {0} empty. Moved to {1}".format(", ".join(book_to_move.failed_fields), new_path))

    def _check_destination(self, destination_path):
        """ Checks the destination path for errors

        Args:
            destination_path (str):

        Returns:
            FileInfo object of the destination_path
        """
        _log.Info("Checking the destination path for errors")
        result = self._check_file(destination_path, True)
        _log.Info("No errors")
        return result

    def _check_for_duplicates(self, destination_file, book, check_different_extensions):
        """ Checks if a duplicate exists by checking the file path

        If check_different_extensions is set to true then other files in the directory are checked for similar file
        names.
        The book_to_move object is modified with the correct information in this
        case.

        Args:
            destination_file (FileInfo):
            book (BookToMove): The book to move object. Required so that fields can be modified is a duplicate exists
            check_different_extensions(bool): If True then the destination folder will be checked for files with the
                same name but different file extensions

        Raises:
            DuplicateExistsException: If a duplicate exists

        """
        _log.Info("Checking for Duplicates")
        duplicate = False
        if destination_file.Exists:
            _log.Info("An exact duplicate was found at the destination, holding to ask user")
            duplicate = True

        # Test for duplicates with different extensions
        if check_different_extensions:
            files = get_files_with_different_ext(destination_file.FullName)
            # get_files also returns the exact duplicate so we don't have to worry about the file not showing up if
            # both are found, however check_different is true then it will always report duplicate_extensions exist
            if len(files) == 1 and duplicate:  # We have the same file as an exact duplicate
                pass
            elif files:
                duplicate = True
                book.duplicate_ext_files = files
                book.duplicate_different_extension = True
                _log.Info(
                    "A duplicate with a different extension was found at the destination. Holding to ask the user")

        if duplicate:
            raise DuplicateExistsException("")

        _log.Info("No Duplicates found")

    def _check_file(self, file_path, ask_about_long_path):
        """ Checks a file path for common errors by creating a FileInfo object

        Args:
            file_path (str): The file path to check
            ask_about_long_path (bool): If true and the path is too long the user is asked to shorten the path,
                otherwise it raises MoveFailedError

        Returns:
            FileInfo object of the file_path

        Raises:
            MoveFailedException if an irrecoverable error is found
            MoveSkippedException if the user declines to rename a too long path
        """
        _log.Debug("Starting _check_file")
        try:
            file_object = FileInfo(file_path)
        except ArgumentException as e:
            raise MoveFailedException("Destination path invalid: " + e.Message)
        except (UnauthorizedAccessException, SecurityException) as e:
            raise MoveFailedException("Unable to access destination path: " + e.Message)
        except PathTooLongException as e:
            if ask_about_long_path:
                new_path = self._get_shorter_path(file_path)
                return self._check_file(new_path, ask_about_long_path)
            else:
                raise MoveFailedException("The file path is too long".format(file_path))
        return file_object

    def _check_source(self, book, destination_path):
        """ Checks the source file to verify there isn't any unrecoverable issues

        Creates a FileInfo Object from the source_path and checks if:
            * The file exists
            * The file is accessible
            * The file is already at the destination

        Args:
            book (ComicBook): The ComicBook to check
            destination_path (str): The full file path of the destination file

        Returns (FileInfo): The FileInfo object referring to the source file.

        Raises:
            MoveSkippedException: if the file is already at the correct location
            MoveFailedException: if the file does not exist or is not accessible

        """
        source_path = book.FilePath
        _log.Info("Checking the source path for errors")
        if source_path == destination_path:
            raise MoveSkippedException("The file is already at the correct location")

        source_file = self._check_file(source_path, False)

        if not source_file.Exists:
            raise MoveFailedException("The file does not exist".format(source_path))

        if source_path.lower() == destination_path.lower():  # Different capitalization
            # We rename the file to the correct
            _log.Info("The book is already at the correct location but with different capitalization. Trying to rename")
            self._move_book(book, source_file, destination_path)
            raise MoveSkippedException(
                "The file is already at the correct location but was renamed to fix the capitalization")

        _log.Info("No errors")
        return source_file

    def _do_operation(self, book, source_file, destination_file, profile):
        """ The operation for this processor. Override in subclasses
        Args:
            book (ComicBook): The book to move
            source_file (FileInfo): The source file FileInfo
            destination_file (FileInfo): The destination FileInfo
            profile (Profile): The profile to use

        Raises:
            MoveFailedException: If the copy or move operations fails.
        """
        _log.Debug("Starting _do_operation")

    def _create_folder(self, destination_directory):
        """ Creates a folder if required.

        Args:
            destination_directory (DirectoryInfo): A DirectoryInfo Object of the destination path

        Raises: MoveFailedException if the folder cannot be created.
        """
        _log.Debug("_create_folder: args: {0}", destination_directory)
        try:
            create_folder(destination_directory)
        except (IOException, UnauthorizedAccessException, ArgumentException, NotSupportedException) as e:
            raise MoveFailedException(
                "Failed to create folder {0}: {1}".format(destination_directory.FullName, e.Message))

    def _get_shorter_path(self, long_path):
        """Asks the user to fix the too long path by invoking a dialog.

        Args:
            long_path (str): The fully qualified path to check as a string.

        Returns:
            The shortened, correct path.

        Raises:
            MoveSkippedException if the user cancels fixing the path.
        """
        _log.Debug("Starting _get_shorter_path")
        with PathTooLongForm(long_path) as p:
            if p.ShowDialog() == DialogResult.OK:
                return p.get_path()
        raise MoveSkippedException("The calculated path was too long and the user skipped shortening it")

    def _get_book_text(self, book):
        _log.Debug("_get_book_text: args: {0]", book)
        return book.FilePath if book.FilePath else book.Caption

    def _clean_up_empty_folders(self, directory, excluded_folders):
        _log.Debug("Starting _clean_up_empty_folders")
        delete_empty_folders(directory, excluded_folders)

    def _add_tag_on_success(self, book, profile):
        """  Adds the success tags to the book.
        Args:
            book: The ComicBook to add the tags too
            profile: The profile being used

        """
        if profile.SuccessTags:
            tags = convert_tags_to_list(book.Tags)
            tags.extend(profile.SuccessTags)
            tags.sort()
            book.Tags = ", ".join(tags)

    def _add_custom_value_on_success(self, book, profile):
        for key in profile.SuccessCustomValues:
            book.SetCustomValue(key, profile.SuccessCustomValues[key])


class MoveBookProcessor(BaseBookProcessor):
    """ Processes a book that should be moved or copied to a new path."""

    process_text = "Beginning to move {1}.\nDestination: {2}"

    def _do_operation(self, book, source_file, destination_file, profile):
        """ Does the move operation. Assumes all the paths are valid
        Args:
            book (ComicBook): The book to move
            source_file (FileInfo): The source file FileInfo
            destination_file (FileInfo): The destination FileInfo
            profile (Profile): The profile to use

        Raises:
            MoveFailedException: If the copy or move operations fails.
        """
        _log.Debug("Starting _do_operation")
        original_directory = source_file.Directory
        self._move_book(book, source_file, destination_file.FullName)
        if profile.RemoveEmptyFolder:
            _log.Info("Trying to remove the source folder if empty")
            self._clean_up_empty_folders(original_directory, profile.ExcludedEmptyFolder)

    def _move_book(self, book, source, destination_path):
        """ Moves the source book to the destination file path and updates the ComicBook's FilePath

        Additionally adds the lo_previous_path custom value to the book for use in the undo move script

        Args:
            book (ComicBook): The book to move
            source (FileInfo): The FileInfo of the source book
            destination_path (str): The fully qualified destination path

        Raises:
            MoveFailedException if the move fails
        """
        _log.Info("Starting to move")
        try:
            source.MoveTo(destination_path)
        except (IOException, SecurityException, UnauthorizedAccessException) as e:
            raise MoveFailedException(e.Message)
        else:
            _log.Info("Successfully moved to {0}", destination_path)
            book.SetCustomValue("lo_previous_path", book.FilePath)
            book.FilePath = destination_path


class CopyBookProcessor(BaseBookProcessor):
    """ Processes a book that should be moved or copied to a new path."""

    process_text = "Beginning to copy {1}.\nDestination: {2}"

    def _do_operation(self, book, source_file, destination_file, profile):
        """ Does the move operation. Assumes all the paths are valid
        Args:
            book (ComicBook): The book to move
            source_file (FileInfo): The source file FileInfo
            destination_file (FileInfo): The destination FileInfo
            profile (Profile): The profile to use

        Raises:
            MoveFailedException: If the copy or move operations fails.
        """
        _log.Debug("Starting _do_operation")
        self._copy_book(book,source_file,destination_file, profile.CopyMode)

    def _copy_book(self, book, source, destination_path, copy_mode):
        """ Copies the book to the destination

        Args:
            book (ComicBook):
            source (FileInfo):
            destination_path (str):
            copy_mode (Bool): If true, adds the copied book to the library.

        Raises: MoveFailedException if the copy fails.

        """
        _log.Info("Starting copy")
        try:
            source.CopyTo(destination_path)
        except (IOException, SecurityException, UnauthorizedAccessException) as e:
            raise MoveFailedException(e.Message)
        else:
            _log.Info("Successfully copied the book to {0}", destination_path)
            if copy_mode:
                new_book = self._comic_rack.App.AddNewBook(False)
                new_book.FilePath = destination_path
                copy_data_to_new_book(book, new_book)
                _log.Info("Copied the data from the original book to the copy and added it to the library.")
                new_book.SetCustomValue("lo_copied", "Copied {0} from {1}".format(DateTime.Now, source.FullName))


class FilelessBookProcessor(BaseBookProcessor):
    """ A subclass of BookProcessor designed to handle fileless books."""

    process_text = "Beginning to save custom thumbnail of {0} to destination {1}"

    def process_book(self, book_to_move, profile, ignore_duplicates=False):
        """ Overrides process_book so that the file can be skipped right away if fileless are not being created

        Args:
            book_to_move (BookToMove):
            profile (Profile):
            ignore_duplicates:

        Raises:
            MoveSkippedException
            MoveFailedException
            DuplicateExistsException
        """
        if not profile.MoveFileless:
            raise MoveSkippedException(
                "{0} is a fileless book and images are not being created".format(
                    self._get_book_text(book_to_move.book)))

        return super(FilelessBookProcessor, self).process_book(book_to_move, profile, ignore_duplicates)

    def _do_operation(self, book, source_file, destination_file, profile):
        """ Runs the save fileless image methods
        Args:
            book (ComicBook): The book to move
            source_file (FileInfo): The source file FileInfo
            destination_file (FileInfo): The destination FileInfo
            profile (Profile): The profile to use

        Raises:
            MoveFailedException: If an error occurs creating the image
        """
        self._save_fileless_image(book, destination_file.FullName)

    def _check_source(self, source_path, destination_path):
        """Overrides _check_source since for fileless books there is no source file"""
        return None

    def _save_fileless_image(self, book, path):
        """Creates and saves the fileless image to file.

        Args:
            book (ComicBook): The ComicBook from which to save the image from.
            path (str): The fully qualified path to save the image to.

        Raises:
            MoveFailedException: If the fileless image couldn't be created.
        """
        _log.Trace("Starting _save_fileless_image with {0}", path)
        try:
            self._comic_rack.App.GetComicThumbnail(book, 0).Save(path)
            _log.Info("Created fileless image at: {0}", path)
        except (ArgumentNullException, ExternalException) as e:
            raise MoveFailedException("Failed to create the fileless image {0}: {1}".format(path, e.Message))


class SimulatedBookProcessor(BaseBookProcessor):
    process_text = "Beginning to (simulate) process {0} with destination {1}"

    def __init__(self, comicrack, reporter):
        super(SimulatedBookProcessor, self).__init__(comicrack)
        self._reporter = reporter
        self._created_folders = []
        self.moved_books = []

    def _clean_up_empty_folders(self, directory, excluded_folders):
        """Overrides so that nothing is deleted"""
        pass

    def _move_book(self, book, source, destination_path):
        _log.Info("Moved (simulated) {0} to {1} successfully", self._get_book_text(book), destination_path)
        self._reporter.success_simulated("moved (simulated) to {0}".format(destination_path))
        self.moved_books.append(destination_path)

    def _check_source(self, source_path, destination_path):
        """ Checks the source path for errors.

        Args:
            source_path (str): The source path. If empty the book is fileless and the function return immediately
            destination_path (str): The destination path

        Returns:
            FileInfo if not a fileless book
            None if a fileless book

        Raises:
            MoveSkippedException
            MoveFailedException
        """
        if not source_path:
            return None
        return super(SimulatedBookProcessor, self)._check_source(source_path, destination_path)

    def _check_for_duplicates(self, destination_file, book, check_different_extensions):
        if destination_file.FullName in self.moved_books:
            _log.Info("A duplicate was found at the destination, holding to ask user")
            raise DuplicateExistsException("")
        return super(SimulatedBookProcessor, self)._check_for_duplicates(destination_file, book,
                                                                         check_different_extensions)

    def _do_operation(self, book, source_file, destination_file, profile):
        if source_file is None:
            return self._save_fileless_image(book, source_file.FullName)
        else:
            self._move_book(book, source_file, destination_file.FullName)

    def _create_folder(self, destination_directory):
        """Creates the directory tree (Simulated).

        If the profile mode is simulate it doesn't create any folders, instead
        it saves the folders that would need to be created to _created_folders

        Args:
            new_directory (DirectoryInfo): A DirectoryInfo object of the folder that is needed.
        """
        _log.Trace("Starting create folder: {0}", destination_directory.FullName)
        if destination_directory.Exists:
            _log.Info("Folder already exists")
            return

        if destination_directory.FullName not in self._created_folders:
            self._reporter.log("Created Folder (Simulated)", destination_directory.FullName)
            self._created_folders.append(destination_directory.FullName)
            _log.Info("Created Folder (Simulated) {0}", destination_directory.FullName)
        else:
            _log.Info("Folder {0} has already been created (simulated)", destination_directory.FullName)

    def _save_fileless_image(self, book, path):
        """Simulates saving the fileless image to the file path.

        Args:
            book (ComicBook): The ComicBook from which to save the image from.
            path (str): The fully qualified path to save the image to.
        """
        _log.Trace("Starting _save_fileless_image with {0}", path)
        _log.Info("Simulated creating image {0}", path)
        self._reporter.success_simulated("Created image {0}".format(path))
        self.moved_books.append(path)

    def _add_custom_value_on_success(self, book, profile):
        return

    def _add_tag_on_success(self, book, profile):
        return



