"""
bookmover.py

This contains all the book moving function the script uses.

Version 2.1:


Copyright 2010 Stonepaw

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""
import clr
clr.AddReference("Microsoft.VisualBasic")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("NLog.dll")
from Microsoft.VisualBasic import FileIO
from System import ArgumentException, ArgumentNullException, NotSupportedException, UnauthorizedAccessException, OperationCanceledException  # @UnresolvedImport
from System.Drawing.Imaging import ImageFormat
from System.IO import Path, File, FileInfo, DirectoryInfo, IOException, PathTooLongException, DirectoryNotFoundException
from System.Runtime.InteropServices import ExternalException
from System.Security import SecurityException
from System.Windows.Forms import DialogResult  # @UnresolvedImport

import re

from NLog import LogManager
from duplicatewindow import DuplicateAction, DuplicateWindow
from locommon import Mode, check_metadata_rules, check_excluded_folders, UndoCollection
from loforms import PathTooLongForm
from pathmaker import PathMaker
from movereporter import MoveReporter

ComicRack = None
log = LogManager.GetLogger("BookMover")


class MoveResult(object):
    """A class to report different statuses during a move operation"""
    Success = 1
    Failed = 2
    Skipped = 3
    Duplicate = 4


class BookToMove(object):
    """A wrapper class for a ComicBook object to provide more information
    during a move operation
    """
    def __init__(self, book, path, profile_index, failed_fields):
        self.book = book
        self.path = path
        self.profile_index = profile_index
        self.failed_fields = failed_fields
        self.duplicate_different_extension = False
        self.duplicate_ext_files = []


class BookMoverResult(object):

    def __init__(self, report_text, failed_or_skipped):
        self.report_text = report_text
        self.failed_or_skipped = failed_or_skipped


class ModeText(object):

    @staticmethod
    def get_mode_text(mode):
        if mode == Mode.Move:
            return "move"
        elif mode == Mode.Copy:
            return "copy"
        else:
            return "move (simulated)"

    @staticmethod
    def get_mode_present(mode):
        if mode == Mode.Copy:
            return "copying"

        elif mode == Mode.Move:
            return "moving"

        else:
            return "moving (simulated)"

    @staticmethod
    def get_mode_past(mode):
        if mode == Mode.Copy:
            return "copied"

        elif mode == Mode.Move:
            return "moved"

        else:
            return "moved (simulated)"


class BookMover(object):
    """The class that manages and moves comics based on a profiles

    Use process_books as the entry point to pass books and start moving books
    """
    def __init__(self, worker, form, reporter):
        # Various variables needed to report progress and ask questions
        self._worker = worker
        self.form = form
        self._report = MoveReporter()
        self.profile = None
        self._pathmaker = PathMaker(form, None)

        self.failed_or_skipped = False

        # Hold books that are duplicates so they can be all asked at the end.
        self._held_duplicate_books = []
        self._held_duplicate_count = 0

        # These variables are for when the profile is in simulate mode
        self._created_paths = []
        self._sim_moved_books = []

        # This holds a list of the books moved and is saved in undo.txt for the
        # undo script
        self.undo_collection = UndoCollection()

        self.report_book_name = ""

        self._progress_increment = 1.0
        self._progress = 0.0

        self._duplicate_handler = DuplicateHandler(self._report,
                                                   self._sim_moved_books)

    def _get_canceled_report(self):
        header_text = "\n\n".join([profile_report.get_report(True)
                                   for profile_report in self.profile_reports])
        report = BookMoverResult(header_text, self.failed_or_skipped)
        self._report.add_header(header_text)
        return report

    def _set_book_and_profile(self, book, profile):
        self._report.current_book = book
        self._report.set_profile(profile)
        self.profile = profile

    def _increase_progress(self):
        """Increases the progress percentage and reports to the worker form,
        if any.
        """
        self._progress += self._progress_increment
        try:
            self._worker.ReportProgress(int(round(self._progress)))
        except AttributeError:
            pass

    def process_books(self, books, profiles):
        """The entry point to start a list of books with a list of profiles.

        Args:
            books: An array of ComicBook objects.
            profiles: An array of Profile objects.
        """
        log.Info("Starting to process {0} books", len(books))

        self._report.create_profile_reports(profiles, len(books))
        books_to_move = self._create_book_paths(books, profiles)

        if not books_to_move:
            log.Info("No books need to be moved, exiting.")
            # TODO: Fix this with new report system.
            return self._get_canceled_report()

        self._progress_percent = 1.0 / len(books_to_move) * 100
        count = 0

        def handle_cancelation():
            self._report.log(
                "Cancelled",
                "{0} operations".format(len(books_to_move) - count),
                "User cancelled the script")
            return self._get_canceled_report()

        # Main moving loop
        for book in books_to_move:
            if self._worker.CancellationPending:
                return handle_cancelation()

            self._set_book_and_profile(book.book, profiles[book.profile_index])

            log.Info("\nStarting to process {0} with destination {1}",
                     self._report.current_book, book.path)

            result = self._process_book(book)

            if result == MoveResult.Success:
                self._report.success()

            elif result == MoveResult.Duplicate:
                self._held_duplicate_books.append(book)
                self._held_duplicate_count += 1
                continue

            count += 1
            self._increase_progress()

        # Duplicate loop
        for book in self._held_duplicate_books:
            if self._worker.CancellationPending:
                return handle_cancelation()

            self._set_book_and_profile(book.book, profiles[book.profile_index])

            log.Info("\nStarting to process {0} with a duplicate at {1}",
                     self._report.current_book, book.path)

            result = self._process_duplicate_book(book)
            self._held_duplicate_count -= 1

            if result == MoveResult.Success:
                self._report.success()

            count += 1
            self._increase_progress()

        header_text = "\n\n".join([profile_report.get_report(True)
                                   for profile_report in self.profile_reports])
        report = BookMoverResult(header_text, self.failed_or_skipped)
        self._report.add_header(header_text)
        return report

    def _create_book_paths(self, books, profiles):
        """Finds the destination path for each book given a set of profiles.

        It goes through each profile in order and only the last possible
        profile is used to create a new path.

        Args:
            books: A list of ComicBook objects
            profiles: An ordered list of Profile objects

        Returns:
            A list of BookToMove objects.
        """
        log.Info("\nFinding paths for all the books using {0} profiles",
                 len(profiles))
        books_to_move = []
        books_profiles = {}  # Storing key: book, value: profile to use

        # Find profiles.
        for book in books:
            self._report.current_book = book
            log.Info("\nFinding profile for {0}", self._report.current_book)

            for index, profile in enumerate(profiles):
                self._report.set_profile(profile)

                if not self._book_should_be_moved(book):
                    continue

                if book in books_profiles:
                    self._report.skip("The book is moved by a later profile",
                                      profiles[books_profiles[book]].Name)
                books_profiles[book] = index

            if book in books_profiles:
                log.Info("Moving with profile {0}",
                         profiles[books_profiles[book]].Name)
            else:
                log.Info("Couldn't find a profile. Skipping.")

        log.Info("\nBuilding paths for {0}", len(books_profiles))
        for book in books_profiles:
            profile_index = books_profiles[book]
            self._set_book_and_profile(books, profiles[profile_index])
            log.Info("Finding path for: {0}", self._report.current_book)
            result = self._create_book_path(book)

            if result == MoveResult.Skipped or result == MoveResult.Failed:
                continue
            path = result
            profile_index = index
            failed_fields = self._pathmaker.failed_fields
            books_to_move.append(BookToMove(book, path, profile_index,
                                            failed_fields))

        return books_to_move

    def _set_book_name_for_report(self, book):
        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

    def _create_book_path(self, book):
        """Creates the new path and checks it for some problems.
        Returns the path or a MoveResult.Failed if something goes wrong.
        """
        log.Trace("Beginning create_book_path for {0}",
                  self._report.current_book)

        # Check if a fileless book should have a path made
        if not book.FilePath:
            if not self.profile.MoveFileless:
                log.Info("Skipping because fileless images are not being "
                         "created")
                self._report.skip("Fileless images are not being created")
                return MoveResult.Skipped

            elif self.profile.MoveFileless and not book.CustomThumbnailKey:
                log.Warn("Skipping because the fileless book does not have a "
                         "custom thumbnail")
                self._report.fail("The fileless book does not have a custom "
                                  "thumbnail")
                return MoveResult.Failed

        folder_path, file_name, failed = self._pathmaker.make_path(
            book, self.profile.FolderTemplate, self.profile.FileTemplate)

        full_path = Path.Combine(folder_path, file_name)
        log.Info("Created: {0}", full_path)

        if failed:
            self._report.failed_or_skipped = True
            if not self.profile.MoveFailed:
                self._report.fail(
                    "{0} {1} empty."
                    .format(
                        ",".join(self._pathmaker.failed_fields),
                        "is" if len(self._pathmaker.failed_fields) == 1
                        else "are")
                )
                return MoveResult.Failed

        if not file_name:
            log.Warn("Skipping: The created filename was blank")
            self._report.fail("The created filename was blank")
            return MoveResult.Failed

        return full_path

    def _process_book(self, book_to_move):
        """ Checks a book for various path problems and then calls the move
        method.

        This fixes or catches possible problems with the
        paths including PathTooLong, various security problems and catches
        duplicates.

        If a duplicate is found the book is not moved, instead
        MoveResult.Duplicate is returned as duplicates are held to be dealt
        with at the end.

        Args:
            book_to_move: A BookToMove object with the book to move.

        Returns:
            A MoveResult enum:
                MoveResult.Failed if an unfixable error occurred.
                MoveResult.Skipped if an error could be fixed but the user
                    declined to fix it.
                MoveResult.Duplicate if a duplicate is found.
                MoveResult.Success if the operation succeeded.

        """
        book = book_to_move.book
        new_path = book_to_move.path
        log.Trace("Starting _process_book for {0} with destination {1}",
                  self._report.current_book, new_path)

        # Catch some errors by using FileInfo on the new_path.
        # Do it here rather then while actually moving the book so that the
        # Simulate function works correctly.
        try:
            dest_info = FileInfo(new_path)
        except (PathTooLongException):
            # The Path is too long or something else is wrong
            # ask the user to shorten it.
            log.Info("{0} is too long. Asking the user to shorten it",
                     new_path)
            result = self._get_fixed_path(new_path)
            if result == MoveResult.Skipped:
                log.Info("The user declined to shorten the path. Skipping")
                return result
            log.Info("Got {0} as a fixed path", result)
            book_to_move.path = result
            return self._process_book(book_to_move)

        except (ArgumentNullException, ArgumentException,
                NotSupportedException) as ex:
            # Something wrong with the path that can't be fixed
            log.Error("The new path {0} is invalid and can't be fixed."
                      " The error is: {1}", new_path, ex.Message)
            self._report.fail("The path {0} is not valid".format(new_path))
            return MoveResult.Failed

        except (UnauthorizedAccessException, SecurityException) as ex:
            log.Error("Cannot access the path {0}. The error is: {1}",
                      new_path, ex.Message)
            self._report.fail("Cannot access the path {0}".format(new_path))
            return MoveResult.Failed

        if self._book_already_at_dest(book, Path.GetFileName(new_path),
                                      new_path):
            return MoveResult.Skipped

        # Checking before the move operation even though it would be better
        # caught there as an error because otherwise a duplicate could be found
        # even though the book can't be moved.
        if book.FilePath and not File.Exists(book.FilePath):
            log.Info("Failed to move {0} because the file does not exist",
                     self._report.current_book)
            self._report.fail("The file does not exist")
            return MoveResult.Failed

        # Check if the file is a Duplicate also check MovedBooks for
        # simulate mode
        if dest_info.Exists or new_path in self._sim_moved_books:
            log.Info("A duplicate was found at the destination, holding"
                     " to ask user")
            return MoveResult.Duplicate

        # Test for duplicates with different extensions
        if self.profile.DifferentExtensionsAreDuplicates:
            files = self._get_files_with_different_ext(new_path)
            if files:
                book_to_move.duplicate_ext_files = files
                book_to_move.duplicate_different_extension = True
                log.Info("A duplicate with a different extension was "
                         "found at the destination. Holding to ask the user")
                return MoveResult.Duplicate

        # Create here because needed for cleaning directories later
        old_folder = DirectoryInfo(book.FileDirectory)
        new_folder = dest_info.Directory

        result = self._move_book(book, new_path)

        # Try to remove the old folders if they are empty
        if self.profile.RemoveEmptyFolder and self.profile.Mode == Mode.Move:
            if old_folder.Exists:
                self._delete_empty_folders(old_folder)
            # If the move operation fails we don't want to leave random folders
            self._delete_empty_folders(new_folder)

        if book_to_move.failed_fields and result == MoveResult.Success:
            self._report.fail(
                "{0} {1} empty. {2} to {3}".format(
                    ",".join(self._pathmaker.failed_fields),
                    "is" if len(self._pathmaker.failed_fields) == 1 else "are",
                    ModeText.get_mode_past(self.profile.Mode),# TODO: get rid of this?
                    new_path)
                )
            return MoveResult.Failed

        return result

    def _get_files_with_different_ext(self, full_path):
        """Looks for files in the same directory of the file with the same name
        but different extensions.

        Args:
            full_path: The full path of the file.

        Returns:
            A list or System.Array containing any FileInfo objects that were
            found, possiblys empty.
        """
        log.Trace("Starting to find files with different extensions.")
        d = DirectoryInfo(Path.GetDirectoryName(full_path))
        if d.Exists:
            file_name = Path.GetFileNameWithoutExtension(full_path)
            files = d.GetFiles(file_name + ".*")
            return files
        return []

    def _process_duplicate_book(self, processing_book):
        """ Process a duplicate book by asking the user what to do.

        The user can answer to Overwrite, Cancel or Rename. If overwriting this
        method then deletes the duplicate and moves/copies the book. If Rename
        it moves the book with the calculated rename path. If cancel it returns

        Args:
            processing_book: The BookToMove object with a duplicate to deal
            with.

        Returns:
            A MoveResult property:
                MoveResult.Success if the function succeeds in moving the book.
                MoveResult.Skipped if the user cancels moving this book
                MoveResult.Failed if an error occurs while processing the book.
        """
        log.Trace("Starting _process_duplicate_book with {0}",
                  self._report.current_book)

        result = self._duplicate_handler.handle_duplicate_book(
            processing_book, self.profile, self._held_duplicate_count)

        if result == DuplicateAction.Cancel:
            return MoveResult.Skipped
        elif result == MoveResult.Failed:
            return result

        processing_book.duplicate_different_extension = False
        processing_book.duplicate_ext_files = []
        result = self._process_book(processing_book)
        if result == MoveResult.Duplicate:
            return self._process_duplicate_book(processing_book)
        else:
            return result

    def _move_book(self, book, new_path):
        """Moves or copies the ComicBook to the new path.

        Directory paths are created if required and IO errors are caught and
        handled gracefully.

        Args:
            book: The ComicBook object to move/copy.
            new_path: The path to move the book too.

        Returns:
            MoveResult.Success if the operation succeeds.
            MoveResult.Failed if the operation fails.
        """
        log.Trace("Starting _move_book")
        try:
            book_file_info = FileInfo(book.FilePath)
        except (ArgumentException, ArgumentNullException) as ex:
            # The file name is empty, contains only white spaces, or contains
            # invalid characters.
            if self.profile.MoveFileless:
                return self._save_fileless_image(book, new_path)
            else:
                log.Error("Failed to move {0}: {1}", self.report_book_name,
                          ex.Message)
                self._report.fail(ex.Message)
                return MoveResult.Failed
        except (SecurityException, UnauthorizedAccessException) as ex:
            log.Error("Can't access {0}: {1}", self.report_book_name,
                      ex.Message)
            self._report.fail(ex.Message)
            return MoveResult.Failed

        try:
            if self.profile.Mode == Mode.Move:
                book_file_info.MoveTo(new_path)
                book.SetCustomValue("lo_previous_path", book.FilePath)
                log.Info("Moved {0} to {1} successfully.",
                         self.report_book_name, new_path)
                book.FilePath = new_path

            elif self.profile.Mode == Mode.Simulate:
                log.Info("Moved (simulated) {0} to {1} successfully",
                         self.report_book_name, new_path)
                self._report.success_simulated("moved (simulated) to {0}"
                                               .format(new_path))
                self._sim_moved_books.append(new_path)

            elif self.profile.Mode == Mode.Copy:
                book_file_info.CopyTo(new_path)
                log.Info("Copied {0} to {1} successfully.",
                         self.report_book_name, new_path)
                if self.profile.CopyMode:
                    newbook = ComicRack.App.AddNewBook(False)
                    newbook.FilePath = new_path
                    CopyData(book, newbook)
                    log.Info("Copied the data from the original book to the "
                             " copy and added it to the library.")
            return MoveResult.Success

        except DirectoryNotFoundException:
            if (self._create_folder(FileInfo(new_path).Directory) ==
               MoveResult.Success):
                return self._move_book(book, new_path)
            return MoveResult.Failed

        except (UnauthorizedAccessException, SecurityException,
                IOException) as ex:
            log.Error("Failed {0} : {1}", self.report_book_name, ex.Message)
            self._report.fail(ex.Messsage)
            return MoveResult.Failed

    def _save_fileless_image(self, book, path):
        """Creates and saves the fileless image to file.

        Args:
            book: The ComicBook from which to save the image from.
            path: The fully qualified path to save the image to.

        Returns:
            MoveResult.Success or MoveResult.Failed
        """
        log.Trace("Starting _save_fileless_image with {0}", path)
        if self.profile.Mode == Mode.Simulate:
            log.Info("Simulated creating image {0}", path)
            self._report.success_simulated("Created image {0}".format(path))
            self._sim_moved_books.append(path)
            return MoveResult.Success
        try:
            image = ComicRack.App.GetComicThumbnail(book, 0)
            if self.profile.FilelessFormat == ".jpg":
                img_format = ImageFormat.Jpeg
            elif self.profile.FilelessFormat == ".png":
                img_format = ImageFormat.Png
            elif self.profile.FilelessFormat == ".bmp":
                img_format = ImageFormat.Bmp
            image.Save(path, img_format)
            return MoveResult.Success
        except (ArgumentNullException, ExternalException) as ex:
            log.Error("Unable to save fileless image {0}: {1}", path,
                      ex.Message)
            self._report.fail("Failed to create the fileless image {0}: {1}"
                              .format(path, ex.Message))
            return MoveResult.Failed

    def _book_should_be_moved(self, book):
        """Checks the excluded folders and metadata rules to see if the book
        should be moved.

        Args:
            book: The ComicBook object to check with.
        Returns:
            True if the book should be moved. False otherwise
        """
        if not check_excluded_folders(book.FilePath, self.profile):
            log.Info("Skipped because the book is located in an excluded path.")
            self._report.skip("The book is located in an excluded path")
            return False

        if not check_metadata_rules(book, self.profile):
            log.Info("Skipped because the book was excluded under the exclude "
                     "rules")
            self._report.skip("The book qualified under the exclude rules")
            return False
        return True

    def _book_already_at_dest(self, book, new_file_name, new_full_path):
        """Checks if the book is already at the correct location

        Have to do several checks because some strange things can happen
        if the capitalizations are different.

        Args:
            book: The ComicBook object to move.
            new_file_name: The new file name for the book.
            new_full_path: The new fully qualified path for the book.

        Returns:
            True if the book is already where it should be or False if it is
            not.
        """
        if new_full_path == book.FilePath:
            self._report.skip("The book is already located at the calculated "
                              "path")
            return True

        # In some cases the file path is the same but has different cases.
        # The FileInfo object doesn't catch this but the File.Move function
        # still thinks that it is a duplicate.
        if new_full_path.lower() == book.FilePath.lower():
            # In that case, better rename it to the correct case
            log.Info("Renamed {0} to {1} to fix capitalization"
                     .format(book.FilePath, new_full_path))
            if self.profile.Mode != Mode.Simulate:
                book.RenameFile(new_file_name)
            self._report.log("Renaming",
                             "to {0} to fix capitalization"
                             .format(new_full_path))
            self._report.skip("The book is already located at the calculated "
                              "path")
            return True
        return False

    def _get_shorter_path(self, long_path):
        """Asks the user to fix the path by invoking a dialog.

        Args:
            long_path: The fully qualified path to check as a string.

        Returns:
            MoveResult.Skipped if the user doesn't want to rename the path.
            The shortened, correct path otherwise.
        """
        with PathTooLongForm(long_path) as p:
            if p.ShowDialog() == DialogResult.OK:
                return p.get_path()
        log.Info("Skipped: The calculated path was too long and the user "
                 "skipped shortening it")
        self._report.skip("The calculated path was too long and the user "
                          "skipped shortening it")
        return MoveResult.Skipped

    def _create_folder(self, new_directory):
        """Creates the directory tree.

        If the profile mode is simulate it doesn't create any folders, instead
        it saves the folders that would need to be created to _created_paths

        Args:
            new_directory: A DirectoryInfo object of the folder that is needed.

        Returns:
            MoveResult.Succes if the creation succeeded or was not needed.
            MoveResult.Failed if something went wrong.
        """
        log.Trace("Creating folders {0}", new_directory.FullName)
        if new_directory.Exists:
            return MoveResult.Success

        if self.profile.Mode == Mode.Simulate:
            if new_directory.FullName not in self._created_paths:
                self._report.log("Created Folder (Simulated)",
                                 new_directory.FullName)
                self._created_paths.append(new_directory.FullName)
                log.Info("Created Folder (Simulated) {0}",
                         new_directory.FullName)
            return MoveResult.Success
        try:
            new_directory.Create()
            new_directory.Refresh()
            log.Info("Created folder {0}", new_directory.FullName)
            return MoveResult.Success
        except IOException as ex:
            self._report.fail("Failed to create folder {0}: {1}"
                              .format(new_directory.FullName, ex.Message))
            log.Error("Failed to create folder {0}: {1}",
                      new_directory.FullName, ex.Message)
            return MoveResult.Failed

    def _delete_empty_folders(self, directory):
        """Recursively deletes directories until an non-empty directory is
        found or the directory is in the excluded list.

        Args:
            directory: A DirectoryInfo object which starts at the first folder
                to remove.
        """
        log.Trace("Starting to delete empty folders.")
        directory.Refresh()
        if not directory.Exists:
            return
        # Only delete if no file or folder and not in folder never to delete
        if directory.FullName not in self.profile.ExcludedEmptyFolder:
            try:
                parent = directory.Parent
                directory.Delete()
                log.Info("Deleted empty folder {0}", directory.FullName)
                self._delete_empty_folders(parent)
            except (UnauthorizedAccessException, DirectoryNotFoundException,
                    IOException, SecurityException) as ex:
                log.Warn("Unable to delete {0}: {1}", directory.FullName,
                         ex.Message)


class DuplicateHandler(object):
    """Handles functions for processing a duplicate book."""

    def __init__(self, reporter, moved_books):
        self._reporter = reporter
        self._moved_books = moved_books
        self._log = LogManager.GetLogger("DuplicateHandler")
        self._duplicate_window = None
        self.duplicate_action = None
        self.always_do_duplicate_action = None

    def _create_rename_path(self, path):
        """Creates a Windows formatted duplicate path in the manner of
        Filename (#).cbz. It searches until it finds the first unused number.

        Originally written by pescuma and modified slightly.

        Args:
            path: The fully qualified path to find the duplicate for.

        Returns:
            The fully qualified path that was created.
        """
        self._log.Trace("Trying to find a rename path")
        extension = Path.GetExtension(path)
        base = path[:-len(extension)]
        base = re.sub(" \([0-9]\)$", "", base)
        i = 0
        while (True):
            i += 1
            newpath = "{0} ({1}){2}".format(base, i + 1, extension)
            # For simulated mode we check _moved_books
            if newpath in self._moved_books or File.Exists(newpath):
                continue
            else:
                self._log.Info("Found rename path: {0}", newpath)
                return newpath

    def _delete_duplicate(self, path):
        """Deletes a duplicate file by sending it to the Recycle bin
        and handles the errors.

        Args:
            path: The string file path to try and move to the recycle bin.

        Returns: MoveResult.Success if it succeeds or MoveResult.Failed if it
            fails
        """
        self._log.Trace("Trying to delete {0}", path)
        try:
            FileIO.FileSystem.DeleteFile(
                path, FileIO.UIOption.OnlyErrorDialogs,
                FileIO.RecycleOption.SendToRecycleBin)
            self._log.Info("Deleted duplicate at {0}", path)
            return MoveResult.Success
        except (IOException, OperationCanceledException, SecurityException,
                UnauthorizedAccessException) as ex:
            self._log.Error("Skipped. Could not delete existing duplicate {0}: {1}",
                            path, ex.Message)
            self._reporter.fail("Could not delete the existing duplicate {0} : {1}"
                                .format(path, ex.Message))
            return MoveResult.Failed

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
        self._log.Trace("Trying to find the duplicate book in the library")
        path = path.lower()
        for book in ComicRack.App.GetLibraryBooks():
            # Fix Issue #5: different capitalization in path names.
            if book.FilePath.lower() == path:
                self._log.Info("Found a duplicate in the library")
                return book
        self._log.Info("Could not find a duplicate in the library")
        return None

    def handle_duplicate_book(self, processing_book, profile, count):
        """ Process a duplicate book by asking the user what to do.

        The user can answer to Overwrite, Cancel or Rename. If overwriting this
        method then deletes the duplicate and moves/copies the book. If Rename
        it moves the book with the calculated rename path. If cancel it returns

        Args:
            processing_book: The BookToMove object with a duplicate to deal
                with.
            profile: The profile being used.
            count: The number of held duplicate books to display in the
                duplicate form.

        Returns:
                MoveResult.Failed if an error occurs while processing the book.
                DuplicateAction.Cancel if the user cancels moving the book.
                DuplicateAction.Rename if the user chooses to rename the book.
                DuplicateAction.Overwrite if the user chooses to overwrite the
                    book.
        """
        book = processing_book.book
        move_path = processing_book.path
        dup_path = move_path

        if processing_book.duplicate_different_extension:
            # Note we only use the first one since this gets called recursively
            # if more exist when _process_book is called.
            dup_path = processing_book.duplicate_ext_files[0].FullName

        self._log.Info(
            "Starting duplicate handling for {0} with a duplicate at {1}",
            self._reporter.current_book, dup_path)

        existing_book = self._find_book_in_library(dup_path)

        if existing_book is None:
            existing_book = FileInfo(dup_path)

        if processing_book.duplicate_different_extension:
            rename_path = move_path
        else:
            rename_path = self._create_rename_path(move_path)

        if (not self.always_do_duplicate_action or
           processing_book.duplicate_different_extension):

            if self._duplicate_window is None:
                self._duplicate_window = DuplicateWindow()

            result = self._duplicate_window.ShowDialog(
                processing_book,
                existing_book,
                rename_path,
                profile.Mode,
                count)

            self.duplicate_action = result.action
            self.always_do_duplicate_action = result.always_do_action

        if self.duplicate_action == DuplicateAction.Cancel:
            self._log.Info("User declined to overwrite or rename the "
                           "book and it was _skipped")
            self._reporter.skip(
                "A file already exists at: {0} and the user declined to "
                " or rename it".format(dup_path))
            return DuplicateAction.Cancel

        elif self.duplicate_action == DuplicateAction.Rename:
            self._log.Info("User chose to rename the processing ComicBook")
            processing_book.path = rename_path
            return DuplicateAction.Rename

        elif self.duplicate_action == DuplicateAction.Overwrite:
            self._log.Info("User choose to overwrite the duplicate")

            if profile.Mode == Mode.Simulate:
                self.log.Info("Deleted (simulated) {0}", move_path)
                self._reporter.log("Deleted (simulated)", move_path)
                if book.FilePath:
                    self._log.Info("Moved (simulated) {0} to {1}",
                                   book.FilePath, move_path)
                    self._reporter.success_simulated("to: {0}"
                                                     .format(move_path))
                else:
                    self._log.Info("Created image for {0} at {1} (simulated)",
                                   book.Caption, move_path)
                    self._reporter.log("Created image (simulated)", move_path)
                self._moved_books.append(move_path)

            else:
                if not self._delete_duplicate(dup_path):
                    return MoveResult.Failed

                if book.FilePath and type(existing_book) is not FileInfo:
                    ComicRack.App.RemoveBook(existing_book)
                    self._log.Info(
                        "Removed the duplicate book from the library")
            return DuplicateAction.Overwrite

# class UndoMover(BookMover):
#
#     def __init__(self, worker, form, undo_collection, profiles, report):
#         self.worker = worker
#         self.form = form
#         self.undo_collection = undo_collection
#         self.AlwaysDoAction = False
#         self.HeldDuplicateBooks = {}
#         self.report = report
#         self.profiles = profiles
#         self.failed_or_skipped = False
#
#
#     def process_books(self):
#         books, notfound = self.get_library_books()
#
#         _success = 0
#         _failed = 0
#         _skipped = 0
#         count = 0
#
#         for book in books + notfound:
#             count += 1
#
#             if self._worker.CancellationPending:
#                 _skipped = len(books) + len(notfound) - _success - _failed
#                 self.report.Add("Canceled", str(_skipped) + " files", "User cancelled the script")
#                 report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (_success, _failed, _skipped), _failed > 0 or _skipped > 0)
#                 #self.report.SetCountVariables(_failed, _skipped, _success)
#                 return report
#
#             result = self.process_book(book)
#
#             if result is MoveResult.Duplicate:
#                 count -= 1
#                 continue
#
#             elif result is MoveResult.Skipped:
#                 _skipped += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#             elif result is MoveResult.Failed:
#                 _failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#             elif result is MoveResult.Success:
#                 _success += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#         self.HeldDuplicateCount = len(self._held_duplicate_books)
#         for book in self._held_duplicate_books:
#
#             if self._worker.CancellationPending:
#                 _skipped = len(books) + len(notfound) - _success - _failed
#                 self.report.Add("Canceled", str(_skipped) + " files", "User cancelled the script")
#                 report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (_success, _failed, _skipped), _failed > 0 or _skipped > 0)
#                 #self.report.SetCountVariables(_failed, _skipped, _success)
#                 return report
#
#             result = self.process_duplicate_book(book, self._held_duplicate_books[book])
#             self._held_duplicate_count -= 1
#
#             if result is MoveResult.Skipped:
#                 _skipped += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#             elif result is MoveResult.Failed:
#                 _failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#             elif result is MoveResult.Success:
#                 _success += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#         report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (_success, _failed, _skipped), _failed > 0 or _skipped > 0)
#         #self.report.SetCountVariables(_failed, _skipped, _success)
#         return report
#
#
#     def MoveBooks(self):
#
#         _success = 0
#         _failed = 0
#         _skipped = 0
#         count = 0
#         #get a list of the books
#         books, notfound = self.get_library_books()
#
#
#
#         for book in books + notfound:
#             if type(book) == str:
#                 path = self.undo_collection[book]
#                 oldfile = book
#             else:
#                 path = self.undo_collection[book.FilePath]
#                 oldfile = book.FilePath
#             count += 1
#
#             if self._worker.CancellationPending:
#                 #User pressed cancel
#                 _skipped = len(books) + len(notfound) - _success - _failed
#                 self.report.Append("\n\nOperation cancelled by user.")
#                 break
#
#             if not File.Exists(oldfile):
#                 self.report.Append("\n\nFailed to move\n%s\nbecause the file does not exist." % (oldfile))
#                 _failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#
#             if path == oldfile:
#                 self.report.Append("\n\nSkipped moving book\n%s\nbecause it is already located at the calculated path." % (oldfile))
#                 _skipped += 1
#                 self._worker.ReportProgress(count)
#                 continue
#
#             #Created the directory if need be
#             f = FileInfo(path)
#
#             if f.Exists:
#                 self._held_duplicate_books.append(book)
#                 count -= 1
#                 continue
#
#             d = f.Directory
#             if not d.Exists:
#                 d.Create()
#
#             if type(book) == str:
#                 oldpath = FileInfo(book).DirectoryName
#             else:
#                 oldpath = book.FileDirectory
#
#             result = self.MoveBook(book, path)
#             if result == MoveResult.Success:
#                 _success += 1
#
#             elif result == MoveResult.Failed:
#                 _failed += 1
#
#             elif result == MoveResult.Skipped:
#                 _skipped += 1
#
#
#             #If cleaning directories
#             if self.settings.RemoveEmptyFolder:
#                 self.CleanDirectories(DirectoryInfo(oldpath))
#                 self.CleanDirectories(DirectoryInfo(f.DirectoryName))
#             self._worker.ReportProgress(count)
#
#         #Deal with the duplicates
#         for book in self._held_duplicate_books[:]:
#             if type(book) == str:
#                 path = self.undo_collection[book]
#                 oldpath = book
#             else:
#                 path = self.undo_collection[book.FilePath]
#                 oldpath = book.FilePath
#             count += 1
#
#             if self._worker.CancellationPending:
#                 #User pressed cancel
#                 _skipped = len(books) + len(notfound) - _success - _failed
#                 self.report.Append("\n\nOperation cancelled by user.")
#                 break
#
#             #Created the directory if need be
#             f = FileInfo(path)
#
#             d = f.Directory
#             if not d.Exists:
#                 d.Create()
#
#             if type(book) == str:
#                 oldpath = FileInfo(book).DirectoryName
#             else:
#                 oldpath = book.FileDirectory
#
#             result = self.MoveBook(book, path)
#             if result == MoveResult.Success:
#                 _success += 1
#
#             elif result == MoveResult.Failed:
#                 _failed += 1
#
#             elif result == MoveResult.Skipped:
#                 _skipped += 1
#
#
#             #If cleaning directories
#             if self.settings.RemoveEmptyFolder:
#                 self.CleanDirectories(DirectoryInfo(oldpath))
#                 self.CleanDirectories(DirectoryInfo(f.DirectoryName))
#
#             self._held_duplicate_books.remove(book)
#
#             self._worker.ReportProgress(count)
#
#         #Return the report to the worker thread
#         report = "Successfully moved: %s\nFailed to move: %s\nSkipped: %s" % (_success, _failed, _skipped)
#         return [_failed + _skipped, report, self.report.ToString()]
#
#
#     def process_book(self, book):
#
#         if type(book) == str:
#             undo_path = self.undo_collection.undo_path(book)
#             current_path = book
#             self.report_book_name = book
#         else:
#             undo_path = self.undo_collection.undo_path(book.FilePath)
#             current_path = book.FilePath
#             self.report_book_name = book.FilePath
#
#         self.profile = self.profiles[self.undo_collection.profile(current_path)]
#
#         if not File.Exists(current_path):
#             self.report.Add("Failed", self.report_book_name, "The file does not exist")
#             return MoveResult.Failed
#
#
#         result = self.check_path_problems(current_path, undo_path)
#         if result is not None:
#             return result
#
#         #Don't need to check path to long
#
#         #Duplicate
#         if File.Exists(undo_path):
#             self._held_duplicate_books[book] = undo_path
#             return MoveResult.Duplicate
#
#         #Create here because needed for cleaning directories later
#         old_folder_path = FileInfo(current_path).DirectoryName
#
#         result = self._create_folder(old_folder_path, book)
#         if result is not MoveResult.Success:
#             return result
#
#         result = self._move_book(book, undo_path)
#
#         if self.profile.RemoveEmptyFolder:
#             self._remove_empty_folders(DirectoryInfo(old_folder_path))
#             self._remove_empty_folders(FileInfo(undo_path).Directory)
#
#         return result
#
#
#     def process_duplicate_book(self, book, undo_path):
#
#         if type(book) == str:
#             current_path = book
#             self.report_book_name = book
#         else:
#             current_path = book.FilePath
#             self.report_book_name = book.FilePath
#
#         self.profile = self.profiles[self.undo_collection.profile(current_path)]
#
#         #Since the duplicate is checked for last in the orginal process_book
#         #function there is no need to check for path errors.
#         if File.Exists(undo_path):
#
#             #Find the existing book if it occurs in the library
#             oldbook = self._find_duplicate_book(undo_path)
#             if oldbook == None:
#                 oldbook = FileInfo(undo_path)
#
#             rename_path = self._create_rename_path(undo_path)
#
#             rename_filename = Path.GetFileName(rename_path)
#
#             if not self.always_do_duplicate_action:
#                 result = self.form.Invoke(Func[type(self.profile), type(book), type(oldbook), str, int, DuplicateResult](self.form.ShowDuplicateForm), System.Array[object]([self.profile, book, oldbook, rename_filename, self._held_duplicate_count]))
#
#                 self.duplicate_action = result.action
#
#                 if result.always_do_action:
#                     self.always_do_duplicate_action = True
#
#             if self.duplicate_action is DuplicateAction.Cancel:
#                 self.report.Add("Skipped", self.report_book_name, "A file already exists at: " + undo_path + " and the user declined to overwrite it")
#                 return MoveResult.Skipped
#
#             elif self.duplicate_action is DuplicateAction.Rename:
#                 #Check if the created path is too long
#                 if len(rename_path) > 259:
#                     result = self.form.Invoke(Func[str, object](self._get_smaller_path), System.Array[System.Object]([rename_path]))
#                     if result is None:
#                         self.report.Add("Skipped", self.report_book_name, "The path was too long and the user _skipped shortening it")
#                         return MoveResult.Skipped
#
#                     return self.process_duplicate_book(book, result)
#
#                 return self.process_duplicate_book(book, rename_path)
#
#             elif self.duplicate_action is DuplicateAction.Overwrite:
#                 try:
#                     if self.profile.CopyReadPercentage and type(oldbook) is not FileInfo:
#                         book.LastPageRead = oldbook.LastPageRead
#                     FileIO.FileSystem.DeleteFile(undo_path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
#
#                 except Exception, ex:
#                     self.report.Add("Failed", self.report_book_name, "Failed to overwrite " + undo_path + ". The error was: " + str(ex))
#                     return MoveResult.Failed
#
#                 #Since we are only working with images there is no need to
#                 #remove a book from the library
#                 if type(oldbook) is not FileInfo:
#                         ComicRack.App.RemoveBook(oldbook)
#
#                 return self.process_duplicate_book(book, undo_path)
#
#         old_folder_path = FileInfo(current_path).DirectoryName
#
#         result = self._move_book(book, undo_path)
#
#         if self.profile.RemoveEmptyFolder:
#             self._remove_empty_folders(DirectoryInfo(old_folder_path))
#             self._remove_empty_folders(FileInfo(undo_path).Directory)
#
#         return result
#
#
#     def get_library_books(self):
#         """Finds which books are in the ComicRack Library.
#         Returns a list of books which are in the library and a list of file paths that are not in the library.
#         """
#         books = []
#         found = set()
#         #Note possible that the book in not in the library.
#         for b in ComicRack.App.GetLibraryBooks():
#             if b.FilePath in self.undo_collection._current_paths:
#                 books.append(b)
#                 found.add(b.FilePath)
#         all = set(self.undo_collection._current_paths)
#         notfound = all.difference(found)
#         return books, list(notfound)
#
#
#     def check_path_problems(self, current_path, undo_path):
#
#         if undo_path == current_path:
#             self.report.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
#             return MoveResult.Skipped
#
#         #In some cases the filepath is the same but has different cases.  The
#         #FileInfo object dosn't catch this but the File.Move function
#         #Thinks that it is a duplicate.
#         if undo_path.lower() == current_path.lower():
#             #In that case, better rename it to the correct case
#             book.RenameFile(file_name)
#             self.report.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
#             return MoveResult.Skipped
#
#         return None
#
#
#     def _move_book(self, book, path):
#         #Finally actually move the book
#         try:
#             if type(book) is str:
#                 File.Move(book, path)
#             else:
#                 File.Move(book.FilePath, path)
#                 book.FilePath = path
#
#             return MoveResult.Success
#
#         except Exception, ex:
#             self.report.Add("Failed", self.report_book_name, "because an error occured. The error was: " + str(ex))
#             return MoveResult.Failed
#
#
#     def MoveBook(self, book, path):
#         """
#         Book is the book to be moved
#
#         Path is the absolute destination with extension
#
#         """
#
#         #Keep the current book path, just incase
#         if type(book) == str:
#             oldpath = book
#         else:
#             oldpath = book.FilePath
#
#
#         if File.Exists(path):
#
#             #Find the existing book if it occurs in the library
#             oldbook = self.FindDuplicate(path)
#
#             if oldbook == None:
#                 oldbook = FileInfo(path)
#
#             if type(book) == str:
#                 dupbook = FileInfo(book)
#             else:
#                 dupbook = book
#
#             if not self.AlwaysDoAction:
#
#                 renamepath = self.CreateRenamePath(path)
#
#                 renamefilename = FileInfo(renamepath).Name
#
#                 #Ask the user:
#                 result = self.form.Invoke(Func[type(dupbook), type(oldbook), str, int, list](self.form.ShowDuplicateForm), System.Array[object]([dupbook, oldbook, renamefilename, len(self._held_duplicate_books)]))
#                 self.Action = result[0]
#                 if result[1] == True:
#                     #User checked always do this opperation
#                     self.AlwaysDoAction = True
#
#             if self.Action == DuplicateResult.Cancel:
#                 self.report.Append("\n\nSkipped moving\n%s\nbecause a file already exists at\n%s\nand the user declined to overwrite it." % (oldpath, path))
#                 return MoveResult.Skipped
#
#             elif self.Action == DuplicateResult.Rename:
#                 return self.MoveBook(book, renamepath)
#
#             elif self.Action == DuplicateResult.Overwrite:
#
#                 try:
#                     FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
#                     #File.Delete(path)
#                 except Exception, ex:
#                     self.report.Append("\n\nFailed to delete\n%s. The error was: %s. File\n%s\nwas not moved." % (path, ex, oldpath))
#                     return MoveResult.Failed
#
#                 #if the book is in the library, delete it
#                 if oldbook:
#                     ComicRack.App.RemoveBook(oldbook)
#
#                 return self.MoveBook(book, path)
#
#
#         #Finally actually move the book
#         try:
#             File.Move(oldpath, path)
#             if not type(book) == str:
#                 book.FilePath = path
#             return MoveResult.Success
#         except PathTooLongException:
#             #Too long path.  Add a way to ask what path you want to use instead
#             print "path was to long"
#             print oldpath
#             print path
#             return MoveResult.Failed
#         except Exception, ex:
#             self.report.Append("\n\nFailed to move\n%s\nbecause an error occured. The error was: %s. The book was not moved." % (book.FilePath, ex))
#             return MoveResult.Failed


def CopyData(book, newbook):
    """This helper function copies all relevant metadata from a book to another book.

    Args:
        book: The ComicBook to copy from.
        newbook: The ComicBook to copy to.
    """
    newbook.SetInfo(book, False, False)  # This copies most fields
    newbook.CustomValuesStore = book.CustomValuesStore
    newbook.SeriesComplete = book.SeriesComplete  # Not copied by SetInfo
    newbook.Rating = book.Rating  # Not copied by SetInfo
    newbook.customThumbnailKey = book.CustomThumbnailKey
