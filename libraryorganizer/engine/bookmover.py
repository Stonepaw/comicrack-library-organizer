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

import re
import sys
from Microsoft.VisualBasic import FileIO
from System import Func, ArgumentException, ArgumentNullException, NotSupportedException, UnauthorizedAccessException, OperationCanceledException
import System
from System.Drawing.Imaging import ImageFormat
from System.IO import Path, File, FileInfo, DirectoryInfo, IOException, PathTooLongException, DirectoryNotFoundException
from System.Security import SecurityException
from System.Windows.Forms import DialogResult

from NLog import LogManager

from duplicatewindow import DuplicateAction, DuplicateWindow
from locommon import Mode, check_metadata_rules, check_excluded_folders, UNDOFILE, UndoCollection
from loforms import PathTooLongForm
from pathmaker import PathMaker



nlog_logger = LogManager.GetLogger("BookMover")
sim_logger = LogManager.GetLogger("Simulate")
DUPLICATE_EXT_CANCEL = "A file with a different extension already exists at: {0} and the user declined to overwrite it"
DUPLICATE_EXT_CANCEL_FILELESS = "The image already exists with a different extension at: {0} and the user declined to overwrite it"

class MoveResult(object):
    """A class to report different statuses during a move operation"""
    Success = 1
    Failed = 2
    Skipped = 3
    Duplicate = 4
    FailedMove = 5
    DuplicateDifferentExtension = 6


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


class ProfileReport(object):
    """Provides way to report on each profile's results for the move operation
    """
    def __init__(self, total, name, mode):
        self.success = 0
        self.failed = 0
        self.skipped = 0
        self._total = total
        self._name = name
        self._mode = mode

    def get_report(self, cancelled):
        """Creates the report to display to the user of this profile's move
        operation results

        Args:
            cancelled: A Boolean if the move process was cancelled, a different
                report is created if true.

        Returns:
            A string of the total success, failed, and skipped operations
        """
        if cancelled:
            self.skipped = self._total - self.success - self.failed
        return "%s:\nSuccessfully %s: %s\tSkipped: %s\tFailed: %s" % (self._name, ModeText.get_mode_past(self._mode), self.success,
                                                                      self.skipped, self.failed)


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
    def __init__(self, worker, form, logger):
        #Various variables needed to report progress and ask questions
        self._worker = worker
        self.form = form
        self.logger = logger

        self.pathmaker = PathMaker(form, None)

        self._duplicate_window = DuplicateWindow()
        self.failed_or_skipped = False
        
        #Hold books that are duplicates so they can be all asked at the end.
        self._held_duplicate_books = []
        self._held_duplicate_count = 0
        
        #These variables are for when the profile is in test mode
        self._created_paths = []
        self._moved_books = []

        #This holds a list of the book moved and is saved in undo.txt for the
        #undo script
        self.undo_collection = UndoCollection()

        self.report_book_name = ""
        
        #Duplicate options
        self.always_do_duplicate_action = False
        self.duplicate_action = None

    def _get_canceled_report(self):
        header_text = "\n\n".join([profile_report.get_report(True) 
                                   for profile_report in self.profile_reports])
        report = BookMoverResult(header_text, self.failed_or_skipped)
        self.logger.add_header(header_text)
        return report

    def process_books(self, books, profiles):
        """The entry point to start moving books.

        Args:
            books: An array of ComicBook objects.
            profiles: An array of Profile objects.
        """
        print "Starting to process %s books" % len(books)

        books_to_move = self._create_book_paths(books, profiles)

        if not books_to_move:
            nlog_logger.Info("No books need to be moved, exiting.")
            return self._get_canceled_report()

        progress_percent = 1.0 / len(books_to_move) * 100
        progress = 0.0
        count = 0

        def handle_result(result):
            if result is MoveResult.Skipped:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].skipped += 1
                
            elif result is MoveResult.Failed:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].failed += 1

            elif result is MoveResult.Success:
                self.profile_reports[book.profile_index].success += 1

            self._worker.ReportProgress(int(round(progress)))

        for book in books_to_move:

            #Handle the cancellation
            if self._worker.CancellationPending:
                self.logger.Add("Cancelled", 
                    "{0} operations".format(len(books_to_move) - count), 
                    "User cancelled the script")
                return self._get_canceled_report()

            count += 1
            progress += progress_percent

            self.profile = profiles[book.profile_index]
            self.logger.SetProfile(self.profile.Name)

            result = self._process_book(book)

            if result == MoveResult.Duplicate:
                count -= 1
                progress -= progress_percent
                self._held_duplicate_books.append(book)
                continue
            else:
                handle_result(result)

        self._held_duplicate_count = len(self._held_duplicate_books)
        for book in self._held_duplicate_books:

            if self._worker.CancellationPending:
                self.logger.Add("Cancelled", str(len(books_to_move) - count) + " operations", "User cancelled the script")
                return self._get_canceled_report()
            
            count += 1
            progress += progress_percent

            self.profile = profiles[book.profile_index]
            self.logger.SetProfile(self.profile.Name)
            self._set_book_name_for_report(book.book)
            result = self._process_duplicate_book(book)
            self._held_duplicate_count -= 1

            handle_result(result)

        if len(self.undo_collection) > 0:
            self.undo_collection.save(UNDOFILE)

        header_text = "\n\n".join([profile_report.get_report(True) for profile_report in self.profile_reports])
        report = BookMoverResult(header_text, self.failed_or_skipped)
        self.logger.add_header(header_text)
        return report

    def _create_book_paths(self, books, profiles):
        """Finds the destination path for each book given a set of profiles.

        It goes through each profile in order and only the last possible 
        path is used. However, if several profiles are in copy mode, the book 
        will be copied several times.
        
        Args:
            books: A list of ComicBook objects
            profiles: An ordered list of Profile objects

        Returns:
            A list of BookToMove objects.
        """
        print "Finding paths for all the books using %s profiles" % len(profiles)
        books_to_move = []
        self.profile_reports = [ProfileReport(len(books), 
                                profile.Name, profile.Mode) 
                                for profile in profiles]

        for book in books:
            path = ""
            profile_index = None
            failed_fields = []
            self._set_book_name_for_report(book)

            for index, profile in enumerate(profiles):
                self.profile = profile
                self.pathmaker.profile = profile
                self.logger.SetProfile(profile.Name)

                result = self.create_book_path(book)

                if result == MoveResult.Skipped:
                    self.profile_reports[index].skipped += 1
                    self.failed_or_skipped = True
                    continue

                elif result == MoveResult.Failed:
                    self.profile_reports[index].failed += 1
                    self.failed_or_skipped = True
                    continue

                # If copying then do it for every possible profile.  TODO: why?
                #if profile.Mode == Mode.Copy:
                #    books_to_move.append(BookToMove(
                #        book, result, index, self.pathmaker.failed_fields))
                #    continue
                else:
                    if path:
                        # Only use the most recent path so mark the old one as
                        # Skipped
                        self.profile_reports[profile_index].skipped +=1
                        self.logger.Add("Skipped", self.report_book_name, 
                            "The book is moved by a later profile", 
                            profiles[profile_index].Name)
                        self.failed_or_skipped = True
                    path = result
                    profile_index = index
                    failed_fields = self.pathmaker.failed_fields

            if path:
                self.profile = profiles[profile_index]
                self.logger.SetProfile(self.profile.Name)
                # Because the path can already be at location the final profile
                # says the book may be moved with the wrong profile if this is
                # checked earlier.
                result = self._check_bookpath_same_as_newpath(book, Path.GetFileName(path), path)
                if not result:
                    self.profile_reports[profile_index].skipped += 1
                    self.failed_or_skipped = True
                else:
                    books_to_move.append(BookToMove(book, path, profile_index, failed_fields))
        print "Created paths are:"
        for i in books_to_move:
            print "%s:%s" % (i.book.FilePath, i.path)

        return books_to_move

    def _set_book_name_for_report(self, book):
        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

    def create_book_path(self, book):
        """Creates the new path and checks it for some problems.
        Returns the path or a MoveResult if something goes wrong.
        """
        #TODO: Put this somewhere else
        #Check if the book isn't fileless and doesn't exist
        if book.FilePath and not File.Exists(book.FilePath):
            self.logger.Add("Failed", self.report_book_name, 
                            "The file does not exist")
            return MoveResult.Failed

        if not self._book_should_be_moved_with_rules(book):
            return MoveResult.Skipped

        #Check if a fileless book should have a path made
        if not book.FilePath:
            if not self.profile.MoveFileless:
                self.logger.Add("Skipped", self.report_book_name, "The book is fileless and fileless images are not being created")
                return MoveResult.Skipped
            
            elif self.profile.MoveFileless and not book.CustomThumbnailKey:
                self.logger.Add("Failed", self.report_book_name, "The fileless book does not have a custom thumbnail")
                return MoveResult.Failed

        
        folder_path, file_name, failed = self.pathmaker.make_path(book, self.profile.FolderTemplate, self.profile.FileTemplate)

        full_path = Path.Combine(folder_path, file_name)

        if failed:
            self.failed_or_skipped = True
            failed_report_verb = " are"
            if len(self.pathmaker.failed_fields) == 1:
                failed_report_verb = " is"
            if not self.profile.MoveFailed:
                self.logger.Add("Failed", self.report_book_name, ",".join(self.pathmaker.failed_fields) + failed_report_verb + " empty.")
                return MoveResult.Failed


        if not file_name:
            self.logger("Failed", self.report_book_name, 
                        "The created filename was blank")
            return MoveResult.Failed

        return full_path

    def _process_book(self, book_to_move):
        """ Checks a book for various path problems and then calls the move method.
        
        This fixes or catches several possible problems with the
        paths including PathTooLong, various security problems and catches
        duplicates. 

        If a duplicate is found the book is not moved, instead MoveResult.Duplicate
        is returned as duplicates are held to be dealt with at the end.

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
        self._set_book_name_for_report(book)
        new_path = book_to_move.path

        # Catch some errors by using FileInfo on the new_path.
        # Do it here rather then while actually moving the book so that the
        # Simulate function works correctly.
        try:
            dest_info = FileInfo(new_path)
        except (PathTooLongException, NotSupportedException):
            # The Path is too long or something else is wrong
            # ask the user to shorten it.
            nlog_logger.Info("Path of {0} is too long. Asking the user to shorten"
                             " it...", new_path)
            result = self._get_fixed_path(new_path)
            if result is MoveResult.Skipped:
                nlog_logger.Info("The user declined to shorten the path. Skipping")
                return result
            nlog_logger.Info("Got {0} as a fixed path", result)
            book_to_move.path = result
            return self._process_book(book_to_move)

        except (ArgumentNullException, ArgumentException) as ex:
            # Something wrong with the path that can't be fixed
            nlog_logger.Error("The new path {0} is invalid and can't be fixed."
                              " The error is: {1}", new_path, ex.Message)
            self.logger.Add("Failed", self.report_book_name, 
                            "The path %s is not valid" % new_path)
            return MoveResult.Failed

        except (UnauthorizedAccessException, SecurityException) as ex:
            nlog_logger.Error("Cannot access the path {0}. The error is: {1}", new_path, ex.Message)
            self.logger.Add("Failed", self.report_book_name, 
                            "Cannot access the path %s" % new_path)
            return MoveResult.Failed

        # Check if the file is a Duplicate also check MovedBooks for
        # simulate mode
        if dest_info.Exists or new_path in self._moved_books:
            return MoveResult.Duplicate
        
        # Test for duplicates with different extensions
        if self.profile.DifferentExtensionsAreDuplicates:
            files = self._get_files_with_different_ext(new_path)
            if files:
                book_to_move.duplicate_ext_files = files
                book_to_move.duplicate_different_extension = True
                return MoveResult.Duplicate

        # Create here because needed for cleaning directories later
        old_folder = DirectoryInfo(book.FileDirectory)
        new_folder = dest_info.Directory

        # Create the new folder paths
        result = self._create_folder(new_folder)
        if result is not MoveResult.Success:
            return result
        
        result = self.move_book(book, new_path)

        # Try to remove the old folders if they are empty
        if self.profile.RemoveEmptyFolder and self.profile.Mode == Mode.Move:
            if old_folder.Exists:
                self._delete_empty_folders(old_folder)
            # If the move operation fails we don't want to leave random folders
            self._delete_empty_folders(new_folder)

        if book_to_move.failed_fields and result == MoveResult.Success:
            if len(book_to_move.failed_fields) > 1:
                failed_report_verb = " are"
            else:
                failed_report_verb = " is"
            #TODO: Fix this gibberish
            self.logger.Add("Failed", self.report_book_name, ",".join(book_to_move.failed_fields) + failed_report_verb + " empty. " + ModeText.get_mode_past(self.profile.Mode) + " to " + new_path)
            return MoveResult.Failed

        return result

    def _get_files_with_different_ext(self, full_path):
        """Looks for files in the same directory of the file with the same name
        but different extensions.

        Args:
            full_path: The full path of the file.

        Returns:
            An Array containing any FileInfo objects that were found.
        """
        d = DirectoryInfo(Path.GetDirectoryName(full_path))
        if d.Exists:
            file_name = Path.GetFileNameWithoutExtension(full_path)                
            files = d.GetFiles(file_name + ".*")
            return files
        return []

#     def _old_process_duplicate_book(self, book_to_move):
#         
#         book = book_to_move.book
#         full_path = book_to_move.path
#         dde = book_to_move.duplicate_different_extension
#         self._set_book_name_for_report(book)
# 
#         #Since the duplicate is checked for in the original process_book
#         #function there is no need to check for path errors.
#         if File.Exists(full_path) or full_path in self._moved_books or dde:
# 
#             # Find the existing book if it occurs in the library
#             # TODO: dde is unnessarary
#             dupbook = self._find_duplicate_book(full_path, dde)
# 
#             # Book does not exist in the library
#             if dupbook is None:
#                 # If we are checking for a different extension then we have to
#                 # find what that path is
#                 if dde:
#                     # TODO: Only do this once when the first one is found
#                     # TODO: This is unnessary since they are saved into the book_to_move
#                     files = self._get_files_with_different_ext(full_path)
#                     if files:
#                         dupbook = files[0]
#                     else:
#                         return MoveResult.Failed
#                 #If not a book with a different extension then just use the
#                 #existing file.
#                 else:
#                     dupbook = FileInfo(full_path)
#                 dup_path = dupbook.FullPath
#             else:
#                 dup_path = dupbook.FilePath
#                 
#             if book_to_move.duplicate_different_extension:
#                 rename_path = full_path
#                 rename_filename = Path.GetFileName(rename_path)
#             else:
#                 rename_path = self._create_rename_path(full_path)
#                 rename_filename = Path.GetFileName(rename_path)
#         
#             if not self.always_do_duplicate_action or book_to_move.duplicate_different_extension:
# 
#                 #result = self.form.Invoke(Func[type(self.profile), type(book),
#                 #type(dupbook), str, int,
#                 #DuplicateResult](self.form.ShowDuplicateForm),
#                 #System.Array[object]([self.profile, book, dupbook,
#                 #rename_filename, self._held_duplicate_count]))
# 
#                 self.duplicate_action = result.action
# 
#                 if result.always_do_action:
#                     self.always_do_duplicate_action = True
# 
#             if self.duplicate_action is DuplicateAction.Cancel:
#                 if book.FilePath:
#                     self.logger.Add("Skipped", self.report_book_name, "A file already exists at: " + full_path + " and the user declined to overwrite it")
#                 else:
#                     self.logger.Add("Skipped", self.report_book_name, "The image already exists at: " + full_path + " and the user declined to overwrite it")
#                 return MoveResult.Skipped
# 
#             elif self.duplicate_action is DuplicateAction.Rename:
#                 #Check if the created path is too long
#                 book_to_move.duplicate_different_extension = False
#                 #TODO: Bad!!!!
#                 if len(rename_path) > 259:
#                     result = self.form.Invoke(Func[str, object](self._get_smaller_path), System.Array[System.Object]([rename_path]))
#                     if result is None:
#                         self.logger.Add("Skipped", self.report_book_name, "The path was too long and the user skipped shortening it")
#                         return MoveResult.Skipped
# 
#                     return self._process_duplicate_book(BookToMove(book, result, book_to_move.profile_index, book_to_move.failed_fields))
#                 return self._process_duplicate_book(BookToMove(book, rename_path, book_to_move.profile_index, book_to_move.failed_fields))
# 
#             elif self.duplicate_action is DuplicateAction.Overwrite:
#                 book_to_move.duplicate_different_extension = False
#                 try:
#                     if self.profile.Mode == Mode.Simulate:
#                         #Because the script goes into a loop if in test mode
#                         #here since no files are actually changed.  return a
#                         #success
#                         self.logger.Add("Deleted (simulated)", full_path)
#                         if book.FilePath:
#                             self.logger.Add(ModeText.get_mode_past(self.profile.Mode), book.FilePath, "to: " + full_path)
#                         else:
#                             self.logger.Add("Created image", full_path)
# 
#                         self._moved_books.append(full_path)
#                         return MoveResult.Success
#                     else:
#                         if self.profile.CopyReadPercentage and type(dupbook) is not FileInfo:
#                             book.LastPageRead = dupbook.LastPageRead
#                         if type(dupbook) is FileInfo:
#                             FileIO.FileSystem.DeleteFile(dupbook.FullPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
#                         else:
#                             FileIO.FileSystem.DeleteFile(dupbook.FilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
# 
#                 except Exception, ex:
#                     self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + full_path + ". The error was: " + str(ex))
#                     return MoveResult.Failed
# 
#                 #Since we are only working with images there is no need to
#                 #remove a book from the library
#                 if book.FilePath and type(dupbook) is not FileInfo:
#                         ComicRack.App.RemoveBook(dupbook)
# 
#                 return self._process_duplicate_book(book_to_move)
# 
#         old_folder_path = book.FileDirectory
#         
#         result = self.move_book(book, full_path)
# 
# 
#         if self.profile.RemoveEmptyFolder and self.profile.Mode == Mode.Move:
#             if old_folder_path:
#                 self._remove_empty_folders(DirectoryInfo(old_folder_path))
#             self._remove_empty_folders(FileInfo(full_path).Directory)
# 
#         if book_to_move.failed_fields and result == MoveResult.Success:
#             if len(book_to_move.failed_fields) > 1:
#                 failed_report_verb = " are"
#             else:
#                 failed_report_verb = " is"
#             self.logger.Add("Failed", self.report_book_name, ",".join(book_to_move.failed_fields) + failed_report_verb + " empty. " + ModeText.get_mode_present(self.profile.Mode) + " to " + full_path)
#             return MoveResult.Failed
# 
# 
#         return result

    def _process_duplicate_book(self, processing_book):
        """ Process a duplicate book by asking the user what to do.

        The user can answer to Overwrite, Cancel or Rename. If overwriting this
        method then deletes the duplicate and moves/copies the book. If Rename
        it moves the book with the calculated rename path. If cancel it returns.

        Args:
            processing_book: The BookToMove object with a duplicate to deal with.

        Returns:
            A MoveResult property:
                MoveResult.Success if the function succeddes in moving the book.
                MoveResult.Skipped if the user cancels moving this book
                MoveResult.Failed if an error occurs while processing the book.
        """
        book = processing_book.book
        move_path = processing_book.path
        dup_path = move_path
        self._set_book_name_for_report(book)
        
        if processing_book.duplicate_different_extension:
            #Note we only use the first one since this gets called recursively if
            #more exist when _process_book is called.
            dup_path = processing_book.duplicate_ext_files[0]

        nlog_logger.Info("Starting to process {0} with a duplicate at {1}",
                         self.report_book_name, dup_path)

        existing_book = self._find_duplicate_book(dup_path)

        if existing_book is None:
            existing_book = FileInfo(dup_path)

        if processing_book.duplicate_different_extension:
            rename_path = move_path
        else:
            rename_path = self._create_rename_path(move_path)

        if (not self.always_do_duplicate_action or 
            processing_book.duplicate_different_extension):
            result = self._duplicate_window.ShowDialog(
                processing_book, existing_book, rename_path, self.profile.Mode, 
                self._held_duplicate_count)
            self.duplicate_action = result.action
            self.always_do_duplicate_action = result.always_do_action

        if self.duplicate_action == DuplicateAction.Cancel:
            nlog_logger.Info("User declined to overwrite or rename the "
                             "book and it was skipped")
            if book.FilePath:
                self.logger.Add(
                    "Skipped", self.report_book_name, "A file already exists "
                    "at: {0} and the user declined to overwrite it or rename "
                    "the book".format(dup_path))
            else:
                self.logger.Add(
                    "Skipped", self.report_book_name, 
                    "The image already exists at: {0} and the user declined to"
                    " overwrite it or rename the image".format(dup_path))
            return MoveResult.Skipped

        elif self.duplicate_action == DuplicateAction.Rename:
            nlog_logger.Info("User choose to rename the processing ComicBook")
            processing_book.path = rename_path

        elif self.duplicate_action is DuplicateAction.Overwrite:
            nlog_logger.Info("User choose to overwrite the duplicate")

            if self.profile.Mode == Mode.Simulate:
                nlog_logger.Info("Deleted (simulated) {0}".format(move_path))
                self.logger.Add("Deleted (simulated)", move_path)
                if book.FilePath:
                    nlog_logger.Info("Moved {0} to {1} (simulated)"
                                     .format(book.FilePath, move_path))
                    self.logger.Add(ModeText.get_mode_past(self.profile.Mode), book.FilePath, "to: " + move_path)
                else:
                    nlog_logger.Info("Created image for {0} at {1} (simulated)"
                                     .format(book.Caption, move_path))
                    self.logger.Add("Created image", move_path)

                self._moved_books.append(move_path)
                return MoveResult.Success

            r = self._delete_duplicate(dup_path)
            if not r: return MoveResult.Failed
            
            if self.profile.CopyReadPercentage and type(existing_book) != FileInfo:
                book.LastPageRead = existing_book.LastPageRead
                nlog_logger.Info("Copied the read percentage from the duplicate"
                                 " to the processing book")

            if book.FilePath and type(existing_book) is not FileInfo:
                ComicRack.App.RemoveBook(existing_book)
                nlog_logger.Info("Removed the duplicate book from the library")

        processing_book.duplicate_different_extension = False
        processing_book.duplicate_ext_files = []
        result = self._process_book(processing_book)
        if result == MoveResult.Duplicate:
            return self.process_duplicate(processing_book)
        else:
            return result

    def _delete_duplicate(self, path):
        """Deletes a duplicate file by sending it to the recycle bin
        and handles the errors.

        Args:
            path: The string file path to try and move to the recycle bin.

        Returns: True if it succeeds and False if it fails.
        """
        try:
            FileIO.FileSystem.DeleteFile(
                path, FileIO.UIOption.OnlyErrorDialogs, 
                FileIO.RecycleOption.SendToRecycleBin)
            return True
        # Not catching any other exceptions because they should never occur.
        except (IOException, OperationCanceledException, SecurityException, 
                UnauthorizedAccessException) as ex:
            nlog_logger.Error("Could not delete {0}: {1}", path, ex.Message)
            self.logger.Add("Failed", self.report_book_name,"Could not delete the "
                            "existing duplicate {0} : {1}"
                            .format(path, ex.Message))
            return False

#     def _handle_duplicate_different_ext(self, book_to_move):
#         """Handles what to do with different"""
#         book = book_to_move.book
#         new_path = book_to_move.path
#         dde = book_to_move.duplicate_different_extension
#         results = []
#         mode = self.profile.Mode
#         self._set_book_name_for_report(book)
# 
#         print ("Starting to deal with the duplicate with different extension %s"
#                " at new path %s" % (self.report_book_name, new_path))
#         print ("There are %s books with the same destination path" % len(book_to_move.duplicate_ext_files))
#         print book_to_move.duplicate_ext_files
# 
#         dup_books = []
# 
#         for dup_file in book_to_move.duplicate_ext_files:
#          Handle each duplicate file.
#          Do we need to do this?
#          If "Overwriting" then we should check every file.
#          If "Renaming" Then we can assume the user wants to keep all of them
#          if multiple.
#          if cancel, same.
#          Better code path: Check the first one, then recursively call this
#          function
#          On the overwrite path.
#          Or....  not
#             print "Looking for the duplicate in the library"
#             dup_book = self._find_duplicate_book(dup_file.FullName)
# 
#             if dup_book is None:
#                 print "Couldn't find the duplicate in the library"
#                 dup_book = dup_file
#             else:
#                 print "Found the duplicate in the library"
#             rename_path = new_path
#             rename_filename = Path.GetFileName(rename_path)
#             dup_books.append(dup_book)
#             TODO: Fix/improve this gibberish.  Package up the args?
#              Don't even need really since we don't need to block the thread.
#             print "Asking what to do about this duplicate"
#             result = self.form.Invoke(Func[type(self.profile), 
#                      type(book), 
#                      type(dup_book), 
#                      str, 
#                      int, 
#                      DuplicateResult](self.form.ShowDuplicateForm), 
#                 System.Array[object]([self.profile, book, dup_book, 
#                                       rename_filename, 
#                                       self._held_duplicate_count]))
# 
#              By design I ignore the duplicate action for book with different
#              extensions.  I could set up a way to remember which extensions
#              should be kept but I feel that LO is not a duplicate finder and
#              this functionality would be beyond the goal of the program
#              self.duplicate_action = result.action
# 
#             if self.duplicate_action == DuplicateAction.Cancel:
#                 if book.FilePath:
#                     self.logger.Add("Skipped", self.report_book_name,
#                         DUPLICATE_EXT_CANCEL.format(dup_file.FullName))
#                 else:
#                     self.logger.Add("Skipped", self.report_book_name, 
#                         DUPLICATE_EXT_CANCEL_FILELESS.format(dup_file.FullName))
#                  If cancel then don't check other dup_books
#                 return MoveResult.Skipped
# 
#             elif self.duplicate_action == DuplicateAction.Rename:
#                  Move to the correct path, if there are multiple duplicates
#                  with different extensions then we ignore them.
#                 return self.move_book(book, new_path)
# 
#             elif self.duplicate_action == DuplicateAction.Overwrite:
#                  For over writing, delete the duplicate, move the book, then
#                  ask to delete the others.
# 
#                 Better yet, ask about the others all at once.
#                 if mode == Mode.Simulate:
#                     Because the script goes into a loop if in test mode here
#                     since no files are actually changed.  return a success
#                     self.logger.Add("Deleted (simulated)", full_path)
#                     if book.FilePath:
#                         self.logger.Add(ModeText.get_mode_past(self.profile.Mode), book.FilePath, "to: " + full_path)
#                     else:
#                         self.logger.Add("Created image", full_path)
# 
#                     self._moved_books.append(full_path)
#                     return MoveResult.Success
#                 try:
#                     if type(dup_book) is FileInfo:
#                         FileIO.FileSystem.DeleteFile(dup_book.FullPath, 
#                             FileIO.UIOption.OnlyErrorDialogs, 
#                             FileIO.RecycleOption.SendToRecycleBin)
#                     else:
#                         FileIO.FileSystem.DeleteFile(dup_book.FilePath, 
#                             FileIO.UIOption.OnlyErrorDialogs, 
#                             FileIO.RecycleOption.SendToRecycleBin)
# 
#                 except Exception, ex:
#                     self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + full_path + ". The error was: " + str(ex))
#                     return MoveResult.Failed
# 
#                 Since we are only working with images there is no need to
#                 remove a book from the library
#                 if book.FilePath and type(dupbook) is not FileInfo:
#                         ComicRack.App.RemoveBook(dupbook)
# 
#                 return self._process_duplicate_book(book_to_move)

    def move_book(self, book, new_path):
        nlog_logger.Trace("Starting move_book")
        try:
            book_file_info = FileInfo(book.FilePath)
        except (ArgumentException, ArgumentNullException):
            # The file name is empty, contains only white spaces, or contains
            # invalid characters.  Since if the file already exists in
            # ComicRack
            # then we can assume that it is a fileless book.
            # During the path creation process files without paths that aren't
            # going to have images created were pulled out.
            return self._save_fileless_image(book, new_path)
        except (SecurityException, UnauthorizedAccessException) as ex:
            # The caller does not have the required permission or Access to
            # fileName is denied.
            nlog_logger.Error("Can't access {0}: {1}", self.report_book_name,
                              ex.Message)
            self.logger.Add("Failed", self.report_book_name, 
                            ("An error occurred accessing the book file." 
                            "The error was: %s" % ex.Message))
        try:
            if self.profile.Mode == Mode.Move:
                book_file_info.MoveTo(new_path)
                self.undo_collection.append(book.FilePath, new_path, 
                                            self.profile.Name)
                nlog_logger.Info("Moved {0} to {1} succesffully", 
                                 book.FilePath, new_path)
                book.FilePath = new_path

            elif self.profile.Mode == Mode.Simulate:
                self.logger.Add(ModeText.get_mode_past(self.profile.Mode), 
                                book.FilePath, "to: %s" % new_path)
                self._moved_books.append(new_path)

            elif self.profile.Mode == Mode.Copy:
                book_file_info.CopyTo(new_path)
                if self.profile.CopyMode:
                    newbook = ComicRack.App.AddNewBook(False)
                    newbook.FilePath = new_path
                    CopyData(book, newbook)
            return MoveResult.Success

        except DirectoryNotFoundException:
            result = self._create_folder(FileInfo(new_path).Directory)
            if result == MoveResult.Success:
                return self.move_book(book, new_path)
            return MoveResult.Failed
        #TODO: finish nlog conversions
        except (UnauthorizedAccessException, SecurityException) as ex:
            print "Can't access %s" % book.FilePath
            self.logger.Add("Failed", self.report_book_name, 
                            ("An error occurred accessing the book file." 
                            "The error was: %s" % ex))
            return MoveResult.Failed

        except IOException as ex:
            print "Can't access %s" % book.FilePath
            self.logger.Add("Failed", self.report_book_name, 
                            "An IO error occurred. The error was: %s" % ex)
            return MoveResult.Failed

    def _save_fileless_image(self, book, path):
        """Creates and saves the fileless image to file.

        Args:
            book: The ComicBook from which to save the image from.
            path: The fully qualified path to save the image to.

        Returns:
            MoveResult.Success or MoveResult.Failed
        """
        try:
            if self.profile.Mode == Mode.Simulate:
                self.logger.Add("Created image", path)
                self._moved_books.append(path)
            else:      
                #TODO: preferred front cover
                image = ComicRack.App.GetComicThumbnail(book, 0)
                img_format = None
                if self.profile.FilelessFormat == ".jpg":
                    img_format = ImageFormat.Jpeg
                elif self.profile.FilelessFormat == ".png":
                    img_format = ImageFormat.Png
                elif self.profile.FilelessFormat == ".bmp":
                    img_format = ImageFormat.Bmp
                image.Save(path, img_format)
            return MoveResult.Success
        #TODO: better exception handling
        except Exception, ex:
            self.logger.Add("Failed", self.report_book_name, "Failed to create the image because an error occurred. The error was: " + str(ex))
            return MoveResult.Failed

    def _book_should_be_moved_with_rules(self, book):
        """Checks the excluded folders and metadata rules to see if the book 
        should be moved.

        Args:
            book: The ComicBook object to check with.
        Returns:
            True if the book should be moved. False otherwise
        """
        if not check_excluded_folders(book.FilePath, self.profile):
            self.logger.Add("Skipped", self.report_book_name, 
                            "The book is located in an excluded path")
            return False
            
        if not check_metadata_rules(book, self.profile):
            self.logger.Add("Skipped", self.report_book_name, 
                            "The book qualified under the exclude rules")
            return False
        return True

    def _check_bookpath_same_as_newpath(self, book, new_file_name, 
                                        new_full_path):
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
            self.logger.Add("Skipped", 
                self.report_book_name, 
                "The book is already located at the calculated path")
            return True

        # In some cases the file path is the same but has different cases.
        # The FileInfo object doesn't catch this but the File.Move function
        # still thinks that it is a duplicate.
        if new_full_path.lower() == book.FilePath.lower():
            #In that case, better rename it to the correct case
            print "Renaming %s to %s to fix capitalization" % (book.FilePath,
                                                               new_full_path)
            if self.profile.Mode == Mode.Simulate:
                self.logger.Add("Renaming", self.report_book_name, "to: " + new_full_path)
            else:
                book.RenameFile(new_file_name)
                self.logger.Add("Renaming", self.report_book_name, 
                                "to %s to fix capitalization" % new_full_path)
            self.logger.Add("Skipped", self.report_book_name, 
                "The book is already located at the calculated path")
            return True
        return False

    def _get_fixed_path(self, long_path):
        """Asks the user to fix the path by invoking a dialog.

        Args:
            long_path: The fully qualified path to check as a string.

        Returns:
            MoveResult.Skipped if the user doesn't want to rename the path.
            The shortened, correct path otherwise.
        """
        #TODO:Doesn't need invoke
        result = self.form.Invoke(Func[str, object](self._get_smaller_path), 
                                  System.Array[System.Object]([long_path]))
        if result is None:
            self.logger.Add("Skipped", self.report_book_name, 
                "The calculated path was too long and the user skipped shortening it")
            return MoveResult.Skipped
        return result

    def _find_duplicate_book(self, path):
        """Tries to find a ComicBook in the ComicRack library by checking the 
        path against all the books in the library.

        It is possible to have different cases in the stored path and
        created paths, therefore both paths are converted to lowercase despite
        the possible performance hit this may have.
        
        If ignore_extension is set to True then it will attempt to match file 
        names without the extensions. 

        Args:
            path: The string path to search in the library for.

        Returns:
            A ComicBook object if located in the library, None otherwise.
        """
        nlog_logger.Info("Trying to find the duplicate book in the library:")
        path = path.lower()
        #path_no_ext = path[:-len(Path.GetExtension(path))]
        #TODO: Make sure commented code isn't needed and fix description
        for book in ComicRack.App.GetLibraryBooks():
            #Fix Issues #5: different capitalization in path names.
            if book.FilePath.lower() == path:
                nlog_logger.Info("Found a duplicate in the library")
                return book
            #elif ignore_extension:
            #    if lower(book.FilePath)[:-len(Path.GetExtension(book.FilePath))] == path_no_ext:
            #        nlog_logger.Info("-->Found a duplicate in the library with"
            #                         " a different extension")
            #        return book
        nlog_logger.Info("Could not find a duplicate in the library")
        return None

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
        nlog_logger.Trace("Creating folders")
        if new_directory.Exists:
            return MoveResult.Success
        
        if self.profile.Mode == Mode.Simulate:
            if new_directory.FullName not in self._created_paths:
                self.logger.Add("Created Folder", new_directory.FullName)
                self._created_paths.append(new_directory.FullName)
                nlog_logger.Info("-->Created Folder {0} (Simulated)", new_directory.FullName)
            return MoveResult.Success
        try:            
            new_directory.Create()
            new_directory.Refresh()
            nlog_logger.Info("Created folder(s) {0}", new_directory.FullName)
            return MoveResult.Success
        except IOException as ex:
            self.logger.Add("Failed to create folder", new_directory, 
                ("Book %s  was not moved because an error occurred "
                "creating the folder. The error was: %s" % (self.report_book_name, ex.Message)))
            nlog_logger.Error("Failed to create folder {0}: {1}", 
                              new_directory.FullName, ex.Message)
            return MoveResult.Failed

    def _create_rename_path(self, path):
        """Creates a Windows formatted duplicate path in the manner of
        Filename (#).cbz. It searches until it finds the first unused number.

        Originally written by pescuma and modified slightly.

        Args:
            path: The fully qualified path to find the duplicate for.

        Returns:
            The fully qualified path that was created.
        """
        nlog_logger.Trace("Trying to find a rename path")
        extension = Path.GetExtension(path)
        base = path[:-len(extension)]
        base = re.sub(" \([0-9]\)$", "", base)
        i = 0 
        while (True):
            i += 1
            newpath = "{0} ({1}){2}".format(base, i + 1, extension)
            #For simulated mode we check _moved_books
            if newpath in self._moved_books or File.Exists(newpath):
                continue
            else:
                nlog_logger.Info("Found rename path: {0}", newpath)
                return newpath

    def _delete_empty_folders(self, directory):
        """Recursively deletes directories until an non-empty directory is 
        found or the directory is in the excluded list.
        
        Args:
            directory: A DirectoryInfo object which starts at the first folder
                to remove.
        """
        directory.Refresh()
        if not directory.Exists:
            return        
        #Only delete if no file or folder and not in folder never to delete
        if directory.FullName not in self.profile.ExcludedEmptyFolder:
            try:
                parent = directory.Parent
                directory.Delete()
                nlog_logger.Info("Deleteed empty folder {0}", directory.FullName)
                self._delete_empty_folders(parent)
            except (UnauthorizedAccessException, 
                    DirectoryNotFoundException, 
                    IOException, SecurityException) as ex:
                # It's fine if we can't delete a folder, this is usually
                # because
                # there is something still in it.
                nlog_logger.Warn("Unable to delete {0} {1}", 
                                 directory.FullName, ex.Message)

    def _get_smaller_path(self, path):
        result = None
        with PathTooLongForm(path) as  p:
            r = p.ShowDialog()
            if r == DialogResult.OK:
                #TODO: Fix this bad variable accessing
                result = p._Path.Text
        return result

    #def _ask_about_duplicate(self, profile, newbook, oldbook, renamefile,
    #count):
    #    if self._DuplicateForm == None:
    #        self._DuplicateForm = DuplicateForm(profile.Mode)
    #        helper = WindowInteropHelper(self._DuplicateForm.win)
    #        helper.Owner = self.Handle
    #    return self._DuplicateForm.ShowForm(newbook, oldbook, renamefile,
    #    count)

# class UndoMover(BookMover):
#     
#     def __init__(self, worker, form, undo_collection, profiles, logger):
#         self.worker = worker
#         self.form = form
#         self.undo_collection = undo_collection
#         self.AlwaysDoAction = False
#         self.HeldDuplicateBooks = {}
#         self.logger = logger
#         self.profiles = profiles
#         self.failed_or_skipped = False
# 
# 
#     def process_books(self):
#         books, notfound = self.get_library_books()
# 
#         success = 0
#         failed = 0
#         skipped = 0
#         count = 0
# 
#         for book in books + notfound:
#             count += 1
# 
#             if self._worker.CancellationPending:
#                 skipped = len(books) + len(notfound) - success - failed
#                 self.logger.Add("Canceled", str(skipped) + " files", "User cancelled the script")
#                 report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0 or skipped > 0)
#                 #self.logger.SetCountVariables(failed, skipped, success)
#                 return report
# 
#             result = self.process_book(book)
# 
#             if result is MoveResult.Duplicate:
#                 count -= 1
#                 continue
# 
#             elif result is MoveResult.Skipped:
#                 skipped += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#             elif result is MoveResult.Failed:
#                 failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#             elif result is MoveResult.Success:
#                 success += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#         self.HeldDuplicateCount = len(self._held_duplicate_books)
#         for book in self._held_duplicate_books:
# 
#             if self._worker.CancellationPending:
#                 skipped = len(books) + len(notfound) - success - failed
#                 self.logger.Add("Canceled", str(skipped) + " files", "User cancelled the script")
#                 report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0 or skipped > 0)
#                 #self.logger.SetCountVariables(failed, skipped, success)
#                 return report
# 
#             result = self.process_duplicate_book(book, self._held_duplicate_books[book])
#             self._held_duplicate_count -= 1
# 
#             if result is MoveResult.Skipped:
#                 skipped += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#             elif result is MoveResult.Failed:
#                 failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#             elif result is MoveResult.Success:
#                 success += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#         report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0 or skipped > 0)
#         #self.logger.SetCountVariables(failed, skipped, success)
#         return report
# 
# 
#     def MoveBooks(self):
# 
#         success = 0
#         failed = 0
#         skipped = 0
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
#                 skipped = len(books) + len(notfound) - success - failed
#                 self.report.Append("\n\nOperation cancelled by user.")
#                 break
# 
#             if not File.Exists(oldfile):
#                 self.report.Append("\n\nFailed to move\n%s\nbecause the file does not exist." % (oldfile))
#                 failed += 1
#                 self._worker.ReportProgress(count)
#                 continue
# 
#         
#             if path == oldfile:
#                 self.report.Append("\n\nSkipped moving book\n%s\nbecause it is already located at the calculated path." % (oldfile))
#                 skipped += 1
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
#                 success += 1
#             
#             elif result == MoveResult.Failed:
#                 failed += 1
#             
#             elif result == MoveResult.Skipped:
#                 skipped += 1
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
#                 skipped = len(books) + len(notfound) - success - failed
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
#                 success += 1
#             
#             elif result == MoveResult.Failed:
#                 failed += 1
#             
#             elif result == MoveResult.Skipped:
#                 skipped += 1
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
#         report = "Successfully moved: %s\nFailed to move: %s\nSkipped: %s" % (success, failed, skipped)
#         return [failed + skipped, report, self.report.ToString()]
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
#             self.logger.Add("Failed", self.report_book_name, "The file does not exist")
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
#         result = self.move_book(book, undo_path)
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
#                 self.logger.Add("Skipped", self.report_book_name, "A file already exists at: " + undo_path + " and the user declined to overwrite it")
#                 return MoveResult.Skipped
# 
#             elif self.duplicate_action is DuplicateAction.Rename:
#                 #Check if the created path is too long
#                 if len(rename_path) > 259:
#                     result = self.form.Invoke(Func[str, object](self._get_smaller_path), System.Array[System.Object]([rename_path]))
#                     if result is None:
#                         self.logger.Add("Skipped", self.report_book_name, "The path was too long and the user skipped shortening it")
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
#                     self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + undo_path + ". The error was: " + str(ex))
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
#         result = self.move_book(book, undo_path)
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
#             self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
#             return MoveResult.Skipped
# 
#         #In some cases the filepath is the same but has different cases.  The
#         #FileInfo object dosn't catch this but the File.Move function
#         #Thinks that it is a duplicate.
#         if undo_path.lower() == current_path.lower():
#             #In that case, better rename it to the correct case
#             book.RenameFile(file_name)
#             self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
#             return MoveResult.Skipped
# 
#         return None
# 
# 
#     def move_book(self, book, path):
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
#             self.logger.Add("Failed", self.report_book_name, "because an error occured. The error was: " + str(ex))
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
    newbook.SetInfo(book, False, False) # This copies most fields
    newbook.CustomValuesStore = book.CustomValuesStore
    newbook.SeriesComplete = book.SeriesComplete # Not copied by SetInfo
    newbook.Rating = book.Rating #Not copied by SetInfo
    newbook.customThumbnailKey = book.CustomThumbnailKey