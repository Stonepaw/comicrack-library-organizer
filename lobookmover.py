"""
lobookmover.py

This contains all the book moving fuction the script uses. Also the path generator

Version 2.1:


Copyright 2010-2012 Stonepaw

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

import re

import System

from System import Func, Action, ArgumentException, ArgumentNullException, NotSupportedException

from System.Text import StringBuilder

from System.IO import Path, File, FileInfo, DirectoryInfo, Directory, IOException, PathTooLongException, DirectoryNotFoundException

import loforms

from loforms import PathTooLongForm, MultiValueSelectionFormArgs, MultiValueSelectionFormResult, MultiValueSelectionForm

import locommon

from locommon import Mode, get_earliest_book, name_to_field, field_to_name, check_metadata_rules, check_excluded_folders, UNDOFILE, UndoCollection, get_last_book

from loduplicate import DuplicateResult, DuplicateForm, DuplicateAction


import lologger

clr.AddReference("System.Drawing")
from System.Drawing.Imaging import ImageFormat

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult

clr.AddReferenceByPartialName('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

clr.AddReference("Microsoft.VisualBasic")

from Microsoft.VisualBasic import FileIO


class MoveResult(object):
    Success = 1
    Failed = 2
    Skipped = 3
    Duplicate = 4
    FailedMove = 5



class BookToMove(object):

    def __init__(self, book, path, index, failed_fields):
        self.book = book
        self.path = path
        self.profile_index = index
        self.failed_fields = failed_fields



class BookMoverResult(object):
    
    def __init__(self, report_text, failed_or_skipped):
        self.report_text = report_text
        self.failed_or_skipped = failed_or_skipped



class ProfileReport(object):
    def __init__(self, total, name, mode):
        self.success = 0
        self.failed = 0
        self.skipped = 0
        self._total = total
        self._name = name
        self._mode = mode


    def get_report(self, cancelled):
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
    
    def __init__(self, worker, form, logger):
        self.worker = worker
        self.form = form
        self.logger = logger
        self.pathmaker = PathMaker(form, None)

        self.failed_or_skipped = False
        
        #Hold books that are duplicates so they can be all asked at the end.

        self.HeldDuplicateBooks = []
        self.HeldDuplicateCount = 0
        
        #These variables are for when the script is in test mode

        self.CreatedPaths = []
        self.MovedBooks = []

        #This hold a list of the book moved and is saved in undo.txt for the undo script
        self.undo_collection = UndoCollection()
        
        #For duplicates
        self.always_do_duplicate_action = False
        self.duplicate_action = None


    def create_book_paths(self, books, profiles):
        """Find the destination paths for all the books given a set of profiles.
        Only the last file path is found if the book can be moved under several profiles.
        If several profiles are in copy mode, the book will be copied several times.
        Returns a list of BookToMove objects.
        """
        books_and_paths = []

        self.profile_reports = [ProfileReport(len(books), profile.Name, profile.Mode) for profile in profiles]

        for book in books:
            path = ""
            profile_index = None
            failed_fields = []
            if book.FilePath:
                self.report_book_name = book.FilePath
            else:
                self.report_book_name = book.Caption

            for profile in profiles:
                index = profiles.index(profile)
                self.profile = profile
                self.pathmaker.profile = profile
                self.logger.SetProfile(profile.Name)

                result = self.create_book_path(book)

                if result is MoveResult.Skipped:
                    self.profile_reports[index].skipped += 1
                    self.failed_or_skipped = True
                    continue

                elif result is MoveResult.Failed:
                    self.profile_reports[index].failed += 1
                    self.failed_or_skipped = True
                    continue

                else:
                    if profile.Mode == Mode.Copy:
                        books_and_paths.append(BookToMove(book, result, index, self.pathmaker.failed_fields))
                        continue
                    else:
                        if path:
                            self.profile_reports[profile_index].skipped +=1
                            self.logger.Add("Skipped", self.report_book_name, "The book is moved by a later profile", profiles[profile_index].Name)
                            self.failed_or_skipped = True
                        path = result
                        profile_index = index
                        failed_fields = self.pathmaker.failed_fields

            if path:
                self.profile = profiles[profile_index]
                self.logger.SetProfile(self.profile.Name)
                #Because the path can already be at location the final profile says the book may be moved with the wrong profile is this is checked earlier.
                result = self.check_path_problems(book, Path.GetFileName(path), path)
                if result is MoveResult.Skipped:
                    self.profile_reports[profile_index].skipped = True
                    self.failed_or_skipped = True
                else:
                    books_and_paths.append(BookToMove(book, path, profile_index, failed_fields))

        for item in books_and_paths:
            print item.book.FilePath + " : " + item.path

        return books_and_paths

    
    def create_book_path(self, book):
        """Creates the new path and checks it for some problems.
        Returns the path or a MoveResult if something goes wrong.
        """
        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

        if book.FilePath and not File.Exists(book.FilePath):
            self.logger.Add("Failed", self.report_book_name, "The file does not exist")
            return MoveResult.Failed

        if not self.book_should_be_moved_with_rules(book):
            return MoveResult.Skipped

        #Fileless
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
            self.logger("Failed", self.report_book_name, "The created filename was blank")
            return MoveResult.Failed


        return full_path

        
    def process_books(self, books, profiles):
        
        

        books_to_move = self.create_book_paths(books, profiles)

        if not books_to_move:
            header_text = "\n\n".join([profile_report.get_report(True) for profile_report in self.profile_reports])
            report = BookMoverResult(header_text, self.failed_or_skipped)
            self.logger.add_header(header_text)
            return report

        percentage = 1.0/len(books_to_move)*100
        progress = 0.0

        count = 0

        for book in books_to_move:

            if self.worker.CancellationPending:
                self.logger.Add("Canceled", str(len(books_to_move) - count) + " operations", "User cancelled the script")
                header_text = "\n\n".join([profile_report.get_report(True) for profile_report in self.profile_reports])
                report = BookMoverResult(header_text, self.failed_or_skipped)
                self.logger.add_header(header_text)
                return report

            count += 1
            progress += percentage

            self.profile = profiles[book.profile_index]
            self.logger.SetProfile(self.profile.Name)

            result = self.process_book(book)

            if result is MoveResult.Duplicate:
                count -= 1
                progress -= percentage
                self.HeldDuplicateBooks.append(book)
                continue

            elif result is MoveResult.Skipped:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].skipped += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

            elif result is MoveResult.Failed:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].failed += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

            elif result is MoveResult.Success:
                self.profile_reports[book.profile_index].success += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

        self.HeldDuplicateCount = len(self.HeldDuplicateBooks)
        for book in self.HeldDuplicateBooks:

            if self.worker.CancellationPending:
                self.logger.Add("Canceled", str(len(books_to_move) - count) + " operations", "User cancelled the script")
                header_text = "\n\n".join([profile_report.get_report(True) for profile_report in self.profile_reports])
                report = BookMoverResult(header_text, self.failed_or_skipped)
                self.logger.add_header(header_text)
                return report
            
            count += 1
            progress += percentage

            self.profile = profiles[book.profile_index]
            self.logger.SetProfile(self.profile.Name)

            result = self.process_duplicate_book(book)
            self.HeldDuplicateCount -= 1

            if result is MoveResult.Skipped:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].skipped += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

            elif result is MoveResult.Failed:
                self.failed_or_skipped = True
                self.profile_reports[book.profile_index].failed += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

            elif result is MoveResult.Success:
                self.profile_reports[book.profile_index].success += 1
                self.worker.ReportProgress(int(round(progress)))
                continue

        if len(self.undo_collection) > 0:
            self.undo_collection.save(UNDOFILE)

        header_text = "\n\n".join([profile_report.get_report(True) for profile_report in self.profile_reports])
        report = BookMoverResult(header_text, self.failed_or_skipped)
        self.logger.add_header(header_text)
        return report


    def process_book(self, book_to_move):
        
        book = book_to_move.book

        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

        full_path = book_to_move.path

        result, full_path = self.check_path_to_long(book, full_path)
        if result is not None:
            return result

        #Duplicate
        if File.Exists(full_path) or full_path in self.MovedBooks:
            return MoveResult.Duplicate

        #Create here because needed for cleaning directories later
        old_folder_path = book.FileDirectory
        folder_path = Path.GetDirectoryName(full_path)

        result = self.create_folder(folder_path, book)
        if result is not MoveResult.Success:
            return result


        if not book.FilePath:
            result = self.create_fileless_image(book, full_path)
                
        else:
            result = self.move_book(book, full_path)

        if self.profile.RemoveEmptyFolder and self.profile.Mode == Mode.Move:
            if old_folder_path:
                self.remove_empty_folders(DirectoryInfo(old_folder_path))
            self.remove_empty_folders(DirectoryInfo(folder_path))

        if book_to_move.failed_fields and result == MoveResult.Success:
            if len(book_to_move.failed_fields) > 1:
                failed_report_verb = " are"
            else:
                failed_report_verb = " is"
            self.logger.Add("Failed", self.report_book_name, ",".join(book_to_move.failed_fields) + failed_report_verb + " empty. " + ModeText.get_mode_past(self.profile.Mode) + " to " + full_path)
            return MoveResult.Failed

        return result


    def process_duplicate_book(self, book_to_move):

        book = book_to_move.book

        full_path = book_to_move.path

        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

        #Since the duplicate is checked for last in the orginal process_book function there is no need to check for path errors.
        if File.Exists(full_path) or full_path in self.MovedBooks:

            #Find the existing book if it occurs in the library
            oldbook = self.find_duplicate_book(full_path)
            if oldbook == None:
                oldbook = FileInfo(full_path)

            rename_path = self.create_rename_path(full_path)
            rename_filename = Path.GetFileName(rename_path)
        
            if not self.always_do_duplicate_action:
                result = self.form.Invoke(Func[type(self.profile), type(book), type(oldbook), str, int, DuplicateResult](self.form.ShowDuplicateForm), System.Array[object]([self.profile, book, oldbook, rename_filename, self.HeldDuplicateCount]))

                self.duplicate_action = result.action

                if result.always_do_action:
                    self.always_do_duplicate_action = True

            if self.duplicate_action is DuplicateAction.Cancel:
                if book.FilePath:
                    self.logger.Add("Skipped", self.report_book_name, "A file already exists at: " + full_path + " and the user declined to overwrite it")
                else:
                    self.logger.Add("Skipped", self.report_book_name, "The image already exists at: " + full_path + " and the user declined to overwrite it")
                return MoveResult.Skipped

            elif self.duplicate_action is DuplicateAction.Rename:
                #Check if the created path is too long
                if len(rename_path) > 259:
                    result = self.form.Invoke(Func[str, object](self.get_smaller_path), System.Array[System.Object]([rename_path]))
                    if result is None:
                        self.logger.Add("Skipped", self.report_book_name, "The path was too long and the user skipped shortening it")
                        return MoveResult.Skipped

                    return self.process_duplicate_book(BookToMove(book, result, book_to_move.profile_index, book_to_move.failed_fields))
                return self.process_duplicate_book(BookToMove(book, rename_path, book_to_move.profile_index, book_to_move.failed_fields))

            elif self.duplicate_action is DuplicateAction.Overwrite:
                try:
                    if self.profile.Mode == Mode.Simulate:
                        #Because the script goes into a loop if in test mode here since no files are actually changed. return a success
                        self.logger.Add("Deleted (simulated)", full_path)
                        if book.FilePath:
                            self.logger.Add(ModeText.get_mode_past(self.profile.Mode), book.FilePath, "to: " + full_path)
                        else:
                            self.logger.Add("Created image", full_path)

                        self.MovedBooks.append(full_path)
                        return MoveResult.Success
                    else:
                        if self.profile.CopyReadPercentage and type(oldbook) is not FileInfo:
                            book.LastPageRead = oldbook.LastPageRead
                        FileIO.FileSystem.DeleteFile(full_path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                except Exception, ex:
                    self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + full_path + ". The error was: " + str(ex))
                    return MoveResult.Failed

                #Since we are only working with images there is no need to remove a book from the library
                if book.FilePath and type(oldbook) is not FileInfo:
                        ComicRack.App.RemoveBook(oldbook)

                return self.process_duplicate_book(book_to_move)

        old_folder_path = book.FileDirectory
        
        if book.FilePath:
            result = self.move_book(book, full_path)
            
        else:
            result = self.create_fileless_image(book, full_path)

        if self.profile.RemoveEmptyFolder and self.profile.Mode == Mode.Move:
            if old_folder_path:
                self.remove_empty_folders(DirectoryInfo(old_folder_path))
            self.remove_empty_folders(FileInfo(full_path).Directory)

        if book_to_move.failed_fields and result == MoveResult.Success:
            if len(book_to_move.failed_fields) > 1:
                failed_report_verb = " are"
            else:
                failed_report_verb = " is"
            self.logger.Add("Failed", self.report_book_name, ",".join(book_to_move.failed_fields) + failed_report_verb + " empty. " + ModeText.get_mode_present(self.profile.Mode) + " to " + full_path)
            return MoveResult.Failed


        return result


    def move_book(self, book, path):
        #Finally actually move the book
        try:
            if self.profile.Mode == Mode.Move:
                File.Move(book.FilePath, path)
                self.undo_collection.append(book.FilePath, path, self.profile.Name)
                book.FilePath = path

            elif self.profile.Mode == Mode.Simulate:
                self.logger.Add(ModeText.get_mode_past(self.profile.Mode), book.FilePath, "to: " + path)
                self.MovedBooks.append(path)

            elif self.profile.Mode == Mode.Copy:
                File.Copy(book.FilePath, path)
                if self.profile.CopyMode:
                    newbook = ComicRack.App.AddNewBook(False)
                    newbook.FilePath = path
                    CopyData(book, newbook)

            return MoveResult.Success

        except Exception, ex:
            self.logger.Add("Failed", self.report_book_name, "because an error occured. The error was: " + str(ex))
            return MoveResult.Failed


    def create_fileless_image(self, book, path):
        #Finally actually move the book
        try:            
            image = ComicRack.App.GetComicThumbnail(book, 0)
            format = None
            if self.profile.FilelessFormat == ".jpg":
                format = ImageFormat.Jpeg
            elif self.profile.FilelessFormat == ".png":
                format = ImageFormat.Png
            elif self.profile.FilelessFormat == ".bmp":
                format = ImageFormat.Bmp
                
            if self.profile.Mode == Mode.Simulate:
                self.logger.Add("Created image", path)
                self.MovedBooks.append(path)
            else:
                image.Save(path, format)
            return MoveResult.Success

        except Exception, ex:
            self.logger.Add("Failed", self.report_book_name, "Failed to create the image because an error occured. The error was: " + str(ex))
            return MoveResult.Failed

    
    def book_should_be_moved_with_rules(self, book):
        """Checks the exlcuded folders and metadata rules to see if the book should be moved.
        Returns True if the book should be moved.
        """
        if not check_excluded_folders(book.FilePath, self.profile):
            self.logger.Add("Skipped", self.report_book_name, "The book is located in an excluded path")
            return False
            

        if not check_metadata_rules(book, self.profile):
            self.logger.Add("Skipped", self.report_book_name, "The book qualified under the exclude rules")
            return False

        return True


    def check_path_problems(self, book, file_name, full_path):


        if full_path == book.FilePath:
            self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
            return MoveResult.Skipped

        #In some cases the filepath is the same but has different cases. The FileInfo object dosn't catch this but the File.Move function
        #Thinks that it is a duplicate.
        if full_path.lower() == book.FilePath.lower():
            #In that case, better rename it to the correct case
            if self.profile.Mode == Mode.Simulate:
                self.logger.Add("Renaming", self.report_book_name, "to: " + full_path)
            else:
                book.RenameFile(file_name)
            self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
            return MoveResult.Skipped

        return None


    def check_path_to_long(self, book, full_path):

        if len(full_path) > 259:
            result = self.form.Invoke(Func[str, object](self.get_smaller_path), System.Array[System.Object]([full_path]))
            if result is None:
                self.logger.Add("Skipped", self.report_book_name, "The calculated path was too long and the user skipped shortening it")
                return MoveResult.Skipped, ""
            full_path = result

        return None, full_path


    def find_duplicate_book(self, path):
        """
        Trys to find a book in the CR library via a path
        """
        for book in ComicRack.App.GetLibraryBooks():
            if book.FilePath == path:
                return book
        return None


    def create_folder(self, folder_path, book):
        """Creates the folder path.

        Returns MoveResult.Succes if the creation succeeded, MoveResult.Failed if something went wrong.
        """
        if not Directory.Exists(folder_path):
            try:
                if self.profile.Mode == Mode.Simulate:
                    if not folder_path in self.CreatedPaths:
                        self.logger.Add("Created Folder", folder_path)
                        self.CreatedPaths.append(folder_path)
                else:
                    Directory.CreateDirectory(folder_path)

            except (IOException, ArgumentException, ArgumentNullException, PathTooLongException, DirectoryNotFoundException, NotSupportedException), ex:
                self.logger.Add("Failed to create folder", folder_path, "Book " + self.report_book_name + " was not moved.\nThe error was: " + str(type(ex)) + ": " + ex.Message)
                return MoveResult.Failed

        return MoveResult.Success


    def create_rename_path(self, path):

        #By pescuma. modified slightly
        extension = Path.GetExtension(path)
        base = path[:-len(extension)]

        base = re.sub(" \([0-9]\)$", "", base)
                
        for i in range(100):
            newpath = base + " (" + str(i+1) + ")" + extension

            #For test mode
            if newpath in self.MovedBooks:
                continue

            if File.Exists(newpath):
                continue

            else:
                return newpath

    
    def remove_empty_folders(self, directory):
        """
        Recursivly deletes directories until an non-empty directory is found or the directory is in the excluded list
        
        directory should be a DirectoryInfo object
        
        """
        if not directory.Exists:
            return
        
        #Only delete if no file or folder and not in folder never to delete
        if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.profile.ExcludedEmptyFolder:
            parent = directory.Parent
            directory.Delete()
            self.remove_empty_folders(parent)


    def get_smaller_path(self, path):

        p = PathTooLongForm(path)
        r = p.ShowDialog()

        if r != DialogResult.OK:
            return None

        return p._Path.Text



class UndoMover(BookMover):
    
    def __init__(self, worker, form, undo_collection, profiles, logger):
        self.worker = worker
        self.form = form
        self.undo_collection = undo_collection
        self.AlwaysDoAction = False
        self.HeldDuplicateBooks = {}
        self.logger = logger
        self.profiles = profiles
        self.failed_or_skipped = False


    def process_books(self):
        books, notfound = self.get_library_books()

        success = 0
        failed = 0
        skipped = 0
        count = 0

        for book in books + notfound:
            count += 1

            if self.worker.CancellationPending:
                skipped = len(books) + len(notfound) - success - failed
                self.logger.Add("Canceled", str(skipped) + " files", "User cancelled the script")
                report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0  or skipped > 0)
                #self.logger.SetCountVariables(failed, skipped, success)
                return report

            result = self.process_book(book)

            if result is MoveResult.Duplicate:
                count -= 1
                continue

            elif result is MoveResult.Skipped:
                skipped += 1
                self.worker.ReportProgress(count)
                continue

            elif result is MoveResult.Failed:
                failed += 1
                self.worker.ReportProgress(count)
                continue

            elif result is MoveResult.Success:
                success += 1
                self.worker.ReportProgress(count)
                continue

        self.HeldDuplicateCount = len(self.HeldDuplicateBooks)
        for book in self.HeldDuplicateBooks:

            if self.worker.CancellationPending:
                skipped = len(books) + len(notfound) - success - failed
                self.logger.Add("Canceled", str(skipped) + " files", "User cancelled the script")
                report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0  or skipped > 0)
                #self.logger.SetCountVariables(failed, skipped, success)
                return report

            result = self.process_duplicate_book(book, self.HeldDuplicateBooks[book])
            self.HeldDuplicateCount -= 1

            if result is MoveResult.Skipped:
                skipped += 1
                self.worker.ReportProgress(count)
                continue

            elif result is MoveResult.Failed:
                failed += 1
                self.worker.ReportProgress(count)
                continue

            elif result is MoveResult.Success:
                success += 1
                self.worker.ReportProgress(count)
                continue

        report = BookMoverResult("Successfully moved: %s\tFailed to move: %s\tSkipped: %s\n\n" % (success, failed, skipped), failed > 0  or skipped > 0)
        #self.logger.SetCountVariables(failed, skipped, success)
        return report


    def MoveBooks(self):

        success = 0
        failed = 0
        skipped = 0
        count = 0
        #get a list of the books
        books, notfound = self.get_library_books()



        for book in books + notfound:
            if type(book) == str:
                path = self.undo_collection[book]
                oldfile = book
            else:
                path = self.undo_collection[book.FilePath]
                oldfile = book.FilePath
            count += 1

            if self.worker.CancellationPending:
                #User pressed cancel
                skipped = len(books) + len(notfound) - success - failed
                self.report.Append("\n\nOperation cancelled by user.")
                break

            if not File.Exists(oldfile):
                self.report.Append("\n\nFailed to move\n%s\nbecause the file does not exist." % (oldfile))
                failed += 1
                self.worker.ReportProgress(count)
                continue

        
            if path == oldfile:
                self.report.Append("\n\nSkipped moving book\n%s\nbecause it is already located at the calculated path." % (oldfile))
                skipped += 1
                self.worker.ReportProgress(count)
                continue

            #Created the directory if need be
            f = FileInfo(path)

            if f.Exists:
                self.HeldDuplicateBooks.append(book)
                count -= 1
                continue

            d = f.Directory
            if not d.Exists:
                d.Create()

            if type(book) == str:
                oldpath = FileInfo(book).DirectoryName
            else:
                oldpath = book.FileDirectory

            result = self.MoveBook(book, path)
            if result == MoveResult.Success:
                success += 1
            
            elif result == MoveResult.Failed:
                failed += 1
            
            elif result == MoveResult.Skipped:
                skipped += 1

            
            #If cleaning directories
            if self.settings.RemoveEmptyFolder:
                self.CleanDirectories(DirectoryInfo(oldpath))
                self.CleanDirectories(DirectoryInfo(f.DirectoryName))
            self.worker.ReportProgress(count)

        #Deal with the duplicates
        for book in self.HeldDuplicateBooks[:]:
            if type(book) == str:
                path = self.undo_collection[book]
                oldpath = book
            else:
                path = self.undo_collection[book.FilePath]
                oldpath = book.FilePath
            count += 1

            if self.worker.CancellationPending:
                #User pressed cancel
                skipped = len(books) + len(notfound) - success - failed
                self.report.Append("\n\nOperation cancelled by user.")
                break

            #Created the directory if need be
            f = FileInfo(path)

            d = f.Directory
            if not d.Exists:
                d.Create()

            if type(book) == str:
                oldpath = FileInfo(book).DirectoryName
            else:
                oldpath = book.FileDirectory

            result = self.MoveBook(book, path)
            if result == MoveResult.Success:
                success += 1
            
            elif result == MoveResult.Failed:
                failed += 1
            
            elif result == MoveResult.Skipped:
                skipped += 1

            
            #If cleaning directories
            if self.settings.RemoveEmptyFolder:
                self.CleanDirectories(DirectoryInfo(oldpath))
                self.CleanDirectories(DirectoryInfo(f.DirectoryName))

            self.HeldDuplicateBooks.remove(book)

            self.worker.ReportProgress(count)

        #Return the report to the worker thread
        report = "Successfully moved: %s\nFailed to move: %s\nSkipped: %s" % (success, failed, skipped)
        return [failed + skipped, report, self.report.ToString()]


    def process_book(self, book):
        
        if type(book) == str:
            undo_path = self.undo_collection.undo_path(book)
            current_path = book
            self.report_book_name = book
        else:
            undo_path = self.undo_collection.undo_path(book.FilePath)
            current_path = book.FilePath
            self.report_book_name = book.FilePath

        self.profile = self.profiles[self.undo_collection.profile(current_path)]

        if not File.Exists(current_path):
            self.logger.Add("Failed", self.report_book_name, "The file does not exist")
            return MoveResult.Failed


        result = self.check_path_problems(current_path, undo_path)
        if result is not None:
            return result

        #Don't need to check path to long

        #Duplicate
        if File.Exists(undo_path):
            self.HeldDuplicateBooks[book] = undo_path
            return MoveResult.Duplicate

        #Create here because needed for cleaning directories later
        old_folder_path = FileInfo(current_path).DirectoryName

        result = self.create_folder(old_folder_path, book)
        if result is not MoveResult.Success:
            return result

        result = self.move_book(book, undo_path)

        if self.profile.RemoveEmptyFolder:
            self.remove_empty_folders(DirectoryInfo(old_folder_path))
            self.remove_empty_folders(FileInfo(undo_path).Directory)

        return result


    def process_duplicate_book(self, book, undo_path):

        if type(book) == str:
            current_path = book
            self.report_book_name = book
        else:
            current_path = book.FilePath
            self.report_book_name = book.FilePath

        self.profile = self.profiles[self.undo_collection.profile(current_path)]

        #Since the duplicate is checked for last in the orginal process_book function there is no need to check for path errors.
        if File.Exists(undo_path):

            #Find the existing book if it occurs in the library
            oldbook = self.find_duplicate_book(undo_path)
            if oldbook == None:
                oldbook = FileInfo(undo_path)

            rename_path = self.create_rename_path(undo_path)
            
            rename_filename = Path.GetFileName(rename_path)
        
            if not self.always_do_duplicate_action:
                result = self.form.Invoke(Func[type(self.profile), type(book), type(oldbook), str, int, DuplicateResult](self.form.ShowDuplicateForm), System.Array[object]([self.profile, book, oldbook, rename_filename, self.HeldDuplicateCount]))

                self.duplicate_action = result.action

                if result.always_do_action:
                    self.always_do_duplicate_action = True

            if self.duplicate_action is DuplicateAction.Cancel:
                self.logger.Add("Skipped", self.report_book_name, "A file already exists at: " + undo_path + " and the user declined to overwrite it")
                return MoveResult.Skipped

            elif self.duplicate_action is DuplicateAction.Rename:
                #Check if the created path is too long
                if len(rename_path) > 259:
                    result = self.form.Invoke(Func[str, object](self.get_smaller_path), System.Array[System.Object]([rename_path]))
                    if result is None:
                        self.logger.Add("Skipped", self.report_book_name, "The path was too long and the user skipped shortening it")
                        return MoveResult.Skipped

                    return self.process_duplicate_book(book, result)

                return self.process_duplicate_book(book, rename_path)

            elif self.duplicate_action is DuplicateAction.Overwrite:
                try:
                    if self.profile.CopyReadPercentage and type(oldbook) is not FileInfo:
                        book.LastPageRead = oldbook.LastPageRead
                    FileIO.FileSystem.DeleteFile(undo_path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                except Exception, ex:
                    self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + undo_path + ". The error was: " + str(ex))
                    return MoveResult.Failed

                #Since we are only working with images there is no need to remove a book from the library
                if type(oldbook) is not FileInfo:
                        ComicRack.App.RemoveBook(oldbook)
                
                return self.process_duplicate_book(book, undo_path)

        old_folder_path = FileInfo(current_path).DirectoryName
        
        result = self.move_book(book, undo_path)

        if self.profile.RemoveEmptyFolder:
            self.remove_empty_folders(DirectoryInfo(old_folder_path))
            self.remove_empty_folders(FileInfo(undo_path).Directory)

        return result


    def get_library_books(self):
        """Finds which books are in the ComicRack Library.
        Returns a list of books which are in the library and a list of file paths that are not in the library.
        """
        books = []
        found = set()
        #Note possible that the book in not in the library.
        for b in ComicRack.App.GetLibraryBooks():
            if b.FilePath in self.undo_collection._current_paths:
                books.append(b)
                found.add(b.FilePath)
        all = set(self.undo_collection._current_paths)
        notfound = all.difference(found)
        return books, list(notfound)


    def check_path_problems(self, current_path, undo_path):

        if undo_path == current_path:
            self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
            return MoveResult.Skipped

        #In some cases the filepath is the same but has different cases. The FileInfo object dosn't catch this but the File.Move function
        #Thinks that it is a duplicate.
        if undo_path.lower() == current_path.lower():
            #In that case, better rename it to the correct case
            book.RenameFile(file_name)
            self.logger.Add("Skipped", self.report_book_name, "The book is already located at the calculated path")
            return MoveResult.Skipped

        return None


    def move_book(self, book, path):
        #Finally actually move the book
        try:
            if type(book) is str:
                File.Move(book, path)
            else:
                File.Move(book.FilePath, path)
                book.FilePath = path

            return MoveResult.Success

        except Exception, ex:
            self.logger.Add("Failed", self.report_book_name, "because an error occured. The error was: " + str(ex))
            return MoveResult.Failed


    def MoveBook(self, book, path):
        """
        Book is the book to be moved
        
        Path is the absolute destination with extension
        
        """
        
        #Keep the current book path, just incase
        if type(book) == str:
            oldpath = book
        else:
            oldpath = book.FilePath
        
    
        if File.Exists(path):
            
            #Find the existing book if it occurs in the library
            oldbook = self.FindDuplicate(path)

            if oldbook == None:
                oldbook = FileInfo(path)

            if type(book) == str:
                dupbook = FileInfo(book)
            else:
                dupbook = book
            
            if not self.AlwaysDoAction:
                
                renamepath = self.CreateRenamePath(path)

                renamefilename = FileInfo(renamepath).Name

                #Ask the user:
                result = self.form.Invoke(Func[type(dupbook), type(oldbook), str, int, list](self.form.ShowDuplicateForm), System.Array[object]([dupbook, oldbook, renamefilename, len(self.HeldDuplicateBooks)]))
                self.Action = result[0]
                if result[1] == True:
                    #User checked always do this opperation
                    self.AlwaysDoAction = True
            
            if self.Action == DuplicateResult.Cancel:
                self.report.Append("\n\nSkipped moving\n%s\nbecause a file already exists at\n%s\nand the user declined to overwrite it." % (oldpath, path))
                return MoveResult.Skipped
            
            elif self.Action == DuplicateResult.Rename:
                return self.MoveBook(book, renamepath)
            
            elif self.Action == DuplicateResult.Overwrite:

                try:
                    FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    #File.Delete(path)
                except Exception, ex:
                    self.report.Append("\n\nFailed to delete\n%s. The error was: %s. File\n%s\nwas not moved." % (path, ex, oldpath))
                    return MoveResult.Failed
                
                #if the book is in the library, delete it
                if oldbook:
                    ComicRack.App.RemoveBook(oldbook)
                
                return self.MoveBook(book, path)
            
        
        #Finally actually move the book
        try:
            File.Move(oldpath, path)
            if not type(book) == str:
                book.FilePath = path
            return MoveResult.Success
        except PathTooLongException:
            #Too long path. Add a way to ask what path you want to use instead
            print "path was to long"
            print oldpath
            print path
            return MoveResult.Failed
        except Exception, ex:
            self.report.Append("\n\nFailed to move\n%s\nbecause an error occured. The error was: %s. The book was not moved." % (book.FilePath, ex))
            return MoveResult.Failed



def CopyData(book, newbook):
    """This helper function copies all revevent metadata from a book to another book"""
    list = ["Series", "Number", "Count", "Month", "Year", "Format", "Title", "Publisher", "AlternateSeries", "AlternateNumber", "AlternateCount",
            "Imprint", "Writer", "Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "AgeRating", "Manga", "LanguageISO", "BlackAndWhite",
            "Genre", "Tags", "SeriesComplete", "Summary", "Characters", "Teams", "Locations", "Notes", "Web", "ScanInformation"]

    for i in list:
        setattr(newbook, i, getattr(book, i))



class PathMaker(object):
    """A class to create directory and file paths from the passed book.
    
    Some of the functions are based on functions in wadegiles's guided rename script. (c) wadegiles. Most have been heavily modified.
    """

    template_to_field = {"series" : "ShadowSeries", "number" : "ShadowNumber", "count" : "ShadowCount", "Day" : "Day", "ReleasedDate": "ReleasedTime",
                         "AddedDate" : "AddedTime", "EndYear" : "EndYear", "EndMonth" : "EndMonth", "EndMonth#" : "EndMonth",
                         "month" : "Month", "month#" : "Month", "year" : "ShadowYear", "imprint" : "Imprint", "publisher" : "Publisher",
                         "altSeries" : "AlternateSeries", "altNumber" : "AlternateNumber", "altCount" : "AlternateCount",
                         "volume" : "ShadowVolume", "title" : "ShadowTitle", "ageRating" : "AgeRating", "language" : "LanguageAsText",
                         "format" : "ShadowFormat", "startyear" : "StartYear", "writer" : "Writer", "tags" : "Tags", "genre" : "Genre",
                         "characters" : "Characters", "altSeries" : "AlternateSeries", "teams" : "Teams", "scaninfo" : "ScanInformation",
                         "manga" : "Manga", "seriesComplete" : "SeriesComplete", "first" : "FirstLetter", "read" : "ReadPercentage",
                         "counter" : "Counter", "startmonth" : "StartMonth", "startmonth#" : "StartMonth", "colorist" : "Colorist", "coverartist" : "CoverArtist",
                         "editor" : "Editor", "inker" : "Inker", "letterer" : "Letterer", "locations" : "Locations", "penciller" : "Penciller", "storyarc" : "StoryArc",
                         "seriesgroup" : "SeriesGroup", "maincharacter" : "MainCharacterOrTeam", "firstissuenumber" : "FirstIssueNumber", "lastissuenumber" : "LastIssueNumber"}

    template_regex = re.compile("{(?P<prefix>[^{}<]*)<(?P<name>[^\d\s(>]*)(?P<args>\d*|(?:\([^)]*\))*)>(?P<postfix>[^{}]*)}")

    yes_no_fields = ["Manga", "SeriesComplete"]


    def __init__(self, parentform, profile):

        self._counter = None
        self.profile = profile
        self.failed_fields = []
        self.failed = False

        #Need to store the parent form so it can use the muilt select form
        self.form = parentform


    def make_path(self, book, folder_template, file_template):
        """Creates a path from a book with the given folder and file templates.
        Returns folder path, file name, and a bool if any values were empty.
        """
        self.book = book

        self.failed = False
        self.failed_fields = []


        #if self.profile.FailEmptyValues:
        #    for field in self.profile.FailedFields:
        #        if getattr(self.book, field) in ("", -1, MangaYesNo.Unknown, YesNo.Unknown):
        #            self.failed_fields.append(field)
        #            self.failed = True
            #
            #if self.failed and not self.profile.MoveFailed:
            #    return book.FileDirectory, book.FileNameWithExtension, True

        #Do filename first so that if MoveFailed is true the base folder is used correctly.
        file_path = self.book.FileNameWithExtension
        if self.profile.UseFileName:
            file_path = self.make_file_name(file_template)

        folder_path = book.FileDirectory
        if self.profile.UseFolder:

            folder_path = self.make_folder_path(folder_template)

        if self.failed and not self.profile.MoveFailed:
            return book.FileDirectory, book.FileNameWithExtension, True
            
        return folder_path, file_path, self.failed
        
 
    def make_folder_path(self, template):
        
        folder_path = ""

        template = template.strip()
        template = template.strip("\\")

        if template:
 
            rough_path = self.insert_fields_into_template(template)
            
            #Split into seperate directories for fixing empty paths and other problems.
            lines = rough_path.split("\\")
            
            for line in lines:
                if not line.strip():
                    line = self.profile.EmptyFolder
                line = self.replace_illegal_characters(line)
                #Fix for illegal periods at the end of folder names
                line = line.strip(".")
                folder_path = Path.Combine(folder_path, line.strip())
        
        if self.failed and self.profile.MoveFailed:
            folder_path = Path.Combine(self.profile.FailedFolder, folder_path)
        else:
            folder_path = Path.Combine(self.profile.BaseFolder, folder_path)


        if self.profile.ReplaceMultipleSpaces:
            folder_path = re.sub("\s\s+", " ", folder_path)

        return folder_path

    
    def make_file_name(self, template):
        """Creates file name with the template.

        template->The template to use.
        Returns->The created file name with extension and a bool if any values were empty.
        """
     
        file_name = self.insert_fields_into_template(template)
        file_name = file_name.strip()
        file_name = self.replace_illegal_characters(file_name)

        if not file_name:
            return ""


        extension = self.profile.FilelessFormat

        if self.book.FilePath:
            extension = FileInfo(self.book.FilePath).Extension

        #replace occurences of multiple spaces with a single space.
        if self.profile.ReplaceMultipleSpaces:
            file_name = re.sub("\s\s+", " ", file_name)


        return file_name + extension

    
    def insert_fields_into_template(self, template):
        """Replaces fields in the template with the correct field text."""
        self.invalid = 0
        while self.template_regex.search(template):
            if self.invalid == len(self.template_regex.findall(template)):
                break
            else:
                self.invalid = 0
                template = self.template_regex.sub(self.insert_field, template)

        return template


    def insert_field(self, match):
        """Replaces a regex match with the correct text. Returns the original match if something is not valid."""
        match_groups = match.groupdict()
        result = ""
        conditional = False
        inversion = False
        inversion_args = ""

        name = match_groups["name"]
        args = match_groups["args"]

        #Inversions
        if name.startswith("!"):
            inversion = True
            name = name.lstrip("!")
            #Inversions can optionally have args.
            if args:
                #Get the last arg and removing from the other args
                r = re.search("(\([^(]*\))$", args)
                if r is not None:
                    args = args[:-len(r.group(0))]
                    inversion_args = r.group(0)[1:-1]

        #Conditionals
        if name.startswith("?"):
            conditional = True
            name = name.lstrip("?")
            if args:
                #Get the last arg and removing from the other args
                r = re.search("(\([^(]*\))$", args)
                if r is None:
                    self.invalid += 1
                    return match.group(0)
                args = args[:-len(r.group(0))]
                conditional_args = r.group(0)[1:-1]
            else:
                self.invalid += 1
                return match.group(0)

        #Checking template names
        if name in self.template_to_field:
            field = self.template_to_field[name]
        else:
            self.invalid += 1
            return match.group(0)



        #Get the fields
        result = self.get_field_text(field, name, args)
        
        #Invalid field result (possibly wrong number of arguments)
        if result is None:
            self.invalid += 1
            return match.group(0)

        #Conditionals
        if conditional:
            #Regex conditional
            if conditional_args.startswith("!"):
                if conditional_args[1:] == "":
                    return ""
                #Insert prefix and suffix if there is a match.
                if re.match(conditional_args[1:], result) is not None:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""
            else:
                #Text argument. Insert if matching the result
                if result == conditional_args:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""

        #Inversions
        if inversion:
            if not inversion_args:
                #No args so only insert the prefix and suffix if the result is empty
                if not result:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""

            elif inversion_args.startswith("!"):
                #Regex so only insert the prefix and suffix if there is no matches
                if re.match(inversion_args[1:], result) is None:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""
            else:
                #Text to match to. Only insert if the result doesn't match the arg
                if result != inversion_args:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""



        #Empty results
        if not result:
            if self.profile.FailEmptyValues and field in self.profile.FailedFields:
                if field not in self.failed_fields:
                    self.failed_fields.append(field)
                self.failed = True

            if field in self.profile.EmptyData and self.profile.EmptyData[field]:
                return self.profile.EmptyData[field]
            else:
                return ""


        return match_groups["prefix"] + result + match_groups["postfix"]
    

    def get_field_text(self, field, template_name, args_match):
        """Gets the text of a field or operation.
        
        field->The name of the field to use.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        args_match->The args matched from the regex.

        Returns a string or None if an error occurs or the field isn't valid.
        """
        args = re.findall("\(([^)]*)\)", args_match)

        try:
            if field in ("StartYear", "StartMonth", "EndYear", "EndMonth"):
                return self.insert_start_value(field, args_match, template_name)

            elif field == "ReadPercentage" and len(args) == 3:
                return self.insert_read_percentage(args)

            elif field == "FirstLetter":
                if len(args) == 0:
                   return None
                elif args[0] in name_to_field or args[0] in field_to_name:
                    return self.insert_first_letter(args[0])

            elif field == "Counter" and len(args) == 3:
                return self.insert_counter(args)

            elif field == "FirstIssueNumber":
                return self.insert_first_issue_number(args_match)

            elif field == "LastIssueNumber":
                return self.insert_last_issue_number(args_match)

            elif type(getattr(self.book, field)) is System.DateTime:
                if args:
                    return self.insert_formated_datetime(args[0], field)
                else:
                    return self.insert_formated_datetime("", field)

            #Yes/no fields can have 1 or 2 args
            elif field in self.yes_no_fields and 0 < len(args) < 3:
                return self.insert_yes_no_field(field, args)

            elif len(args) == 2:
                return self.insert_multi_value_field(field, args)

            elif not args_match:
                return self.insert_text_field(field, template_name)

            elif args_match.isdigit():
                return self.insert_number_field(field, int(args_match))

            else:
                return None
        except (AttributeError, ValueError), ex:
            print ex
            return None


    def insert_text_field(self, field, template_name):
        """Get the string of any field and stips it of any illegal characters.

        field->The name of the field to get.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        Returns: The unicode string of the field."""

        if field == "Month" and not template_name.endswith("#"):
            return self.insert_month_as_name()

        text = getattr(self.book, field)

        if not text or text == -1:
            return ""

        return self.replace_illegal_characters(unicode(text))

    
    def insert_number_field(self, field, padding):
        """Get the padded value of a number field. Replaces illegal character in the number field.

        field->The name of the field to use.
        padding->integer of the ammount of padding to use.
        Returns-> Unicode string with the padded number value or empty string if the field is empty.
        """
        number = getattr(self.book, field)

        if number == -1:
            return ""

        if number == "":
            return ""

        if padding == 0:
            value = getattr(get_last_book(self.book), field)
            padding = len(str(value))

        return self.replace_illegal_characters(self.pad(number, padding))


    def insert_yes_no_field(self, field, args):
        """Gets a value using a yes/no field.

        field->The name of the field to check.
        args->A list with one or two items. 
              First item should be a string of what text to insert when the value is Yes.
              Second item (optional) should be a ! or null. ! means the text is inserted when the value is No.
        Returns the user text or an empty string."""
              

        text = args[0]

        no = False

        result = ""

        if len(args) == 2 and args[1] == "!":
            no = True

        field_value = getattr(self.book, field)

        if not no and field_value in (MangaYesNo.Yes, MangaYesNo.YesAndRightToLeft, YesNo.Yes):
            result = text

        elif no and field_value in (MangaYesNo.No, YesNo.No):
            result = text

        return self.replace_illegal_characters(result)


    def insert_read_percentage(self, args):
        """Get a value from the book's readpercentage.

        args should be a list with 3 items:
            1. Should be a string of the text to insert if the readpercentage matches the caclulations.
            2. Should be an operator: < > =
            3. Should be the percentage to match.
        Returns the user text or and empty string.
        """

        text = args[0]
        operator = args[1]
        percent = args[2]
        result = ""

        if operator == "=":
            if self.book.ReadPercentage == int(percent):
                result = text
        elif operator == ">":
            if self.book.ReadPercentage > int(percent):
                result = text
        elif operator == "<":
            if self.book.ReadPercentage < int(percent):
                result = text

        return self.replace_illegal_characters(result)


    def insert_first_letter(self, field):
        """Gets the first letter of a field not counting articles.

        field->The name of the field to find the first letter of.
        Returns a single character string.
        """
        if field in name_to_field:
            field = name_to_field[field]

        property = unicode(getattr(self.book, field))

        result = ""

        match_result = re.match(r"(?:(?:the|a|an|de|het|een|die|der|das|des|dem|der|ein|eines|einer|einen|la|le|l'|les|un|une|el|las|los|las|un|una|unos|unas|o|os|um|uma|uns|umas|en|et|il|lo|uno|gli)\s+)?(?P<letter>.).+", property, re.I)

        if match_result:
            result = match_result.group("letter").capitalize()

        return self.replace_illegal_characters(result)


    def insert_counter(self, args):
        """Get the (padded) next number in a counter given a start and increment.

        args should be a list with 3 items:
            1. An integer with the start number.
            2. An integer of the increment.
            3. The amount of padding to use.
        Returns a string of the number.
        """
        start = int(args[0])
        increment = int(args[1])
        pad = args[2]

        result = ""

        if pad == "":
            pad = 0

        else:
            pad = int(pad)

        if self._counter is None:
            self._counter = start
            result = self.pad(self._counter, pad)

        else:
            self._counter = self._counter + increment
            result = self.pad(self._counter, pad)

        return result


    def insert_month_as_name(self):
        """Get the month name from a month number."""
        month_number = self.book.Month

        if month_number in self.profile.Months:
            return self.replace_illegal_characters(self.profile.Months[month_number])

        return ""


    def insert_start_value(self, field, args, template_name):
        """Get the value for StartYear or StartMonth.
        
        field->The name of the field to get.
        args->args is padding for the startmonth as number. Can be null.
        template_name->The name of the field as entered in the template. This is for keeping startmonth and startmonth# seperate.

        returns the string of the field or an empty string.
        """

        if field in ("StartYear", "EndYear"):
            return self.insert_start_year(field == "EndYear")

        elif field in ("StartMonth", "EndMonth"):
            return self.insert_start_month(args, template_name, field == "EndMonth")


    def insert_start_year(self, end=False):
        """Gets the start year of the earliest book in the series of the current issue."""
        if end:
            year = get_last_book(self.book).ShadowYear
        else:
            year = get_earliest_book(self.book).ShadowYear

        if year == -1:
            return ""

        return self.replace_illegal_characters(unicode(year))


    def insert_start_month(self, args, template_name, end=False):
        """Gets the start month of from the earliest issues in the series.
        Depending on the template_name the month as name or month as a number is inserted.
        args->can be either null or a number.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        Returns a string."""
        if end:
            month = get_last_book(self.book).Month
        else:
            month = get_earliest_book(self.book).Month

        if month == -1:
            return ""
        if template_name.endswith('#'):
            if args and args.isdigit():
                month = self.pad(month, int(args))
            return self.replace_illegal_characters(unicode(month))

        else:
            if month in self.profile.Months:
                return self.replace_illegal_characters(self.profile.Months[month])
            return ""

    def insert_formated_datetime(self, time_format, field):

        date_time = getattr(self.book, field)
        if time_format:
            return date_time.ToString(time_format)
        return date_time.ToString()

    def insert_first_issue_number(self, padding):
        """
        padding is the padding used, can be none.
        """
        number = get_earliest_book(self.book).ShadowNumber
        
        if padding is not None and padding.isdigit():
            return self.replace_illegal_characters(self.pad(number, int(padding)))
        else:
            return self.replace_illegal_characters(number)


    def insert_last_issue_number(self, padding):
        """
        padding is the padding used, can be none.
        """
        number = get_last_book(self.book).ShadowNumber
        
        if padding is not None and padding.isdigit():
            return self.replace_illegal_characters(self.pad(number, int(padding)))
        else:
            return self.replace_illegal_characters(number)


    def insert_multi_value_field(self, field, args):
        """Gets the value from a multiple value field.

        args should be a list of two items:
            1. Should be the seperator between the items.
            2. A string of either issues or series.
        Returns a string containing the values.
        """
        seperator = args[0]
        mode = args[1]

        if mode == "series":
            return self.insert_multi_value_series(field, seperator)

        elif mode == "issue":
            return self.insert_multi_value_issue(field, seperator)

        else:
            return None


    def insert_multi_value_issue(self, field, seperator):
        """
        Finds which values to use in a multiple value field per issue. Asks the user which values to use and offers
        options to save that selection for following issues. Checks the stored selections and only asks the user if required.

        field->The string name of the field to use.
        seperator->The string to seperate every value with.
        Returns a string of the values or an empty string if no values are found or selected.
        """

        #field_dict stores the alwaysused collections and also the selected values for each issue used.
        try:
            field_dict = getattr(self, field)
        except AttributeError:
            field_dict = {}
            setattr(self, field, field_dict)


        index = self.book.Publisher + self.book.ShadowSeries + str(self.book.ShadowVolume) + self.book.ShadowNumber
        booktext = self.book.ShadowSeries + " vol. " + str(self.book.ShadowVolume) + " #" + self.book.ShadowNumber


        if index in field_dict:
            #This particular issue has already been done.
            result = field_dict[index]
            return self.make_multi_value_issue_string(result.Selection, seperator, result.Folder)


        if not getattr(self.book, field).strip():
            field_dict[index] = MultiValueSelectionFormResult([])
            return ""


        try:
            always_used_values = getattr(self, field + "AlwaysUse")
        except AttributeError:
            always_used_values = []
            setattr(self, field + "AlwaysUse", always_used_values)


        values = [item.strip() for item in getattr(self.book, field).split(",")]


        if len(values) == 1 and self.profile.DontAskWhenMultiOne:
            return self.replace_illegal_characters(values[0])


        selected_values = []

        #Find which items are set to always use
        for list_of_always_used_values in sorted(always_used_values, key=len, reverse=True):
            count = 0
            for value in list_of_always_used_values:
                if value in values:
                    count +=1

            if count == len(list_of_always_used_values):

                if list_of_always_used_values.do_not_ask:
                    return self.make_multi_value_issue_string(list_of_always_used_values, seperator, list_of_always_used_values.use_folder_seperator)

                selected_values = list_of_always_used_values[:]
                break

        if self.form.InvokeRequired:
            result = self.form.Invoke(Func[MultiValueSelectionFormArgs, MultiValueSelectionFormResult](self.show_multi_value_selection_form), System.Array[object]([MultiValueSelectionFormArgs(values, selected_values, field, booktext, False)]))

        else:
            result = self.show_multi_value_selection_form(MultiValueSelectionFormArgs(values, selected_values, field, booktext, False))

        field_dict[index] = result

        if len(result.Selection) == 0:
            return ""

        if result.AlwaysUse:
            always_used_values.append(MultiValueAlwaysUsedValues(result.AlwaysUseDontAsk, result.Folder, result.Selection))
        
        return self.make_multi_value_issue_string(result.Selection, seperator, result.Folder)


    def make_multi_value_issue_string(self, values, seperator, usefolder):
        #When folders are being used we need to make sure there are no "\" characters in any of the values or it will mess up the folders. 
        if usefolder:
            seperator = self.replace_illegal_characters(seperator)
            seperator += "\\"
            values = [self.replace_illegal_characters(value) for value in values]

        
            return seperator.join(values)

        return self.replace_illegal_characters(seperator.join(values))


    def insert_multi_value_series(self, field, seperator):
        """
        Finding a multiple value field via a series operation will find all the possible multiple values of the field
        from every issue of the series in the library. The user can then pick which values to use and those chosen values will
        be used for every issue encountered.

        The user can choose to use every chossen value for the every issue in the series, even if the issue doesn't have that value.
        Or the user can choose to only use the value if it is in the issue.
        
        field->The string name of the field to use.
        seperator->The string to seperate every value with
        """

        #Chosen values for series are stored in the field_dict.
        try:
            field_dict = getattr(self, field)
        except AttributeError:
            field_dict = {}
            setattr(self, field, field_dict)


        index = self.book.Publisher + self.book.ShadowSeries + str(self.book.ShadowVolume)
        booktext = self.book.ShadowSeries + " vol. " + str(self.book.ShadowVolume) 


        #See if this series has been done before:
        if index in field_dict:
            return self.make_multi_value_series_string(field_dict[index], seperator, field)



        values = self.get_all_multi_values_from_series(field)


        if not values:
            field_dict[index] = MultiValueSelectionFormResult([])
            return ""


        #Since this can be shown from the configform...
        if self.form.InvokeRequired:
            result = self.form.Invoke(Func[MultiValueSelectionFormArgs, MultiValueSelectionFormResult](self.show_multi_value_selection_form), System.Array[object]([MultiValueSelectionFormArgs(values, [], field, booktext, True)]))

        else:
            result = self.show_multi_value_selection_form(MultiValueSelectionFormArgs(values, [], field, booktext, True))

        field_dict[index] = result

        return self.make_multi_value_series_string(result, seperator, field)


    def make_multi_value_series_string(self, selection_result, seperator, field):
        """
        Makes the correctly formated inserted string based on the input values:

        selection_result is a SelectionFormResult object
        seperator is the string seperator
        field is the correct name of the field to find from the book
        """

        #If using folder sperator we need to make sure that there are no "\" characters in any of the fields.
        if selection_result.Folder:
            seperator = self.replace_illegal_characters(seperator)
            seperator += "\\"
            values = [self.replace_illegal_characters(value.strip()) for value in getattr(self.book, field).split(",")]
            selection = [self.replace_illegal_characters(value) for value in selection_result.Selection]
        else:
            values = [value.strip() for value in getattr(self.book, field).split(",")]
            selection = selection_result.Selection

        #Using every issue
        if selection_result.EveryIssue:
            result = seperator.join(selection)

        #Not using every issue so just use the ones the particular issue has.
        else:
            items_to_use = [item for item in selection if item in values]
            result = seperator.join(items_to_use)
        
        if selection_result.Folder:
            return result

        return self.replace_illegal_characters(result)


    def replace_illegal_characters(self, text):
        """Replaces illegal path characters in a string and retures the cleaned string."""
        for illegal_character in sorted(self.profile.IllegalCharacters.keys(), key=len, reverse=True):
            text = text.replace(illegal_character, self.profile.IllegalCharacters[illegal_character])

        return text

    
    def pad(self, value, padding):
        #value as string, padding as int
        remainder = ""

        try:
            numberValue = float(value)
        except ValueError:
            return value

        if type(value) == int:
            value = str(value)

        if numberValue >= 0:
            #To make sure that the item is padded correctly when a decimal such as 7.1
            if value.Contains("."):
                value, remainder = value.split(".")
                remainder = "." + remainder
            return value.PadLeft(padding, '0') + remainder
        else:
            value = value[1:]
            #To make sure that the item is padded correctly when a decimal such as 7.1
            if value.Contains("."):
                value, remainder = value.split(".")
                remainder = "." + remainder
            value = value.PadLeft(padding, '0')
            return '-' + value + remainder

    
    def show_multi_value_selection_form(self, args):
        form = MultiValueSelectionForm(args)
        form.ShowDialog()
        result = form.GetResults()
        form.Dispose()
        return result


    def get_all_multi_values_from_series(self, field):
        """
        Helper to get all the unique values from a multivale field in an entire series.

        field is the name of the field to find values from.
        Returns a list of the values.
        """
        allbooks = ComicRack.App.GetLibraryBooks()
        #using a set to avoid duplicate entries
        results = set()
        for book in allbooks:
            if book.ShadowSeries == self.book.ShadowSeries and book.ShadowVolume == self.book.ShadowVolume and book.Publisher == self.book.Publisher:
                results.update([value.strip() for value in getattr(book, field).split(",") if value.strip()])

        return list(results)



class MultiValueAlwaysUsedValues(list):

    def __init__(self, do_not_ask=False, folder=False, values=None):
        self.do_not_ask = do_not_ask
        self.use_folder_seperator = folder
        if values is None:
            values = []
        list.__init__(self, values)