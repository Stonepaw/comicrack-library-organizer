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
import re

clr.AddReference("Microsoft.VisualBasic")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

import System
from Microsoft.VisualBasic import FileIO
from System import Func, ArgumentException, ArgumentNullException, NotSupportedException
from System.Drawing.Imaging import ImageFormat
from System.IO import Path, File, FileInfo, DirectoryInfo, Directory, IOException, PathTooLongException, DirectoryNotFoundException
from System.Windows.Forms import DialogResult

from loforms import PathTooLongForm
from locommon import Mode, check_metadata_rules, check_excluded_folders, UNDOFILE, UndoCollection
from loduplicate import DuplicateResult, DuplicateAction
from pathmaker import PathMaker

class MoveResult(object):
    Success = 1
    Failed = 2
    Skipped = 3
    Duplicate = 4
    FailedMove = 5
    DuplicateDifferentExtension= 6


class BookToMove(object):

    def __init__(self, book, path, index, failed_fields):
        self.book = book
        self.path = path
        self.profile_index = index
        self.failed_fields = failed_fields
        self.duplicate_different_extension = False



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

            if result is MoveResult.DuplicateDifferentExtension:
                count -= 1
                progress -= percentage
                self.HeldDuplicateBooks.append(book)
                book.duplicate_different_extension = True
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
                self.logger.Add("Cancelled", str(len(books_to_move) - count) + " operations", "User cancelled the script")
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

        #Check if the file is a Duplicate
        if File.Exists(full_path) or full_path in self.MovedBooks:
            return MoveResult.Duplicate
        
        #Create here because needed for cleaning directories later
        old_folder_path = book.FileDirectory
        folder_path = Path.GetDirectoryName(full_path)

        #Duplicates with different extensions
        if self.profile.DifferentExtensionsAreDuplicates:
            if self.get_same_files_with_different_ext(full_path):
                return MoveResult.DuplicateDifferentExtension

        #Create the new path
        result = self.create_folder(folder_path, book)
        if result is not MoveResult.Success:
            return result
        
        #If fileless book
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


    def get_same_files_with_different_ext(self, full_path):
        d = DirectoryInfo(Path.GetDirectoryName(full_path))
        if d.Exists:
            file_name = Path.GetFileNameWithoutExtension(full_path)                
            files = d.GetFiles(file_name + ".*")
            return files
        return []

    def process_duplicate_book(self, book_to_move):

        book = book_to_move.book
        full_path = book_to_move.path

        if book.FilePath:
            self.report_book_name = book.FilePath
        else:
            self.report_book_name = book.Caption

        #Since the duplicate is checked for last in the orginal process_book function there is no need to check for path errors.
        if File.Exists(full_path) or full_path in self.MovedBooks or book_to_move.duplicate_different_extension:

            #Find the existing book if it occurs in the library
            dupbook = self.find_duplicate_book(full_path, book_to_move.duplicate_different_extension)

            #Book does not exist in the library
            if dupbook == None:
                #If we are checking for a different extension then we have to 
                #find what that path is
                if book_to_move.duplicate_different_extension:
                    files = self.get_same_files_with_different_ext(full_path)
                    if files:
                        dupbook = files[0]
                    else:
                        return MoveResult.Failed
                #If not a book with a different extension then just use the existing file.
                else:
                    dupbook = FileInfo(full_path)
                dup_path=dupbook.FullPath
            else:
                dup_path = dupbook.FilePath
                
            if book_to_move.duplicate_different_extension:
                rename_path = full_path
                rename_filename = Path.GetFileName(rename_path)
            else:
                rename_path = self.create_rename_path(full_path)
                rename_filename = Path.GetFileName(rename_path)
        
            if not self.always_do_duplicate_action or book_to_move.duplicate_different_extension:
                result = self.form.Invoke(Func[type(self.profile), type(book), type(dupbook), str, int, DuplicateResult](self.form.ShowDuplicateForm), System.Array[object]([self.profile, book, dupbook, rename_filename, self.HeldDuplicateCount]))

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
                book_to_move.duplicate_different_extension = False
                if len(rename_path) > 259:
                    result = self.form.Invoke(Func[str, object](self.get_smaller_path), System.Array[System.Object]([rename_path]))
                    if result is None:
                        self.logger.Add("Skipped", self.report_book_name, "The path was too long and the user skipped shortening it")
                        return MoveResult.Skipped

                    return self.process_duplicate_book(BookToMove(book, result, book_to_move.profile_index, book_to_move.failed_fields))
                return self.process_duplicate_book(BookToMove(book, rename_path, book_to_move.profile_index, book_to_move.failed_fields))

            elif self.duplicate_action is DuplicateAction.Overwrite:
                book_to_move.duplicate_different_extension = False
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
                        if self.profile.CopyReadPercentage and type(dupbook) is not FileInfo:
                            book.LastPageRead = dupbook.LastPageRead
                        if type(dupbook) is FileInfo:
                            FileIO.FileSystem.DeleteFile(dupbook.FullPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        else:
                            FileIO.FileSystem.DeleteFile(dupbook.FilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                except Exception, ex:
                    self.logger.Add("Failed", self.report_book_name, "Failed to overwrite " + full_path + ". The error was: " + str(ex))
                    return MoveResult.Failed

                #Since we are only working with images there is no need to remove a book from the library
                if book.FilePath and type(dupbook) is not FileInfo:
                        ComicRack.App.RemoveBook(dupbook)

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


    def find_duplicate_book(self, path, ignore_extension=False):
        """Trys to find ComicBook in the ComicRack library by checking the path
            against all books in the library.

        Args:
            path: The string path to search in the library for.
            ignore_extension: If set to true the function will search for the 
                same path and file name ignoring the extension.

        Returns:
            A ComicBook object if located in the library, None otherwise.
        """
        path = path.lower()
        path_no_ext = path[:-len(Path.GetExtension(path))]
        
        for book in ComicRack.App.GetLibraryBooks():
            #Fix for different captilization in path names.
            if book.FilePath.lower() == path:
                return book
            elif ignore_extension:
                if book.FilePath.lower()[:-len(Path.GetExtension(book.FilePath))] == path_no_ext:
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
        Recursively deletes directories until an non-empty directory is found or the directory is in the excluded list
        
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
    """This helper function copies all relevant metadata from a book to another book"""
    list = ["Series", "Number", "Count", "Month", "Year", "Format", "Title", "Publisher", "AlternateSeries", "AlternateNumber", "AlternateCount",
            "Imprint", "Writer", "Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "AgeRating", "Manga", "LanguageISO", "BlackAndWhite",
            "Genre", "Tags", "SeriesComplete", "Summary", "Characters", "Teams", "Locations", "Notes", "Web", "ScanInformation"]

    for i in list:
        setattr(newbook, i, getattr(book, i))