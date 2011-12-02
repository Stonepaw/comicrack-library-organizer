"""
lobookmover.py

This contains all the book moving fuction the script uses. Also the path generator

Version 1.7.17
	
	Fixed error with fileless images
	Added Alternate Series into the multi selection

Copyright Stonepaw 2011. Some code copyright wadegiles Anyone is free to use code from this file as long as credit is given.
"""


import clr

import re

import System

from System import Func, Action

from System.Text import StringBuilder

from System.IO import Path, File, FileInfo, DirectoryInfo, Directory

import loforms

from loforms import PathTooLongForm, SelectionFormArgs, SelectionFormResult, SelectionForm

import locommon

from locommon import Mode

from loduplicate import DuplicateResult, DuplicateForm


import lologger

clr.AddReference("System.Drawing")
from System.Drawing.Imaging import ImageFormat

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult

clr.AddReferenceByPartialName('ComicRack.Engine')
from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

clr.AddReference("Microsoft.VisualBasic")

from Microsoft.VisualBasic import FileIO



class MoveResult:
	Success = 1
	Failed = 2
	Skipped = 3
	

class OverwriteAction:
	Overwrite = 1
	Cancel = 2
	Rename = 3

class BookMover(object):
	
	def __init__(self, worker, form, books, settings):
		self.worker = worker
		self.form = form
		self.books = books
		self.settings = settings
		self.logger = lologger.logger()
		self.pathmaker = PathMaker(form)
		
		#Hold books that are duplicates so they can be all asked at the end.

		self.HeldDuplicateBooks = []
		
		#These variables are for when the script is in test mode

		self.CreatedPaths = []

		self.MovedBooks = []

		#This hold a list of the book moved and is saved in undo.txt for the undo script
		#Dict with the key of the newpath and value of the old path
		self.MovedBooksUndo = {}
		
		self.AlwaysDoAction = False
		self.Action = None
		
		if self.settings.Mode == Mode.Copy:
			self.modetext = "copy"
			self.modetextplural = "copying"
			self.modetextpast = "copied"

		if self.settings.Mode == Mode.Move:
			self.modetext = "move"
			self.modetextplural = "moving"
			self.modetextpast = "moved"

		if self.settings.Mode == Mode.Test:
			self.modetext = "move (simulated)"
			self.modetextplural = "moving (simulated)"
			self.modetextpast = "moved (simulated)"

		
	def MoveBooks(self):
		
		#Variable to keep track of progress
		count = 0
		skipped = 0
		failed = 0
		success = 0
		
		for book in self.books:

			#Put the count to the current book being worked on ie book 1 for first loop
			count += 1 
			
			
			#Before making the paths check if the book needs to be moved at all
			
			if self.worker.CancellationPending:
				#User pressed cancel
				skipped = len(self.books) - success - failed
				self.logger.Add("Canceled", str(skipped) + " files", "User cancelled the script")
				break
			

			#Check if the file exists if it should exist

			if not File.Exists(book.FilePath) and book.FilePath != "":
				self.logger.Add("Failed", book.FilePath, "The file does not exist")
				#self.report.Append("\n\nFailed to %s\n%s\nbecause the file does not exist." % (self.modetext, book.FilePath))
				failed += 1
				self.worker.ReportProgress(count)
				continue
			
			#Check the exclude paths first

			if ExcludePath(book, self.settings.ExcludeFolders):
				skipped += 1
				#self.report.Append("\n\nSkipped %s\n%s\nbecause it is located in an excluded path." % (self.modetextplural, book.FilePath))
				self.logger.Add("Skipped", book.FilePath, "The book is located in an excluded path")
				self.worker.ReportProgress(count)
				continue
			
			
			
			# Metadata Exclude Rules
			"""
			Possible results are True or False
				True is when the book should not be moved
				False is when the book should be moved
				
				The function does the checking for the correct mode and returns the correct move/don't move result
			"""
			if ExcludeMeta(book, self.settings.ExcludeRules, self.settings.ExcludeOperator, self.settings.ExcludeMode):
				skipped += 1
				self.logger.Add("Skipped", book.FilePath, "The book qualified under the exclude rules")
				#self.report.Append("\n\nSkipped %s\n%s \nbecause it qualified under the exclude rules." % (self.modetextplural, book.FilePath))
				self.worker.ReportProgress(count)
				continue


			
			# Fileless and not creating fileless images

			if book.FilePath == "" and self.settings.MoveFileless == False:
				#Fileless comic
				skipped += 1
				self.logger.Add("Skipped", book.ShadowSeries + " " + str(book.ShadowNumber), "The book is fileless and fileless images are not being created")
				#self.report.Append("\n\nSkipped %s book\n%s %s\nbecause it is a fileless book." % (self.modetextplural, book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			
			# Fileless and moving but no custom thumbnail
			
			if book.FilePath == "" and self.settings.MoveFileless and not book.CustomThumbnailKey:
				#Fileless comic
				skipped += 1
				self.logger.Add("Failed", book.ShadowSeries + " " + str(book.ShadowNumber), "The fileless book does not have a custom thumbnail")
				#self.report.Append("\n\nSkipped creating image for filess book\n%s %s\nbecause it does not have a custom thumbnail." % (book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			# Create the path from the template seperated as directory path and the file path
			# The directory and the file path are kept seperate because they are needed individually for checks later.
			
			dirpath, filepath = self.CreatePath(book)

			#Combine them

			fullpath = Path.Combine(dirpath, filepath)

		
			# Empty file name

			if filepath == "":
				failed += 1
				self.logger("Failed", book.FilePath, "The created filename was empty")
				#self.report.Append("\n\nFailed to %s\n%s\nbecause the filename created was empty. The book was not moved" % (self.modetext, book.FilePath))
				self.worker.ReportProgress(count)
				continue


			# Book is already located at that path

			if fullpath == book.FilePath:
				self.logger.Add("Skipped", book.FilePath, "The book is already located at the calculated path")
				#self.report.Append("\n\nSkipped %s book\n%s\nbecause it is already located at the calculated path." % (self.modetextplural, book.FilePath))
				skipped += 1
				self.worker.ReportProgress(count)
				continue

			#In some cases the new path and the book path are the same but have different cases.
			#This can cause inccorect behaviour when looking at duplicates
			if fullpath.lower() == book.FilePath.lower():

				#In that case, better rename it to the correct case
				if self.settings.Mode == Mode.Test:
					self.logger.Add("Renaming", book.FilePath, "to: " + fullpath)
				else:
					book.RenameFile(filepath)

				success += 1
				self.worker.ReportProgress(count)
				continue


			# Check for too long path error

			if len(fullpath) > 259:
				result = self.form.Invoke(Func[str, object](self.GetSmallerPath), System.Array[System.Object]([fullpath]))
				if result == None:
					skipped += 1
					self.logger.Add("Skipped", book.FilePath, "The calculated path was too long and the user skipped shortening it")
					#self.report.Append("\n\nSkipped %s book\n%s\nbecause the path was too long and the user skipped shortening it." % (self.modetextplural, book.FilePath))
					self.worker.ReportProgress(count)
					continue

				fullpath = result
				f = FileInfo(fullpath)
				dirpath = f.DirectoryName
				filepath = f.Name



			# Duplicate. Hold for later

			if File.Exists(fullpath) or fullpath in self.MovedBooks:
				self.HeldDuplicateBooks.append(book)
				count -= 1
				continue


			
			#Create here because needed for cleaning directories later
			oldpath = book.FileDirectory
			
			
			# Create the directory

			if not Directory.Exists(dirpath):
				try:
					if not self.settings.Mode == Mode.Test:
						Directory.CreateDirectory(dirpath)
					else:
						if not dirpath in self.CreatedPaths:
							self.logger.Add("Created Folder", dirpath)
							#self.report.Append("\n\nCreating Folder:\n%s" % (dirpath))
							self.CreatedPaths.append(dirpath)

				except Exception, ex:
					failed += 1
					self.logger.Add("Failed to create Folder", fullpath, "Book " + book.FilePath + " was not moved.\nThe error was: " + str(ex))
					#self.report.Append("\n\nFailed to create path\n%s\nbecause an error occured. The error was: %s. Book was not %s." % (fullpath, ex, self.modetextpast))
					self.worker.ReportProgress(count)
					continue
			

			# If fileless. Since all fileless exceptions and rules were filtered out above. It is safe to create
					
			if book.FilePath == "":
				result = self.MoveFileless(book, fullpath)
				
				if result == MoveResult.Success:
					success += 1
			
				elif result == MoveResult.Failed:
					failed += 1
				
				elif result == MoveResult.Skipped:
					skipped += 1
				self.worker.ReportProgress(count)
				continue
				

#-------------- Normal file ----------------------------------------------------------------------------------------------------------------#
			
			result = self.MoveBook(book, fullpath)
			if result == MoveResult.Success:
				success += 1
			
			elif result == MoveResult.Failed:
				failed += 1
			
			elif result == MoveResult.Skipped:
				skipped += 1
				

			# Cleanup old directories
			if self.settings.RemoveEmptyFolder and self.settings.Mode == Mode.Move:
				self.CleanDirectories(DirectoryInfo(oldpath))
				self.CleanDirectories(DirectoryInfo(dirpath))
			
			self.worker.ReportProgress(count)
			

#----------- Now process the duplicate books --------------------------------------------------------------#

		#Making a copy of the array to iterate over since items are removed from the arry
		for book in self.HeldDuplicateBooks[:]:
			count += 1


			#Before making the paths check if the book needs to be moved at all
			
			if self.worker.CancellationPending:
				#User pressed cancel
				skipped = len(self.books) - success - failed
				self.logger.Add("Skipped", str(skipped) + " files", "Operation cancelled by user")
				#self.report.Append("\n\nOperation cancelled by user.")
				break


			#Since pretty much all the errors have already been dealt with the first time round. We can go straight to the move.

			#Create the path from pattern seperated as directory path and the file path
			
			dirpath, filepath = self.CreatePath(book)

			fullpath = Path.Combine(dirpath, filepath)

			#Create here because needed for cleaning directories later

			oldpath = book.FileDirectory
			
			#Note don't need to create the directry since the destination file already exists
				
			#If fileless. Since fileless and not creating fileless was filtered out above. It is safe to create.
					
			if book.FilePath == "":
				result = self.MoveFileless(book, fullpath)
				
				if result == MoveResult.Success:
					success += 1
			
				elif result == MoveResult.Failed:
					failed += 1
				
				elif result == MoveResult.Skipped:
					skipped += 1
				self.worker.ReportProgress(count)
				self.HeldDuplicateBooks.remove(book)
				continue
				

			#Normal file
			
			result = self.MoveBook(book, fullpath)
			if result == MoveResult.Success:
				success += 1
			
			elif result == MoveResult.Failed:
				failed += 1
			
			elif result == MoveResult.Skipped:
				skipped += 1
				
			#Cleanup old directories
			if self.settings.RemoveEmptyFolder and self.settings.Mode == Mode.Move:
				self.CleanDirectories(DirectoryInfo(oldpath))
				self.CleanDirectories(DirectoryInfo(dirpath))

			#Remove the book from the list so the duplicate dialog will show the correct number

			self.HeldDuplicateBooks.remove(book)

			self.worker.ReportProgress(count)


		#If we moved some comics write the undo file
		if len(self.MovedBooksUndo) > 0:
			locommon.SaveDict(self.MovedBooksUndo, locommon.UNDOFILE)

		#Return the report to the worker thread
		report = "Successfully %s: %s\nFailed to %s: %s\nSkipped: %s" % (self.modetextpast, success, self.modetext, failed, skipped)
		self.logger.SetCountVariables(failed, skipped, success)
		return [failed + skipped, report, self.logger]
	
	def MoveFileless(self, book, path):
		"""
		Book is the fileless book whose custom thumbnail will be created
		
		Path is the absolute path (with extension) where the images should be created
		"""
		
		#If something is already there. Note this will only occur on the final pass with held duplicate books
		if File.Exists(path) or path in self.MovedBooks:
			
			#Find the existing book in the library (if it exists)
			oldbook = self.FindDuplicate(path)
			
			if oldbook == None:
				oldbook = FileInfo(path)

			#Get the rename file name here so it can be put into the form
			renamepath = self.CreateRenamePath(path, book)
			
			renamefilename = FileInfo(renamepath).Name

			if not self.AlwaysDoAction:
				
				#Ask user what they want to do:
				
				result = self.form.Invoke(Func[type(book), type(oldbook), str, int, list](self.form.ShowDuplicateForm), System.Array[object]([book, oldbook, renamefilename, len(self.HeldDuplicateBooks)]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == DuplicateResult.Cancel:
				self.logger.Add("Skipped", path, "The image already exists and the user declined to overwrite it")
				#self.report.Append("\n\nSkipped creating image\n%s\nbecause a file already exists there and the user declined to overwrite it." % (path))
				return MoveResult.Skipped
			
			elif self.Action == DuplicateResult.Rename:
				if len(renamepath) > 259:
					result = self.form.Invoke(Func[str, object](self.GetSmallerPath), System.Array[System.Object]([renamepath]))
					if result == None:
						self.logger.Add("Skipped", path, "The path was too long and the user skipped shortening it")
						#self.report.Append("\n\nSkipped %s book\n%s\nbecause the path was too long and the user skipped shortening it." % (self.modetextplural, book.FilePath))
						return MoveResult.Skipped

					return self.MoveFileless(book, result)

				return self.MoveFileless(book, renamepath)
			
			elif self.Action == OverwriteAction.Overwrite:
				try:
					if self.settings.Mode == Mode.Test:
						#Because the script goes into a loop if in test mode here since no files are actually changed. return a success
						self.logger.Add("Deleted", path)
						#self.report.Append("\n\nDeleting:\n%s" % (path))
						self.logger.Add("Created image", path)
						#self.report.Append("\n\nCreating image %s" % (path))
						self.MoveBooks.append(path)
						return MoveResult.Success
					else:
						FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
						#File.Delete(path)
				except Exception, ex:
					self.logger.Add("Failed", path, "Failed to overwrite. The error was: " + str(ex))
					#self.report.Append("\n\nFailed to overwrite %s. The error was: %s. Image for fileless book\n%s %s\nwas not created" % (path, ex, book.ShadowSeries, book.ShadowNumber))
					return MoveResult.Failed

				#Since we are only working with images there is no need to remove a book from the library
				
				return self.MoveFileless(book, path)
			
		
		#Finally actually move the book
		try:			
			image = ComicRack.App.GetComicThumbnail(book, 0)
			format = None
			if self.settings.FilelessFormat == ".jpg":
				format = ImageFormat.Jpeg
			elif self.settings.FilelessFormat == ".png":
				format = ImageFormat.Png
			elif self.settings.FilelessFormat == ".bmp":
				format = ImageFormat.Bmp
				
			if self.settings.Mode == Mode.Test:
				self.logger.Add("Created image", path)
				#self.report.Append("\n\nCreating image %s" % (path))
				self.MovedBooks.append(path)
			else:
				image.Save(path, format)
			return MoveResult.Success

		except Exception, ex:
			self.logger.Add("Failed", path, "Failed to create the image because an error occured. The error was: " + str(ex))
			#self.report.Append("\n\nFailed to create image\n%s\nbecause an error occured. The error was: %s." % (path, ex))
			return MoveResult.Failed
	
	def MoveBook(self, book, path):
		"""
		Moves a book to the path. Checks for overwrite and catches a few errors

		Book is the book to be moved
		
		Path is the absolute destination with extension
		
		"""
		
		#Keep the current book path, just incase
		oldpath = book.FilePath
		
		#If something is already there. Note this will only occur on the final pass with held duplicate books
		if File.Exists(path) or path in self.MovedBooks:
			
			#Find the existing book if it occurs in the library
			oldbook = self.FindDuplicate(path)


			if oldbook == None:
				oldbook = FileInfo(path)

			#Get the rename file name here so it can be put into the form
			renamepath = self.CreateRenamePath(path, book)
			
			renamefilename = FileInfo(renamepath).Name

			if not self.AlwaysDoAction:

				result = self.form.Invoke(Func[type(book), type(oldbook), str, int, list](self.form.ShowDuplicateForm), System.Array[object]([book, oldbook, renamefilename, len(self.HeldDuplicateBooks)]))

				self.Action = result[0]

				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == DuplicateResult.Cancel:
				self.logger.Add("Skipped", book.FilePath, "A file already exists at: " + path + " and the user declined to overwrite it")
				#self.report.Append("\n\nSkipped %s\n%s\nbecause a file already exists at\n%s\nand the user declined to overwrite it." % (self.modetextplural, book.FilePath, path))
				return MoveResult.Skipped
			
			elif self.Action == DuplicateResult.Rename:
					#user choose to use the new name. Now we have to check and make sure it isn't too long:
					#-------------- Check for too long path error --------------------------------------------------------------------------#

				if len(renamepath) > 259:
					result = self.form.Invoke(Func[str, object](self.GetSmallerPath), System.Array[System.Object]([renamepath]))
					if result == None:
						self.logger.Add("Skipped", book.FilePath, "The path was too long and the user skipped shortening it")
						#self.report.Append("\n\nSkipped %s book\n%s\nbecause the path was too long and the user skipped shortening it." % (self.modetextplural, book.FilePath))
						return MoveResult.Skipped

					return self.MoveBook(book, result)

				return self.MoveBook(book, renamepath)
			
			elif self.Action == DuplicateResult.Overwrite:
				
				if self.settings.Mode == Mode.Test:
					self.logger.Add("Deleting (simulated)", path)
					#self.report.Append("\n\nDeleting:\n%s" % (path))

					#Because the script will get into an infitite loop if allowed to be called recursivly as down below since the book is not actualy deleted:
					#Fake the copy and return a success.
					self.logger.Add(self.modetextplural.capitalize(), book.FilePath, "to: " + path)
					#self.report.Append("\n\n%s book from\n%s\nto\n%s." % (self.modetextplural.capitalize(), book.FilePath, path))
					self.MovedBooks.append(path)
					return MoveResult.Success
				
				#Not in test mode so can actually make changes
				else:
					try:
						FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
						#File.Delete(path)
					except Exception, ex:
						self.logger.Add("Failed", book.FilePath, "Because " + path + " was unable to be deleted. The error was: " + str(ex))
						#self.report.Append("\n\nFailed to delete\n%s. The error was: %s. File\n%s\nwas not %s." % (path, ex, book.FilePath, self.modetextpast))
						return MoveResult.Failed
				
					#if the book is in the library, delete it!
					if type(oldbook) != FileInfo:
						ComicRack.App.RemoveBook(oldbook)
				
				return self.MoveBook(book, path)
			
		
		#Finally actually move the book
		try:

			if self.settings.Mode == Mode.Move:
				File.Move(book.FilePath, path)
				self.MovedBooksUndo[path] = book.FilePath
				book.FilePath = path

			elif self.settings.Mode == Mode.Test:
				self.logger.Add(self.modetextplural.capitalize(), book.FilePath, "to: " + path)
				#self.report.Append("\n\n%s book from\n%s\nto\n%s." % (self.modetextplural.capitalize(), book.FilePath, path))
				self.MovedBooks.append(path)

			elif self.settings.Mode == Mode.Copy:
				File.Copy(book.FilePath, path)
				if self.settings.CopyMode:
					newbook = ComicRack.App.AddNewBook(False)
					newbook.FilePath = path
					CopyData(book, newbook)

			return MoveResult.Success

		except Exception, ex:
			self.logger.Add("Failed", book.FilePath, "because an error occured. The error was: " + str(ex))
			#self.report.Append("\n\nFailed to %s\n%s\nbecause an error occured. The error was: %s. The book was not %s." % (self.modetext, book.FilePath, ex, self.modetextpast))
			return MoveResult.Failed
	
	def CreatePath(self, book):
		"""
		This function create a path with the passed book.
		
		Returns the directory path and the filepath as two strings
		"""
		
		dirpath = ""
		filepath = ""
		
		#Create the directory
		if self.settings.UseFolder:
			dirpath = self.pathmaker.CreateDirectoryPath(book, self.settings.FolderTemplate, self.settings.BaseFolder, self.settings.EmptyFolder, self.settings.EmptyData, self.settings.DontAskWhenMultiOne, self.settings.IllegalCharacters, self.settings.Months, self.settings.ReplaceMultipleSpaces)
		
		#Or use the books current directory
		else:
			dirpath = book.FileDirectory
		
		#Create filename
		if self.settings.UseFileName:
			filepath = self.pathmaker.CreateFileName(book, self.settings.FileTemplate, self.settings.EmptyData, self.settings.FilelessFormat, self.settings.DontAskWhenMultiOne, self.settings.IllegalCharacters, self.settings.Months, self.settings.ReplaceMultipleSpaces)
		
		#Or use current filename
		else:
			filepath = book.FileNameWithExtension
			
		return dirpath, filepath
	
	def FindDuplicate(self, path):
		"""
		Trys to find a book in the CR library via a path
		"""
		for b in ComicRack.App.GetLibraryBooks():
			if b.FilePath == path:
				return b
		return None

	def CreateRenamePath(self, path, book):

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
	
	def CleanDirectories(self, directory):
		"""
		Recursivly deletes directories until an non-empty directory is found or the directory is in the excluded list
		
		directory should be a DirectoryInfo object
		
		"""
		if not directory.Exists:
			return
		
		#Only delete if no file or folder and not in folder never to delete
		if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.settings.ExcludedEmptyFolder:
			parent = directory.Parent
			directory.Delete()
			self.CleanDirectories(parent)

	def GetSmallerPath(self, path):
		p = PathTooLongForm(path)
		r = p.ShowDialog()

		if r != DialogResult.OK:
			return None

		return p._Path.Text

class UndoMover(object):
	
	def __init__(self, worker, form, dict, settings):
		self.worker = worker
		self.form = form
		self.BookDict = dict
		self.report = StringBuilder()
		self.AlwaysDoAction = False
		self.settings = settings

		self.HeldDuplicateBooks = []

	def MoveBooks(self):

		success = 0
		failed = 0
		skipped = 0
		count = 0
		#get a list of the books
		books, notfound = self.GetBooks()



		for book in books + notfound:
			if type(book) == str:
				path = self.BookDict[book]
				oldfile = book
			else:
				path = self.BookDict[book.FilePath]
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
				path = self.BookDict[book]
				oldpath = book
			else:
				path = self.BookDict[book.FilePath]
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

	def GetBooks(self):
		books = []
		found = set()
		#Note possible that the book in not in the library.
		for b in ComicRack.App.GetLibraryBooks():
			if b.FilePath in self.BookDict:
				books.append(b)
				found.add(b.FilePath)
		all = set(self.BookDict.keys())
		notfound = all.difference(found)
		return books, list(notfound)

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

	def FindDuplicate(self, path):
		"""
		Trys to find a book in the CR library via a path
		"""
		for b in ComicRack.App.GetLibraryBooks():
			if b.FilePath == path:
				return b
		return None

	def CreateRenamePath(self, path):

		#By pescuma. modified slightly
		extension = Path.GetExtension(path)
		base = path[:-len(extension)]

		base = re.sub(" \([0-9]\)$", "", base)
				
		for i in range(100):
			newpath = base + " (" + str(i+1) + ")" + extension

			if File.Exists(newpath):
				continue
			else:
				return newpath

	def CleanDirectories(self, directory):
		"""
		Recursivly deletes directories until an non-empty directory is found or the directory is in the excluded list
		
		directory should be a DirectoryInfo object
		
		"""
		if not directory.Exists:
			return
		
		#Only delete if no file or folder and not in folder never to delete
		if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.settings.ExcludedEmptyFolder:
			parent = directory.Parent
			directory.Delete()
			self.CleanDirectories(parent)

def CopyData(book, newbook):
	"""This helper function copies all revevent metadata from a book to another book"""
	list = ["Series", "Number", "Count", "Month", "Year", "Format", "Title", "Publisher", "AlternateSeries", "AlternateNumber", "AlternateCount",
			"Imprint", "Writer", "Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "AgeRating", "Manga", "LanguageISO", "BlackAndWhite",
			"Genre", "Tags", "SeriesComplete", "Summary", "Characters", "Teams", "Locations", "Notes", "Web", "ScanInformation"]

	for i in list:
		setattr(newbook, i, getattr(book, i))

def ExcludePath(book, paths):
	"""
	Checks if a book is located in any directory or sub-directory in the passed list of paths
	
	book is the book to check
	paths is the list of paths
	
	retuns True if book is located in one of the paths or false if it is not
	"""
	
	for path in paths:
		if path in book.FilePath:
			return True
	return False

def ExcludeMeta(book, rules, operator, mode):
	"""
	Goes through each list of rules and finds if it applies or not
	
	book is the book to check
	rules is the python list of ExcludeRule and ExcludeGroup objects
	operator is the operator to use (all or any)
	mode is the mode to use. It determine is the book is moved when it falls under the rules or if it does not
	
	return True if the book should be skipped or false if it should be moved
	
	"""
	
	#Keep track of how many rules the book fell under.
	count = 0
	
	#keeps track of the total valid rules.
	total = 0
	
	qualifies = False
	
	for i in rules:
		
		#Empty list item. Doesn't count towards total. Have to check because of how the gui handles deleted rules
		if i == None:
			continue
		
		r = i.Calculate(book)
		
		#Possibly empty group. Do not count
		if r == None:
			continue
		
		count += r
		total += 1
	
	if total == 0:
		#No rules so the book should be moved regardless
		return False
	
	if operator == "Any":
		if count > 0:
			qualifies = True
	
	elif operator == "All":
		if count == total:
			qualifies = True

	if mode == "Only":
		#book falls under rules and should be moved if it does so return false to have it moved
		if qualifies == True:
			return False
		
		#book doesn't fall under rules and should not be moved. So return True to have it skipped
		else:
			return True
	
	elif mode == "Do not":
		#book falls under rules and should not be moved so return true to have it skipped
		if qualifies == True:
			return True
		#book doe not fall under rules and should be moved, so return false to have it moved
		else:
			return False

class PathMaker:
	"""A class to create directory and file paths from the passed book"""
	
	#To speed up repeated calculations of start year. Note this is a class variable since years won't change between the gui and worker using this
	startbook = {}

	def __init__(self, parentform):
		#These are for keeping track of the mutiple select options the user selected
		self.Characters = {}
		self.Tags = {}
		self.Genre = {}
		self.Writer = {}
		self.Teams = {}
		self.ScanInformation = {}
		self.AlternateSeries = {}
		self.Counter = None

		#Need to store the parent form so it can use the muilt select form
		self.form = parentform

		self._illegalCharactersRegEx = re.compile("[" + "".join(Path.GetInvalidPathChars()) + "]")
	
	def CreateDirectoryPath(self, insertedbook, template, basepath, emptypath, emptyreplace, DontAskWhenMultiOne, illegals, months, replacespaces):
	#To let the re.sub functions access the book object.
		global book 
		book = insertedbook
		
		global emptyreplacements
		emptyreplacements = emptyreplace

		self.Illegals = illegals
		#Sort by len so the script tries to replace the largest character first. ie. space?space is matched before ?
		self.IllegalsIterator = sorted(self.Illegals.keys(), key=len, reverse=True)

		self.Months = months
		self.DontAskWhenMultiOne = DontAskWhenMultiOne
		
		path = ""
		if not template.strip() == "":
			#In one possible case the template ends with \. This causes emptypath to be input at the end so remove that last \ if it exists
			template.strip()
			if template.endswith("\\"):
				template = template[:-1]
		
			roughpath = self.ReplaceValues(template)
			
			#Split into seperate directories for fixing empty paths and other problems.
			lines = roughpath.split("\\")
			
			for line in lines:
				if line.strip() == "":
					line = emptypath
				line = self.replaceIllegal(line)
				path = Path.Combine(path, line.strip())
		
		
		path = Path.Combine(basepath, path)

		#replace occurences of multiple spaces with a single space.
		if replacespaces:
			path = re.sub("\s\s+", " ", path)

		return path
	
	def CreateFileName(self, ibook, template, emptyreplace, filelessformat, DontAskWhenMultiOne, illegals, months, replacespaces):
		global book
		book = ibook
		
		
		global emptyreplacements
		emptyreplacements = emptyreplace	
		#Sort by len so the script tries to replace the largest character first. ie. space?space is matched before ?
		self.Illegals = illegals
		self.IllegalsIterator = sorted(self.Illegals.keys(), key=len, reverse=True)
		self.Months = months
		self.DontAskWhenMultiOne = DontAskWhenMultiOne
		
		r = self.ReplaceValues(template)
		r = r.strip()
		r = self.replaceIllegal(r)

		if not r:
                #template was empty
			return ""
                        
		
		#Do this check in the case of a fileless book for moving the thumbnails to the correct place.
		extension = filelessformat
		if book.FilePath:
			extension = FileInfo(book.FilePath).Extension

		#replace occurences of multiple spaces with a single space.
		if replacespaces:
			r = re.sub("\s\s+", " ", r)

		return r+extension
	
	def ReplaceValues(self, templateText):
		#Much of this modified from wadegiles's guided rename script.
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>series)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>number)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>count)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>month)#(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<month\>(?P<post>[^}]*)(?P<end>})', self.insertMonth, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<year\>(?P<post>[^}]*)(?P<end>})', self.insertYear, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>imprint)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>publisher)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altSeries)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altNumber)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altCount)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>volume)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>title)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>ageRating)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>language)\>(?P<post>[^}]*)(?P<end>})', self.insertLanguage, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>format)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>startyear)\>(?P<post>[^}]*)(?P<end>})', self.insertStartYear, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>writer)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>tags)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>genre)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>characters)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altSeries)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>teams)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>scaninfo)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>manga)\((?P<text>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertManga, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>seriesComplete)\((?P<text>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertSeriesComplete, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>first)\((?P<property>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertFirstLetter, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>read)\((?P<text>[^\)]*?)\)\((?P<operator>[^\)]*?)\)\((?P<percent>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertReadPercentage, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>counter)\((?P<start>\d*)\)\((?P<increment>\d*)\)\((?P<pad>\d*)\)\>(?P<post>[^}]*)(?P<end>})', self.insertCounter, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<startmonth\>(?P<post>[^}]*)(?P<end>})', self.insertStartMonth, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<startmonth#(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertStartMonthNumber, templateText)
		return templateText
	
	#Most of these functions are copied from wadegiles's guided rename script. (c) wadegiles. Most have been heavily modified by Stonepaw
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
	
	def replaceIllegal(self, text):
		for i in self.IllegalsIterator:
			text = text.replace(i, self.Illegals[i])

		#Replace any other illegal chracters that slip through
		text = self._illegalCharactersRegEx.sub("", text)
		return text
	
	def insertYear(self, matchObj):
		result = ""
		if book.ShadowYear > 0	:
			result = matchObj.group("pre") + str(book.ShadowYear) + matchObj.group("post")
		else:
			if emptyreplacements["Year"] != "":
				result = matchObj.group("pre") + emptyreplacements["Year"] + matchObj.group("post")
			else:
				result = ""
		
		return self.replaceIllegal(result)

	def insertShadowText(self, matchObj):
		result = None
		property =  matchObj.group("name").capitalize()
		
		sproperty = "Shadow" + property

		if getattr(book, sproperty) != "":
			result = matchObj.group("pre") + getattr(book, sproperty) + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				result = matchObj.group("pre") + emptyreplacements[property] + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)

	def insertText(self, matchObj):
		result = None
		property =  matchObj.group("name").capitalize()
		
		#Small change for the alternate series
		if property == "Altseries":
			property = "AlternateSeries"
		
		#Small change for age rating. probably a better way to do this
		if property == "Agerating":
			property = "AgeRating"

		if getattr(book, property) != "":
			result = matchObj.group("pre") + getattr(book, property) + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				result = matchObj.group("pre") + emptyreplacements[property] + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)


	def insertLanguage(self, matchObj):

		if book.LanguageAsText != "":
			result = matchObj.group("pre") + book.LanguageAsText + matchObj.group("post")
		else:
			if emptyreplacements["Language"] != "":
				result = matchObj.group("pre") + emptyreplacements["Language"] + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)

	def insertShadowPadded(self, matchObj):
		property = matchObj.group("name").capitalize()
		sproperty = "Shadow" + property
		
		result = None
		if getattr(book, sproperty) != "" and getattr(book, sproperty) > 0:
			if matchObj.group("pad") != None:
				try:
					result = self.pad(getattr(book, sproperty), int(matchObj.group("pad")))
				except ValueError:
					result = unicode(getattr(book, sproperty))
			else:
				result = unicode(getattr(book, sproperty))
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				try:
					int(emptyreplacements[property])
					if matchObj.group("pad") != None:
						result = self.pad(emptyreplacements[property], int(matchObj.group("pad")))
				except ValueError:
					result = emptyreplacements[property]
					
				result = matchObj.group("pre") + result + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)

	def insertPadded(self, matchObj):
		property = matchObj.group("name").capitalize()
		if property == "Altnumber":
			property = "AlternateNumber"
		if property == "Altcount":
			property = "AlternateCount"
	
		result = None
		if getattr(book, property) != "" and getattr(book, property) > 0:
			if matchObj.group("pad") != None:
				try:
					result = self.pad(getattr(book, property), int(matchObj.group("pad")))
				except ValueError:
					result = unicode(getattr(book, property))
			else:
				result = unicode(getattr(book,property))
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				try:
					int(emptyreplacements[property])
					if matchObj.group("pad") != None:
						result = self.pad(emptyreplacements[property], int(matchObj.group("pad")))
				except ValueError:
					result = emptyreplacements[property]
					
				result = matchObj.group("pre") + result + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)
		
	
	def insertMonth(self, matchObj):
		result = None
		if book.Month != -1:
			if book.Month in self.Months:
				result = self.Months[book.Month]
			else:
				#Something else, no set string for it so return it as nothing
				return ""
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements["Month"] != "":
				result = matchObj.group("pre") + emptyreplacements["Month"] + matchObj.group("post")
			else:
				result = ""
	
		return self.replaceIllegal(result)
	
	def insertStartYear(self, matchObj):
		#Find the start year by going through the whole list of comics in the library find the earliest year field of the same series and volume
		
		#index = book.Publisher+book.ShadowSeries+str(book.ShadowVolume)
		#
		#if self.startyear.has_key(index):
		#	startyear = self.startyear[index]
		#else:
		#	startyear = book.ShadowYear
		#	
		#	for b in ComicRack.App.GetLibraryBooks():
		#		if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume and b.Publisher == book.Publisher:
		#			
		#			#In case the initial values is bad
		#			if startyear == -1 and b.ShadowYear != 1:
		#				startyear = b.ShadowYear
		#			
	 #
		#			if b.ShadowYear != -1 and b.ShadowYear < startyear:
		#				startyear = b.ShadowYear
		#	
		#	#Store this final result in the dict so no calculation require for others of the series.
		#	self.startyear[index] = startyear

		startyear = self.GetEarliestBook().ShadowYear

		if	startyear != -1:
			result = matchObj.group("pre") + str(startyear) + matchObj.group("post")
		else:
			if emptyreplacements["StartYear"] != "":
				result = matchObj.group("pre") + emptyreplacements["StartYear"] + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)


	def insertStartMonth(self, matchObj):
		startmonth = self.GetEarliestBook().Month
		result = None
		if startmonth != -1:
			if startmonth in self.Months:
				result = self.Months[startmonth]
			else:
				#Something else, no set string for it so return it as nothing
				return ""
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements["StartMonth"] != "":
				result = matchObj.group("pre") + emptyreplacements["StartMonth"] + matchObj.group("post")
			else:
				result = ""
	
		return self.replaceIllegal(result)

	def insertStartMonthNumber(self, matchObj):

		startmonth = self.GetEarliestBook().Month
		result = None
		if startmonth > 0:
			if matchObj.group("pad") != None:
				try:
					result = self.pad(startmonth, int(matchObj.group("pad")))
				except ValueError:
					result = unicode(startmonth)
			else:
				result = unicode(startmonth)
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements["StartMonth"] != "":
				try:
					int(emptyreplacements["StartMonth"])
					if matchObj.group("pad") != None:
						result = self.pad(emptyreplacements["StartMonth"], int(matchObj.group("pad")))
				except ValueError:
					result = emptyreplacements["StartMonth"]
					
				result = matchObj.group("pre") + result + matchObj.group("post")
			else:
				result = ""

		return self.replaceIllegal(result)


	def GetEarliestBook(self):
		"""
		Finds the first published issue of a series in the library
		Returns a ComicBook object
		"""
		#Find the Earliest by going through the whole list of comics in the library find the earliest year field and month field of the same series and volume
		
		index = book.Publisher+book.ShadowSeries+str(book.ShadowVolume)
		
		if self.startbook.has_key(index):
			startbook = self.startbook[index]
		else:
			startbook = book
			
			for b in ComicRack.App.GetLibraryBooks():
				if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume and b.Publisher == book.Publisher:
					
					#Notes:
					#Year can be empty (-1)
					#Month can be empty (-1)

					#In case the initial value is bad
					if startbook.ShadowYear == -1 and b.ShadowYear != 1:
						startbook = b
					
					#Check if the current book's year 
					if b.ShadowYear != -1 and b.ShadowYear < startbook.ShadowYear:
						startbook = b

					#Check if year the same and a valid month
					if b.ShadowYear == startbook.ShadowYear and b.Month != -1:

						#Current book has empty month
						if startbook.Month == -1:
							startbook = b
						
						#Month is earlier
						elif b.Month < startbook.Month:
							startbook = b
			
			#Store this final result in the dict so no calculation require for others of the series.
			self.startbook[index] = startbook

		return startbook

	def insertManga(self, matchObj):

		if emptyreplacements["Manga"].strip() != "":
			empty = matchObj.group("pre") + emptyreplacements["Manga"] + matchObj.group("post")
		else:
			empty = ""
		
		if book.Manga == MangaYesNo.Yes or book.Manga == MangaYesNo.YesAndRightToLeft:
			r = matchObj.group("text")
			if r.strip() != "":
				return self.replaceIllegal(matchObj.group("pre") + r + matchObj.group("post"))

		return self.replaceIllegal(empty)


	def insertSeriesComplete(self, matchObj):
		if emptyreplacements["SeriesComplete"].strip() != "":
			empty = matchObj.group("pre") + emptyreplacements["SeriesComplete"] + matchObj.group("post")
		else:
			empty = ""
		
		if book.SeriesComplete == YesNo.Yes:
			r = matchObj.group("text")
			if r.strip() != "":
				return self.replaceIllegal(matchObj.group("pre") + r + matchObj.group("post"))

		return self.replaceIllegal(empty)

	def insertReadPercentage(self, matchObj):
		text = matchObj.group("text")
		percent = matchObj.group("percent")
		operator = matchObj.group("operator")
		post = matchObj.group("post")
		pre = matchObj.group("pre")


		result = ""

		if operator == "=":
			if book.ReadPercentage == int(percent):
				result = pre + text + post
		elif operator == ">":
			if book.ReadPercentage > int(percent):
				result = pre + text + post
		elif operator == "<":
			if book.ReadPercentage < int(percent):
				result = pre + text + post

		if result == "":
			if emptyreplacements["ReadPercentage"].strip() != "":
				result = pre + emptyreplacements["ReadPercentage"] + post
			else:
				return ""

		return self.replaceIllegal(result)

	def insertMulti(self, matchObj):
		#Get a bool for if using series. Just simplifies the code a bit
		if matchObj.group("series") == "series":
			return self.InsertMultiSeries(matchObj)
		else:
			return self.InsertMultiIssue(matchObj)


	def InsertMultiIssue(self, matchObj):

		#Initial setup

		field = matchObj.group("name").capitalize()

		#Small catch
		if field == "Scaninfo":
			field = "ScanInformation"

		if field == "Altseries":
			field = "AlternateSeries"

		#Get the correct list storage based on the field
		list = getattr(self, field)

		#Easier access in the code
		post = matchObj.group("post")
		pre = matchObj.group("pre")
		sep = matchObj.group("sep")


		#What to return if empty. Simplifies the later code.

		if emptyreplacements[field].strip() != "":
			empty = pre + emptyreplacements[field] + post
		else:
			empty = ""

		index = book.ShadowSeries + str(book.ShadowVolume) + book.ShadowNumber
		booktext = book.ShadowSeries + " vol. " + str(book.ShadowVolume) + " #" + book.ShadowNumber

		#No items
		if getattr(book, field).strip() == "":
			list[index] = SelectionFormResult([])
			return empty

		alwaysuseditems = []
		if "LibraryOrgaizerAlwaysUse" in list:
			alwaysuseditems = list["LibraryOrgaizerAlwaysUse"]

		#Already done this one so retry it. This won't occur in normal moving operation but is needed for the gui preview.
		if index in list:
			r = self.MakeMultiSelectionString(list[index], sep, field)

			if r == "":
				return empty

			return pre + r + post

		items = []
		for i in getattr(book, field).split(","):
			items.append(i.strip())

		#Only one result, check if asking user
		if len(items) == 1 and self.DontAskWhenMultiOne == True:
			return pre + items[0] + post


		#Find which we already need to use:
		used = []
		#Go through each set of alwaysuseditems
		for set in alwaysuseditems:
			count = 0
			#Go through each item in the set. First two items are bools for the two options
			for item in set[2:]:
				if item in items:
					count +=1
			#All items are in the field
			if count == len(set) -2 :

				#Not asking if there are additional items
				if set[0] == True:
					if set[1] == True:
						#Using folder speration
						return pre + "\\".join(set[2:]) + post
					else:
						#Not using folder speration
						return pre + sep.join(set[2:]) + post

				#Same number of items but asking if there are additonal items
				else:
					used = set[2:]
					break

		#Since this can be shown from the configform:
		if self.form.InvokeRequired:
			result = self.form.Invoke(Func[SelectionFormArgs, SelectionFormResult](self.ShowSelectionForm), System.Array[object]([SelectionFormArgs(items, used, field, booktext, False)]))

		else:
			result = self.ShowSelectionForm(SelectionFormArgs(items, used, field, booktext, False))


		if result.AlwaysUse and len(result.Selection) > 0:
			if len(alwaysuseditems) == 0:
				list["LibraryOrgaizerAlwaysUse"] = []

			rlist = []
			rlist.append(result.AlwaysUseDontAsk)
			rlist.append(result.Folder)
			rlist.extend(result.Selection)


			#Check if the same list already exists. This shouldn't be encountered but it doesn't hurt to check
			if not rlist in list["LibraryOrgaizerAlwaysUse"]:
				list["LibraryOrgaizerAlwaysUse"].append(rlist)

				alwaysuseditems = list["LibraryOrgaizerAlwaysUse"]

		list[index] = result
		r = self.MakeMultiSelectionString(result, sep, field)

		if r == "":
			return empty

		return pre + r + post

	def InsertMultiSeries(self, matchObj):

		"""
		A series based calculation will find all of the filed it can find in the series then ask the user which ones to use.

		The user has the option to select to use all these characters for every item in the series, even if the item doesn't have that
		"""

		#Initial setup

		field = matchObj.group("name").capitalize()

		#Small catch
		if field == "Scaninfo":
			field = "ScanInformation"


		#Get the correct list storage based on the field
		list = getattr(self, field)

		#Easier access in the code
		post = matchObj.group("post")
		pre = matchObj.group("pre")
		sep = matchObj.group("sep")


		#What to return if empty. Simplifies the later code.
		#Temporarily disabled becacue those empty replacements aren't implemented yet

		if emptyreplacements[field].strip() != "":
			empty = pre + emptyreplacements[field] + post
		else:
			empty = ""



		#Note to make it simpler, a series based caclulation will never access the always use option


		index = book.ShadowSeries + str(book.ShadowVolume)
		booktext = book.ShadowSeries + " vol. " + str(book.ShadowVolume) 

		#First see if anything is saved

		if index in list:
			#This series was encountered before so calculate based off of that
			r = self.MakeMultiSelectionString(list[index], sep, field)

			if r == "":
				return empty
			
			return pre + r + post

		#New series so calculate from scratch:


		#Find all the items in the entire series

		items = self.GetAllFromSeriesField(field, book.ShadowSeries, book.ShadowVolume)

		#No items at all in the entire series
		if len(items) == 0:
			list[index] = SelectionFormResult([])
			return empty


		#Ask the user which ones to use

		#Since this can be shown from the configform...
		if self.form.InvokeRequired:
			result = self.form.Invoke(Func[SelectionFormArgs, SelectionFormResult](self.ShowSelectionForm), System.Array[object]([SelectionFormArgs(items, [], field, booktext, True)]))

		else:
			result = self.ShowSelectionForm(SelectionFormArgs(items, [], field, booktext, True))

		list[index] = result

		r = self.MakeMultiSelectionString(result, sep, field)
		
		if r == "":
			return empty

		return pre + r + post


	def ShowSelectionForm(self, args):
		f = SelectionForm(args)
		f.ShowDialog()
		result = f.GetResults()
		f.Dispose()
		return result

	def MakeMultiSelectionString(self, results, sep, field):
		"""
		Makes the correctly formated inserted string based on the input values:

		results is a SelectionFormResult object
		sep is the string seperator
		field is the correct name of the field to find from the book
		"""

		itemsToUse = ""

		result = ""

		#First check to see if using a folder. If so no need to use the user entered sperator
		if results.Folder:
			sep += "\\"

		#Get a list of all the items in the book
		items = []
		for i in getattr(book, field).split(","):
			items.append(i.strip())

		#Using every issue
		if results.EveryIssue:
			result = sep.join(result.Selection)

		#Not using every issue so just use the ones it has
		else:
			itemsToUse = []		
			for i in results.Selection:
				if i in items:
					itemsToUse.append(i)
			result = sep.join(itemsToUse)
		
		if results.Folder == False:
			return self.replaceIllegal(result)

		return result
		

	def GetAllFromSeriesField(self, field, series, volume):
		"""
		Helper to get all the unique fields from a series

		field is a string correctly named field to find
		series is the string series to find
		volume is the string volume to find
		"""
		allbooks = ComicRack.App.GetLibraryBooks()
		#using a set to avoid duplicate entries
		results = set()
		for book in allbooks:
			if book.ShadowSeries == series and book.ShadowVolume == volume:
				for i in getattr(book, field).split(","):
					i = i.strip()
					if not i == "":
						results.add(i)

		return list(results)

	def insertCounter(self, matchObj):
		post = matchObj.group("post")
		pre = matchObj.group("pre")

		pad = matchObj.group("pad")

		if pad == "":
			pad = 0

		else:
			pad = int(pad)

		if self.Counter == None:
			self.Counter = int(matchObj.group("start"))
			result = self.pad(self.Counter, pad)

		else:
			self.Counter = self.Counter + int(matchObj.group("increment"))
			result = self.pad(self.Counter, pad)

		if result.strip() == "":
			if emptyreplacements["Counter"].strip() != "":
				return pre + self.pad(emptyreplacements["Counter"], pad) + post
			else:
				return ""

		return pre + result + post


	def insertFirstLetter(self, matchObj):

		propertyname = matchObj.group("property").capitalize()

		if propertyname in ["Series"]:
			propertyname = "Shadow" + propertyname

		try:
			property = unicode(getattr(book, propertyname))
		except System.MissingMemberException:
			return ""



		r = re.match(r"(?:(?:the|a|an|de|het|een|die|der|das|des|dem|der|ein|eines|einer|einen|la|le|l'|les|un|une|el|las|los|las|un|una|unos|unas|o|os|um|uma|uns|umas|en|et|il|lo|uno|gli)\s+)?(?P<letter>.).+", property, re.I)

		if r:
			result = r.group("letter").capitalize()

		else:
			if emptyreplacements["FirstLetter"].strip() != "":
				result = emptyreplacements["FirstLetter"]
			else:
				return ""

		return matchObj.group("pre") + result + matchObj.group("post")