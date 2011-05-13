"""
lobookmover.py

This contains all the book moving fuction the script uses. Also the path generator

Version 1.4

Copyright Stonepaw 2011. Some code copyright wadegiles Anyone is free to use code from this file as long as credit is given.
"""


import clr

import re

import System

from System import Func

from System.Text import StringBuilder

from System.IO import Path, File, FileInfo, DirectoryInfo, Directory

import loforms

from loforms import SelectionFormArgs, SelectionFormResult, SelectionForm

import locommon

from locommon import Mode

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
		self.report = StringBuilder()
		self.pathmaker = PathMaker(form)
		
		
		#This is for when the script is in Test mode. It hold created paths so it won't be writing that it created the path for every comic
		self.CreatedPaths = []
		self.MovedBooks = []

		#This is for creating the undo script later
		#Dict with the key of the newpath and value of the old path
		self.MovedBooksUndo = {}
		
		self.AlwaysDoAction = False
		self.Action = None
		
		self.report.Append("Library Organizer Report:")
		
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
				self.report.Append("\n\nOperation cancelled by user.")
				break
			
			
			#Check the exclude paths first.
			if ExcludePath(book, self.settings.ExcludeFolders):
				skipped += 1
				self.report.Append("\n\nSkipped moving\n%s\nbecause it is located in an excluded path." % (book.FilePath))
				self.worker.ReportProgress(count)
				continue
			
			
			
			#
			#   Metadata Exclude Rules
			#
			"""
			Possible results are True or False
				True is when the book should not be moved
				False is when the book should be moved
				
				The function does the checking for the correct mode and returns the correct move/don't move result
			"""
			if ExcludeMeta(book, self.settings.ExcludeRules, self.settings.ExcludeOperator, self.settings.ExcludeMode):
				skipped += 1
				self.report.Append("\n\nSkipped moving\n%s \nbecause it qualified under the exclude rules." % (book.FilePath))
				self.worker.ReportProgress(count)
				continue


			
			# Fileless and not creating fileless images
			if book.FilePath == "" and not self.settings.MoveFileless:
				#Fileless comic
				skipped += 1
				self.report.Append("\n\nSkipped moving book\n%s %s\nbecause it is a fileless book." % (book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			
			#Fileless and moving but no custom thumbnail
			
			if book.FilePath == "" and self.settings.MoveFileless and not book.CustomThumbnailKey:
				#Fileless comic
				skipped += 1
				self.report.Append("\n\nSkipped creating image for filess book\n%s %s\nbecause it does not have a custom thumbnail." % (book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			#Create the path from pattern seperated as directory path and the file path
			
			dirpath, filepath = self.CreatePath(book)
			
			
			#Empty file name
			if filepath == "":
				failed += 1
				self.report.Append("\n\nFailed to move\n%s\nbecause the filename created was empty. The book was not moved" % (book.FilePath))
				self.worker.ReportProgress(count)
				continue
			
			#Create here because needed for cleaning directories later
			oldpath = book.FileDirectory
			
			#Create the directory
			if not Directory.Exists(dirpath):
				try:
					if not self.settings.Mode == Mode.Test:
						Directory.CreateDirectory(dirpath)
					else:
						if not dirpath in self.CreatedPaths:
							self.report.Append("\n\nCreating directory:\n%s" % (dirpath))
							self.CreatedPaths.append(dirpath)

				except Exception, ex:
					failed += 1
					self.report.Append("\n\nFailed to create path\n%s\nbecause an error occured. The error was: %s. Book was not moved." % (newpath, ex, book.FilePath))
					self.worker.ReportProgress(count)
					continue
				
			#If fileless. Since filess and not creating fileless was filtered out above. It is save to create.
					
			if book.FilePath == "":
				result = self.MoveFileless(book, Path.Combine(dirpath, filepath))
				
				if result == MoveResult.Success:
					success += 1
			
				elif result == MoveResult.Failed:
					failed += 1
				
				elif result == MoveResult.Skipped:
					skipped += 1
				self.worker.ReportProgress(count)
				continue
				

			#Normal file
			
			result = self.MoveBook(book, Path.Combine(dirpath, filepath))
			if result == MoveResult.Success:
				success += 1
			
			elif result == MoveResult.Failed:
				failed += 1
			
			elif result == MoveResult.Skipped:
				skipped += 1
				
			#Cleanup old directories
			if self.settings.RemoveEmptyDir and self.settings.Mode == Mode.Move:
				self.CleanDirectories(DirectoryInfo(oldpath))
				self.CleanDirectories(DirectoryInfo(dirpath))
			
			self.worker.ReportProgress(count)
			

		#If we moved some comics write the undo file
		if len(self.MovedBooksUndo) > 0:
			locommon.SaveDict(self.MovedBooksUndo, locommon.UNDOFILE)
		#Return the report to the worker thread
		report = "Successfully moved: %s\nFailed to move: %s\nSkipped: %s" % (success, failed, skipped)
		return [failed + skipped, report, self.report.ToString()]
	
	def MoveFileless(self, book, path):
		"""
		Book is the fileless book whose custom thumbnail will be created
		
		Path is the absolute path (with extension) where the images should be created
		"""
		
		
		#TODO: required? Since this is a fileless comic oldpath will always be empty
		#oldpath = book.FilePath
		
		#If something is already there
		if File.Exists(path) or path in self.MovedBooks:
			
			#Find the existing book in the library (if it exists)
			oldbook = self.FindDuplicate(path)
			
			if not self.AlwaysDoAction:
				
				#Ask user what they want to do:
				
				result = self.form.Invoke(Func[System.String, type(book), type(oldbook), list](self.form.AskOverwrite), System.Array[object]([path, book, oldbook]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == OverwriteAction.Cancel:
				self.report.Append("\n\nSkipped creating image\n%s\nbecause a file already exists there and the user declined to overwrite it." % (path))
				return MoveResult.Skipped
			
			elif self.Action == OverwriteAction.Rename:
				#This bit writen by pescuma. Thanks!
				# Find an available name
				extension = Path.GetExtension(path)
				base = path[:-len(extension)]
				
				for i in range(100):
					newpath = base + " (" + str(i+1) + ")" + extension

					#For test mode
					if newpath in self.MovedBooks:
						continue

					if not File.Exists(newpath):
						return self.MoveFileless(book, newpath)
				
				self.report.Append("\n\nFailed to find an available name to rename image. Image %s was not created." % (path))
				return MoveResult.Failed
			
			elif self.Action == OverwriteAction.Overwrite:
				try:
					if self.settings.Mode == Mode.Test:
						self.report.Append("\n\nDeleting:\n%s" % (path))
						self.report.Append("\n\nCreating image %s" % (path))
						self.MoveBooks.append(path)
					else:
						FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
						#File.Delete(path)
				except Exception, ex:
					self.report.Append("\n\nFailed to overwrite %s. The error was: %s. Image for fileless book\n%s %s\nwas not created" % (path, ex, book.ShadowSeries, book.ShadowNumber))
					return MoveResult.Failed
				
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
				self.report.Append("\n\nCreating image %s" % (path))
				self.MoveBooks.append(path)
			else:
				image.Save(path, format)
			return MoveResult.Success
		except Exception, ex:
			self.report.Append("\n\nFailed to create image\n%s\nbecause an error occured. The error was: %s." % (path, ex))
			return MoveResult.Failed
	
	def MoveBook(self, book, path):
		"""
		Book is the book to be moved
		
		Path is the absolute destination with extension
		
		"""
		
		#Keep the current book path, just incase
		oldpath = book.FilePath
		
		if not File.Exists(book.FilePath):
			self.report.Append("\n\nFailed to move\n%s\nbecause the file does not exist." % (book.FilePath))
			return MoveResult.Failed
		
		if path == book.FilePath:
			self.report.Append("\n\nSkipped moving book\n%s\nbecause it is already located at the calculated path." % (book.FilePath))
			return MoveResult.Skipped
		
		if File.Exists(path) or path in self.MovedBooks:
			
			#Find the existing book if it occurs in the library
			oldbook = self.FindDuplicate(path)
			
			if not self.AlwaysDoAction:
				
				#Ask the user:
				result = self.form.Invoke(Func[System.String, type(book), type(oldbook), list](self.form.AskOverwrite), System.Array[object]([path, book, oldbook]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == OverwriteAction.Cancel:
				self.report.Append("\n\nSkipped moving\n%s\nbecause a file already exists at\n%s\nand the user declined to overwrite it." % (book.FilePath, path))
				return MoveResult.Skipped
			
			elif self.Action == OverwriteAction.Rename:
				#This bit writen by pescuma. Thanks!
				# Find an available name
				extension = Path.GetExtension(path)
				base = path[:-len(extension)]
				
				for i in range(100):
					newpath = base + " (" + str(i+1) + ")" + extension

					#For test mode
					if newpath in self.MovedBooks:
						continue

					if not File.Exists(newpath) or newpath == book.FilePath:
						return self.MoveBook(book, newpath)
				
				self.report.Append("\n\nFailed to find an available name to rename book. Book\n%s\nwas not moved." % (path))
				return MoveResult.Failed
			
			elif self.Action == OverwriteAction.Overwrite:
				
				if self.settings.Mode == Mode.Test:
					self.report.Append("\n\nDeleting:\n%s" % (path))

					#Because the script will get into an infitite loop if allowed to be called recursivly as down below since the book is not actualy deleted:
					#Fake the copy and return a success.
					self.report.Append("\n\nMoving book from\n%s\nto\n%s." % (book.FilePath, path))
					self.MovedBooks.append(path)
					return MoveResult.Success
				
				#Not in test mode so can actually make changes
				else:
					try:
						FileIO.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
						#File.Delete(path)
					except Exception, ex:
						self.report.Append("\n\nFailed to delete\n%s. The error was: %s. File\n%s\nwas not moved." % (path, ex, book.FilePath))
						return MoveResult.Failed
				
					#if the book is in the library, delete it
					if oldbook:
						ComicRack.App.RemoveBook(oldbook)
				
				return self.MoveBook(book, path)
			
		
		#Finally actually move the book
		try:
			if self.settings.Mode == Mode.Move:
				File.Move(book.FilePath, path)
				self.MovedBooksUndo[path] = book.FilePath
				book.FilePath = path
			elif self.settings.Mode == Mode.Test:
				self.report.Append("\n\nMoving book from\n%s\nto\n%s." % (book.FilePath, path))
				self.MovedBooks.append(path)
			elif self.settings.Mode == Mode.Copy:
				File.Copy(book.FilePath, path)
				if self.settings.CopyMode:
					newbook = ComicRack.App.AddNewBook(False)
					newbook.FilePath = path
					CopyData(book, newbook)
			return MoveResult.Success
		except Exception, ex:
			self.report.Append("\n\nFailed to move\n%s\nbecause an error occured. The error was: %s. The book was not moved." % (book.FilePath, ex))
			return MoveResult.Failed
	
	def CreatePath(self, book):
		"""
		This function create a path with the passed book.
		
		Returns the directory path and the filepath as two strings
		"""
		
		dirpath = ""
		filepath = ""
		
		#Create the directory
		if self.settings.UseDirectory:
			dirpath = self.pathmaker.CreateDirectoryPath(book, self.settings.DirTemplate, self.settings.BaseDir, self.settings.EmptyDir, self.settings.EmptyData)
		
		#Or use the books current directory
		else:
			dirpath = book.FileDirectory
		
		#Create filename
		if self.settings.UseFileName:
			filepath = self.pathmaker.CreateFileName(book, self.settings.FileTemplate, self.settings.EmptyData, self.settings.FilelessFormat)
		
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
	
	def CleanDirectories(self, directory):
		"""
		Recursivly deletes directories until an non-empty directory is found or the directory is in the excluded list
		
		directory should be a DirectoryInfo object
		
		"""
		if not directory.Exists:
			return
		
		#Only delete if no file or folder and not in folder never to delete
		if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.settings.ExcludedEmptyDir:
			parent = directory.Parent
			directory.Delete()
			self.CleanDirectories(parent)

class UndoMover(object):
	
	def __init__(self, worker, form, dict, settings):
		self.worker = worker
		self.form = form
		self.BookDict = dict
		self.report = StringBuilder()
		self.AlwaysDoAction = False
		self.settings = settings

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
			else:
				path = self.BookDict[book.FilePath]
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
			if self.settings.RemoveEmptyDir:
				self.CleanDirectories(DirectoryInfo(oldpath))
				self.CleanDirectories(DirectoryInfo(f.DirectoryName))
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
		
		if not File.Exists(oldpath):
			self.report.Append("\n\nFailed to move\n%s\nbecause the file does not exist." % (oldpath))
			return MoveResult.Failed
		
		if path == oldpath:
			self.report.Append("\n\nSkipped moving book\n%s\nbecause it is already located at the calculated path." % (oldpath))
			return MoveResult.Skipped
		
		if File.Exists(path):
			
			#Find the existing book if it occurs in the library
			oldbook = self.FindDuplicate(path)
			
			if not self.AlwaysDoAction:
				
				#Ask the user:
				result = self.form.Invoke(Func[System.String, type(book), type(oldbook), list](self.form.AskOverwrite), System.Array[object]([path, book, oldbook]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == OverwriteAction.Cancel:
				self.report.Append("\n\nSkipped moving\n%s\nbecause a file already exists at\n%s\nand the user declined to overwrite it." % (oldpath, path))
				return MoveResult.Skipped
			
			elif self.Action == OverwriteAction.Rename:
				#This bit writen by pescuma. Thanks!
				# Find an available name
				extension = Path.GetExtension(path)
				base = path[:-len(extension)]
				
				for i in range(100):
					newpath = base + " (" + str(i+1) + ")" + extension

					if not File.Exists(newpath) or newpath == oldpath:
						return self.MoveBook(book, newpath)
				
				self.report.Append("\n\nFailed to find an available name to rename book. Book\n%s\nwas not moved." % (path))
				return MoveResult.Failed
			
			elif self.Action == OverwriteAction.Overwrite:

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

	def CleanDirectories(self, directory):
		"""
		Recursivly deletes directories until an non-empty directory is found or the directory is in the excluded list
		
		directory should be a DirectoryInfo object
		
		"""
		if not directory.Exists:
			return
		
		#Only delete if no file or folder and not in folder never to delete
		if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.settings.ExcludedEmptyDir:
			parent = directory.Parent
			directory.Delete()
			self.CleanDirectories(parent)

def CopyData(book, newbook):
	"""This helper function copies all revevent metadata from a book to another book"""
	list = ["Series", "Number", "Count", "Month", "Year", "Format", "Title", "Publisher", "AlternateSeries", "AlternateNumber", "AlternateCount",
			"Imprint", "Writer", "Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "AgeRating", "Manga", "LanguageISO", "BlackAndWhite",
			"Genre", "Tags", "SeriesComplete", "Summary", "Characters", "Teams", "Locations", "Notes", "Web"]

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
	startyear = {}

	def __init__(self, parentform):
		#These are for keeping track of the mutiple select options the user selected
		self.Characters = {}
		self.Tags = {}
		self.Genre = {}
		self.Writer = {}
		self.Teams = {}

		#Need to store the parent form so it can use the muilt select form
		self.form = parentform
	
	def CreateDirectoryPath(self, insertedbook, template, basepath, emptypath, emptyreplace):
	#To let the re.sub functions access the book object.
		global book 
		book = insertedbook
		
		global emptyreplacements
		emptyreplacements = emptyreplace
		
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
		return path
	
	def CreateFileName(self, ibook, template, emptyreplace, filelessformat):
		global book
		book = ibook
		
		
		global emptyreplacements
		emptyreplacements = emptyreplace	
		
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

		return r+extension
	
	def ReplaceValues(self, templateText):
		#Much of this modified from wadegiles's guided rename script.
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<series\>(?P<post>[^}]*)(?P<end>})', self.insertSeries, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>number)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>count)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<year\>(?P<post>[^}]*)(?P<end>})', self.insertYear, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>imprint)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>publisher)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altSeries)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altNumber)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>altCount)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<month\>(?P<post>[^}]*)(?P<end>})', self.insertMonth, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>month)#(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>volume)(?P<pad>\d*)\>(?P<post>[^}]*)(?P<end>})', self.insertShadowPadded, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>title)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>format)\>(?P<post>[^}]*)(?P<end>})', self.insertText, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>startyear)\>(?P<post>[^}]*)(?P<end>})', self.insertStartYear, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>writer)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>tags)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>genre)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>characters)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>teams)\((?P<sep>[^\)]*?)\)\((?P<series>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertMulti, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>manga)\((?P<text>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertManga, templateText)
		templateText = re.sub(r'(?i)(?P<start>{)(?P<pre>[^{]*)\<(?P<name>seriesComplete)\((?P<text>[^\)]*?)\)\>(?P<post>[^}]*)(?P<end>})', self.insertSeriesComplete, templateText)
		return templateText
	
	#Most of these functions are copied from wadegiles's guided rename script. (c) wadegiles. Most have been modified by Stonepaw
	def pad(self, value, padding):
		if value >= 0:
			paddedValue = str(value)
			return paddedValue.PadLeft(padding, '0')
		else:
			paddedValue = str(-value)
			paddedValue = paddedValue.PadLeft(padding, '0')
			return '-' + paddedValue
	
	def replaceIllegal(self, text):
		text = text.replace('?', '')
		text = text.replace('/', '')
		text = text.replace('\\', '')
		text = text.replace('*', '')
		text = text.replace(':', ' -')
		text = text.replace('<', '[')
		text = text.replace('>', ']')
		text = text.replace('|', '!')
		text = text.replace('"', '\'')
		return text
	
	def insertSeries(self, matchObj):
		result = None
		series = book.ShadowSeries
		if series != "":
				result = matchObj.group("pre") + series + matchObj.group("post")
		else:
			if emptyreplacements["Series"] != "":
				result = matchObj.group("pre") + emptyreplacements["Series"] + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result
	
	def insertYear(self, matchObj):
		result = None
		if book.ShadowYear > 0	:
			result = matchObj.group("pre") + str(book.ShadowYear) + matchObj.group("post")
		else:
			if emptyreplacements["Year"] != "":
				result = matchObj.group("pre") + emptyreplacements["Year"] + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result
	
	def insertText(self, matchObj):
		result = None
		property =  matchObj.group("name").capitalize()
		
		#Small change for the alternate series
		if property == "Altseries":
			property = "AlternateSeries"
		
		if getattr(book, property) != "":
			result = matchObj.group("pre") + getattr(book, property) + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				result = matchObj.group("pre") + emptyreplacements[property] + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result

	def insertShadowPadded(self, matchObj):
		property = matchObj.group("name").capitalize()
		sproperty = "Shadow" + property
		
		result = None
		if getattr(book, sproperty) != "" and getattr(book, sproperty) > 0:
			if matchObj.group("pad") != None:
				try:
					result = self.pad(int(getattr(book, sproperty)), int(matchObj.group("pad")))
				except ValueError:
					result = str(getattr(book, sproperty))
			else:
				result = str(getattr(book, sproperty))
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				try:
					int(emptyreplacements[property])
					if matchObj.group("pad") != None:
						result = self.pad(int(emptyreplacements[property]), int(matchObj.group("pad")))
				except ValueError:
					result = emptyreplacements[property]
					
				result = matchObj.group("pre") + result + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result

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
					result = self.pad(int(getattr(book, property)), int(matchObj.group("pad")))
				except ValueError:
					result = str(getattr(book, property))
			else:
				result = str(getattr(book,property))
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements[property] != "":
				try:
					int(emptyreplacements[property])
					if matchObj.group("pad") != None:
						result = self.pad(int(emptyreplacements[property]), int(matchObj.group("pad")))
				except ValueError:
					result = emptyreplacements[property]
					
				result = matchObj.group("pre") + result + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result
		
	
	def insertMonth(self, matchObj):
		result = None
		if book.Month > 0:
			if book.Month == 1:
				result = "January"
			elif book.Month == 2:
				result = "February"
			elif book.Month == 3:
				result = "March"
			elif book.Month == 4:
				result = "April"
			elif book.Month == 5:
				result = "May"
			elif book.Month == 6:
				result = "June"
			elif book.Month == 7:
				result = "July"
			elif book.Month == 8:
				result = "August"
			elif book.Month == 9:
				result = "September"
			elif book.Month == 10:
				result = "October"
			elif book.Month == 11:
				result = "November"
			elif book.Month == 12:
				result = "December"
			else:
				#Something else, no set string for it so return it as nothing
				return ""
			result = matchObj.group("pre") + result + matchObj.group("post")
		else:
			if emptyreplacements["Month"] != "":
				result = matchObj.group("pre") + emptyreplacements["Month"] + matchObj.group("post")
			else:
				result = ""
	
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result
	
	def insertStartYear(self, matchObj):
		#Find the start year by going through the whole list of comics in the library find the earliest year field of the same series and volume
		
		
		if self.startyear.has_key(book.ShadowSeries+str(book.ShadowYear)):
			startyear = self.startyear[book.ShadowSeries+str(book.ShadowYear)]
		else:
			startyear = book.ShadowYear
			
			for b in ComicRack.App.GetLibraryBooks():
				if b.ShadowSeries == book.ShadowSeries and b.ShadowVolume == book.ShadowVolume:
					
					#In case the initial values is bad
					if startyear == -1 and b.ShadowYear != 1:
						startyear = b.ShadowYear
					
	
					if b.ShadowYear != -1 and b.ShadowYear < startyear:
						startyear = b.ShadowYear
			
			#Store this final result in the dict so no calculation require for others of the series.
			self.startyear[book.ShadowSeries+str(book.ShadowYear)] = startyear
			
		if	startyear != -1:
			result = matchObj.group("pre") + str(startyear) + matchObj.group("post")
		else:
			if emptyreplacements["StartYear"] != "":
				result = matchObj.group("pre") + emptyreplacements["StartYear"] + matchObj.group("post")
			else:
				result = ""
		result = result.replace('/', ' ')
		result = result.replace('\\', ' ')
		return result

	def insertManga(self, matchObj):
		if emptyreplacements["Manga"].strip() != "":
			empty = matchObj.group("pre") + emptyreplacements["Manga"] + matchObj.group("post")
		else:
			empty = ""
		
		if book.Manga == MangaYesNo.Yes or book.Manga == MangaYesNo.YesAndRightToLeft:
			r = matchObj.group("text")
			if r.strip() != "":
				return matchObj.group("pre") + r + matchObj.group("post")

		return empty


	def insertSeriesComplete(self, matchObj):
		if emptyreplacements["SeriesComplete"].strip() != "":
			empty = matchObj.group("pre") + emptyreplacements["SeriesComplete"] + matchObj.group("post")
		else:
			empty = ""
		
		if book.SeriesComplete == YesNo.Yes:
			r = matchObj.group("text")
			if r.strip() != "":
				return matchObj.group("pre") + r + matchObj.group("post")

		return empty

	def insertMulti(self, matchObj):

		field = matchObj.group("name").capitalize()

		list = getattr(self, field)

		#Get a bool for if using series. Just simplifies the code a bit
		if matchObj.group("series") == "series":
			series = True
			index = book.ShadowSeries + str(book.ShadowVolume)
			booktext = book.ShadowSeries + " vol. " + str(book.ShadowVolume) 
		else:
			series = False
			index = book.ShadowSeries + str(book.ShadowVolume) + book.ShadowNumber
			booktext = book.ShadowSeries + " vol. " + str(book.ShadowVolume) + " #" + book.ShadowNumber

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

		#Try and get an existing value
		try:
			if series:
				return self.MakeMultiSelectionStringSeries(list[index], pre, post, sep, empty, field)
			else:
				return self.MakeMultiSelectionStringIssue(list[index], pre, post, sep, empty)

		except KeyError:
			#key not there so continue to find the writer to use
			pass
		
		#No writers and going by issue
		if getattr(book, field).strip() == "" and series == False :
			list[index] = SelectionFormResult([])
			return empty
		
		#There is writers or going by series
		else:
			items = []

			#Finding for whole series
			if series:
				items = self.GetAllFromSeriesField(field, book.ShadowSeries, book.ShadowVolume)

				#No items at all in the entire series
				if len(items) == 0:
					list[index] = SelectionFormResult([])
					return empty

			#finding just for the issue
			else:
				for i in getattr(book, field).split(","):
					items.append(i.strip())
			if self.form.InvokeRequired:
				result = self.form.Invoke(Func[SelectionFormArgs, SelectionFormResult](self.ShowSelectionForm), System.Array[object]([SelectionFormArgs(items, field, booktext, series)]))

			else:
				result = self.ShowSelectionForm(SelectionFormArgs(items, field, booktext, series))

			if series:
				list[index] = result
				return self.MakeMultiSelectionStringSeries(result, pre, post, sep, empty, field)

			else:
				list[index] = result
				return self.MakeMultiSelectionStringIssue(result, pre, post, sep, empty)


	def MakeMultiSelectionStringIssue(self, results, pre, post, sep, empty):
		"""
		Makes the correctly formated inserted string based on the input values:

		results is a SelectionFormResult class
		pre is the string prefix
		post is the string postfix
		sep is the string seperator
		empty is the string to return when the calculated string is empty
		"""
		if results.Folder:
			sep += "\\"

		string = sep.join(results.Selection)
		if string == "":
			return empty

		else:
			return pre + string + post

	def ShowSelectionForm(self, args):
		f = SelectionForm(args)
		f.ShowDialog()
		result = f.GetResults()
		f.Dispose()
		return result

	def MakeMultiSelectionStringSeries(self, results, pre, post, sep, empty, field):
		"""
		Makes the correctly formated inserted string based on the input values:

		results is a SelectionFormResult class
		pre is the string prefix
		post is the string postfix
		sep is the string seperator
		empty is the string to return when the calculated string is empty
		field is the correct name of the field to find from the book
		"""
		if results.Folder:
			sep += "\\"

		if results.EveryIssue:
			string = sep.join(results.Selection)
			if string == "":
				return empty
			else:
				return pre + string + post

		else:
			string = ""
			items = []
			for i in getattr(book, field).split(","):
				items.append(i.strip())
			
			itemsToUse = []
			for i in results.Selection:
				if i in items:
					itemsToUse.append(i)

			string = sep.join(itemsToUse)

			
			#Possible for the 

			if string == "":
				return empty
			else:
				return pre + string + post
			

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

		return results
