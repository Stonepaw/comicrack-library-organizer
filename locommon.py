"""
locommon.py

Author: Stonepaw

Version: 1.3
				Change on how duplicate paths are calculated: Thanks pescuma
                Changes from 1.1: Added catch for empty filename.

Contains several classes and functions. The book mover, path calculator and various misscellanous functions.


Copyright Stonepaw 2011. Anyone is free to use code from this file as long as credit is given.

"""




import clr

import re

import System

from System.Text import StringBuilder

from System.IO import Path, File, FileInfo, DirectoryInfo, Directory

from System import Func

clr.AddReference("System.Drawing")
from System.Drawing.Imaging import ImageFormat



SCRIPTDIRECTORY = __file__[0:-len("locommon.py")]

SETTINGSFILE = Path.Combine(SCRIPTDIRECTORY, "losettingsx.dat")

OLDSETTINGSFILE = Path.Combine(SCRIPTDIRECTORY, "losettings.dat")

ICON = Path.Combine(SCRIPTDIRECTORY, "libraryorganizer.ico")


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
		self.pathmaker = PathMaker()
		
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
			
			if Exclude(book, self.settings):
				#See if we need to skip the book, Exclude should return true if this is the case
				skipped += 1
				self.report.Append("\n\nSkipped moving %s because it qualified under the exclude rules." % (book.FilePath))
				self.worker.ReportProgress(count)
				continue
				
			if book.FilePath == "" and not self.settings.MoveFileless:
				#Fileless comic
				skipped += 1
				self.report.Append("\n\nSkipped moving book %s %s because it is a fileless book." % (book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			if book.FilePath == "" and self.settings.MoveFileless and not book.CustomThumbnailKey:
				#Fileless comic
				skipped += 1
				self.report.Append("\n\nSkipped creating image for filess book %s %s because it does not have a custom thumbnail." % (book.ShadowSeries, book.ShadowNumber))
				self.worker.ReportProgress(count)
				continue
			
			dirpath, filepath = self.CreatePath(book)

			if filepath == "":
                                failed += 1
                                self.report.Append("\n\nFailed to move %s because the filename created was empty. The book was not moved" % (book.FilePath))
                                self.worker.ReportProgress(count)
                                continue
			
			#Create here because needed for cleaning directories later
			oldpath = book.FileDirectory
			
			if not Directory.Exists(dirpath):
				try:
					Directory.CreateDirectory(dirpath)
				except Exception, ex:
					failed += 1
					self.report.Append("\n\nFailed to create path %s because an error occured. The error was: %s. Book %s was not moved." % (newpath, ex, book.FilePath))
					self.worker.ReportProgress(count)
					continue
					
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
				

			#Now that the directories are created the book can be moved
			result = self.MoveBook(book, Path.Combine(dirpath, filepath))
			if result == MoveResult.Success:
				success += 1
			
			elif result == MoveResult.Failed:
				failed += 1
			
			elif result == MoveResult.Skipped:
				skipped += 1
				
			#Cleanup old directories
			if self.settings.RemoveEmptyDir:
				self.CleanDirectories(DirectoryInfo(oldpath))
				self.CleanDirectories(DirectoryInfo(dirpath))
			
			self.worker.ReportProgress(count)
			
		#Return the report to the worker thread
		report = "Successfully moved: %s\nFailed to move: %s\nSkipped: %s" % (success, failed, skipped)
		return [failed + skipped, report, self.report.ToString()]
	
	def MoveFileless(self, book, path):
		
		oldpath = book.FilePath
		if File.Exists(path):
			oldbook = self.FindDuplicate(path)
			if not self.AlwaysDoAction:
				result = self.form.Invoke(Func[System.String, type(book), type(oldbook), list](self.form.AskOverwrite), System.Array[object]([path, book, oldbook]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == OverwriteAction.Cancel:
				self.report.Append("\n\nSkipped creating %s because a file already exists there and the user declined to overwrite it." % (path))
				return MoveResult.Skipped
			
			elif self.Action == OverwriteAction.Rename:
				#This bit writen by pescuma. Thanks!
				# Find an available name
				extension = Path.GetExtension(path)
				base = path[:-len(extension)]
				
				for i in range(100):
					newpath = base + " (" + str(i+1) + ")" + extension
					if not File.Exists(newpath):
						return self.MoveFileless(book, newpath)
				
				self.report.Append("\n\nFailed to find an available name to rename book. Book %s was not moved." % (path, ex, book.FilePath))
				return MoveResult.Failed
			
			elif self.Action == OverwriteAction.Overwrite:
				try:
					File.Delete(path)
				except Exception, ex:
					self.report.Append("\n\nFailed to overwrite %s. The error was: %s. Image was not created" % (path, ex, book.FilePath))
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
			image.Save(path, format)
			return MoveResult.Success
		except Exception, ex:
			self.report.Append("\n\nFailed to create image %s because an error occured. The error was: %s." % (path, ex))
			return MoveResult.Failed
	
	def MoveBook(self, book, path):
		#Return a result, 0 for success, 1 for failed, 2 for skipped
		oldpath = book.FilePath
		
		if not File.Exists(book.FilePath):
			self.report.Append("\n\nFailed to move %s because the file does not exist." % (book.FilePath))
			return MoveResult.Failed
		
		if path == book.FilePath:
			self.report.Append("\n\nSkipped moving book %s because it is already located at the calculated path." % (book.FilePath))
			return MoveResult.Skipped
		
		if File.Exists(path):
			oldbook = self.FindDuplicate(path)
			if not self.AlwaysDoAction:
				result = self.form.Invoke(Func[System.String, type(book), type(oldbook), list](self.form.AskOverwrite), System.Array[object]([path, book, oldbook]))
				self.Action = result[0]
				if result[1] == True:
					#User checked always do this opperation
					self.AlwaysDoAction = True
			
			if self.Action == OverwriteAction.Cancel:
				self.report.Append("\n\nSkipped moving %s because a file already exists at %s and the user declined to overwrite it." % (book.FilePath, path))
				return MoveResult.Skipped
			
			elif self.Action == OverwriteAction.Rename:
				#This bit writen by pescuma. Thanks!
				# Find an available name
				extension = Path.GetExtension(path)
				base = path[:-len(extension)]
				
				for i in range(100):
					newpath = base + " (" + str(i+1) + ")" + extension
					if not File.Exists(newpath) or newpath == book.FilePath:
						return self.MoveBook(book, newpath)
				
				self.report.Append("\n\nFailed to find an available name to rename book. Book %s was not moved." % (path, ex, book.FilePath))
				return MoveResult.Failed
			
			elif self.Action == OverwriteAction.Overwrite:
				try:
					File.Delete(path)
				except Exception, ex:
					self.report.Append("\n\nFailed to overwrite %s. The error was: %s. Book %s was not moved." % (path, ex, book.FilePath))
					return MoveResult.Failed
				
				#if the book is in the library, delete it
				if oldbook:
					ComicRack.App.RemoveBook(oldbook)
				
				return self.MoveBook(book, path)
			
		
		#Finally actually move the book
		try:
			File.Move(book.FilePath, path)
			book.FilePath = path
			return MoveResult.Success
		except Exception, ex:
			self.report.Append("\n\nFailed to move %s because an error occured. The error was: %s. The book was not moved." % (book.FilePath, ex))
			return MoveResult.Failed
	
	def CreatePath(self, book):
		dirpath = ""
		filepath = ""
		
		if self.settings.UseDirectory:
			dirpath = self.pathmaker.CreateDirectoryPath(book, self.settings.DirTemplate, self.settings.BaseDir, self.settings.EmptyDir, self.settings.EmptyData)
		
		else:
			dirpath = book.FileDirectory
			
		if self.settings.UseFileName:
			filepath = self.pathmaker.CreateFileName(book, self.settings.FileTemplate, self.settings.EmptyData, self.settings.FilelessFormat)
			
		else:
			filepath = book.FileNameWithExtension
			
		return dirpath, filepath
	
	def FindDuplicate(self, path):
		for b in ComicRack.App.GetLibraryBooks():
			if b.FilePath == path:
				return b
		return None
	
	def CleanDirectories(self, directory):
		#Driectory should be a directoryinfo object
		if not directory.Exists:
			return
		if len(directory.GetFiles()) == 0 and len(directory.GetDirectories()) == 0 and not directory.FullName in self.settings.ExcludedEmptyDir:
			parent = directory.Parent
			directory.Delete()
			self.CleanDirectories(parent)


def Exclude(book, settings):
	if ExcludePath(book, settings.ExcludeFolders):
		return True
	
	if ExcludeMeta(book, settings.ExcludeMetaData, settings.ExcludeOperator):
		return True
	
	return False

def ExcludePath(book, paths):
	for path in paths:
		if path in book.FilePath:
			return True
	return False

def ExcludeMeta(book, MetaSets, opperator):
	count = 0
	for set in MetaSets:
		#0 is the field, 1 is the opeartor, 2 is the text to match
		if set[1] == "is":
			#Convert to string just in case
			#Replace a space with nothing in the case of Alternate fields
			if str(getattr(book, set[0].replace(" ", ""))) == set[2]:
				count += 1
		elif set[1] == "does not contain":
			if set[2] not in str(getattr(book, set[0].replace(" ", ""))):
				count += 1
		elif set[1] == "contains":
			if set[2] in str(getattr(book, set[0].replace(" ", ""))):
				count += 1
		elif set[1] == "is not":
			if set[2] != str(getattr(book, set[0].replace(" ", ""))):
				count += 1
	
	if opperator == "Any":
		if count > 0:
			return True
		else:
			return False
	
	elif opperator == "All":
		if count == len(MetaSets):
			return True
		else:
			return False

class PathMaker:
	"""A class to create directory and file paths from the passed book"""
	
	#To speed up repeated calculations of start year
	startyear = {}
	
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
		
						
