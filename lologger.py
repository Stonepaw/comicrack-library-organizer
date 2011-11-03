"""
lologger.py

Contains a class for logging what the bookmover does.

1.7.12: Fix for unicode characters

Copyright Stonepaw
"""

import clr
import System

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import MessageBox

class logger():
	
	def __init__(self):
		self._log = []
		self._failedcount = 0
		self._skippedcount = 0
		self._successcount = 0

	def Add(self, action, path, message = ""):
		self._log.append([unicode(action), unicode(path), unicode(message)])

	def ToArray(self):
		return self._log

	def Clear(self):
		del(self._log[:])

	def SaveLog(self, filepath):
		try:
			f = open(filepath,'w')
			f.write("Library Organizer Report:\n\nSkipped: %s\nFailed: %s\nSuccess: %s" % (self._failedcount, self._skippedcount, self._successcount))
			for array in self._log:
				f.write("\n\n%s: %s" % (array[0], array[1]))
				if array[2] != "":
					f.write("\nMessage: " + array[2])
			f.close()

		except Exception, ex:
			MessageBox.Show("something went wrong saving the file. The error is: " + str(ex))

	def SetCountVariables(self, failed, skipped, success):
		self._failedcount = failed
		self._skippedcount = skipped
		self._successcount = success