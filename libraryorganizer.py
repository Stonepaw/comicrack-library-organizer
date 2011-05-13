"""
libraryorganizer.py

The main script file. Some code is this file is based off of wadegiles's Guided eComic file renaming script. Credit is very much due to him

Version 1.4:
				
				Added Undo script

Author: Stonepaw

Copyright Stonepaw 2011. Some code copyright wadegiles Anyone is free to use code from this file as long as credit is given.
"""

import cPickle
import clr
import System

import System.IO
from System.IO import File, StreamReader, StreamWriter

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult

clr.AddReference("System.Xml")
import System.Xml
from System.Xml import XmlWriter, Formatting, XmlTextWriter, XmlWriterSettings, XmlDocument

import loconfigForm
from loconfigForm import ConfigForm

import losettings
from losettings import settings

import loworkerform
from loworkerform import ProfileSelector, WorkerForm, WorkerFormUndo

import locommon
from locommon import SETTINGSFILE, OLDSETTINGSFILE

import lobookmover

#@Name Library Organizer
#@Hook Books
#@Image libraryorganizer.png

def LibraryOrganizer(books):
	if books:
		try:
			settings, lastused = LoadSettings()
			loworkerform.ComicRack = ComicRack
			locommon.ComicRack = ComicRack
			lobookmover.ComicRack = ComicRack
			#Create the config form
			print "Creating config form"
			config = ConfigForm(books, settings, lastused)
			result = config.ShowDialog()
			if result == DialogResult.Cancel:
				#Dont save settings and quit the script
				return
			else:
				config.SaveSettings()
				lastused = config._cmbProfiles.SelectedItem
				#Now save the settings
				SaveSettings(settings, lastused)
				#Create a worker form
				#Note that all the moving files code is in the background worker of the worker 
				#form, It would be better it this could be done elsewhere but I don't have a 
				#very good understanding of threads so this was the easiest way for me to get 
				#this to work
				workerForm = WorkerForm(books, settings[lastused])
				workerForm.ShowDialog()
				workerForm.Dispose()
	
		except Exception, ex:
			print "The following error occured"
			print Exception
			print str(ex)

#@Name Library Organizer (Quick)
#@Hook Books
#@Image libraryorganizerquick.png
def LibraryOrganizerQuick(books):
	if books:
		try:
			loworkerform.ComicRack = ComicRack
			locommon.ComicRack = ComicRack
			lobookmover.ComicRack = ComicRack
			settings, lastused = LoadSettings()
			selected = lastused
			if len(settings) > 1:
				
				profileselector = ProfileSelector(settings.keys(), lastused)
				profileselector.ShowDialog()
				if profileselector.DialogResult == DialogResult.OK:
					selected = profileselector.Profile.SelectedItem
				else:
					print "user cancled opperation"
					return
			workerForm = WorkerForm(books, settings[selected])
			workerForm.ShowDialog()
			workerForm.Dispose()
				
		except Exception, ex:
			print "The following error occured"
			print Exception
			print str(ex)
		
#@Name Configure Library Organizer
#@Hook Library
#@Image libraryorganizer.png
def ConfigureLibraryOrganizer(books):
	if books:
		try:
			loworkerform.ComicRack = ComicRack
			locommon.ComicRack = ComicRack
			lobookmover.ComicRack = ComicRack
			settings, lastused = LoadSettings()
			#Get a random book to use as an example
			config = ConfigForm(books, settings, lastused)
			result = config.ShowDialog()
			if result == DialogResult.Cancel:
				#Dont save settings and quit the script
				return
			else:
				config.SaveSettings()
				lastused = config._cmbProfiles.SelectedItem
				#Now save the settings
				SaveSettings(settings, lastused)
		except Exception, ex:
			print "The Following error occured"
			print Exception
			print str(ex)

#@Name Library Organizer - Undo last move
#@Hook Library
#@Image libraryorganizer.png
def LibraryOrganizerUndo(books):
	if books:
		try:
			loworkerform.ComicRack = ComicRack
			locommon.ComicRack = ComicRack
			lobookmover.ComicRack = ComicRack
			settings, lastused = LoadSettings()
			if File.Exists(locommon.UNDOFILE):
				dict = locommon.LoadDict(locommon.UNDOFILE)
				if dict:
					workerForm = WorkerFormUndo(dict, settings[lastused])
					workerForm.ShowDialog()
					workerForm.Dispose()
					File.Delete(locommon.UNDOFILE)
				else:
					MessageBox.Show("Error loading Undo file")
			else:
				MessageBox.Show("Nothing to Undo")
		except Exception, ex:
			print "The following error occured"
			print Exception
			print str(ex)

def LoadSettings():
	settings = {}
	lastused = ""
	#Clean up old settings
	if File.Exists(OLDSETTINGSFILE):
		f = open(OLDSETTINGSFILE, 'r')
		settings["Default"] = cPickle.load(f)
		f.close()
		lastused = "Default"
		settings["Default"].Update()
		File.Delete(OLDSETTINGSFILE)
		return settings, lastused

	if File.Exists(SETTINGSFILE):
		try:
			file = StreamReader(SETTINGSFILE)
			xml = XmlDocument()
			xml.Load(file)
			nodes = xml.SelectNodes("Settings/Setting")
			if nodes.Count > 0:
				for i in nodes:
					settings[i.Attributes["Name"].Value] = losettings.settings()
					settings[i.Attributes["Name"].Value].Name = i.Attributes["Name"].Value
					settings[i.Attributes["Name"].Value].Load(i)
			lu = xml.SelectSingleNode("Settings")
			if lu.Attributes.Count > 0:
				lastused = lu.Attributes["LastUsed"].Value
			else:
				lastused = settings.keys()[0]
		except Exception, ex:
			print "Something seems to have gone wrong loading the xml file. Creating a blank settings file"
			print ex
			settings["Default"] = losettings.settings()
			settings["Default"].Name = "Default"
			lastused = "Default"
		finally:
			file.Close()
		
	if len(settings) == 0:
		#Just in case
		settings["Default"] = settings()
		settings["Default"].Name = "Default"
		lastused = "Default"
		
	return settings, lastused

def SaveSettings(settings, lastused):
	try:
		file = StreamWriter(SETTINGSFILE)
		xSettings = XmlWriterSettings()
		xSettings.Indent = True
		xWriter = XmlWriter.Create(file, xSettings)
		xWriter.WriteStartElement("Settings")
		xWriter.WriteAttributeString("LastUsed", lastused)
		for setting in settings:
			settings[setting].Save(xWriter)
		xWriter.WriteEndElement()
		xWriter.Close()
	except Exception, ex:
		print "An error occured writing the settings file. The error was:"
		print type(Exception)
		print str(ex)
	finally:
		file.Close()
	

	
