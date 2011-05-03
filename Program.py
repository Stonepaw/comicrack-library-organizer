import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Application, MessageBox, DialogResult
import loconfigForm
import System.IO
from System.IO import FileStream, FileMode, File, Directory, Path, StreamWriter, StreamReader
from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter
clr.AddReferenceToFile("ComicRack.Engine")
clr.AddReferenceToFile("cYo.Common")
import cYo.Common
import cYo.Projects.ComicRack.Engine
from cYo.Projects.ComicRack.Engine import *
import libraryorganizer
import cPickle
import losettings
import loworkerform
import loduplicate
import System.Text
from System.Text import StringBuilder
clr.AddReference("System.Xml")
import System.Xml
from System.Xml import XmlWriter, Formatting, XmlTextWriter, XmlWriterSettings, XmlDocument
clr.AddReference("System.Core")
from System import Func
import loconfigForm

try:
	#sfrom libraryorganizer import OverwriteAction
#	i.ShowDialog()
#	"""	print __file__[0:-len("Program.py")]
	
	s = libraryorganizer.LoadSettings()
	
	f = FileStream("Sample.dat", FileMode.Open)
	s = BinaryFormatter()
	books = s.Deserialize(f)
	#for book in bbooks:
	    #print book.ShadowSeries
	    #print book.Series
	    #print book.Number
	f.Close()
	
	settings, lastused = libraryorganizer.LoadSettings()
	#Get a random book to use as an example
	config = loconfigForm.ConfigForm(books, settings, lastused)
	result = config.ShowDialog()
	if result == DialogResult.OK:
		config.SaveSettings()
		lastused = config._cmbProfiles.SelectedItem
		#Now save the settings
		libraryorganizer.SaveSettings(settings, lastused)

	System.Console.ReadLine()

except Exception, ex:
	print "An error occured"
	print Exception
	print ex
	print type(ex)
	System.Console.ReadLine()
