import clr
import System
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\ComicRack.Engine.dll")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\cYo.Common.dll")

from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter
from ComicRack import ComicRack
from System.IO import Path, FileInfo, FileStream, FileMode
import i18n
i18n.resourcespath = ".\\resources\\"
import duplicateform
duplicateform.ICONDIRECTORY = Path.Combine(Path.GetDirectoryName(__file__), "GUI")
from duplicateform import DuplicateForm
i18n.setup(ComicRack())
print "Loading test data"
f = FileStream(".\\vs runner\\sample data.dat", FileMode.Open)
bf = BinaryFormatter()
books = bf.Deserialize(f)
print books
d = DuplicateForm()
d.ShowForm(books[1], FileInfo("C:\\test.txt"), "", 1)
System.Console.ReadLine()