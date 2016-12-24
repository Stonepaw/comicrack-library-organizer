import clr
import System
import sys

clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\ComicRack.Engine.dll")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\cYo.Common.dll")

from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter
from ComicRack import ComicRack
from System.IO import Path, FileInfo, FileStream, FileMode
import i18n

i18n.resourcespath = "..\\resources\\"
import duplicatewindow
gui = Path.GetFullPath("..\\gui")
duplicatewindow.ICONDIRECTORY = gui
from duplicatewindow import DuplicateWindow

c = ComicRack()
duplicatewindow.ComicRack = c
i18n.setup(c)
print "Loading test data"
f = FileStream("sample data.dat", FileMode.Open)
bf = BinaryFormatter()
books = bf.Deserialize(f)
print books
try:
    d = DuplicateWindow(c)
    c = d.ShowDialog(books[1], FileInfo(
        r"E:\Users\Andrew\Downloads\Ultimate Marvel\Miles Morales Ultimate Spider-Man\Miles Morales Ultimate "
        r"Spider-Man 001 (2014) (Digital) (Zone-Empire).cbr"),
                     "sadf", "Move", 2)
    c = d.ShowDialog(books[2], books[3], "asdfa", "Copy", 6)
    System.Console.ReadLine()
except Exception as ex:
    trace = sys.exc_info()[2].print_exc()
    print ex
    System.Console.ReadLine()
