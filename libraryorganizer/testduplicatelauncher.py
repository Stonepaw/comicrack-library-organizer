import clr
import System
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\ComicRack.Engine.dll")
clr.AddReferenceToFileAndPath("C:\\Program Files\\ComicRack\\cYo.Common.dll")

from System.Runtime.Serialization.Formatters.Binary import BinaryFormatter
from ComicRack import ComicRack
from System.IO import Path, FileInfo, FileStream, FileMode
import i18n
i18n.resourcespath = ".\\resources\\"
import duplicatewindow
duplicatewindow.ICONDIRECTORY = Path.Combine(Path.GetDirectoryName(__file__), "GUI")
from duplicatewindow import DuplicateWindow
duplicatewindow.ComicRack = ComicRack()
i18n.setup(ComicRack())
print "Loading test data"
f = FileStream(".\\vs runner\\sample data.dat", FileMode.Open)
bf = BinaryFormatter()
books = bf.Deserialize(f)
print books
try:
    d = DuplicateWindow()
    c = d.ShowDialog(books[1], FileInfo(r"E:\Users\Andrew\Downloads\Ultimate Marvel\Miles Morales Ultimate Spider-Man\Miles Morales Ultimate Spider-Man 001 (2014) (Digital) (Zone-Empire).cbr"), "sadf", "Move", 2)
    c = d.ShowDialog(books[2],books[3] , "asdfa", "Copy", 6)
    System.Console.ReadLine()
except Exception as ex:
    print ex
    System.Console.ReadLine()