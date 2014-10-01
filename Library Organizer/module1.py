import clr
import System
clr.AddReference("Stonepaw.LibraryOrganizer")
clr.AddReference("System.Xml")
import System
from System.Xml.Serialization import XmlSerializer, XmlSerializationWriter
from System.Xml import XmlWriter, XmlWriterSettings
from System.IO import TextWriter, FileStream, StreamWriter, StreamReader
from System import Console
from Stonepaw.LibraryOrganizer import Profile

#class Profile2(Profile):

#    def foo(self):
#        print "foo"

#p = Profile();
#p.FolderTemplate = "kasjdflkajsf"
#p.ExclcudedEmptyFolders.Add("laskjdfaslkjf")
#p.ExclcudedEmptyFolders.Add("lkajsfslkj")

x = StreamReader("text.txt")
s = XmlSerializer(Profile)
p = s.Deserialize(x)
print p
x.Close()
Console.WriteLine()
Console.ReadLine()
