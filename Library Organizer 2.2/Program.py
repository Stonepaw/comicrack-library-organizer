import clr
import System
clr.AddReference("System.Core")
clr.AddReference("System.Xml")
from System.Xml.Serialization import XmlSerializer
from System.IO import StreamReader, StreamWriter, File

clr.AddReference("Stonepaw.LibraryOrganizer")
from Stonepaw.LibraryOrganizer import Profile, Profiles

from System.Collections.ObjectModel import ObservableCollection

x = XmlSerializer(Profiles)
#Load the settings file if it exists
if File.Exists("profiles.dat"):
    s = StreamReader("profiles.dat")
    profiles = x.Deserialize(s)
    s.Close()
    print "loaded profiles"
else:
    profiles = Profiles()
    p = Profile()
    p.Name = "Default"
    p.FolderTemplate = "laksjdf;laskjf'"
    profiles.Add(p)
    print "created new profiles"

print profiles
print type(profiles)
print profiles.__class__
System.Console.ReadLine()

s = StreamWriter("profiles.dat")
x.Serialize(s,profiles)
s.Close()

