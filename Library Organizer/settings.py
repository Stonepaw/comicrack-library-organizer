import clr

clr.AddReference("System.Core")
clr.AddReference("System.Xml")
clr.AddReference("Stonepaw.LibraryOrganizer")

from System.Collections.ObjectModel import ObservableCollection
from System.IO import StreamWriter, File, StreamReader
from System.Xml.Serialization import XmlSerializer

from Stonepaw.LibraryOrganizer import Profile

from common import VERSION


def get_profiles():
    """Loads or creates new profiles and returns them.

    Loads the Stonepaw.LibraryOrganizer.Profile profiles from profiles.dat,
    if it exits. If the file does not exist then it creates a new default 
    profile.

    Returns:
        An ObservableCollection[Profile] containing the 
        Stonepaw.LibraryOrgnizer.Profile objects.
    """

    if File.Exists('profiles.dat'):
        x = XmlSerializer(ObservableCollection[Profile])
        with StreamReader('profiles.dat') as sr:
            profiles = x.Deserialize(sr)
        if profiles.Count == 0:
            profiles.Add(_create_default_profile())
    else:
        profiles = ObservableCollection[Profile]()
        profiles.Add(_create_default_profile())
    return profiles


def save_profiles(profiles):
    """Saves the profiles to profiles.dat.

    Args:
        profiles: The ObesrvableCollection containing all the profiles.
    """
    x = XmlSerializer(ObservableCollection[Profile])
    with StreamWriter('profiles.dat', False) as sw:
        x.Serialize(sw, profiles)


def _create_default_profile():
    """Creates a default Profile.

    Returns:
        A Stonepaw.LibraryOrganizer.Profile with a default name.
    """
    p = Profile('Default', VERSION)
    return p