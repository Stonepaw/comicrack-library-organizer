using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stonepaw.LibraryOrganizer
{
    public class Profile
    {
        public const string Version = "2.2.0";
        public string Name;
        public string FolderTemplate;
        public string BaseFolder;
        public string FileTemplate;
        public string FilelessFormat;

        public bool UseFolderOrganization = true;
        public bool UseFileNaming = true;
        public bool RemoveEmptyFolders = true;
        public bool MoveFileless = false;

        public List<string> ExcludedFolders = new List<string>();
        public List<string> ExclcudedEmptyFolders = new List<string>();
    }

    public class Profiles : ObservableCollection<Profile>
    {
    }
}
