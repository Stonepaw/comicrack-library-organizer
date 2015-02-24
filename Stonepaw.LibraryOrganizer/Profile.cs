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
        public string Version;
        public string Name { get; set; }
        public string BaseFolder { get; set; }
        public string FolderTemplate;
        public string FileTemplate;
        public string FilelessFormat;

        public bool UseFileNaming { get; set; }
        public bool UseFolderOrganization { get; set; }
        public bool AddCopiedBooksToLibrary { get; set; }
        public bool CopyFilelessThumbnails { get; set; }
        public bool RemoveEmptyFolders = true;

        public LibraryOrganizerMode Mode { get; set; }

        public List<string> ExcludedFolders = new List<string>();
        public List<string> ExclcudedEmptyFolders = new List<string>();
        public Dictionary<String, String> IllegalCharacters = new Dictionary<string, string>();

        public Profile()
        {
            this.UseFolderOrganization = true;
            this.UseFileNaming = true;
            this.Mode = LibraryOrganizerMode.Move;
        }

        public Profile(string name, string version) : this()
        {
            this.Name = name;
            this.Version = version;
        }

    }

    public class Profiles : ObservableCollection<Profile>
    {
    }
}
