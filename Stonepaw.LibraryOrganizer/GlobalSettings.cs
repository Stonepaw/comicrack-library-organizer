using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Stonepaw.LibraryOrganizer
{
    public class GlobalSettings
    {
        public string Version;
        public List<string> LastUsedProfiles = new List<string>();
        public bool GenerateLocalizedMonths = true;
        public bool RemoveEmptyFolders = true;
        public List<string> ExcludedEmptyFolders = new List<string>();
        public Dictionary<int,string> MonthNames = new Dictionary<int,string>();
        public Dictionary<string,string> IllegalCharacterReplacements = new Dictionary<string,string>();
    }
}
