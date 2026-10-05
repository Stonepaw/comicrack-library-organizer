using System.Collections.Generic;
using System.Linq;

namespace LibraryOrganizer.Data
{
    public class FailedEmptyField
    {
        public FailedEmptyField(string name, bool enabled)
        {
            Name = name;
            Enabled = enabled;
        }

        public bool Enabled { get; set; }

        public string Name { get; set; }

        private static readonly string[] _fieldNames =
        {
            "Age Rating",
            "Alternate Count",
            "Alternate Number",
            "Alternate Series",
            "Black And White",
            "Characters",
            "Colorist",
            "Count",
            "Cover Artist",
            "Editor",
            "Format",
            "Genre",
            "Imprint",
            "Inker",
            "Language",
            "Letterer",
            "Locations",
            "Main Character Or Team",
            "Manga",
            "Month",
            "Notes",
            "Number",
            "Penciller",
            "Publisher",
            "Rating",
            "Read Percentage",
            "Review",
            "Scan Information",
            "Series Complete",
            "Series Group",
            "Series",
            "Start Month",
            "Start Year",
            "Story Arc",
            "Tags",
            "Teams",
            "Title",
            "Volume",
            "Web",
            "Writer",
            "Year",
        };

        public static List<FailedEmptyField> DefaultList()
        {
            return new List<FailedEmptyField>(
                _fieldNames.Select((field) => new FailedEmptyField(field, false))
            );
        }
    }
}
