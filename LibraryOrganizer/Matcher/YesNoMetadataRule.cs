using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.Data
{
    public class YesNoMetadataRule : IMetadataRule
    {
        public YesNoField Field { get; set; }

        public YesNo Value { get; set; }
    }
}
