using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.Data
{
    public class MangaYesNoMetadataRule : IMetadataRule
    {
        public MangaYesNoField Field { get; set; }

        public MangaYesNo Value { get; set; }
    }
}
