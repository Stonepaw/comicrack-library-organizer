using cYo.Projects.ComicRack.Engine;
using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldMangaYesNoMatcher : IBookFieldMatcher
    {
        IComicBookYesNoField Field { get; }

        MangaYesNo Value { get; }
    }
}
