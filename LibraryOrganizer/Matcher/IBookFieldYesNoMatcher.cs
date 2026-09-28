using cYo.Projects.ComicRack.Engine;
using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldYesNoMatcher
    {
        IComicBookYesNoField Field { get; }

        YesNo Value { get; }
    }
}
