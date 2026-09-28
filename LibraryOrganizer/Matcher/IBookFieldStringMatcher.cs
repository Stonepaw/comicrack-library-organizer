using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldStringMatcher : IBookFieldMatcher
    {
        IComicBookStringField Field { get; }

        BookFieldStringMatcherMode Mode { get; }

        string Value { get; }
    }
}
