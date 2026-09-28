using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldStringMatcher : IBookFieldMatcher
    {
        IComicBookStringField Field { get; }

        StringMatcherMode Mode { get; }

        string Value { get; }
    }
}
