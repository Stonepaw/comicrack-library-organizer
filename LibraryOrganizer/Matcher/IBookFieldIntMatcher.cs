using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldIntMatcher : IBookFieldMatcher
    {
        IComicBookIntField Field { get; }

        NumberMatcherMode Mode { get; }

        string Value { get; }
    }
}
