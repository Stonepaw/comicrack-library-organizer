using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Matcher
{
    internal interface IBookFieldIntMatcher : IBookFieldMatcher
    {
        IComicBookIntField Field { get; }

        NumberMatcherMode Mode { get; }

        string Value { get; }

        /// <summary>
        /// The second value to use when using the range operator.
        /// </summary>
        string Value2 { get; }
    }
}
