using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class BookFieldStringMatcherViewModel
        : BookFieldMatcherViewModel<IComicBookStringField>
    {
        public StringMatcherMode Mode { get; set; } = StringMatcherMode.Contains;

        public string Value { get; set; } = string.Empty;

        public BookFieldStringMatcherViewModel(
            IGroupMatcherViewModel parent,
            IComicBookStringField field
        )
            : base(parent, field) { }

        public BookFieldStringMatcherViewModel(
            IGroupMatcherViewModel parent,
            IBookFieldStringMatcher rule
        )
            : base(parent, rule.Field)
        {
            Mode = rule.Mode;
            Value = rule.Value;
            Negated = rule.Negated;
        }
    }
}
