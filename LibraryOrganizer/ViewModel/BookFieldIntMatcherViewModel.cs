using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class BookFieldIntMatcherViewModel : BookFieldMatcherViewModel<IComicBookIntField>
    {
        public NumberMatcherMode Mode { get; set; } = NumberMatcherMode.Equal;

        public string Value { get; set; } = string.Empty;

        public BookFieldIntMatcherViewModel(IGroupMatcherViewModel parent, IComicBookIntField field)
            : base(parent, field) { }

        public BookFieldIntMatcherViewModel(
            IGroupMatcherViewModel parent,
            IBookFieldIntMatcher matcher
        )
            : base(parent, matcher.Field)
        {
            Mode = matcher.Mode;
            Value = matcher.Value;
        }
    }
}
