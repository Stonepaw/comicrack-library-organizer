using cYo.Projects.ComicRack.Engine;
using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class BookFieldYesNoMatcherViewModel : BookFieldMatcherViewModel<IComicBookYesNoField>
    {
        public YesNo Value { get; set; } = YesNo.Yes;

        public BookFieldYesNoMatcherViewModel(
            GroupMatcherViewModel parent,
            IComicBookYesNoField field
        )
            : base(parent, field) { }

        public BookFieldYesNoMatcherViewModel(
            IGroupMatcherViewModel parent,
            IBookFieldYesNoMatcher matcher
        )
            : base(parent, matcher.Field)
        {
            Value = matcher.Value;
        }
    }
}
