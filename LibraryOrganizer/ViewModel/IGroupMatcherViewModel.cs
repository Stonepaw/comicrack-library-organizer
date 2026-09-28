using System.ComponentModel;
using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal interface IGroupMatcherViewModel : IMatcherViewModel
    {
        GroupMatcherMode Mode { get; set; }

        BindingList<IMatcherViewModel> Matchers { get; }

        void AddGroup();

        void AddGroup(IGroupMatcher groupMatcher);

        void AddBookFieldMatcher();

        void AddBookFieldMatcher(IBookFieldMatcher matcher);

        void ChangeFieldType(IBookFieldMatcherViewModel matcher, IComicBookField field);

        void Remove(IMatcherViewModel matcher);
    }
}
