using System.ComponentModel;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal interface IGroupMatcherViewModel : IMatcherViewModel
    {
        GroupMatcherMode Mode { get; set; }

        BindingList<IMatcherViewModel> Matchers { get; }

        void AddBookFieldMatcher();

        void AddBookFieldMatcher(IBookFieldMatcher matcher);

        void AddBookFieldMatcherAfter(IMatcherViewModel after);

        void AddGroup();

        void AddGroup(IGroupMatcher groupMatcher);

        void AddGroupAfter(IMatcherViewModel after);

        void Remove(IMatcherViewModel matcher);
    }
}
