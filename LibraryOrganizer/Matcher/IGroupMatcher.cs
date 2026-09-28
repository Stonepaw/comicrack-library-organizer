using System.Collections.Generic;

namespace LibraryOrganizer.Matcher
{
    internal interface IGroupMatcher
    {
        IEnumerable<IMatcher> Matchers { get; }

        GroupMatcherMode Mode { get; }
    }
}
