using System.Collections.Generic;

namespace LibraryOrganizer.Matcher
{
    internal interface IGroupMatcher : IMatcher
    {
        IEnumerable<IMatcher> Matchers { get; }

        GroupMatcherMode Mode { get; }
    }
}
