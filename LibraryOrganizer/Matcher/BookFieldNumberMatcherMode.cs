using System.ComponentModel;

namespace LibraryOrganizer.Matcher
{
    internal enum BookFieldNumberMatcherMode
    {
        [Description("is")]
        Equal,

        [Description("is greater")]
        Greater,

        [Description("is smaller")]
        Less,

        [Description("is in the range")]
        Range,
    }
}
