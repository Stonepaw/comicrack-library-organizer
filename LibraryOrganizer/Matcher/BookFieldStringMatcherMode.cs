using System.ComponentModel;

namespace LibraryOrganizer.Matcher
{
    internal enum BookFieldStringMatcherMode
    {
        [Description("contains")]
        Contains,

        [Description("contains all of")]
        ContainsAllOf,

        [Description("contains any of")]
        ContainsAnyOf,

        [Description("ends with")]
        EndsWith,

        [Description("is")]
        Equal,

        [Description("regular expression")]
        RegularExpression,

        [Description("starts with")]
        StartsWith,
    }
}
