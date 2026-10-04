using System.Collections.Generic;
using cYo.Projects.ComicRack.Engine;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class YesNoMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        private YesNo _value = YesNo.Yes;

        public YesNo Value
        {
            get => _value;
            set => Set(ref _value, value);
        }

        public static readonly IReadOnlyCollection<KeyValuePair<YesNo, string>> Options = new List<
            KeyValuePair<YesNo, string>
        >
        {
            new KeyValuePair<YesNo, string>(YesNo.Yes, "is Yes"),
            new KeyValuePair<YesNo, string>(YesNo.No, "is No"),
            new KeyValuePair<YesNo, string>(YesNo.Unknown, "is Unknown"),
        };

        public YesNoMatcherValueViewModel() { }

        public YesNoMatcherValueViewModel(IBookFieldYesNoMatcher matcher)
        {
            _value = matcher.Value;
        }
    }
}
