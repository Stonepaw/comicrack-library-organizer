using System.Collections.Generic;
using cYo.Projects.ComicRack.Engine;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class MangaYesNoMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        private MangaYesNo _value = MangaYesNo.Yes;

        public MangaYesNo Value
        {
            get => _value;
            set => Set(ref _value, value);
        }

        public static readonly IReadOnlyCollection<KeyValuePair<MangaYesNo, string>> Options =
            new List<KeyValuePair<MangaYesNo, string>>
            {
                new KeyValuePair<MangaYesNo, string>(MangaYesNo.Yes, "is Yes"),
                new KeyValuePair<MangaYesNo, string>(
                    MangaYesNo.YesAndRightToLeft,
                    "is Yes (Right to Left)"
                ),
                new KeyValuePair<MangaYesNo, string>(MangaYesNo.No, "is No"),
                new KeyValuePair<MangaYesNo, string>(MangaYesNo.Unknown, "is Unknown"),
            };

        public MangaYesNoMatcherValueViewModel() { }

        public MangaYesNoMatcherValueViewModel(IBookFieldMangaYesNoMatcher matcher)
        {
            _value = matcher.Value;
        }
    }
}
