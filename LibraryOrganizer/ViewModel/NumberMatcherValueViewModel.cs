using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class NumberMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        private string _value = string.Empty;
        private NumberMatcherMode _mode = NumberMatcherMode.Equal;

        public NumberMatcherMode Mode
        {
            get => _mode;
            set => Set(ref _mode, value);
        }

        public string Value
        {
            get => _value;
            set => Set(ref _value, value);
        }

        public NumberMatcherValueViewModel() { }

        public NumberMatcherValueViewModel(IBookFieldIntMatcher bookFieldIntMatcher)
        {
            _mode = bookFieldIntMatcher.Mode;
            _value = bookFieldIntMatcher.Value;
        }
    }
}
