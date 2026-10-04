using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class NumberMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        private string _value = string.Empty;
        private string _value2 = string.Empty;
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

        public string Value2
        {
            get => _value2;
            set => Set(ref _value2, value);
        }

        public NumberMatcherValueViewModel() { }

        public NumberMatcherValueViewModel(IBookFieldIntMatcher bookFieldIntMatcher)
        {
            _mode = bookFieldIntMatcher.Mode;
            _value = bookFieldIntMatcher.Value;
            _value2 = bookFieldIntMatcher.Value2;
        }
    }
}
