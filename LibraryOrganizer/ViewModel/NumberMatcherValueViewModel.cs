using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class NumberMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        public static readonly IReadOnlyCollection<KeyValuePair<NumberMatcherMode, string>> Modes =
            Enum.GetValues(typeof(NumberMatcherMode))
                .Cast<NumberMatcherMode>()
                .Select(matcherMode => new KeyValuePair<NumberMatcherMode, string>(
                    matcherMode,
                    matcherMode
                        .GetType()
                        .GetField(matcherMode.ToString())
                        .GetCustomAttribute<DescriptionAttribute>()
                        ?.Description
                        ?? matcherMode.ToString()
                ))
                .ToList();

        private string _value = string.Empty;
        private string _value2 = string.Empty;
        private NumberMatcherMode _mode = NumberMatcherMode.Equal;

        public NumberMatcherMode Mode
        {
            get => _mode;
            set
            {
                bool beforeValue2Enabled = Value2Enabled;

                if (Set(ref _mode, value) && Value2Enabled != beforeValue2Enabled)
                {
                    NotifyPropertyChanged(nameof(Value2Enabled));
                }
            }
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

        public bool Value2Enabled => Mode == NumberMatcherMode.Range;

        public NumberMatcherValueViewModel() { }

        public NumberMatcherValueViewModel(IBookFieldIntMatcher bookFieldIntMatcher)
        {
            _mode = bookFieldIntMatcher.Mode;
            _value = bookFieldIntMatcher.Value;
            _value2 = bookFieldIntMatcher.Value2;
        }
    }
}
