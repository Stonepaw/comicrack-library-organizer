using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class StringMatcherValueViewModel : ViewModelBase, IMatcherValueViewModel
    {
        public static readonly IReadOnlyCollection<
            KeyValuePair<StringMatcherMode, string>
        > StringOperators = Enum.GetValues(typeof(StringMatcherMode))
            .Cast<StringMatcherMode>()
            .Select(stringMatcherMode => new KeyValuePair<StringMatcherMode, string>(
                stringMatcherMode,
                stringMatcherMode
                    .GetType()
                    .GetField(stringMatcherMode.ToString())
                    .GetCustomAttribute<DescriptionAttribute>()
                    ?.Description
                    ?? stringMatcherMode.ToString()
            ))
            .ToList();

        private string _value = string.Empty;

        private StringMatcherMode _mode = StringMatcherMode.Contains;

        public StringMatcherMode Mode
        {
            get => _mode;
            set => Set(ref _mode, value);
        }

        public string Value
        {
            get => _value;
            set => Set(ref _value, value);
        }
    }
}
