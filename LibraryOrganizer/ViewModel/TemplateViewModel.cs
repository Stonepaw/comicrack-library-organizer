using System;
using LibraryOrganizer.Data;

namespace LibraryOrganizer.ViewModel
{
    internal interface ITemplateFormatViewModel { }

    internal class CsvTemplateFormatViewModel : ICsvTemplateFormat, ITemplateFormatViewModel
    {
        public bool SelectAll { get; set; } = false;

        public string Separator { get; set; } = string.Empty;
    }

    internal class EmptyInsertTemplateFormatViewModel
        : IEmptyTemplateFormat,
            ITemplateFormatViewModel { }

    internal class NumberInsertTemplateFormatConfig
        : INumberTemplateFormat,
            ITemplateFormatViewModel
    {
        public int Padding { get; set; }
    }

    internal class YesNoInsertTemplateFormatConfig : IYesNoTemplateFormat, ITemplateFormatViewModel
    {
        public bool Invert { get; set; } = false;

        public string Text { get; set; } = string.Empty;
    }

    internal class TemplateViewModel : ViewModelBase
    {
        public string Field { get; }

        public string Label { get; }

        public ITemplateFormatViewModel Format { get; }

        private string _suffix = string.Empty;

        public string Prefix
        {
            get => _suffix;
            set => Set(ref _suffix, value);
        }

        public string Suffix { get; set; } = string.Empty;

        public TemplateViewModel(string field, string label, ITemplateFormatViewModel format)
        {
            Field = field;
            Format = format;
            Label = label;
        }
    }
}
