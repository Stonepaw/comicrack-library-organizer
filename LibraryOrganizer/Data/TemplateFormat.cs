namespace LibraryOrganizer.Data
{
    internal interface ITemplateFormat { }

    internal interface ICsvTemplateFormat : ITemplateFormat
    {
        bool SelectAll { get; }

        string Separator { get; }
    }

    internal interface IEmptyTemplateFormat : ITemplateFormat { }

    internal interface INumberTemplateFormat : ITemplateFormat
    {
        int Padding { get; }
    }

    internal interface IYesNoTemplateFormat : ITemplateFormat
    {
        bool Invert { get; }

        string Text { get; }
    }
}
