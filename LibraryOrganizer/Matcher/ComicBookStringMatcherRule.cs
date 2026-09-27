using LibraryOrganizer.ComicBook;

namespace LibraryOrganizer.Data
{
    internal interface IComicBookStringMatcherRule : IComicBookMatcherRule
    {
        IComicBookStringField Field { get; }

        ComicBookStringMatcherOperator Operator { get; }

        string Value { get; }
    }

    internal class ComicBookStringMatcherRule : IComicBookStringMatcherRule
    {
        public IComicBookStringField Field { get; }

        public ComicBookStringMatcherOperator Operator { get; }

        public string Value { get; }

        public ComicBookStringMatcherRule(
            IComicBookStringField field,
            ComicBookStringMatcherOperator @operator,
            string value
        )
        {
            Field = field;
            Operator = @operator;
            Value = value;
        }
    }
}
