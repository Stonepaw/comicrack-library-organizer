namespace LibraryOrganizer.ComicBook
{
    internal class ComicBookFieldUsage
    {
        public bool Rule { get; }

        public bool Template { get; }

        public ComicBookFieldUsage(bool rule, bool template)
        {
            Rule = rule;
            Template = template;
        }
    }
}
