namespace LibraryOrganizer.ComicBookField
{
    internal class ComicBookFieldBase : IComicBookField
    {
        protected ComicBookFieldBase(string label)
        {
            Label = label;
        }

        public string Label { get; }
    }
}
