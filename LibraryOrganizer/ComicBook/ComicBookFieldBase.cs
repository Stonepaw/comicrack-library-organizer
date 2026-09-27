using LibraryOrganizer.Data;

namespace LibraryOrganizer.ComicBook
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
