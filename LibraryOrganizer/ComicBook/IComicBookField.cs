using LibraryOrganizer.Data;

namespace LibraryOrganizer.ComicBook
{
    internal interface IComicBookField
    {
        /// <summary>
        /// The label of the field to display in UI elements
        /// </summary>
        string Label { get; }
    }
}
