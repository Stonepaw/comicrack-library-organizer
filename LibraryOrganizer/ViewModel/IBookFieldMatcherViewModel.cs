using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Data;

namespace LibraryOrganizer.ViewModel
{
    internal interface IBookFieldMatcherViewModel : IMatcherViewModel
    {
        IComicBookField Field { get; set; }
    }
}
