using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal interface IComicBookMangaYesNoField : IComicBookField
    {
        MangaYesNo GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics);
    }
}
