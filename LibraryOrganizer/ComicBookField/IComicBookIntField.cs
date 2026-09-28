using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal interface IComicBookIntField : IComicBookField
    {
        int GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics);
    }
}
