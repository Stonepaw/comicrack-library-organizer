using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal interface IComicBookStringField : IComicBookField
    {
        string GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics);
    }
}
