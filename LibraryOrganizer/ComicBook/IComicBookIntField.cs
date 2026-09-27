namespace LibraryOrganizer.ComicBook
{
    internal interface IComicBookIntField : IComicBookField
    {
        int GetValue(
            cYo.Projects.ComicRack.Engine.ComicBook comicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics seriesStatistics
        );
    }
}
