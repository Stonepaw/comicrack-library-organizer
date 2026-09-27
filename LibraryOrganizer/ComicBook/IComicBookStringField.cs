namespace LibraryOrganizer.ComicBook
{
    internal interface IComicBookStringField : IComicBookField
    {
        string GetValue(
            cYo.Projects.ComicRack.Engine.ComicBook comicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics seriesStatistics
        );
    }
}
