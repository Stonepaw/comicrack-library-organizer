using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal interface IComicBookYesNoField : IComicBookField
    {
        YesNo GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics);
    }
}
