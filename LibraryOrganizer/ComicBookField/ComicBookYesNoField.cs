using System;
using System.Collections.Generic;
using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal class ComicBookYesNoField : IComicBookYesNoField
    {
        public static readonly ComicBookYesNoField BlackAndWhite = new ComicBookYesNoField(
            "Black and White",
            (comicBook, seriesStatistics) => comicBook.BlackAndWhite
        );

        public static readonly ComicBookYesNoField SeriesComplete = new ComicBookYesNoField(
            "Series Complete",
            (comicBook, seriesStatistics) => comicBook.SeriesComplete
        );

        public string Label { get; }

        private Func<ComicBook, ComicBookSeriesStatistics, YesNo> _getValue;

        private ComicBookYesNoField(
            string label,
            Func<ComicBook, ComicBookSeriesStatistics, YesNo> getValue
        )
        {
            Label = label;
            _getValue = getValue;
        }

        public YesNo GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics)
        {
            return _getValue(comicBook, seriesStatistics);
        }

        public static readonly IReadOnlyCollection<ComicBookYesNoField> Fields =
            new List<ComicBookYesNoField> { BlackAndWhite, SeriesComplete }.AsReadOnly();
    }
}
