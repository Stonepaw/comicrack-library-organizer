using System;
using System.Collections.Generic;
using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal class ComicBookIntField : ComicBookFieldBase, IComicBookIntField
    {
        public static readonly ComicBookIntField AlternateCount = new ComicBookIntField(
            "Alternate Count",
            ((book, seriesStatistics) => book.AlternateCount)
        );

        public static readonly ComicBookIntField Count = new ComicBookIntField(
            "Count",
            (book, seriesStatistics) => book.ShadowCount
        );

        public static readonly ComicBookIntField Day = new ComicBookIntField(
            "Day",
            (book, seriesStatistics) => book.Day
        );

        public static readonly ComicBookIntField Month = new ComicBookIntField(
            "Month",
            (book, seriesStatistics) => book.Month
        );

        public static readonly ComicBookIntField ReadPercentage = new ComicBookIntField(
            "ReadPercentage",
            (book, seriesStatistics) => book.ReadPercentage
        );

        public static readonly ComicBookIntField Year = new ComicBookIntField(
            "Year",
            (book, seriesStatistics) => book.ShadowYear
        );

        private readonly Func<ComicBook, ComicBookSeriesStatistics, int> _getValue;

        private ComicBookIntField(
            string label,
            Func<ComicBook, ComicBookSeriesStatistics, int> getValue
        )
            : base(label)
        {
            _getValue = getValue;
        }

        public int GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics) =>
            _getValue(comicBook, seriesStatistics);

        public static readonly IReadOnlyCollection<ComicBookIntField> Fields =
            new List<ComicBookIntField>
            {
                AlternateCount,
                Count,
                Day,
                Month,
                ReadPercentage,
                Year,
            }.AsReadOnly();
    }
}
