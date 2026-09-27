using System;
using System.Collections.Generic;
using System.Linq;

namespace LibraryOrganizer.ComicBook
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

        private readonly Func<
            cYo.Projects.ComicRack.Engine.ComicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics,
            int
        > _getValue;

        private ComicBookIntField(
            string label,
            Func<
                cYo.Projects.ComicRack.Engine.ComicBook,
                cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics,
                int
            > getValue
        )
            : base(label)
        {
            _getValue = getValue;
        }

        public int GetValue(
            cYo.Projects.ComicRack.Engine.ComicBook comicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics seriesStatistics
        ) => _getValue(comicBook, seriesStatistics);

        public static IReadOnlyCollection<ComicBookIntField> Fields = new List<ComicBookIntField>
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
