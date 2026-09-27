using System;
using System.Collections.Generic;

namespace LibraryOrganizer.ComicBook
{
    internal class ComicBookStringField : ComicBookFieldBase, IComicBookStringField
    {
        public static readonly ComicBookStringField AgeRating = new ComicBookStringField(
            "Age Rating",
            (book, seriesStatistics) => book.AgeRating
        );

        public static readonly ComicBookStringField AlternateNumber = new ComicBookStringField(
            "Alternate Number",
            (book, seriesStatistics) => book.AlternateNumberAsText
        );

        public static readonly ComicBookStringField AlternateSeries = new ComicBookStringField(
            "Alternate Series",
            (book, seriesStatistics) => book.AlternateSeries
        );

        public static readonly ComicBookStringField FileFormat = new ComicBookStringField(
            "File Format",
            (book, seriesStatistics) => book.FileFormat
        );

        public static readonly ComicBookStringField FilePath = new ComicBookStringField(
            "File Path",
            (book, seriesStatistics) => book.FilePath
        );

        public static readonly ComicBookStringField Format = new ComicBookStringField(
            "Format",
            (book, seriesStatistics) => book.ShadowFormat
        );

        public static readonly ComicBookStringField Imprint = new ComicBookStringField(
            "Imprint",
            (book, seriesStatistics) => book.Imprint
        );

        public static readonly ComicBookStringField LanguageISO = new ComicBookStringField(
            "Language ISO",
            (book, seriesStatistics) => book.LanguageISO
        );

        public static readonly ComicBookStringField Notes = new ComicBookStringField(
            "Notes",
            (book, seriesStatistics) => book.Notes
        );

        public static readonly ComicBookStringField Number = new ComicBookStringField(
            "Number",
            (book, seriesStatistics) => book.ShadowNumber
        );

        public static readonly ComicBookStringField Publisher = new ComicBookStringField(
            "Publisher",
            (book, seriesStatistics) => book.Publisher
        );

        public static readonly ComicBookStringField Review = new ComicBookStringField(
            "Publisher",
            (book, seriesStatistics) => book.Review
        );

        public static readonly ComicBookStringField Series = new ComicBookStringField(
            "Series",
            (book, seriesStatistics) => book.ShadowSeries
        );

        private readonly Func<
            cYo.Projects.ComicRack.Engine.ComicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics,
            string
        > _getValue;

        private ComicBookStringField(
            string label,
            Func<
                cYo.Projects.ComicRack.Engine.ComicBook,
                cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics,
                string
            > getValue
        )
            : base(label)
        {
            _getValue = getValue;
        }

        public string GetValue(
            cYo.Projects.ComicRack.Engine.ComicBook comicBook,
            cYo.Projects.ComicRack.Engine.ComicBookSeriesStatistics seriesStatistics
        ) => _getValue(comicBook, seriesStatistics);

        public static IReadOnlyCollection<ComicBookStringField> Fields =
            new List<ComicBookStringField>
            {
                AgeRating,
                AlternateNumber,
                AlternateSeries,
                FileFormat,
                FilePath,
                Format,
                Imprint,
                LanguageISO,
                Notes,
                Number,
                Publisher,
                Review,
                Series,
            }.AsReadOnly();
    }
}
