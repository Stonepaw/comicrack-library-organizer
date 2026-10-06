using System;
using System.Collections.Generic;
using System.Linq;

namespace LibraryOrganizer.ComicBookField
{
    internal static class ComicBookField
    {
        public static readonly ComicBookIntField AlternateCount = ComicBookIntField.AlternateCount;
        public static readonly ComicBookStringField AlternateSeries =
            ComicBookStringField.AlternateSeries;
        public static readonly ComicBookIntField Count = ComicBookIntField.Count;
        public static readonly ComicBookIntField Day = ComicBookIntField.Day;
        public static readonly ComicBookMangaYesNoField Manga = ComicBookMangaYesNoField.Manga;
        public static readonly ComicBookIntField Month = ComicBookIntField.Month;
        public static readonly ComicBookIntField ReadPercentage = ComicBookIntField.ReadPercentage;
        public static readonly ComicBookStringField Publisher = ComicBookStringField.Publisher;
        public static readonly ComicBookStringField Series = ComicBookStringField.Series;
        public static readonly ComicBookIntField Year = ComicBookIntField.Year;

        /// <summary>
        /// All the supported fields configured for use in the application.
        /// </summary>
        public static IReadOnlyCollection<IComicBookField> Fields = new List<IComicBookField>()
            .Concat(ComicBookIntField.Fields)
            .Concat(ComicBookStringField.Fields)
            .Concat(ComicBookMangaYesNoField.Fields)
            .Concat(ComicBookYesNoField.Fields)
            .OrderBy((field) => field.Label)
            .ToList()
            .AsReadOnly();
    }
}
