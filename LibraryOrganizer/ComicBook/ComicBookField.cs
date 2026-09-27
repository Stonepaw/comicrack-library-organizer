using System.Collections.Generic;
using System.Linq;

namespace LibraryOrganizer.ComicBook
{
    internal class ComicBookField
    {
        public static readonly ComicBookIntField AlternateCount = ComicBookIntField.AlternateCount;
        public static readonly ComicBookStringField AlternateSeries =
            ComicBookStringField.AlternateSeries;
        public static readonly ComicBookIntField Count = ComicBookIntField.Count;
        public static readonly ComicBookIntField Day = ComicBookIntField.Day;
        public static readonly ComicBookIntField Month = ComicBookIntField.Month;
        public static readonly ComicBookIntField ReadPercentage = ComicBookIntField.ReadPercentage;
        public static readonly ComicBookStringField Publisher = ComicBookStringField.Publisher;
        public static readonly ComicBookStringField Series = ComicBookStringField.Series;
        public static readonly ComicBookIntField Year = ComicBookIntField.Year;

        public static IReadOnlyCollection<IComicBookField> Fields = ComicBookIntField
            .Fields.Concat<IComicBookField>(ComicBookStringField.Fields)
            .ToList()
            .AsReadOnly();
    }
}
