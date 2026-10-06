using System;
using System.Collections.Generic;
using cYo.Projects.ComicRack.Engine;

namespace LibraryOrganizer.ComicBookField
{
    internal class ComicBookMangaYesNoField : IComicBookMangaYesNoField
    {
        public static readonly ComicBookMangaYesNoField Manga = new ComicBookMangaYesNoField(
            "Manga",
            (comicBook, seriesStatistics) => comicBook.Manga
        );

        public string Label { get; }

        private Func<ComicBook, ComicBookSeriesStatistics, MangaYesNo> _getValue;

        private ComicBookMangaYesNoField(
            string label,
            Func<ComicBook, ComicBookSeriesStatistics, MangaYesNo> getValue
        )
        {
            Label = label;
            _getValue = getValue;
        }

        public MangaYesNo GetValue(ComicBook comicBook, ComicBookSeriesStatistics seriesStatistics)
        {
            return _getValue(comicBook, seriesStatistics);
        }

        public static readonly IReadOnlyCollection<ComicBookMangaYesNoField> Fields =
            new List<ComicBookMangaYesNoField> { Manga }.AsReadOnly();
    }
}
