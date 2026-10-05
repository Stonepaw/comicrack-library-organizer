using System.Collections.Generic;
using System.Collections.ObjectModel;
using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.Data
{
    public class EmptyFieldReplacement
    {
        internal static List<EmptyFieldReplacement> Default()
        {
            return new List<EmptyFieldReplacement>
            {
                new EmptyFieldReplacement(ComicBookField.ComicBookField.AlternateCount, ""),
                new EmptyFieldReplacement(ComicBookField.ComicBookField.Count, ""),
            };
        }

        public EmptyFieldReplacement(IComicBookField field, string replacement)
        {
            Field = field;
            Replacement = replacement;
        }

        /// <summary>
        /// The field name to substitute.
        /// </summary>
        public IComicBookField Field { get; private set; }

        /// <summary>
        /// The field name extracted from the field.
        /// </summary>
        public string Label => Field.Label;

        /// <summary>
        /// The replacement to use when the field is empty.
        ///
        /// This may be a null or an empty string for no replacement.
        /// </summary>
        public string Replacement { get; set; }
    }
}
