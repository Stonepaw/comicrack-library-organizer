namespace LibraryOrganizer.Data
{
    public class EmptyFieldReplacement
    {
        public EmptyFieldReplacement(string field, string replacement)
        {
            Field = field;
            Replacement = replacement;
        }

        /// <summary>
        /// The field name to substitute.
        /// </summary>
        public string Field { get; private set; }

        /// <summary>
        /// The replacement to use when the field is empty.
        ///
        /// This may be a null or an empty string for no replacement.
        /// </summary>
        public string Replacement { get; set; }
    }
}
