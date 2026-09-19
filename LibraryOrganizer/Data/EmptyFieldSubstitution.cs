namespace LibraryOrganizer.Data
{
    public class EmptyFieldSubstitution
    {
        public EmptyFieldSubstitution(string field, string substitution)
        {
            Field = field;
            Substitution = substitution;
        }

        /// <summary>
        /// The field name to substitute.
        /// </summary>
        public string Field { get; private set; }

        /// <summary>
        /// The substitution to use when the field is empty.
        ///
        /// This may be a null or an empty string for no substitution.
        /// </summary>
        public string Substitution { get; set; }
    }
}
