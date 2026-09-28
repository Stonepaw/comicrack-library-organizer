namespace LibraryOrganizer.ViewModel
{
    internal interface IMatcherViewModel
    {
        /// <summary>
        /// If the rule is negated.
        /// </summary>
        bool Negated { get; set; }

        /// <summary>
        /// Deletes this matcher by removing it from its parent.
        /// </summary>
        void Delete();
    }
}
