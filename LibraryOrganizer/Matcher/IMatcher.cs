namespace LibraryOrganizer.Matcher
{
    public interface IMatcher
    {
        /// <summary>
        /// If the matcher should negate the result of the match
        /// </summary>
        bool Negated { get; }
    }
}
