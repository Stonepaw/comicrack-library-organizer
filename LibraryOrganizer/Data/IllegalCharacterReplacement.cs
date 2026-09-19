namespace LibraryOrganizer.Data
{
    public class IllegalCharacterReplacement
    {

        public IllegalCharacterReplacement(string character, string replacement)
        {
            Character = character;
            Replacement = replacement;
        }

        public string Character { get; set; }

        public string Replacement { get; set; }
    }
}
