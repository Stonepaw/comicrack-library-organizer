namespace LibraryOrganizer.Data
{
    public class MonthReplacement
    {

        public MonthReplacement(int month, string replacement)
        {
            Month = month;
            Replacement = replacement;
        }

        public int Month { get; set; }

        public string Replacement { get; set; }
    }
}
