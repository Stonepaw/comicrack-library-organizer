using System.Collections.Generic;
using System.Linq;

namespace LibraryOrganizer.Data
{
    public class IllegalCharacterReplacement
    {
        public IllegalCharacterReplacement(char character, string replacement)
        {
            Character = character;
            Replacement = replacement;
        }

        public char Character { get; set; }

        public string Replacement { get; set; }

        public bool IsRequired()
        {
            return IsRequiredIllegalCharacter(Character);
        }

        private static readonly char[] DefaultIllegalCharacters = new[]
        {
            '?',
            '/',
            '\\',
            '*',
            ':',
            '<',
            '>',
            '|',
            '"',
        };

        private static readonly IReadOnlyDictionary<char, string> DefaultReplacements =
            new Dictionary<char, string>
            {
                [':'] = " - ",
                ['<'] = "[",
                ['>'] = "]",
                ['|'] = "!",
                ['"'] = "'",
            };

        public static List<IllegalCharacterReplacement> DefaultList()
        {
            return DefaultIllegalCharacters
                .Select(character => new IllegalCharacterReplacement(
                    character,
                    DefaultReplacements.TryGetValue(character, out var value) ? value : ""
                ))
                .ToList();
        }

        private static readonly IReadOnlyCollection<char> RequiredIllegalCharacters =
            new HashSet<char>(DefaultIllegalCharacters);

        public static bool IsRequiredIllegalCharacter(char character)
        {
            return RequiredIllegalCharacters.Contains(character);
        }
    }
}
