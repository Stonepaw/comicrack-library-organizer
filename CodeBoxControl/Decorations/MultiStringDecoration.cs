using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Decoration based on index positions of a list of strings
    /// </summary>
  public   class MultiStringDecoration:Decoration 
    {
        private List<string> mStrings = new List<string>();
        /// <summary>
        /// The list of strings to be searched for 
        /// </summary>
        public List<string> Strings
        {
            get { return mStrings; }
            set { mStrings = value; }
        }
        /// <summary>
        /// The System.StringComparison value to be used in searching 
        /// </summary>
        public StringComparison StringComparison { get; set; }
        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            foreach (string word in mStrings)
            {
                int index = Text.IndexOf(word, 0, StringComparison);
                while (index != -1)
                {
                    pairs.Add(new Pair(index, word.Length));
                    index = Text.IndexOf(word, index + word.Length, StringComparison);
                }
            }
            return pairs;
        }
    }
}
