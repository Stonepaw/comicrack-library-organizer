using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Decoration based on a list of strings sandwiched between word boundaries
    /// </summary>
 public   class MultiRegexWordDecoration:Decoration 
    {
        private List<string> mWords = new List<string>();
        public List<string> Words
        {
            get { return mWords; }
            set { mWords = value; }
        }
        public bool IsCaseSensitive { get; set; }

        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            foreach (string word in mWords)
            {
                string rstring = @"(?i:\b" + word + @"\b)";
                if (IsCaseSensitive)
                {
                    rstring = @"\b" + word + @"\b)";
                }
                Regex rx = new Regex(rstring);
                MatchCollection mc = rx.Matches(Text);
                foreach (Match m in mc)
                {
                    pairs.Add(new Pair(m.Index, m.Length));
                }
            }
            return pairs;
        }
    }
}
