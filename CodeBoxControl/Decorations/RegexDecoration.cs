using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Decoration based on a single regular expression string
    /// </summary>
 public   class RegexDecoration:Decoration 
    {
        public String RegexString { get; set; }

        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            Regex rx = new Regex(RegexString);
            MatchCollection mc = rx.Matches(Text);
            foreach (Match m in mc)
            {

                pairs.Add(new Pair(m.Index, m.Length));
            }
            return pairs;
        }
    }
}
