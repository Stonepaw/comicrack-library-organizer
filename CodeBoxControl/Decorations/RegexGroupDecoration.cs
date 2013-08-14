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
 public   class RegexGroupDecoration:Decoration 
    {
        public String RegexString { get; set; }

        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            Regex rx = new Regex(RegexString);
            MatchCollection mc = rx.Matches(Text);
            foreach (Match m in mc)
            {
                for (int i = 1; i <= m.Groups.Count; i++)
                {
			        foreach (Capture c in m.Groups[i].Captures)
                    {
                        pairs.Add(new Pair(c.Index, c.Length));
                    }
                }
            }
            return pairs;
        }
    }
}
