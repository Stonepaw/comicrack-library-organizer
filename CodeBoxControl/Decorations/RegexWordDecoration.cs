using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Decoration based on a string sandwiched between word boundaries
    /// </summary>
  public  class RegexWordDecoration:Decoration 
    {
        public string Word { get; set; }
        public bool IsCaseSensitive { get; set; }
     

        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            string rstring = "(?i:\b" + Word + "\b)";
            if (IsCaseSensitive)
            {
                rstring = "\b" + Word + "\b)";
            }
            Regex rx = new Regex(rstring);
            MatchCollection mc = rx.Matches(Text);
            foreach (Match m in mc)
            {

                pairs.Add(new Pair(m.Index, m.Length));
            }
            return pairs;
        }
    }
}
