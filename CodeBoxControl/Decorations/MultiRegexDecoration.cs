using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace CodeBoxControl.Decorations
{
 public  class MultiRegexDecoration: Decoration 
    {
     private List<string> mRegexStrings = new List<string>();
     /// <summary>
     /// List of Strings defining the regular expressions
     /// </summary>
        public List<string> RegexStrings
        {
            get { return mRegexStrings; }
            set { mRegexStrings = value; }
        }

        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            foreach (string rString in mRegexStrings)
            {
                
                Regex rx = new Regex(rString);
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
