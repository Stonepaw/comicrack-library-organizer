using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Decoration based on index positions of a single string
    /// </summary>
 public   class StringDecoration:Decoration 
    {
     /// <summary>
     /// The string to be searched for 
     /// </summary>
        public string String { get; set; }
     /// <summary>
     /// The System.StringComparison value to be used in searching 
     /// </summary>
        public StringComparison StringComparison { get; set; }


        public override List<Pair> Ranges(string Text)
        {
            List<Pair> pairs = new List<Pair>();
            int index = Text.IndexOf(String, 0, StringComparison);
            while (index != -1)
            {
                pairs.Add(new Pair(index, String.Length));
                index = Text.IndexOf(String, index +String.Length, StringComparison);
            }
            
            return pairs;
        }
    }
}
