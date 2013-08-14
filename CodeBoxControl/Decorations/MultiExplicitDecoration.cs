using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodeBoxControl.Decorations
{
  public  class MultiExplicitDecoration:Decoration 
    {
        private List<Pair> mColoredRanges = new List<Pair>();
        public List<Pair> ColoredRanges
        {
            get { return mColoredRanges; }
            set { mColoredRanges = value; }

        }


        public override List<Pair> Ranges(string Text)
        {
            return mColoredRanges;
        }
    }
}
