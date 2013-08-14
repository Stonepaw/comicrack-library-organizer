using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodeBoxControl
{   
    /// <summary>
    /// Represents a character range with starting value and length
    /// </summary>
    public class Pair
    {
        /// <summary>
        /// First character in range, zero based
        /// </summary>
        public int Start { get; set; }
        /// <summary>
        /// Number of characters in range
        /// </summary>
        public int Length { get; set; }

        public Pair(){}
        
        public Pair(int start, int length)
        {
            Start = start;
            Length = length;
        }
    }
}
 