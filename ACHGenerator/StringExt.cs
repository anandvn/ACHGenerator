using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACHGenerator
{
    public static class StringExt
    {
        public static string StripSpecialChars(this string str)
        {
            var charsToRemove = new string[] { "@", ",", ".", ";", "'", "\"","<", ">", "+", "=", "[", "]", "^" };
            foreach (var c in charsToRemove)
            {
                str = str.Replace(c, string.Empty);
            }
            return str;
        }
    }
}