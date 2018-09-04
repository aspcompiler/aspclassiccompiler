using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Dlrsoft.VBScript.Runtime
{
    public static class StringExtensionsClass
    {
        public static string RemoveNonNumeric(this string s)
        {
            MatchCollection col = Regex.Matches(s, "[0-9]");
            StringBuilder sb = new StringBuilder();
            foreach (Match m in col)
                sb.Append(m.Value);
            return sb.ToString();
        }

        public static StringBuilder op_Addition(this StringBuilder sb, object o)
        {
            sb.Append(o);
            return sb;
        }
    }
}
