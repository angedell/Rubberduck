using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Ship
{
    public static class ShipCodeText
    {
        private static readonly Regex OptionStatement = new Regex(@"^\s*Option\s+(Explicit|Compare\s+\w+|Base\s+\d+|Private\s+Module)\s*$", RegexOptions.IgnoreCase);

        /// <summary>The code with every Option statement that repeats an earlier one taken out (with its line). VBA refuses a module that says <c>Option Explicit</c> twice.</summary>
        public static string WithoutRepeatedOptions(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return code;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var lines = code.Split(new[] { "\r\n" }, StringSplitOptions.None);
            var kept = lines.Where(line =>
            {
                var match = OptionStatement.Match(line);
                return !match.Success || seen.Add(Regex.Replace(match.Groups[1].Value, @"\s+", " "));
            });
            return string.Join("\r\n", kept);
        }
    }
}
