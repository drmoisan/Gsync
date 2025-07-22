using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Gsync.Utilities.HelperClasses
{
    public static class SimpleRegex
    {
        /// <summary>
        /// Creates a simple Regex Pattern that treats '%' or '*' as wildcards. 
        /// Wildcards are enclosed in parenthesis to aid in the execution of 
        /// <seealso cref="SearchOptions.DeleteFromMatches"/>
        /// </summary>
        /// <param name="searchString"></param>
        /// <returns></returns>
        public static string MakeSearchPattern(string searchString)
        {
            string searchPattern = searchString.Replace("%", "*");
            if (searchPattern.Contains("*"))
            {
                searchPattern = string.Join("(.*)",
                                            searchPattern.Split('*')
                                            .Select(x => Regex.Escape(x)));
            }
            else
            {
                searchPattern = Regex.Escape(searchPattern);
            }
            searchPattern = "^" + searchPattern + "$";
            return searchPattern;
        }

        public static string MakeReplacePattern(string searchPattern)
        {
            int groupNum = searchPattern.Count(x => (x == '*'));
            string replacePattern = string.Join("", (from number in Enumerable.Range(1, groupNum)
                                                     select string.Concat("$", number.ToString())));
            return replacePattern;
        }

        public static (Regex rg, string searchPattern) MakeRegex(string searchString)
        {
            string searchPattern = MakeSearchPattern(searchString);
            Regex rg = new Regex(searchPattern, RegexOptions.IgnoreCase);

            return (rg, searchPattern);
        }

        public static string[] GetRegexGroups(this Regex regex, string input)
        {
            Match match = regex.Match(input);
            if (match.Success)
            {
                // Extract groups from the match
                string[] groups = new string[match.Groups.Count - 1];
                for (int i = 1; i < match.Groups.Count; i++)
                {
                    groups[i - 1] = match.Groups[i].Value;
                }
                return groups;
            }

            return [];
        }

    }
}
