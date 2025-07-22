using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Utilities.Extensions
{
    public static class ArrayExtensions
    {
        #region conversion and slicing extensions

        public static string[,] ToStringArray<T>(this T[,] array)
        {
            int rowCount = array.GetLength(0);
            int columnCount = array.GetLength(1);
            string[,] stringArray = new string[rowCount, columnCount];
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    if (array[i, j] is null) { stringArray[i, j] = ""; }
                    else { stringArray[i, j] = array[i, j].ToString(); }
                }
            }
            return stringArray;
        }
                
        public static string[] ToStringArray<T>(this T[] array)
        {
            int rowCount = array.Length;
            string[] stringArray = new string[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                stringArray[i] = array[i].ToString();
            }
            return stringArray;
        }

        public static string[,] ToStringArray<T>(this T[,] array, string nullReplacement)
        {
            int rowCount = array.GetLength(0);
            int columnCount = array.GetLength(1);
            string[,] stringArray = new string[rowCount, columnCount];
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    if (array[i, j] is null) 
                    {
                        stringArray[i, j] = nullReplacement;
                    }
                    else
                    {
                        stringArray[i, j] = array[i, j].ToString();
                    }
                }
            }
            return stringArray;
        }

        public static string[] ToStringArray<T>(this T[] array, string nullReplacement)
        {
            int rowCount = array.Length;
            string[] stringArray = new string[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                if(array[i] is null)
                {
                    stringArray[i] = nullReplacement;
                }
                else
                {
                    stringArray[i] = array[i].ToString();
                }
            }
            return stringArray;
        }

        public static IEnumerable<T> SliceRow<T>(this T[,] array, int row)
        {
            for (var i = 0; i < array.GetLength(1); i++)
            {
                yield return array[row, i];
            }
        }

        public static IEnumerable<T> SliceColumn<T>(this T[,] array, int column)
        {
            for (var i = 0; i < array.GetLength(0); i++)
            {
                yield return array[i, column];
            }
        }

        public static T[,] To2D<T>(this T[][] source)
        {
            try
            {
                int FirstDim = source.Length;
                int SecondDim = source.GroupBy(row => row.Length).Single().Key; // throws InvalidOperationException if source is not rectangular

                var result = new T[FirstDim, SecondDim];
                for (int i = 0; i < FirstDim; ++i)
                    for (int j = 0; j < SecondDim; ++j)
                        result[i, j] = source[i][j];

                return result;
            }
            catch (InvalidOperationException)
            {
                throw new InvalidOperationException("The given jagged array is not rectangular.");
            }
        }

        #endregion

        #region Allocation and Initialization

        public static bool IsInitialized<T>(this T[,] array)
        {
            if (array == null) { return false; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    if (array[i, j] == null) { return false; }
                }
            }

            return true;
        }

        public static bool IsInitialized<T>(this T[,] array, bool partially)
        {
            if (array == null) { return false; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    if (array[i, j] != null) { return true; }
                }
            }
            return false;
        }

        public static bool IsInitialized<T>(this T[] array)
        {
            if (array == null) { return false; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if (array[i] == null) { return false; }
            }

            return true;
        }

        public static bool IsInitialized<T>(this T[] array, bool partially)
        {
            if (array == null) { return false; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if (array[i] != null) { return true; }
            }
            return false;
        }


        #endregion

        #region Array indexing, lookup and search

        /// <summary>
        /// Function searches a string array and returns matching elements in a new string array.
        /// How matches are performed is defined by the search options. See 
        /// <see cref="StringSearchOptions"/>.
        /// </summary>
        /// <param name="sourceArray">String array to search</param>
        /// <param name="searchString">Target substring to search</param>
        /// <param name="options">Defines how matches are performed. 
        /// See also <see cref="StringSearchOptions"/>.</param>
        /// <returns>String array with elements that match the criteria. Null if no matches</returns>
        public static string[] SearchArry4Str(this string[] sourceArray, 
                                              string searchString = "", 
                                              SearchOptions options = SearchOptions.Standard)
        {
            if (searchString.Trim().Length != 0)
            {
                switch (options)
                {
                    case SearchOptions.Standard:
                        (Regex rg, string searchPattern) = SimpleRegex.MakeRegex(searchString);
                        return sourceArray.Where(x => rg.IsMatch(x)).ToArray();

                    case SearchOptions.Complement:
                        (rg, searchPattern) = SimpleRegex.MakeRegex(searchString);
                        return sourceArray.Where(x => !rg.IsMatch(x)).ToArray();

                    case SearchOptions.DeleteFromMatches:
                        (rg, searchPattern) = SimpleRegex.MakeRegex(searchString);
                        string replacePattern = SimpleRegex.MakeReplacePattern(searchPattern);
                        return sourceArray.Where(x => rg.IsMatch(x))
                                          .Select(x => rg.Replace(x, replacePattern))
                                          .ToArray();

                    case SearchOptions.ExactMatch:
                        return sourceArray.Where(x => x == searchString).ToArray();

                    case SearchOptions.ExactComplement:
                        return sourceArray.Where(x => x != searchString).ToArray();

                    default:
                        return sourceArray;
                }
            }
            return sourceArray;
        }

        /// <summary>
        /// Enumeration with search options.
        /// <list type="number">
        /// <listheader>
        ///     <term>Standard</term>
        ///     <description>Performs a simple regex search using a * or % as a wildcard</description>
        /// </listheader>
        /// <item>
        /// <term>Complement</term>
        /// <description>Elements that do NOT match the regex pattern will be returned</description>
        /// </item>
        /// <item>
        /// <term>DeleteFromMatches</term>
        /// <description>Similar to Standard except that the matching substring is removed from each 
        /// matching element</description>
        /// </item>
        /// <item>
        /// <term>ExactMatch</term>
        /// <description>Return elements that match the literal search string (case sensitive)</description>
        /// </item>
        /// <item>
        /// <term>ExactComplement</term>
        /// <description>Return elements that do Not match the literal search string (case sensitive)</description>
        /// </item>
        /// </list>
        /// </summary>
        public enum SearchOptions
        {
            Standard = 0,
            Complement = 1,
            DeleteFromMatches = 2,
            ExactMatch = 3,
            ExactComplement = 4
        }

        #endregion 

        
        public static T[] FlattenArrayTree<T>(this object node)
        {
            return FlattenArrayTree<T>(node, true).ToArray();
        }

        public static T[] TryFlattenArrayTree<T>(this object node)
        {
            try
            {
                return FlattenArrayTree<T>(node, false).ToArray();
            }
            catch (Exception)
            {
                return default;                
            }
            
        }

        internal static List<T> FlattenArrayTree<T>(this object node, bool strict) 
        { 
            if (strict || node.IsArray() || node is T) 
            { 
                List<T> result = new List<T>();
                FlattenArrayTree(node, strict, ref result);
                return result;
            }
            else { return null; }
        }
        
        internal static void FlattenArrayTree<T>(this object node, bool strict, ref List<T> result)
        {
            if (node.IsArray())
            {
                if (node.IsArray<T>())
                {
                    var branches = (T[])node;
                    result.AddRange(branches);
                }
                else
                {
                    var branches = (object[])node;
                    foreach (var branch in branches) 
                    { 
                        branch.FlattenArrayTree(strict, ref result); 
                    }
                }
            }
            else if (node is T) { result.Append((T)node);}
            else
            {
                if (strict)
                {
                    throw new ArgumentException($"node is of type {node.GetType().Name}. Array elements in " +
                                                $"{nameof(FlattenArrayTree)} must be arrays or of type {typeof(T).Name}.");
                }
                else 
                {
                    result.Add(default(T));
                }
            }
        }
        
        

        public static bool IsArray<T>(this object container) => container.GetType().IsArray && typeof(T).IsAssignableFrom(container.GetType().GetElementType());
        public static bool IsArray(this object container) => container.GetType().IsArray;

        //TODO: Implement IsTringArrayTree
        internal static bool IsStringArrayTree(this object[] branches, bool strictValidation)
        {
            return false;
        }

        public static string SentenceJoin(this IEnumerable<string> array, string separator = ", ", string lastSeparator = " and ")
        {
            var count = array.Count();
            if (count == 0) { return ""; }
            if (count == 1) { return array.ElementAt(0); }
            if (count == 2) { return array.ElementAt(0) + lastSeparator + array.ElementAt(1); }
            return string.Join(separator, array.Take(count - 1)) + lastSeparator + array.ElementAt(count - 1);
        }

        public static string SentenceJoin(this string[] array, string separator = ", ", string lastSeparator = " and ")
        {
            if (array.Length == 0) { return ""; }
            if (array.Length == 1) { return array[0]; }
            if (array.Length == 2) { return array[0] + lastSeparator + array[1]; }
            return string.Join(separator, array.Take(array.Length - 1)) + lastSeparator + array[array.Length - 1];
        }

        public static string SentenceJoin(this char[] array, string separator = ", ", string lastSeparator = " and ")
        {
            if (array.Length == 0) { return ""; }
            if (array.Length == 1) { return char.ToString(array[0]); }
            if (array.Length == 2) { return array[0] + lastSeparator + array[1]; }
            return string.Join(separator, array.Take(array.Length - 1)) + lastSeparator + array[array.Length - 1];
        }

        #region Deprecated

        [Obsolete("Use FlattenArrayTree instead")]
        public static string FlattenStringTree(this object[] branches, bool strictValidation = true)
        {
            if (!Array.TrueForAll(branches, branch => branch is string))
            {
                for (int i = 0; i < branches.Length; i++)
                {
                    if (branches[i] is Array) { branches[i] = FlattenStringTree((object[])branches[i]); }
                    else if (!(branches[i] is string))
                    {
                        if (strictValidation)
                        {
                            throw new ArgumentException($"branches[{i}] is of type {branches[i].GetType().Name}"
                                + $". Array elements in FlattenStringTree must be arrays or strings.");
                        }
                        branches[i] = "Error";
                    }
                }
            }
            string result = string.Join(", ", branches);
            if (result.Contains("Error")) { result = "Error"; }

            return result;
        }


        #endregion
    }

    public static class ArrayIsAllocated
    {
        public static bool IsAllocated(ref Array inArray)
        {
            bool FlagEx = true;
            try
            {
                if (inArray is null)
                {
                    FlagEx = false;
                }
                else if (inArray.Length <= 0)
                {
                    FlagEx = false;
                }
                else if (inArray.GetValue(0) == null)
                {
                    FlagEx = false;
                }
            }
            catch
            {
                FlagEx = false;
            }
            return FlagEx;
        }

        public static bool IsAllocated(ref string[] inArray)
        {
            bool FlagEx = true;
            try
            {
                if (inArray is null)
                {
                    FlagEx = false;
                }
                else if (inArray.Length <= 0)
                {
                    FlagEx = false;
                }
                else if (inArray[0] is null)
                {
                    FlagEx = false;
                }
            }
            catch (Exception)
            {
                FlagEx = false;
            }
            return FlagEx;
        }
    }

}
