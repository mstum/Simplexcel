using System;
using System.Text.RegularExpressions;

namespace Simplexcel
{
    internal static class CellAddressHelper
    {
        private readonly static Regex _cellAddressRegex = new Regex(@"(?<Column>[a-z]+)(?<Row>[0-9]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private readonly static char[] _chars = new[] { 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z' };

        /// <summary>
        /// Convert a zero-based row and column to an Excel Cell Reference, e.g. A1 for [0,0]
        /// </summary>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        internal static string ColRowToReference(int col, int row)
        {
            return ColToReference(col) + (row + 1);
        }

        /// <summary>
        /// Convert a zero-based column number to the proper Reference in Excel (e.g, 0 = A)
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        internal static string ColToReference(int col)
        {
            var dividend = (col + 1);
            string columnName = String.Empty;
            while (dividend > 0)
            {
                var modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier) + columnName;
                dividend = ((dividend - modifier) / 26);
            }
            return columnName;
        }

        /// <summary>
        /// Convert an Excel Cell Reference to a zero-based column/row representation (e.g., A1 => [0,0])
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>A Tuple where Item1 is the Column and Item2 is the Row</returns>
        internal static Tuple<int, int> ReferenceToColRow(string reference)
        {
            var parts = _cellAddressRegex.Match(reference);
            if (parts.Success)
            {
                var rowMatch = parts.Groups["Row"];
                var colMatch = parts.Groups["Column"];
                if (rowMatch.Success && colMatch.Success)
                {
                    int row = Convert.ToInt32(rowMatch.Value) - 1;

                    string colString = colMatch.Value.ToLower();
                    int col = 0;
                    for (int i = 0; i < colString.Length; i++)
                    {
                        char c = colString[colString.Length - i - 1];
                        var ix = Array.IndexOf(_chars, c);
                        if (ix == -1)
                        {
                            throw new ArgumentException("Cell Reference needs to be in Excel format, e.g. A1 or BA321. Invalid: " + reference, "reference");
                        }

                        var ixo = ix + 1;
                        var multiplier = Math.Pow(26, i);
                        col += (int)(ixo * multiplier);
                    }
                    int column = col - 1;

                    return Tuple.Create(column, row);
                }
            }

            throw new ArgumentException("Cell Reference needs to be in Excel format, e.g. A1 or BA321. Invalid: " + reference, "reference");
        }
    }
}
