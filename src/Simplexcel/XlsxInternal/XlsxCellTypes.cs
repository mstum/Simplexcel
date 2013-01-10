namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// ECMA-376, 3rd Edition, Part 1, 18.18.11 ST_CellType (Cell Type)
    /// </summary>
    internal static class XlsxCellTypes
    {
        /// <summary>
        /// b (Boolean) Cell containing a boolean.
        /// </summary>
        internal static readonly string Boolean = "b";
        
        /// <summary>
        /// d (Date) Cell contains a date in the ISO 8601 format.
        /// </summary>
        internal static readonly string Date = "d";

        /// <summary>
        /// e (Error) Cell containing an error.
        /// </summary>
        internal static readonly string Error = "e";

        /// <summary>
        /// inlineStr (Inline String) Cell containing an (inline) rich string, i.e., one not in the shared string table.
        /// Note: Excel expects "str" instead of "inlineStr"
        /// </summary>
        internal static readonly string InlineString = "str";

        /// <summary>
        /// n (Number) Cell containing a number.
        /// </summary>
        internal static readonly string Number = "n";

        /// <summary>
        /// s (Shared String) Cell containing a shared string.
        /// </summary>
        internal static readonly string SharedString = "s";

        /// <summary>
        /// str (String) Cell containing a formula string.
        /// </summary>
        internal static readonly string FormulaString = "str";
    }
}