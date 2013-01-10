namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// A Cell inside the sheet.xml file
    /// ECMA-376, 3rd Edition, Part 1, 18.3.1.4 c (Cell)
    /// </summary>
    internal class XlsxCell
    {
        /// <see cref="XlsxCellTypes"/>
        internal string CellType { get; set; }

        /// <summary>
        /// r (Reference) An "A1" style reference to the location of this cell
        /// The possible values for this attribute are defined by the ST_CellRef simple type (§18.18.7).
        /// </summary>
        internal string Reference { get; set; }

        /// <summary>
        /// s (StyleIndex)
        /// </summary>
        internal int StyleIndex { get; set; }

        /// <summary>
        /// The v element
        /// </summary>
        internal object Value { get; set; }
    }
}
