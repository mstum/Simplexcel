using System.Collections.Generic;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// ECMA-376, 3rd Edition, Part 1, 18.3.1.73 row (Row)
    /// </summary>
    internal class XlsxRow
    {
        /// <summary>
        /// The c elements
        /// </summary>
        internal IList<XlsxCell> Cells { get; set; }

        /// <summary>
        /// This implicitly sets customHeight to 1 and ht to the value
        /// </summary>
        internal double? CustomRowHeight { get; set; }

        /// <summary>
        /// r (Row Index) Row index. Indicates to which row in the sheet this row definition corresponds.
        /// </summary>
        internal int RowIndex { get; set; }

        public XlsxRow()
        {
            Cells = new List<XlsxCell>();
        }
    }
}
