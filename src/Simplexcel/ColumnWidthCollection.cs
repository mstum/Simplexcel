using System.Collections;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// Custom Column Widths within a worksheet
    /// </summary>
    [DataContract]
    public sealed class ColumnWidthCollection : IEnumerable<KeyValuePair<int, double>>
    {
        [DataMember]
        private readonly Dictionary<int, double> _columnWidths = new Dictionary<int, double>();

        /// <summary>
        /// Get or set the width of a column (Zero-based column index, null value = auto)
        /// </summary>
        /// <param name="column">Zero-based column index</param>
        /// <returns></returns>
        public double? this[int column]
        {
            get { return _columnWidths.ContainsKey(column) ? _columnWidths[column] : (double?)null; }
            set
            {
                if (!value.HasValue)
                {
                    _columnWidths.Remove(column);
                }
                else
                {
                    _columnWidths[column] = value.Value;
                }
            }
        }

        /// <summary>
        /// Enumerate over the custom column widths. The Key is the zero-based column, the value is the custom column width.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<KeyValuePair<int, double>> GetEnumerator()
        {
            return _columnWidths.GetEnumerator();
        }

        /// <summary>
        /// Enumerate over the custom column widths. The Key is the zero-based column, the value is the custom column width.
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
