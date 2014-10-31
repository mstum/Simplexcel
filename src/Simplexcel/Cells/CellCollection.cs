using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// A Collection of <see cref="Cell">Cells</see>
    /// </summary>
    [DataContract]
    public sealed class CellCollection : IEnumerable<KeyValuePair<CellAddress, Cell>>
    {
        [DataMember]
        private readonly Dictionary<CellAddress, Cell> _cells = new Dictionary<CellAddress, Cell>();

        /// <summary>
        /// How many Rows are in the sheet? (This counts the maximum row index, so empty rows are counted if they are followed by another row)
        /// </summary>
        public int RowCount
        {
            get { return _cells.Count > 0 ? _cells.Keys.Max(k => k.Row) + 1 : 0; }
        }

        /// <summary>
        /// How many Columns are in the sheet? (This counts the rightmost column, so empty columns are counted if they are followed by another column)
        /// </summary>
        public int ColumnCount
        {
            get { return _cells.Count > 0 ? _cells.Keys.Max(k => k.Column) + 1 : 0; }
        }

        /// <summary>
        /// Get the cell with the given cell reference, e.g. Get the cell "A1". May return NULL.
        /// </summary>
        /// <param name="address"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[string address]
        {
            get { return this[new CellAddress(address)]; }
            set { this[new CellAddress(address)] = value; }
        }

        /// <summary>
        /// Get the cell with the given zero based row and column, e.g. [0,0] returns the A1 cell. May return NULL.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[int row, int column]
        {
            get { return this[new CellAddress(row, column)]; }
            set { this[new CellAddress(row, column)] = value; }
        }

        private Cell this[CellAddress key]
        {
            get
            {
                if (!_cells.ContainsKey(key))
                    _cells[key] = string.Empty;

                return _cells[key];
            }
            set { _cells[key] = value; }
        }

        /// <summary>
        /// Enumerate over the Cells in this Collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator<KeyValuePair<CellAddress, Cell>> GetEnumerator()
        {
            return _cells.GetEnumerator();
        }

        /// <summary>
        /// Enumerate over the Cells in this Collection
        /// </summary>
        /// <returns></returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
