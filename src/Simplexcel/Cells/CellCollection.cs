using System.Collections.Generic;

namespace Simplexcel
{
    /// <summary>
    /// A Collection of <see cref="Cell">Cells</see>
    /// </summary>
    public sealed class CellCollection : IEnumerable<KeyValuePair<CellAddress, Cell>>
    {
        private readonly Dictionary<CellAddress, Cell> _cells = new Dictionary<CellAddress, Cell>();

        // Todo: Ideally, this shouldn't return NULL but an Empty Cell, so that I can set formatting on an "empty" cell without a NullReferenceException.

        /// <summary>
        /// Get the cell with the given cell reference, e.g. Get the cell "A1". May return NULL.
        /// </summary>
        /// <param name="address"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[string address]
        {
            get
            {
                var ca = new CellAddress(address);
                return this[ca.Row, ca.Column];
            }
            set
            {
                var ca = new CellAddress(address);
                this[ca.Row, ca.Column] = value;
            }
        }

        /// <summary>
        /// Get the cell with the given zero based row and column, e.g. [0,0] returns the A1 cell. May return NULL.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[int row, int column]
        {
            get
            {
                var key = new CellAddress(row, column);
                if (_cells.ContainsKey(key)) return _cells[key];
                return null;
            }
            set
            {
                var key = new CellAddress(row, column);
                _cells[key] = value;
            }
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
