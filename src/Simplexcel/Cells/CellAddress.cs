using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// The zero-based Address of a cell
    /// </summary>
    /// <remarks>
    /// This is zero-based, so cell A1 is Row,Column [0,0]
    /// </remarks>
    [DataContract]
    public struct CellAddress
    {
        /// <summary>
        /// The zero-based row
        /// </summary>
        [DataMember]
        public readonly int Row;

        /// <summary>
        /// The zero-based column
        /// </summary>
        [DataMember]
        public readonly int Column;

        /// <summary>
        /// Create a Cell Adress given the zero-based row and column
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        public CellAddress(int row, int column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// Create a CellAddress from a Reference like A1
        /// </summary>
        /// <param name="cellAddress"></param>
        public CellAddress(string cellAddress)
        {
            var cr = CellAddressHelper.ReferenceToColRow(cellAddress);
            Column = cr.Item1;
            Row = cr.Item2;
        }

        /// <summary>
        /// Convert this CellAddress into a Cell Reference, e.g. A1 for [0,0]
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return CellAddressHelper.ColRowToReference(Column, Row);
        }
    }
}
