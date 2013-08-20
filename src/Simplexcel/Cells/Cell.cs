using System;
using System.Drawing;
using System.Runtime.Serialization;
using Simplexcel.XlsxInternal;

namespace Simplexcel
{
    /// <summary>
    /// A cell inside a Worksheet
    /// </summary>
    [DataContract]
    public sealed class Cell
    {
        [DataMember]
        private readonly XlsxCellStyle _xlsxCellStyle;
        internal XlsxCellStyle XlsxCellStyle { get { return _xlsxCellStyle; } }

        /// <summary>
        /// Create a new Cell of the given <see cref="CellType"/>.
        /// You can also implicitly create a cell from a string or number.
        /// </summary>
        /// <param name="cellType"></param>
        public Cell(CellType cellType) : this(cellType, null, BuiltInCellFormat.General) { }

        /// <summary>
        /// Create a new Cell of the given <see cref="CellType"/>, with the given value and format. For some common formats, see <see cref="BuiltInCellFormat"/>.
        /// You can also implicitly create a cell from a string or number.
        /// </summary>
        /// <param name="type"> </param>
        /// <param name="value"> </param>
        /// <param name="format"> </param>
        public Cell(CellType type, object value, string format)
        {
            _xlsxCellStyle = new XlsxCellStyle();
            _xlsxCellStyle.Format = format;

            Value = value;
            CellType = type;
        }

        /// <summary>
        /// The Excel Format for the cell, see <see cref="BuiltInCellFormat"/>
        /// </summary>
        public string Format
        {
            get { return _xlsxCellStyle.Format; }
            set { _xlsxCellStyle.Format = value; }
        }

        /// <summary>
        /// The border around the cell
        /// </summary>
        public CellBorder Border
        {
            get { return _xlsxCellStyle.Border; }
            set { _xlsxCellStyle.Border = value; }
        }

        /// <summary>
        /// The name of the Font (Default: Calibri)
        /// </summary>
        public string FontName
        {
            get { return _xlsxCellStyle.Font.Name; }
            set { _xlsxCellStyle.Font.Name = value; }
        }

        /// <summary>
        /// The Size of the Font (Default: 11)
        /// </summary>
        public int FontSize
        {
            get { return _xlsxCellStyle.Font.Size; }
            set { _xlsxCellStyle.Font.Size = value; }
        }

        /// <summary>
        /// Should the text be bold?
        /// </summary>
        public bool Bold
        {
            get { return _xlsxCellStyle.Font.Bold; }
            set { _xlsxCellStyle.Font.Bold = value; }
        }

        /// <summary>
        /// Should the text be italic?
        /// </summary>
        public bool Italic
        {
            get { return _xlsxCellStyle.Font.Italic; }
            set { _xlsxCellStyle.Font.Italic = value; }
        }

        /// <summary>
        /// Should the text be underlined?
        /// </summary>
        public bool Underline
        {
            get { return _xlsxCellStyle.Font.Underline; }
            set { _xlsxCellStyle.Font.Underline = value; }
        }

        /// <summary>
        /// The font color.
        /// </summary>
        public Color TextColor
        {
            get { return _xlsxCellStyle.Font.TextColor; }
            set { _xlsxCellStyle.Font.TextColor = value; }
        }

        /// <summary>
        /// The Type of the cell.
        /// </summary>
        // TODO: This is immutable because XlsxWriter.GetXlsxRows casts Value from object to whatever. This logic needs to change.
        [DataMember]
        public CellType CellType { get; private set; }

        /// <summary>
        /// The Content of the cell.
        /// </summary>
        [DataMember]
        public object Value { get; set; }

        /// <summary>
        /// Should this cell be a Hyperlink to something?
        /// </summary>
        [DataMember]
        public string Hyperlink { get; set; }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Text from a string.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(string value)
        {
            return new Cell(CellType.Text, value, BuiltInCellFormat.Text);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted without decimal places) from an integer.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(long value)
        {
            return new Cell(CellType.Number, Convert.ToDecimal(value), BuiltInCellFormat.NumberNoDecimalPlaces);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted with 2 decimal places) places from a decimal.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(Decimal value)
        {
            return new Cell(CellType.Number, value, BuiltInCellFormat.NumberTwoDecimalPlaces);
        }

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Number (formatted with 2 decimal places) places from a double.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(double value)
        {
            return new Cell(CellType.Number, Convert.ToDecimal(value), BuiltInCellFormat.NumberTwoDecimalPlaces);
        }
    }
}
