﻿using System;
using Simplexcel.XlsxInternal;

namespace Simplexcel
{
    /// <summary>
    /// A cell inside a Worksheet
    /// </summary>
    public sealed class Cell
    {
        internal XlsxCellStyle XlsxCellStyle { get; private set; }

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
            XlsxCellStyle = new XlsxCellStyle();
            XlsxCellStyle.Format = format;

            Value = value;
            CellType = type;
        }

        /// <summary>
        /// The Excel Format for the cell, see <see cref="BuiltInCellFormat"/>
        /// </summary>
        public string Format
        {
            get { return XlsxCellStyle.Format; }
            set { XlsxCellStyle.Format = value; }
        }

        /// <summary>
        /// The border around the cell
        /// </summary>
        public CellBorder Border
        {
            get { return XlsxCellStyle.Border; }
            set { XlsxCellStyle.Border = value; }
        }

        /// <summary>
        /// The name of the Font (Default: Calibri)
        /// </summary>
        public string FontName
        {
            get { return XlsxCellStyle.Font.Name; }
            set { XlsxCellStyle.Font.Name = value; }
        }

        /// <summary>
        /// The Size of the Font (Default: 11)
        /// </summary>
        public int FontSize
        {
            get { return XlsxCellStyle.Font.Size; }
            set { XlsxCellStyle.Font.Size = value; }
        }

        /// <summary>
        /// Should the text be bold?
        /// </summary>
        public bool Bold
        {
            get { return XlsxCellStyle.Font.Bold; }
            set { XlsxCellStyle.Font.Bold = value; }
        }

        /// <summary>
        /// Should the text be italic?
        /// </summary>
        public bool Italic
        {
            get { return XlsxCellStyle.Font.Italic; }
            set { XlsxCellStyle.Font.Italic = value; }
        }

        /// <summary>
        /// Should the text be underlined?
        /// </summary>
        public bool Underline
        {
            get { return XlsxCellStyle.Font.Underline; }
            set { XlsxCellStyle.Font.Underline = value; }
        }

        /// <summary>
        /// The font color.
        /// </summary>
        public Color TextColor
        {
            get { return XlsxCellStyle.Font.TextColor; }
            set { XlsxCellStyle.Font.TextColor = value; }
        }

        /// <summary>
        /// The Type of the cell.
        /// </summary>
        public CellType CellType { get; private set; }

        /// <summary>
        /// The Content of the cell.
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Should this cell be a Hyperlink to something?
        /// </summary>
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

        /// <summary>
        /// Create a new <see cref="Cell"/> with a <see cref="CellType"/> of Date, formatted as DateAndTime.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static implicit operator Cell(DateTime value)
        {
            return new Cell(CellType.Date, value, BuiltInCellFormat.DateAndTime);
        }

        public static Cell FromObject(object val)
        {
            Cell cell;
            if (val is sbyte || val is short || val is int || val is long || val is byte || val is uint || val is ushort || val is ulong)
            {
                cell = new Cell(CellType.Number, Convert.ToDecimal(val), BuiltInCellFormat.NumberNoDecimalPlaces);
            }
            else if (val is float || val is double || val is decimal)
            {
                cell = new Cell(CellType.Number, Convert.ToDecimal(val), BuiltInCellFormat.General);
            }
            else if (val is DateTime)
            {
                cell = new Cell(CellType.Date, val, BuiltInCellFormat.DateAndTime);
            }
            else
            {
                cell = new Cell(CellType.Text, (val ?? String.Empty).ToString(), BuiltInCellFormat.Text);
            }
            return cell;
        }
    }
}
