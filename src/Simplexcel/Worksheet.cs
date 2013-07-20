using System;
using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// A single Worksheet in an Excel Document
    /// </summary>
    [DataContract]
    public sealed class Worksheet
    {
        [DataMember]
        private readonly CellCollection _cells = new CellCollection();

        [DataMember]
        private readonly ColumnWidthCollection _columnWidth = new ColumnWidthCollection();

        [DataMember]
        private readonly PageSetup _pageSetup = new PageSetup();

        /// <summary>
        /// Get a list of characters that are invalid to use in the Sheet Name
        /// </summary>
        /// <remarks>
        /// These chars are not part of the ECMA-376 standard, but imposed by Excel
        /// </remarks>
        public static readonly char[] InvalidSheetNameChars = new[] {'\\', '/', '?', '*', '[', ']'};

        /// <summary>
        /// Get the maximum allowable length for a sheet name
        /// </summary>
        /// <remarks>
        /// This limit is not part of the ECMA-376 standard, but imposed by Excel
        /// </remarks>
        public static readonly int MaxSheetNameLength = 31;

        /// <summary>
        /// Create a new Worksheet with the given name
        /// </summary>
        /// <see cref="ArgumentException">Thrown if the name is null, empty, longer than <see cref="MaxSheetNameLength"/> characters or contains any character in <see cref="InvalidSheetNameChars"/></see>
        /// <param name="name"></param>
        public Worksheet(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("Sheet name is null or empty");
            }

            if (name.Length > MaxSheetNameLength)
            {
                throw new ArgumentException("Sheet name is longer than " + MaxSheetNameLength + " characters.");
            }

            if (name.IndexOfAny(InvalidSheetNameChars) > -1)
            {
                throw new ArgumentException("Sheet Names can not contain any of these characters: " + string.Join(" ", InvalidSheetNameChars));
            }

            Name = name;
        }

        /// <summary>
        /// Gets or sets the Name of the Worksheet
        /// </summary>
        [DataMember]
        public string Name { get; private set; }

        /// <summary>
        /// The Page Orientation and some other related values
        /// </summary>
        public PageSetup PageSetup
        {
            get { return _pageSetup; }
        }

        /// <summary>
        /// The Cells of the Worksheet (zero based, [0,0] = A1)
        /// </summary>
        public CellCollection Cells
        {
            get { return _cells; }
        }

        /// <summary>
        /// The Width of individual columns (Zero-based, in Excel's Units)
        /// </summary>
        public ColumnWidthCollection ColumnWidths
        {
            get { return _columnWidth; }
        }

        /// <summary>
        /// Get the cell with the given cell reference, e.g. Get the cell "A1". May return NULL.
        /// </summary>
        /// <param name="address"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[string address]
        {
            get { return Cells[address]; }
            set { Cells[address] = value; }
        }

        /// <summary>
        /// Get the cell with the given zero based row and column, e.g. [0,0] returns the A1 cell. May return NULL.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>The Cell, or NULL of the Cell hasn't been created yet.</returns>
        public Cell this[int row, int column]
        {
            get { return Cells[row, column]; }
            set { Cells[row, column] = value; }
        }
    }
}