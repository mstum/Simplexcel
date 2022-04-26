using System;
using System.Collections.Generic;

namespace Simplexcel
{
    /// <summary>
    /// A single Worksheet in an Excel Document
    /// </summary>
    public sealed partial class Worksheet
    {
        private readonly CellCollection _cells = new CellCollection();

        private readonly ColumnWidthCollection _columnWidth = new ColumnWidthCollection();

        private readonly PageSetup _pageSetup = new PageSetup();

        private List<SheetView> _sheetViews;
        private List<PageBreak> _rowBreaks;
        private List<PageBreak> _columnBreaks;

        /// <summary>
        /// Get a list of characters that are invalid to use in the Sheet Name
        /// </summary>
        /// <remarks>
        /// These chars are not part of the ECMA-376 standard, but imposed by Excel
        /// </remarks>
        public static readonly char[] InvalidSheetNameChars = new[] { '\\', '/', '?', '*', '[', ']' };

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
            LargeNumberHandlingMode = LargeNumberHandlingMode.StoreAsText;
        }

        /// <summary>
        /// Gets or sets the Name of the Worksheet
        /// </summary>
        public string Name { get; }

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
        /// How to handle numbers that are larger than <see cref="Cell.LargeNumberPositiveLimit"/> or smaller than <see cref="Cell.LargeNumberNegativeLimit"/>?
        /// </summary>
        public LargeNumberHandlingMode LargeNumberHandlingMode { get; set; }

        /// <summary>
        /// Whether to enable the AutoFilter feature on the first non-empty row.
        /// <para/>
        /// Defaults to <see langword="false"/>.
        /// </summary>
        /// <remarks>See <a href="https://support.microsoft.com/en-us/office/use-autofilter-to-filter-your-data-7d87d63e-ebd0-424b-8106-e2ab61133d92">Use AutoFilter to filter your data</a></remarks>
        public bool AutoFilter { get; set; }

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

        /// <summary>
        /// Freeze the top row, that is, create a <see cref="SheetView"/> that splits the first row (A) into a pane. 
        /// </summary>
        public void FreezeTopRow() => FreezeTopLeft(1, 0, Panes.BottomLeft);

        /// <summary>
        /// Freeze the first column, that is, create a <see cref="SheetView"/> that splits the first column (1) into a pane. 
        /// </summary>
        public void FreezeLeftColumn() => FreezeTopLeft(0, 1, Panes.TopRight);

        /// <summary>
        /// Freezes rows at the top and columns on the left.
        /// </summary>
        /// <param name="rows">The number of rows to freeze.</param>
        /// <param name="columns">The number of columns to freeze.</param>
        /// <param name="activePane">The pane that's selected in the sheet. Leave null to automatically determine this.</param>
        public void FreezeTopLeft(int rows, int columns, Panes? activePane = null)
        {
            // TODO: Eventually, support more SheetView functionality, right now, keep it simple.
            if (_sheetViews != null)
            {
                throw new InvalidOperationException("You have already frozen rows and/or columns on this Worksheet.");
            }

            if (rows < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rows), "Rows cannot be negative.");
            }

            if (columns < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columns), "Columns cannot be negative.");
            }

            if (columns == 0 && rows == 0)
            {
                return;
            }

            if (!activePane.HasValue)
            {
                if (columns == 0)
                {
                    activePane = Panes.BottomLeft;
                }
                else if (rows == 0)
                {
                    activePane = Panes.TopRight;
                }
                else
                {
                    activePane = Panes.BottomRight;
                }
            }

            var sheetView = new SheetView
            {
                Pane = new Pane
                {
                    ActivePane = activePane.Value,
                    XSplit = columns > 0 ? (int?)columns : null,
                    YSplit = rows > 0 ? (int?)rows : null,
                    TopLeftCell = CellAddressHelper.ColRowToReference(columns, rows),
                    State = PaneState.Frozen
                }
            };
            sheetView.AddSelection(new Selection { ActivePane = activePane.Value });
            AddSheetView(sheetView);
        }

        private void AddSheetView(SheetView sheetView)
        {
            if (_sheetViews == null)
            {
                _sheetViews = new List<SheetView>();
            }
            _sheetViews.Add(sheetView);
        }

        internal ICollection<SheetView> GetSheetViews()
        {
            return _sheetViews;
        }

        internal ICollection<PageBreak> GetRowBreaks()
        {
            return _rowBreaks;
        }

        internal ICollection<PageBreak> GetColumnBreaks()
        {
            return _columnBreaks;
        }

        /// <summary>
        /// Insert a manual page break after the row specified by the zero-based index (e.g., 0 for the first row)
        /// </summary>
        /// <param name="row">The zero-based row index (e.g., 0 for the first row)</param>
        public void InsertManualPageBreakAfterRow(int row)
        {
            if (_rowBreaks == null)
            {
                _rowBreaks = new List<PageBreak>();
            }

            _rowBreaks.Add(new PageBreak
            {
                Id = row,
                IsManualBreak = true
            });
        }

        /// <summary>
        /// Insert a manual page break to the left of the column specified by the zero-based index (e.g., 0 for column A)
        /// </summary>
        /// <param name="col">The zero-based index of the column (e.g., 0 for column A)</param>
        public void InsertManualPageBreakAfterColumn(int col)
        {
            if (_columnBreaks == null)
            {
                _columnBreaks = new List<PageBreak>();
            }

            _columnBreaks.Add(new PageBreak
            {
                Id = col,
                IsManualBreak = true
            });
        }
    }
}