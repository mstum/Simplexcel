﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Simplexcel
{
    /// <summary>
    /// A single Worksheet in an Excel Document
    /// </summary>
    public sealed class Worksheet
    {
        private readonly CellCollection _cells = new CellCollection();

        private readonly ColumnWidthCollection _columnWidth = new ColumnWidthCollection();

        private readonly PageSetup _pageSetup = new PageSetup();

        private List<SheetView> _sheetViews;

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
        }

        /// <summary>
        /// Gets or sets the Name of the Worksheet
        /// </summary>
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

        /// <summary>
        /// Freeze the top row, that is, create a <see cref="SheetView"/> that splits the first row (A) into a pane. 
        /// </summary>
        public void FreezeTopRow()
        {
            // TODO: Eventually, support more SheetView functionality, right now, keep it simple.
            if (_sheetViews != null)
            {
                throw new InvalidOperationException("You have already frozen either the Top Row or Left Column.");
            }

            var sheetView = new SheetView();
            sheetView.Pane = new Pane
            {
                ActivePane = Panes.BottomLeft,
                YSplit = 1,
                TopLeftCell = "A2",
                State = PaneState.Frozen
            };
            sheetView.AddSelection(new Selection { ActivePane = Panes.BottomLeft });
            AddSheetView(sheetView);
        }

        /// <summary>
        /// Freeze the first column, that is, create a <see cref="SheetView"/> that splits the first column (1) into a pane. 
        /// </summary>
        public void FreezeLeftColumn()
        {
            // TODO: Eventually, support more SheetView functionality, right now, keep it simple.
            if (_sheetViews != null)
            {
                throw new InvalidOperationException("You have already frozen either the Top Row or Left Column.");
            }

            var sheetView = new SheetView();
            sheetView.Pane = new Pane
            {
                ActivePane = Panes.TopRight,
                XSplit = 1,
                TopLeftCell = "B1",
                State = PaneState.Frozen
            };
            sheetView.AddSelection(new Selection { ActivePane = Panes.TopRight });
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

        /// <summary>
        /// Populate Worksheet with the provided data.
        /// 
        /// Will use the Object Property Names as Column Headers (First Row) and then populate the cells with data.
        /// 
        /// Caveats:
        /// * Does not look at inherited members from a Base Class
        /// * Only looks at strings and value types.
        /// * No way to specify the order of properties
        /// </summary>
        /// <param name="data"></param>
        public void Populate<T>(IEnumerable<T> data) where T : class
        {
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            int row = 0;
            int col = 0;

            var props = typeof(T).GetTypeInfo().DeclaredProperties
                .Where(p => p.GetIndexParameters().Length == 0)
                .Where(p => p.PropertyType.GetTypeInfo().IsValueType || p.PropertyType == typeof(string))
                .ToList();

            foreach (var prop in props)
            {
                Cells[row, col++] = prop.Name;
            }

            foreach (var item in data)
            {
                row++;
                col = 0;

                foreach (var prop in props)
                {
                    Cell cell;

                    object val = prop.GetValue(item);

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

                    Cells[row, col++] = cell;
                }
            }
        }

        public static Worksheet FromData<T>(string sheetName, IEnumerable<T> data) where T : class
        {
            var sheet = new Worksheet(sheetName);
            sheet.Populate(data);
            return sheet;
        }
    }
}