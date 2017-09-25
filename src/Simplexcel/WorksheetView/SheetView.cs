using System;
using System.Collections.Generic;

namespace Simplexcel
{
    /// <summary>
    /// A single sheet view definition. When more than one sheet view is defined in the file, it means that when opening
    /// the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where
    /// each window is showing the particular sheet containing the same workbookViewId value, the last sheetView
    /// definition is loaded, and the others are discarded.
    /// 
    /// When multiple windows are viewing the same sheet, multiple
    /// sheetView elements (with corresponding workbookView entries) are saved.
    /// </summary>
    public sealed class SheetView
    {
        private List<Selection> _selections;

        internal ICollection<Selection> Selections
        {
            get { return _selections; }
        }

        /// <summary>
        /// Flag indicating whether this sheet is selected.
        /// When only 1 sheet is selected and active, this value should be in synch with the activeTab value.
        /// In case of a conflict, the Start Part setting wins and sets the active sheet tab.
        /// Multiple sheets can be selected, but only one sheet shall be active at one time.
        /// </summary>
        public bool? TabSelected { get; set; }

        /// <summary>
        /// Show the ruler in page layout view
        /// </summary>
        public bool? ShowRuler { get; set; }

        /// <summary>
        /// Zero-based index of this workbook view, pointing to a workbookView element in the bookViews collection.
        /// </summary>
        public int WorkbookViewId { get { return 0; } }

        /// <summary>
        /// The pane that this SheetView applies to
        /// </summary>
        public Pane Pane { get; set; }

        /// <summary>
        /// Add the given selection to the SheetView
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="throwOnDuplicatePane">If true, will throw an <see cref="InvalidOperationException"/> if there is already a selection with the same <see cref="Selection.ActivePane"/>. If false, overwrite the previous selection for that pane.</param>
        public void AddSelection(Selection sel, bool throwOnDuplicatePane = true)
        {
            if (sel == null) { throw new ArgumentNullException(nameof(sel)); }
            if (!Enum.IsDefined(typeof(Panes), sel.ActivePane))
            {
                throw new ArgumentOutOfRangeException(nameof(sel), "Not a valid ActivePane: " + sel.ActivePane);
            }

            // Up to 4 Selections, and the Selection Pane cannot already be in use
            if (_selections == null)
            {
                _selections = new List<Selection>(4);
            }

            foreach (var s in _selections)
            {
                if (s.ActivePane == sel.ActivePane)
                {
                    if (throwOnDuplicatePane)
                    {
                        throw new InvalidOperationException("There is already a Selection for ActivePane " + sel.ActivePane);
                    }
                    else
                    {
                        _selections.Remove(s);
                        break;
                    }
                }
            }

            _selections.Add(sel);
        }
    }
}