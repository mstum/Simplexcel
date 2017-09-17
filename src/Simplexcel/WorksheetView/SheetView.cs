using System;
using System.Collections.Generic;

namespace Simplexcel
{
    public class SheetView
    {
        private List<Selection> _selections;

        internal ICollection<Selection> Selections {
            get { return _selections; }
        }

        public bool? TabSelected
        {
            get;
            set;
        }

        public bool? ShowRuler
        {
            get;
            set;
        }

        public int WorkbookViewId
        {
            get { return 0; }
        }

        public Pane Pane
        {
            get;
            set;
        }

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

        public SheetView()
        {
        }
    }
}