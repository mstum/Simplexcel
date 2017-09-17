using System;
namespace Simplexcel
{
    public class Selection
    {
		/// <summary>
		/// The pane to which this selection belongs.
		/// </summary>
		public Panes ActivePane
        {
            get;
            set;
        }

		/// <summary>
		/// Location of the active cell. E.g., "A1"
		/// </summary>
		public string ActiveCell
        {
            get;
            set;
        }

        public Selection()
        {
            ActivePane = Panes.TopLeft;
        }
    }
}
