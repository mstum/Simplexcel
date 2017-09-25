namespace Simplexcel
{
    /// <summary>
    /// Worksheet view selection.
    /// </summary>
    public sealed class Selection
    {
		/// <summary>
		/// The pane to which this selection belongs.
		/// </summary>
		public Panes ActivePane { get; set; }

		/// <summary>
		/// Location of the active cell. E.g., "A1"
		/// </summary>
		public string ActiveCell { get; set; }

        // activeCellId
        // sqref

        /// <summary>
        /// Create a new selection
        /// </summary>
        public Selection()
        {
            ActivePane = Panes.TopLeft;
        }
    }
}
