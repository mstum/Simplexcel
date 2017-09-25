namespace Simplexcel
{
    /// <summary>
    /// Worksheet view pane
    /// </summary>
    public sealed class Pane
    {
        internal const int DefaultXSplit = 0;
        internal const int DefaultYSplit = 0;
        internal const Panes DefaultActivePane = Panes.TopLeft;
        internal const PaneState DefaultState = PaneState.Split;

        /// <summary>
        /// Horizontal position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen,
        /// this value indicates the number of columns visible in the top pane.
        /// </summary>
        public int? XSplit { get; set; }

        /// <summary>
        /// Vertical position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, 
        /// this value indicates the number of rows visible in the left pane
        /// </summary>
        public int? YSplit { get; set; }

        /// <summary>
        /// The pane that is active.
        /// </summary>
        public Panes? ActivePane { get; set; }

        /// <summary>
        /// Indicates whether the pane has horizontal / vertical splits, and whether those splits are frozen.
        /// </summary>
        public PaneState? State { get; set; }

        /// <summary>
        /// Location of the top left visible cell in the bottom right pane (when in Left-To-Right mode)
        /// </summary>
        public string TopLeftCell { get; set; }
    }
}