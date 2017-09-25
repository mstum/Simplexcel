namespace Simplexcel
{
    /// <summary>
    /// State of the sheet's pane.
    /// </summary>
    public enum PaneState
    {
        /// <summary>
        /// Panes are split, but not frozen. In this state, the split bars are adjustable by the user. 
        /// </summary>
        Split,

        /// <summary>
        /// Panes are frozen, but were not split being frozen. In
        /// this state, when the panes are unfrozen again, a single
        /// pane results, with no split.
        /// </summary>
        Frozen,

        /// <summary>
        /// Panes are frozen and were split before being frozen. In
        /// this state, when the panes are unfrozen again, the split
        /// remains, but is adjustable. 
        /// </summary>
        FrozenSplit
    }
}