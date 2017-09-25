namespace Simplexcel
{
    /// <summary>
    /// Defines the names of the four possible panes into which the view of a workbook in the application can be split.
    /// </summary>
    public enum Panes
    {
        /// <summary>
        /// Bottom right pane, when both vertical and horizontal splits are applied.
        /// </summary>
        BottomRight,

        /// <summary>
        /// Top right pane, when both vertical and horizontal splits are applied.
        /// 
        /// This value is also used when only a vertical split has
        /// been applied, dividing the pane into right and left
        /// regions. In that case, this value specifies the right pane.
        /// </summary>
        TopRight,

        /// <summary>
        /// Bottom left pane, when both vertical and horizontal splits are applied.
        /// 
        /// This value is also used when only a horizontal split has
        /// been applied, dividing the pane into upper and lower
        /// regions. In that case, this value specifies the bottom pane
        /// </summary>
        BottomLeft,

        /// <summary>
        /// Top left pane, when both vertical and horizontal splits are applied.
        /// 
        /// This value is also used when only a horizontal split has
        /// been applied, dividing the pane into upper and lower
        /// regions. In that case, this value specifies the top pane.
        /// 
        /// This value is also used when only a vertical split has
        /// been applied, dividing the pane into right and left
        /// regions. In that case, this value specifies the left pane
        /// </summary>
        TopLeft
    }
}