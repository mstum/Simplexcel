namespace Simplexcel
{
    /// <summary>
    /// The Vertical Alignment of content within a Cell
    /// </summary>
    public enum VerticalAlign
    {
        /// <summary>
        /// No specific alignment, use Excel's default
        /// </summary>
        None = 0,

        /// <summary>
        /// The vertical alignment is aligned-to-top.
        /// </summary>
        Top,

        /// <summary>
        /// The vertical alignment is centered across the height of the cell.
        /// </summary>
        Middle,

        /// <summary>
        /// The vertical alignment is aligned-to-bottom.
        /// </summary>
        Bottom,

        /// <summary>
        /// When text direction is horizontal: the vertical alignment of lines of text is
        /// distributed vertically, where each line of text inside the cell is evenly
        /// distributed across the height of the cell, with flush top and bottom margins.
        /// 
        /// When text direction is vertical: similar behavior as horizontal justification.
        /// The alignment is justified (flush top and bottom in this case). For each line
        /// of text, each line of the wrapped text in a cell is aligned to the top and
        /// bottom (except the last line). If no single line of text wraps in the cell,
        /// then the text is not justified.
        /// </summary>
        Justify,

        //Distributed
    }
}
