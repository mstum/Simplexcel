namespace Simplexcel
{
    /// <summary>
    /// The Horizontal Alignment of content within a Cell
    /// </summary>
    public enum HorizontalAlign
    {
        /// <summary>
        /// No specific alignment, use Excel's default
        /// </summary>
        None = 0,

        /// <summary>
        /// The horizontal alignment is general-aligned. Text data
        /// is left-aligned. Numbers, dates, and times are rightaligned.
        /// Boolean types are centered. Changing the
        /// alignment does not change the type of data.
        /// </summary>
        General,

        /// <summary>
        /// The horizontal alignment is left-aligned, even in Right-to-Left mode.
        /// </summary>
        Left,

        /// <summary>
        /// The horizontal alignment is centered, meaning the text is centered across the cell.
        /// </summary>
        Center,

        /// <summary>
        /// The horizontal alignment is right-aligned, meaning that
        /// cell contents are aligned at the right edge of the cell,
        /// even in Right-to-Left mode.
        /// </summary>
        Right,

        //Fill,

        /// <summary>
        /// The horizontal alignment is justified (flush left and
        /// right). For each line of text, aligns each line of the
        /// wrapped text in a cell to the right and left (except the
        /// last line). If no single line of text wraps in the cell, then
        /// the text is not justified.
        /// </summary>
        Justify,
        
        //CenterContinuous,

        //Distributed
    }
}
