namespace Simplexcel
{
    /// <summary>
    /// The brk element, for Page Breaks
    /// </summary>
    public sealed class PageBreak
    {
        /// <summary>
        /// Zero-based row or column Id of the page break.
        /// Breaks occur above the specified row and left of the specified column.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Manual Break flag. true means the break is a manually inserted break.
        /// </summary>
        public bool IsManualBreak { get; set; }

        /// <summary>
        /// Zero-based index of end row or column of the break.
        /// For row breaks, specifies column index;
        /// for column breaks, specifies row index.
        /// </summary>
        public int Max { get; set; }

        /// <summary>
        /// Zero-based index of start row or column of the break.
        /// For row breaks, specifies column index;
        /// for column breaks, specifies row index.
        /// </summary>
        public int Min { get; set; }

        /// <summary>
        /// Flag indicating that a PivotTable created this break.
        /// </summary>
        public bool IsPivotCreatedPageBreak { get; set; }
    }
}
