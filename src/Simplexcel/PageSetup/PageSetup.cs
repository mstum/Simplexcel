namespace Simplexcel
{
    /// <summary>
    /// Page Setup for a Sheet
    /// </summary>
    public sealed class PageSetup
    {
        /// <summary>
        /// How many rows to repeat on each page when printing?
        /// </summary>
        public int PrintRepeatRows { get; set; }

        /// <summary>
        /// How many columns to repeat on each page when printing?
        /// </summary>
        public int PrintRepeatColumns { get; set; }

        /// <summary>
        /// The Orientation of the page, Portrait (default) or Landscape
        /// </summary>
        public Orientation Orientation { get; set; }
    }
}
