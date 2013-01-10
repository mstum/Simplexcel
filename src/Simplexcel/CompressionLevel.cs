namespace Simplexcel
{
    /// <summary>
    /// How much should content in a workbook be compressed?
    /// </summary>
    /// <remarks>
    /// The underlying values intentionally mirror <see cref="System.IO.Packaging.CompressionOption"/>
    /// </remarks>
    public enum CompressionLevel
    {
        /// <summary>
        /// Compression is turned off
        /// </summary>
        NoCompression = -1,

        /// <summary>
        /// Compression is optimized for a balance between size and performance
        /// </summary>
        Balanced = 0,

        /// <summary>
        ///  Compression is optimized for size
        /// </summary>
        Maximum = 1
    }
}
