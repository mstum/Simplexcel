namespace Simplexcel
{
    /// <summary>
    /// The Type of a Cell, important for handling the values corrently (e.g., Excel not converting long numbers into scientific notation or removing leading zeroes)
    /// </summary>
    public enum CellType
    {
        /// <summary>
        /// Cell contains text, so use it exactly as specified
        /// </summary>
        Text,

        /// <summary>
        /// Cell contains a number (may strip leading zeroes but sorts properly)
        /// </summary>
        Number,

        /// <summary>
        /// Cell contains a date, must be greater than 01/01/0100
        /// </summary>
        Date,

        Formula
    }
}
