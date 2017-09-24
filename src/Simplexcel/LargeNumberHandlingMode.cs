namespace Simplexcel
{
    /// <summary>
    /// Excel doesn't want to display large numbers (more than 11 digits) properly:
    /// https://support.microsoft.com/en-us/help/2643223/long-numbers-are-displayed-incorrectly-in-excel
    /// 
    /// This decides how to handle larger numbers.
    /// See <see cref="Cell.LargeNumberNegativeLimit"/> and <see cref="Cell.LargeNumberPositiveLimit"/> for the actual limits.
    /// </summary>
    public enum LargeNumberHandlingMode
    {
        /// <summary>
        /// Force the number to be stored as Text (Default)
        /// </summary>
        StoreAsText,
        /// <summary>
        /// Keep the number as a number, but force scientific notation
        /// </summary>
        UseScientificNotation
    }
}
