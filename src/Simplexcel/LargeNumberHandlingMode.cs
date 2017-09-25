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
        StoreAsText = 0,

        /// <summary>
        /// Do not do anything different, and store numbers as-is.
        /// This may cause Excel to truncate the number.
        /// This was the behavior before Simplexcel Version 2.1.0.
        /// </summary>
        None
    }
}
