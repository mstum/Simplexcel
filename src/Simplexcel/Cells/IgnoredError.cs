namespace Simplexcel
{
    /// <summary>
    /// Information about Errors that are being ignored in a cell
    /// </summary>
    public sealed class IgnoredError
    {
        /// <summary>
        /// Ignore errors when numbers are formatted as text or are preceded by an apostrophe.
        /// </summary>
        public bool NumberStoredAsText { get; set; }

        /// <summary>
        /// Ignore errors when cells contain a value different from a calculated column formula.
        /// In other words, for a calculated column, a cell in that column is considered to have
        /// an error if its formula is different from the calculated column formula, or doesn't
        /// contain a formula at all.
        /// </summary>
        public bool CalculatedColumn { get; set; }

        /// <summary>
        /// Ignore errors when formulas refer to empty cells
        /// </summary>
        public bool EmptyCellReference { get; set; }

        /// <summary>
        /// Ignore errors when cells contain formulas that result in an error.
        /// </summary>
        public bool EvalError { get; set; }

        /// <summary>
        /// Ignore errors when a formula in a region of your worksheet differs from other formulas in the same region.
        /// </summary>
        public bool Formula { get; set; }

        /// <summary>
        /// Ignore errors when formulas omit certain cells in a region.
        /// </summary>
        public bool FormulaRange { get; set; }

        /// <summary>
        /// Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified.
        /// </summary>
        public bool ListDataValidation { get; set; }

        /// <summary>
        /// Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        /// </summary>
        public bool TwoDigitTextYear { get; set; }

        /// <summary>
        /// Ignore errors when unlocked cells contain formulas.
        /// </summary>
        public bool UnlockedFormula { get; set; }

        /// <summary>
        /// Is this instance different from the default (does it need an ignoredError element?)
        /// </summary>
        internal bool IsDifferentFromDefault => GetHashCode() != 0;

        /// <summary>
        /// Get the HashCode for this instance, which is a unique combination based on the boolean properties
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            // This is like a bit field, since an int is 32-Bit and this class is all booleans
            var result = 0;
            if (EvalError) { result |= (1 << 0); }
            if (TwoDigitTextYear) { result |= (1 << 1); }
            if (NumberStoredAsText) { result |= (1 << 2); }
            if (Formula) { result |= (1 << 3); }
            if (FormulaRange) { result |= (1 << 4); }
            if (UnlockedFormula) { result |= (1 << 5); }
            if (EmptyCellReference) { result |= (1 << 6); }
            if (ListDataValidation) { result |= (1 << 7); }
            if (CalculatedColumn) { result |= (1 << 8); }
            return result;
        }

        /// <summary>
        /// Compare this <see cref="IgnoredError"/> to another <see cref="IgnoredError"/>
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj is null) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(IgnoredError)) return false;
            return Equals((IgnoredError)obj);
        }

        /// <summary>
        /// Compare this <see cref="IgnoredError"/> to another <see cref="IgnoredError"/>
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(IgnoredError other)
        {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;

            var thisId = GetHashCode();
            var otherId = other.GetHashCode();

            return otherId.Equals(thisId);
        }

        /// <summary>
        /// Compare a <see cref="IgnoredError"/> to another <see cref="IgnoredError"/>
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator ==(IgnoredError left, IgnoredError right)
        {
            return Equals(left, right);
        }

        /// <summary>
        /// Compare a <see cref="IgnoredError"/> to another <see cref="IgnoredError"/>
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator !=(IgnoredError left, IgnoredError right)
        {
            return !Equals(left, right);
        }
    }
}
