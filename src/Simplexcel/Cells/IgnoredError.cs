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
        /// Is this instance different from the default (does it need an ignoredError element?)
        /// </summary>
        internal bool IsDifferentFromDefault => GetHashCode() != 0;

        /// <summary>
        /// Get the HashCode for this instance, which is a unique combination based on the boolean properties
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            // This is like a bit field, so further properties should be 2, 4, 8, 16, ...
            var result = 0;
            if(NumberStoredAsText) { result |= 1; }

            return result;
        }

        /// <summary>
        /// Compare this <see cref="IgnoredError"/> to another <see cref="IgnoredError"/>
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
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
            if (ReferenceEquals(null, other)) return false;
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
