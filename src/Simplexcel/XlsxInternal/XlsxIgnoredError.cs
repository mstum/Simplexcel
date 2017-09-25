using System;
using System.Collections.Generic;

namespace Simplexcel.XlsxInternal
{
    internal sealed class XlsxIgnoredError
    {
        private readonly IgnoredError _ignoredError;
        internal HashSet<string> Cells { get; }

        public XlsxIgnoredError()
        {
            Cells = new HashSet<string>();
            _ignoredError = new IgnoredError();
        }

        internal IgnoredError IgnoredError
        {
            get
            {
                // Note: This is a mutable reference, but changing it would be... bad.
                return _ignoredError;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                // And because the reference is mutable, this stores a copy so that modifications to the Worksheet don't break this
                _ignoredError.NumberStoredAsText = value.NumberStoredAsText;
                IgnoredErrorId = _ignoredError.GetHashCode();
            }
        }

        internal int IgnoredErrorId { get; private set; }

        internal string GetSqRef()
        {
            // TODO: Support Ranges. Ranges are Rectangular, e.g., A1:B5 (TopLeft:BottomRight)
            return string.Join(" ", Cells);
        }
    }
}
