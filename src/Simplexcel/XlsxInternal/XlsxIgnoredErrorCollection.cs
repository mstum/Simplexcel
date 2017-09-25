using System.Collections.Generic;

namespace Simplexcel.XlsxInternal
{
    internal sealed class XlsxIgnoredErrorCollection
    {
        internal IList<XlsxIgnoredError> DistinctIgnoredErrors { get; }

        public XlsxIgnoredErrorCollection()
        {
            DistinctIgnoredErrors = new List<XlsxIgnoredError>();
        }

        public void AddIgnoredError(CellAddress cellAddress, IgnoredError ignoredErrors)
        {
            if (!ignoredErrors.IsDifferentFromDefault)
            {
                return;
            }

            var id = ignoredErrors.GetHashCode();

            foreach (var distinctIgnored in DistinctIgnoredErrors)
            {
                if (distinctIgnored.IgnoredErrorId == id)
                {
                    distinctIgnored.Cells.Add(cellAddress);
                    return;
                }
            }

            var newIgnoredError = new XlsxIgnoredError
            {
                IgnoredError = ignoredErrors
            };
            newIgnoredError.Cells.Add(cellAddress);
            DistinctIgnoredErrors.Add(newIgnoredError);
        }
    }
}
