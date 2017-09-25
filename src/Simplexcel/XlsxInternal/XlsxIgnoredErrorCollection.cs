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
            var ca = cellAddress.ToString();

            foreach (var distinctIgnored in DistinctIgnoredErrors)
            {
                if (distinctIgnored.IgnoredErrorId == id)
                {
                    distinctIgnored.Cells.Add(ca);
                    return;
                }
            }

            var newIgnoredError = new XlsxIgnoredError
            {
                IgnoredError = ignoredErrors
            };
            newIgnoredError.Cells.Add(ca);
            DistinctIgnoredErrors.Add(newIgnoredError);
        }
    }
}
