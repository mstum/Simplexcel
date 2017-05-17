using Simplexcel.XlsxInternal;
using System.Linq;
using Xunit;

namespace Simplexcel.Tests
{
    public class BugTests
    {
        /// <summary>
        /// https://github.com/mstum/Simplexcel/issues/12
        /// 
        /// SharedStrings _sanitizeRegex makes "&" in CellStrings impossible
        /// </summary>
        [Fact]
        public void Issue12_AmpersandInCellValues()
        {
            var sharedStrings = new SharedStrings();
            sharedStrings.GetStringIndex("Here & Now");
            sharedStrings.GetStringIndex("&");
            var xmlFile = sharedStrings.ToXmlFile();

            var nodeValues = xmlFile.Content.Descendants(Namespaces.x + "t").Select(x => x.Value).ToList();

            Assert.Contains("Here & Now", nodeValues);
            Assert.Contains("&", nodeValues);
        }
    }
}
