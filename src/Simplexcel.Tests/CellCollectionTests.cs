using Xunit;

namespace Simplexcel.Tests
{
    public class CellCollectionTests
    {
        [Fact]
        public void EmptyCellCollection_AccessUninitializedCell_ImplicitlyCreatesCell()
        {
            var cc = new CellCollection();

            cc[0, 0].Bold = true;
        }
    }
}
