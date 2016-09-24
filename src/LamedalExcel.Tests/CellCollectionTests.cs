using LamedalExcel.Cells;
using Xunit;

namespace LamedalExcel.Tests
{
    public sealed class CellCollectionTests
    {
        [Fact]
        public void EmptyCellCollection_AccessUninitializedCell_ImplicitlyCreatesCell()
        {
            var cc = new CellCollection();

            cc[0, 0].Bold = true;
        }
    }
}
