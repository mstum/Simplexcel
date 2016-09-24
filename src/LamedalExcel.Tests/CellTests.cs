using LamedalExcel.Cells;
using LamedalExcel.enums;
using Xunit;

namespace LamedalExcel.Tests
{
    public sealed class CellTests
    {
        [Fact]
        public void ImplicitCreation_FromLong_NoDecimalPlaces()
        {
            Cell cell = 1234;
            Assert.Equal(BuiltInCellFormat.NumberNoDecimalPlaces, cell.Format);
            Assert.Equal(enCellType.Number, cell.CellType);
        }

        [Fact]
        public void ImplicitCreation_FromDouble_TwoDecimalPlaces()
        {
            Cell cell = 1234.56;
            Assert.Equal(BuiltInCellFormat.NumberTwoDecimalPlaces, cell.Format);
            Assert.Equal(enCellType.Number, cell.CellType);
           
        }

        [Fact]
        public void ImplicitCreation_FromDecimal_TwoDecimalPlaces()
        {
            Cell cell = 1234.56m;
            Assert.Equal(BuiltInCellFormat.NumberTwoDecimalPlaces, cell.Format);
            Assert.Equal(enCellType.Number, cell.CellType);
        }

        [Fact]
        public void ImplicitCreation_FromString_CellTypeText()
        {
            Cell cell = "1234";
            Assert.Equal(enCellType.Text, cell.CellType);
        }
    }
}
