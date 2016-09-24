using Xunit;

namespace LamedalExcel.Tests
{
    public sealed class ColRowCountTests
    {
        [Fact]
        public void EmptySheet_ColumnCount_Zero()
        {
            var sheet = new Worksheet("test");
            Assert.Equal(0, sheet.Cells.ColumnCount);
        }

        [Fact]
        public void EmptySheet_RowCount_Zero()
        {
            var sheet = new Worksheet("test");
            Assert.Equal(0, sheet.Cells.RowCount);
        }

        [Fact]
        public void NoOffByOneError()
        {
            var sheet = new Worksheet("Test");
            sheet.Cells[0, 0] = "Test";
            Assert.Equal(1, sheet.Cells.RowCount);
            Assert.Equal(1, sheet.Cells.ColumnCount);
        }
        
        [Fact]
        public void EmptyRows_RowCount_CountsEmpty()
        {
            var sheet = new Worksheet("Test");
            sheet.Cells[0, 0] = "Test";
            sheet.Cells[6, 0] = "Row Seven";
            Assert.Equal(7, sheet.Cells.RowCount);
        }

        [Fact]
        public void EmptyRows_ColumnCount_CountsEmpty()
        {
            var sheet = new Worksheet("Test");
            sheet.Cells[0, 0] = "Test";
            sheet.Cells[0, 6] = "Column Seven";
            Assert.Equal(7, sheet.Cells.ColumnCount);
        }
    }
}
