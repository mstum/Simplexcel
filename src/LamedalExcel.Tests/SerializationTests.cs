using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using LamedalExcel.Cells;
using LamedalExcel.enums;
using Xunit;

namespace LamedalExcel.Tests
{
    public sealed class SerializationTests
    {
        private byte[] Serialize<T>(T input) where T : class
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, input);
                return ms.ToArray();
            }
        }

        private T Deserialize<T>(byte[] data) where T: class
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var ms = new MemoryStream(data))
            {
                return serializer.ReadObject(ms) as T;
            }
        }

        private T Clone<T>(T input) where T : class
        {
            var data = Serialize(input);
            return Deserialize<T>(data);
        }

        [Fact]
        public void EmptyWorkbook_Serializes()
        {
            var wb = new Workbook
                         {
                             Title = "The Title",
                             Author = "The Author"
                         };


            var cloned = Clone(wb);

            Assert.Equal("The Title", cloned.Title);
            Assert.Equal("The Author", cloned.Author);
        }

        [Fact]
        public void EmptyWorksheet_Serializes()
        {
            var ws = new Worksheet("Test");

            var cloned = Clone(ws);

            Assert.Equal("Test", cloned.Name);
        }

        [Fact]
        public void WorkbookWithSheet_Serializes()
        {
            var wb = new Workbook();
            var ws = new Worksheet("Test");
            wb.Add(ws);

            var cloned = Clone(wb);

            Assert.Equal(1, cloned.Sheets.Count());
            Assert.Equal("Test", cloned.Sheets.First().Name);
        }
        

        [Fact]
        public void PageSetup_Serializes()
        {
            var ws = new Worksheet("Test");
            ws.PageSetup.PrintRepeatColumns = 8;
            ws.PageSetup.PrintRepeatRows = 12;
            ws.PageSetup.Orientation = enOrientation.Landscape;

            var cloned = Clone(ws);

            Assert.Equal(8, cloned.PageSetup.PrintRepeatColumns);
            Assert.Equal(12, cloned.PageSetup.PrintRepeatRows);
            Assert.Equal(enOrientation.Landscape, cloned.PageSetup.Orientation);
        }

        [Fact]
        public void ColumnWidths_Serializes()
        {
            var ws = new Worksheet("Test");
            ws.ColumnWidths[0] = 14.8;

            var cloned = Clone(ws);

            Assert.Equal(14.8, cloned.ColumnWidths[0]);
        }

        [Fact]
        public void SheetWithCell_Serializes()
        {
            var ws = new Worksheet("Test");
            ws.Cells["D4"] = 1234.5m;
            ws.Cells["D4"].Bold = true;

            var cloned = Clone(ws);

            Assert.Equal(1234.5m, cloned.Cells["D4"].Value);
            Assert.True(cloned.Cells["D4"].Bold);
            Assert.Equal(enCellType.Number, cloned.Cells["D4"].CellType);
        }

        [Fact]
        public void SheetWithCellBorder_Serializes()
        {
            var ws = new Worksheet("Test");
            ws.Cells["A1"] = "My Test";
            ws.Cells["A1"].Border = enCellBorder.All;

            var cloned = Clone(ws);

            Assert.True(cloned.Cells["A1"].Border.HasFlag(enCellBorder.Bottom));
            Assert.True(cloned.Cells["A1"].Border.HasFlag(enCellBorder.Left));
            Assert.True(cloned.Cells["A1"].Border.HasFlag(enCellBorder.Right));
            Assert.True(cloned.Cells["A1"].Border.HasFlag(enCellBorder.Top));
            Assert.True(cloned.Cells["A1"].Border.HasFlag(enCellBorder.All));
        }
    }
}
