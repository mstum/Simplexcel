using System;
using System.Drawing;
using LamedalCore;
using LamedalCore.zz;
using LamedalExcel.Cells;
using LamedalExcel.enums;
using Xunit;

namespace LamedalExcel.TestsAppWinForms
{
    public sealed class LamedalExcel_
    {
        private readonly LamedalCore_ _lamed = LamedalCore_.Instance; // system library

        [Fact]
        public void Excel_CreateFile()
        {
            //Workbook workbook = Workbook_New("The Author", "Workbook Title");
            Worksheet worksheet = WorkSheet_New("Test123", enOrientation.Landscape);

            // Widths
            worksheet.ColumnWidths[0] = 24.6;
            WorkSheet_ColumnWidth(worksheet, "B", 25.8);
            WorkSheet_ColumnWidth(worksheet, "C1", 26.9);
            Assert.Equal(24.6, worksheet.ColumnWidths[0]);
            Assert.Equal(25.8, worksheet.ColumnWidths[1]);
            Assert.Equal(26.9, worksheet.ColumnWidths[2]);

            // A1
            var cellA1 = WorkSheet_CellSet(worksheet, "A1", "Text", fontName: "Arial Black");
            Assert.Equal("Text", worksheet.Cells["A1"].Value);
            Assert.Equal("Arial Black", worksheet.Cells["A1"].FontName);

            // A1 test 2
            cellA1.Value = "Test2";
            cellA1.FontName = "Arial";
            Assert.Equal("Test2", worksheet.Cells["A1"].Value);
            Assert.Equal("Arial", worksheet.Cells["A1"].FontName);

            Cell cellB1 = WorkSheet_CellSet(worksheet, 2,1, "Another Test");  // This must give an error
            cellB1.Border = enCellBorder.Bottom | enCellBorder.Right;

            // C1
            WorkSheet_CellSet(worksheet,"C1","Bold Red", bold:true, textColor:Color.Red);

            // D1
            WorkSheet_CellSet(worksheet, "D1", "BIU Big Blue", bold:true, underline:true, italic:true, textColor:Color.Blue, fontSize:18, webLink: "https://github.com/mstum/Simplexcel", columnWidth:40);
            //cell.Bold = true;
            //cell.Underline = true;
            //cell.Italic = true;
            //cell.TextColor = Color.Blue;
            //cell.FontSize = 18;
            //cell.Hyperlink = "https://github.com/mstum/Simplexcel";
            //sheet.ColumnWidths[3] = 40;
            //sheet.Cells["D1"] = cell;

            // E1, F1
            worksheet.Cells["E1"] = 13;
            worksheet.Cells["F1"] = 13.58;

            worksheet.Workbook.Save(@"C:\test\testCompressed.xlsx", CompressionLevel.Maximum);
            worksheet.Workbook.Save(@"C:\test\testUncompressed.xlsx", CompressionLevel.NoCompression);
        }

 
    }
}
