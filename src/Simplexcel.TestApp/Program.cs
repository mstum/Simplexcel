using System;
using System.Collections.Generic;

namespace Simplexcel.TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new Workbook();
            wb.Title = "Workbook Title";
            wb.Author = "The Author";


            var sheet = new Worksheet("Test");
            sheet.PageSetup.Orientation = Orientation.Landscape;
            sheet.PageSetup.PrintRepeatRows = 2;
            sheet.PageSetup.PrintRepeatColumns = 2;

            sheet.ColumnWidths[0] = 24.6;

            sheet.Cells["A1"] = "Test";
            sheet.Cells["A1"].FontName = "Arial Black";

            sheet.Cells[0, 1] = "Another Test";
            sheet.Cells[0, 1].Border = CellBorder.Bottom | CellBorder.Right;

            sheet.Cells[0, 2] = "Bold Red";
            sheet.Cells[0, 2].Bold = true;
            sheet.Cells[0, 2].TextColor = Color.Red;

            Cell cell = "BIU & Big & Blue";
            cell.Bold = true;
            cell.Underline = true;
            cell.Italic = true;
            cell.TextColor = Color.Blue;
            cell.FontSize = 18;
            cell.Hyperlink = "https://github.com/mstum/Simplexcel";
            sheet.Cells[0, 3] = cell;
            sheet.ColumnWidths[3] = 40;


            sheet.Cells[0, 4] = 13;
            sheet.Cells[0, 5] = 13.58;

            for (int i = 1; i < 55; i++)
            {
                sheet.Cells[1,i] = Guid.NewGuid().ToString("N");
                sheet.Cells[i,0] = Guid.NewGuid().ToString("N");
            }

            Cell cell2 = "Orange";
            cell2.Bold = true;
            cell2.Italic = true;
            cell2.TextColor = Color.Orange;
            cell2.FontSize = 18;
            sheet.Cells[0, 2] = cell2;

            sheet.Cells[0, 6] = "👪";
            sheet.Cells[0, 7] = "👨‍👩‍👧‍👦";

            wb.Add(sheet);

            var sheet2 = new Worksheet("Sheet 2");
            sheet2[0, 0] = "Sheet Number 2";
            sheet2[0, 0].Bold = true;
            wb.Add(sheet2);

            var populatedSheet = new Worksheet("Populate");
            populatedSheet.Populate(EnumeratePopulateTestData());
            wb.Add(populatedSheet);

            wb.Save("compressed.xlsx", compress: true);
            wb.Save("uncompressed.xlsx", compress: false);
        }

        private static IEnumerable<PopulateTestData> EnumeratePopulateTestData()
        {
            var r = new Random();

            for(int i = 0; i < 500; i++)
            {
                yield return new PopulateTestData
                {
                    Id = i,
                    CreatedUtc = new DateTime(2015, 1, 1, 1, 1, 1).AddHours(r.Next(1, 852)),
                    Name = "Item Number " + i,
                    Value = int.MaxValue - i,
                    Price = r.Next(800, 9400) * 0.52m,
                    Quantity = r.Next(4, 75) * 0.5m
                };
            }
        }

        private class PopulateTestData
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public long Value { get; set; }
            public decimal Price { get; set; }
            public decimal Quantity { get; set; }
            public decimal TotalPrice { get { return Price * Quantity; } }
            public DateTime CreatedUtc { get; set; }
        }
    }
}
