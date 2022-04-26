﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace Simplexcel.TestApp
{
    class Program
    {
        static void Main()
        {
            var wb = new Workbook
            {
                Title = "Workbook Title",
                Author = "The Author"
            };

            var sheet = new Worksheet("Test");
            sheet.PageSetup.Orientation = Orientation.Landscape;
            sheet.PageSetup.PrintRepeatRows = 2;
            sheet.PageSetup.PrintRepeatColumns = 2;

            sheet.ColumnWidths[0] = 24.6;

            // Version Info
            sheet["F11"] = "Version";
            sheet["G11"] = SimplexcelVersion.Version.ToString();
            sheet["F12"] = "VersionString";
            sheet["G12"] = SimplexcelVersion.VersionString;
            sheet["F13"] = "PublicKeyToken";
            sheet["G13"] = SimplexcelVersion.PublicKeyToken;
            sheet["F14"] = "PublicKey";
            sheet["G14"] = SimplexcelVersion.PublicKey;

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
                sheet.Cells[1, i] = Guid.NewGuid().ToString("N");
                sheet.Cells[i, 0] = Guid.NewGuid().ToString("N");
            }

            Cell cell2 = "Orange";
            cell2.Bold = true;
            cell2.Italic = true;
            cell2.TextColor = Color.Orange;
            cell2.FontSize = 18;
            sheet.Cells[0, 2] = cell2;

            // Note that Emoji only work in Excel 2013 or newer, not 2007 or 2010. This is an Excel limitation.
            sheet.Cells[0, 6] = "👪";
            sheet.Cells[0, 7] = "👨‍👩‍👧‍👦";

            var dateTime = new DateTime(2022, 4, 26, 13, 14, 15);
            sheet.Cells["D4"] = dateTime;
            sheet.Cells["D5"] = new Cell(CellType.Date, dateTime, BuiltInCellFormat.DateOnly);
            sheet.Cells["D6"] = new Cell(CellType.Date, dateTime, BuiltInCellFormat.TimeOnly);
            sheet.Cells["D7"] = new Cell(CellType.Date, dateTime, "yyyy\"_\"mm\"_\"dd\"_\"hh\"_\"mm\"_\"ss");
            sheet.Cells["D8"] = long.MaxValue;
            sheet.Cells["D9"] = long.MinValue;
            sheet.Cells["D10"] = decimal.MaxValue;
            sheet.Cells["D11"] = decimal.MinValue;

            sheet.Cells["D12"] = 9999999999L;
            sheet.Cells["D13"] = 99999999999L;
            sheet.Cells["D14"] = 100000000000L;
            sheet.Cells["D15"] = 100000000001L;
            sheet.Cells["D16"] = 1000000000000L;
            sheet.Cells["D17"] = 1000000000001L;
            sheet.Cells["D18"] = Cell.LargeNumberPositiveLimit;
            sheet.Cells["D19"] = Cell.LargeNumberPositiveLimit + 1;
            sheet.Cells["D20"] = Cell.LargeNumberPositiveLimit - 1;

            sheet.Cells["D21"] = -9999999999L;
            sheet.Cells["D22"] = -99999999999L;
            sheet.Cells["D23"] = -100000000000L;
            sheet.Cells["D24"] = -100000000001L;
            sheet.Cells["D25"] = -1000000000000L;
            sheet.Cells["D26"] = -1000000000001L;
            sheet.Cells["D27"] = Cell.LargeNumberNegativeLimit;
            sheet.Cells["D28"] = Cell.LargeNumberNegativeLimit + 1;
            sheet.Cells["D29"] = Cell.LargeNumberNegativeLimit - 1;
            sheet.LargeNumberHandlingMode = LargeNumberHandlingMode.StoreAsText;

            sheet.Cells["C5"] = "ThinDiagonalCrosshatch";
            sheet.Cells["C5"].Fill = new PatternFill { PatternType = PatternType.ThinDiagonalCrosshatch, PatternColor = Color.Yellow, BackgroundColor = Color.Violet };
            sheet.Cells["C5"].TextColor = Color.Green;

            sheet.Cells["C6"] = "ThickDiagonalCrosshatch";
            sheet.Cells["C6"].Fill = new PatternFill { PatternType = PatternType.ThickDiagonalCrosshatch, PatternColor = Color.Violet };
            sheet.Cells["C6"].TextColor = Color.Green;

            sheet.Cells["C7"] = "DiagonalCrosshatch";
            sheet.Cells["C7"].Fill = new PatternFill { PatternType = PatternType.DiagonalCrosshatch, BackgroundColor = Color.Violet };
            sheet.Cells["C7"].TextColor = Color.Green;

            sheet.Cells["C8"] = "FillSolid";
            sheet.Cells["C8"].Fill = new PatternFill { PatternType = PatternType.Solid, PatternColor = Color.Yellow, BackgroundColor = Color.Violet };
            sheet.Cells["C8"].TextColor = Color.Green;

            sheet.Cells["C9"] = "SetFillBgColor";
            sheet.Cells["C9"].Fill.BackgroundColor = Color.Navy;
            sheet.Cells["C9"].TextColor = Color.Green;

            sheet.Cells["F5"] = "Hello,";
            sheet.Cells["G5"] = "World!";
            sheet.Cells["F5"].Fill.BackgroundColor = Color.Blue;
            sheet.Cells["G5"].Fill.BackgroundColor = Color.Red;

            wb.Add(sheet);

            // Prime the Cache...
            var populatedSheet = new Worksheet("Populate");
            populatedSheet.Populate(EnumeratePopulateTestData(), cacheTypeColumns: true);

            // ...and use the cache
            populatedSheet = new Worksheet("Populate");
            populatedSheet.Populate(EnumeratePopulateTestData(), cacheTypeColumns: true);
            wb.Add(populatedSheet);

            var frozenTopRowSheet = new Worksheet("Frozen Top Row");
            frozenTopRowSheet.Cells[0, 0] = "Header 1";
            frozenTopRowSheet.Cells[0, 1] = "Header 2";
            frozenTopRowSheet.Cells[0, 2] = "Header 3";
            foreach (var i in Enumerable.Range(1, 100))
            {
                frozenTopRowSheet.Cells[i, 0] = "Value 1-" + i;
                frozenTopRowSheet.Cells[i, 1] = "Value 2-" + i;
                frozenTopRowSheet.Cells[i, 2] = "Value 3-" + i;
            }
            frozenTopRowSheet.FreezeTopRow();
            wb.Add(frozenTopRowSheet);

            var frozenLeftColumnSheet = new Worksheet("Frozen Left Column");
            frozenLeftColumnSheet.Cells[0, 0] = "Header 1";
            frozenLeftColumnSheet.Cells[1, 0] = "Header 2";
            frozenLeftColumnSheet.Cells[2, 0] = "Header 3";
            foreach (var i in Enumerable.Range(1, 100))
            {
                frozenLeftColumnSheet.Cells[0, i] = "Value 1-" + i;
                frozenLeftColumnSheet.Cells[1, i] = "Value 2-" + i;
                frozenLeftColumnSheet.Cells[2, i] = "Value 3-" + i;
            }
            frozenLeftColumnSheet.FreezeLeftColumn();
            wb.Add(frozenLeftColumnSheet);

            var pageBreakSheet = new Worksheet("Page Breaks");
            for (var row = 0; row < 20; row++)
            {
                for (var col = 0; col < 20; col++)
                {
                    pageBreakSheet.Cells[row, col] = $"Row {row} /Col {col}";
                }
            }
            pageBreakSheet.InsertManualPageBreakAfterRow("B1");
            pageBreakSheet.InsertManualPageBreakAfterRow(5);
            pageBreakSheet.InsertManualPageBreakAfterColumn("B8");
            pageBreakSheet.InsertManualPageBreakAfterColumn(16);
            wb.Add(pageBreakSheet);

            var formulaSheet = new Worksheet("Formula");
            formulaSheet.Cells["A1"] = 5;
            formulaSheet.Cells["A2"] = 10;
            formulaSheet.Cells["A3"] = 15;
            formulaSheet.Cells["B1"] = Cell.Formula("SUM(A:A)");
            formulaSheet.Cells["B2"] = Cell.Formula("SUM(A1:A2)");
            formulaSheet.Cells["B3"] = Cell.Formula("AVERAGE(A:A)");
            formulaSheet.Cells["B4"] = Cell.Formula("MEDIAN(A:A)");
            formulaSheet.Cells["C1"] = Cell.Formula("A1+B1");
            formulaSheet.Cells["C2"] = Cell.Formula("=SUM(B:B)");
            formulaSheet.Cells["C3"] = Cell.Formula("C5+A1");
            formulaSheet.Cells["C5"] = Cell.Formula("SUM(B:B)+C1");
            wb.Add(formulaSheet);

            var freezeTopLeftSheet = new Worksheet("FreezeTopLeft");
            for (var row = 0; row < 10; row++)
                for (var col = 0; col < 10; col++)
                {
                    freezeTopLeftSheet[row, col] = $"{row},{col}";
                }
            freezeTopLeftSheet.FreezeTopLeft(2, 2);
            wb.Add(freezeTopLeftSheet);

            wb.Save("compressed.xlsx", compress: true);
            wb.Save("uncompressed.xlsx", compress: false);
        }

        private static IEnumerable<PopulateTestData> EnumeratePopulateTestData()
        {
            var r = new Random();

            for (int i = 0; i < 500; i++)
            {
                yield return new PopulateTestData
                {
                    Id = i,
                    CreatedUtc = new DateTime(2015, 1, 1, 1, 1, 1).AddDays(r.Next(1, 852)),
                    Name = "Item Number " + i,
                    Value = int.MaxValue - i,
                    Price = r.Next(800, 9400) * 0.52m,
                    Quantity = r.Next(4, 75) * 0.5m,
                    IgnoreMe = Guid.NewGuid().ToString()
                };
            }
        }

        private abstract class PopulateTestDataBase
        {
            public int Id { get; set; }

            [XlsxColumn("Unit Price")]
            public decimal Price { get; set; }

            [XlsxColumn(null)]
            public decimal Quantity { get; set; }
        }

        private class PopulateTestData : PopulateTestDataBase
        {            
            public string Name { get; set; }

            [XlsxColumn("")]
            public long Value { get; set; }

            [XlsxColumn("Total")]
            public decimal TotalPrice { get { return Price * Quantity; } }
            public DateTime CreatedUtc { get; set; }

            [XlsxIgnoreColumn]
            public string IgnoreMe { get; set; }
        }
    }
}
