using System.Collections.Generic;
using System.Linq;

namespace Simplexcel.MvcTestApp
{
    public class ExcelTestActionResult : ExcelResultBase
    {
        private readonly List<string> _data;

        public ExcelTestActionResult(string filename, List<string> data) : base(filename)
        {
            _data = data;
        }

        protected override Workbook GenerateWorkbook()
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

            Cell cell2 = "Orange";
            cell2.Bold = true;
            cell2.Italic = true;
            cell2.TextColor = Color.Orange;
            cell2.FontSize = 18;
            sheet.Cells[0, 2] = cell2;

            sheet.Cells[0, 6] = "👪";
            sheet.Cells[0, 7] = "👨‍👩‍👧‍👦";

            var i = 0;
            foreach(var datum in _data ?? Enumerable.Empty<string>())
            {
                i++;
                sheet.Cells[i, 0] = datum;
            }

            wb.Add(sheet);

            var sheet2 = new Worksheet("Sheet 2");
            sheet2[0, 0] = "Sheet Number 2";
            sheet2[0, 0].Bold = true;
            wb.Add(sheet2);

            return wb;
        }
    }
}