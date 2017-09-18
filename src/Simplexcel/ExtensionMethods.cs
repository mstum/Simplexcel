namespace Simplexcel
{
    public static class SimplexcelExtensionMethods
    {
        /// <summary>
        /// Insert a manual page break after the row specified by the cell address (e.g., B5)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellAddress"></param>
        public static void InsertManualPageBreakAfterRow(this Worksheet sheet, string cellAddress)
        {
            CellAddressHelper.ReferenceToColRow(cellAddress, out int row, out int col);
            sheet.InsertManualPageBreakAfterRow(row);
        }

        /// <summary>
        /// Insert a manual page break to the left of the column specified by the cell address (e.g., B5)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellAddress"></param>
        public static void InsertManualPageBreakAfterColumn(this Worksheet sheet, string cellAddress)
        {
            CellAddressHelper.ReferenceToColRow(cellAddress, out int row, out int col);
            sheet.InsertManualPageBreakAfterColumn(col);
        }
    }
}
