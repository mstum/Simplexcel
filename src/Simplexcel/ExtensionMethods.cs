using System;
using System.Linq;
using System.Reflection;

namespace Simplexcel
{
    /// <summary>
    /// Some utility methods that don't need to be in their respective classes
    /// </summary>
    public static class SimplexcelExtensionMethods
    {
        /// <summary>
        /// Insert a manual page break after the row specified by the cell address (e.g., B5)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellAddress"></param>
        public static void InsertManualPageBreakAfterRow(this Worksheet sheet, string cellAddress)
        {
            CellAddressHelper.ReferenceToColRow(cellAddress, out int row, out _);
            sheet.InsertManualPageBreakAfterRow(row+1);
        }

        /// <summary>
        /// Insert a manual page break to the left of the column specified by the cell address (e.g., B5)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellAddress"></param>
        public static void InsertManualPageBreakAfterColumn(this Worksheet sheet, string cellAddress)
        {
            CellAddressHelper.ReferenceToColRow(cellAddress, out _, out int col);
            sheet.InsertManualPageBreakAfterColumn(col+1);
        }

        private readonly static Type XlsxIgnoreColumnType = typeof(XlsxIgnoreColumnAttribute);
        private readonly static Type XlsxColumnType = typeof(XlsxColumnAttribute);
        internal static bool HasXlsxIgnoreAttribute(this PropertyInfo prop)
        {
            if(prop == null)
            {
                throw new ArgumentNullException(nameof(prop));
            }

            return prop.GetCustomAttributes(XlsxIgnoreColumnType).Any();
        }

        internal static XlsxColumnAttribute GetXlsxColumnAttribute(this PropertyInfo prop)
        {
            if (prop == null)
            {
                throw new ArgumentNullException(nameof(prop));
            }

            return prop.GetCustomAttributes(XlsxColumnType).Cast<XlsxColumnAttribute>().FirstOrDefault();
        }
    }
}
