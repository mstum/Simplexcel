using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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
            sheet.InsertManualPageBreakAfterRow(row+1);
        }

        /// <summary>
        /// Insert a manual page break to the left of the column specified by the cell address (e.g., B5)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellAddress"></param>
        public static void InsertManualPageBreakAfterColumn(this Worksheet sheet, string cellAddress)
        {
            CellAddressHelper.ReferenceToColRow(cellAddress, out int row, out int col);
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

        internal static IEnumerable<PropertyInfo> GetAllProperties(this TypeInfo typeInfo) => GetAllForType(typeInfo, ti => ti.DeclaredProperties);

        private static IEnumerable<T> GetAllForType<T>(TypeInfo typeInfo, Func<TypeInfo, IEnumerable<T>> accessor)
        {
            if (typeInfo == null)
            {
                yield break;
            }

            // The Stack is to make sure that we fetch Base Type Properties first
            var baseTypes = new Stack<TypeInfo>();
            while (typeInfo != null)
            {
                baseTypes.Push(typeInfo);
                typeInfo = typeInfo.BaseType?.GetTypeInfo();
            }

            while (baseTypes.Count > 0)
            {
                var ti = baseTypes.Pop();
                foreach (var t in accessor(ti))
                {
                    yield return t;
                }
            }
        }
    }
}
