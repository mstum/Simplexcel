using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Simplexcel
{
    public sealed partial class Worksheet
    {
        public static Worksheet FromData<T>(string sheetName, IEnumerable<T> data) where T : class
        {
            var sheet = new Worksheet(sheetName);
            sheet.Populate(data);
            return sheet;
        }

        /// <summary>
        /// Populate Worksheet with the provided data.
        /// 
        /// Will use the Object Property Names as Column Headers (First Row) and then populate the cells with data.
        /// 
        /// Caveats:
        /// * Does not look at inherited members from a Base Class
        /// * Only looks at strings and value types.
        /// * No way to specify the order of properties
        /// </summary>
        /// <param name="data"></param>
        public void Populate<T>(IEnumerable<T> data) where T : class
        {
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            int row = 0;
            int col = 0;

            var props = typeof(T).GetTypeInfo().DeclaredProperties
                .Where(p => p.GetIndexParameters().Length == 0)
                .Where(p => p.PropertyType.GetTypeInfo().IsValueType || p.PropertyType == typeof(string))
                .ToList();

            foreach (var prop in props)
            {
                Cells[row, col++] = prop.Name;
            }

            foreach (var item in data)
            {
                row++;
                col = 0;

                foreach (var prop in props)
                {
                    object val = prop.GetValue(item);
                    var cell = Cell.FromObject(val);
                    Cells[row, col++] = cell;
                }
            }
        }
    }
}
