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

            // Key = ColumnIndex, Value = Attribute
            var cols = new Dictionary<int, PopulateCellInfo>();

            var type = typeof(T);
            var props = type.GetTypeInfo().GetAllProperties()
                .Where(p => p.GetIndexParameters().Length == 0)
                .ToList();

            int tempCol = 0; // Just a counter to keep the order of Properties the same
            int maxCol = -1; // Largest Column that has XlsxColumnAttribute.ColumnIndex specified
            foreach (var prop in props)
            {
                if(prop.HasXlsxIgnoreAttribute())
                {
                    continue;
                }

                var pi = new PopulateCellInfo();
                var colInfo = prop.GetXlsxColumnAttribute();
                pi.Name = string.IsNullOrEmpty(colInfo?.Name) ? prop.Name : colInfo.Name;
                pi.ColumnIndex = colInfo?.ColumnIndex != null ? colInfo.ColumnIndex.Value : -1; // -1 will later be reassigned
                pi.TempColumnIndex = tempCol++;

                if(pi.ColumnIndex > -1 && cols.ContainsKey(pi.ColumnIndex))
                {
                    throw new InvalidOperationException($"Type {type.FullName} includes more than one Property with ColumnIndex {pi.ColumnIndex}.");
                }

                if (pi.ColumnIndex > maxCol)
                {
                    maxCol = pi.ColumnIndex;
                }

                pi.Property = prop;
                cols[pi.TempColumnIndex] = pi;
            }

            // Slot any Columns without an order after maxCol
            for(int i = 0; i < tempCol; i++)
            {
                var pi = cols[i];
                if(pi.ColumnIndex == -1)
                {
                    pi.ColumnIndex = ++maxCol;
                }
            }

            foreach (var pi in cols.Values)
            {
                Cells[0, pi.ColumnIndex] = pi.Name;
                Cells[0, pi.ColumnIndex].Bold = true;
            }

            var row = 0;
            foreach (var item in data)
            {
                row++;
                
                foreach (var pi in cols.Values)
                {
                    object val = pi.Property.GetValue(item);
                    var cell = Cell.FromObject(val);
                    Cells[row, pi.ColumnIndex] = cell;
                }
            }
        }

        private class PopulateCellInfo
        {
            public string Name { get; set; }
            public int ColumnIndex { get; set; }
            public int TempColumnIndex { get; set; }
            public PropertyInfo Property { get; set; }
        }
    }
}
