using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Simplexcel
{
    public sealed partial class Worksheet
    {
        private static Lazy<ConcurrentDictionary<Type, Dictionary<int, PopulateCellInfo>>> PopulateCache 
            = new Lazy<ConcurrentDictionary<Type, Dictionary<int, PopulateCellInfo>>>(
                () => new ConcurrentDictionary<Type, Dictionary<int, PopulateCellInfo>>(),
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);

        /// <summary>
        /// Create Worksheet with the provided data.
        /// 
        /// Will use the Object Property Names as Column Headers (First Row) and then populate the cells with data.
        /// You can use <see cref="XlsxColumnAttribute"/> and <see cref="XlsxIgnoreColumnAttribute"/> to control the output.
        /// </summary>
        /// <param name="sheetName">The name of the Sheet</param>
        /// <param name="data">The data, can be empty or null</param>
        /// <param name="cacheTypeColumns">If true, the Column info for the given type is being cached in memory</param>
        public static Worksheet FromData<T>(string sheetName, IEnumerable<T> data, bool cacheTypeColumns = false) where T : class
        {
            var sheet = new Worksheet(sheetName);
            sheet.Populate(data, cacheTypeColumns);
            return sheet;
        }

        /// <summary>
        /// Populate Worksheet with the provided data.
        /// 
        /// Will use the Object Property Names as Column Headers (First Row) and then populate the cells with data.
        /// You can use <see cref="XlsxColumnAttribute"/> and <see cref="XlsxIgnoreColumnAttribute"/> to control the output.
        /// </summary>
        /// <param name="data">The data, can be empty or null</param>
        /// <param name="cacheTypeColumns">If true, the Column info for the given type is being cached in memory</param>
        public void Populate<T>(IEnumerable<T> data, bool cacheTypeColumns = false) where T : class
        {
            data = data ?? Enumerable.Empty<T>();

            var type = typeof(T);

            // Key = TempColumnIndex, Value = Attribute
            var cols = cacheTypeColumns ? TryGetFromCache(type) : null;
            if (cols == null)
            {
                cols = GetColumsFromType(type);
                if (cacheTypeColumns)
                {
                    AddToPopulateCache(type, cols);
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

        private static Dictionary<int, PopulateCellInfo> GetColumsFromType(Type type)
        {
            var cols = new Dictionary<int, PopulateCellInfo>();
            var props = type.GetTypeInfo().GetAllProperties()
                .Where(p => p.GetIndexParameters().Length == 0)
                .ToList();

            int tempCol = 0; // Just a counter to keep the order of Properties the same
            int maxCol = -1; // Largest Column that has XlsxColumnAttribute.ColumnIndex specified
            foreach (var prop in props)
            {
                if (prop.HasXlsxIgnoreAttribute())
                {
                    continue;
                }

                var pi = new PopulateCellInfo();
                var colInfo = prop.GetXlsxColumnAttribute();
                pi.Name = colInfo?.Name == null ? prop.Name : colInfo.Name;
                pi.ColumnIndex = colInfo?.ColumnIndex != null ? colInfo.ColumnIndex : -1; // -1 will later be reassigned
                pi.TempColumnIndex = tempCol++;

                if (pi.ColumnIndex > maxCol)
                {
                    maxCol = pi.ColumnIndex;
                }

                pi.Property = prop;
                cols[pi.TempColumnIndex] = pi;
            }

            // Slot any Columns without an order after maxCol
            for (int i = 0; i < tempCol; i++)
            {
                var pi = cols[i];
                if (pi.ColumnIndex == -1)
                {
                    pi.ColumnIndex = ++maxCol;
                }
            }

            return cols;
        }

        private static void AddToPopulateCache(Type type, Dictionary<int, PopulateCellInfo> cols)
        {
            PopulateCache.Value.AddOrUpdate(type, cols, (u1, u2) => cols);
        }

        private static Dictionary<int, PopulateCellInfo> TryGetFromCache(Type type)
        {
            if (!PopulateCache.IsValueCreated)
            {
                return null;
            }

            if (PopulateCache.Value.TryGetValue(type, out var cached))
            {
                return cached;
            }

            return null;
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