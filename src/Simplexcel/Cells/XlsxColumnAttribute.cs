using System;

namespace Simplexcel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class XlsxColumnAttribute : Attribute
    {
        /// <summary>
        /// The name of the Column, used as the Header row
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The Index of the Column, e.g., "0" for "A".
        /// If there are Columns with and without an Index, the columns
        /// without an Index will be added after the last column with an Index.
        /// </summary>
        public int? ColumnIndex { get; set; }
    }

    /// <summary>
    /// This attribute causes <see cref="Worksheet.FromData"/> to ignore the property completely
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class XlsxIgnoreColumnAttribute : Attribute
    {
    }
}
