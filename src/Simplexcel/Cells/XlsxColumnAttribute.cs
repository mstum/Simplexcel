using System;

namespace Simplexcel
{
    /// <summary>
    /// Used with <see cref="Worksheet.Populate"/> or <see cref="Worksheet.FromData"/>, allows setting how object properties are handled.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class XlsxColumnAttribute : Attribute
    {
        /// <summary>
        /// The name of the Column, used as the Header row.
        /// If this is NULL, the Property Name will be used.
        /// If this is an Empty string, then no text will be in the Cell header.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The Index of the Column, e.g., "0" for "A".
        /// If there are Columns with and without an Index, the columns
        /// without an Index will be added after the last column with an Index.
        /// 
        /// It is recommended that either all object properties or none specify Column Indexes
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// Create a new XlsxColumnAttribute
        /// </summary>
        public XlsxColumnAttribute()
        {
            ColumnIndex = -1;
        }

        /// <summary>
        /// Create a new XlsxColumnAttribute with the given name
        /// </summary>
        /// <param name="name"></param>
        public XlsxColumnAttribute(string name)
        {
            Name = name;
            ColumnIndex = -1;
        }
    }

    /// <summary>
    /// This attribute causes <see cref="Worksheet.FromData"/> to ignore the property completely
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class XlsxIgnoreColumnAttribute : Attribute
    {
    }
}
