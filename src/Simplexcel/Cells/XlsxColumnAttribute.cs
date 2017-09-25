using System;

namespace Simplexcel.Cells
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    internal class XlsxColumnAttribute : Attribute
    {
        /// <summary>
        /// The name of the Column, used as the Header row
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// This attribute causes <see cref="Worksheet.FromData"/> to ignore the property completely
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    internal class XlsxIgnoreColumnAttribute : Attribute
    {
    }
}
