using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Simplexcel.XlsxInternal;

namespace Simplexcel
{
    /// <summary>
    /// An Excel Workbook
    /// </summary>
    public sealed class Workbook
    {
        private readonly List<Worksheet> _sheets = new List<Worksheet>();

        /// <summary>
        /// The Worksheets in this Workbook
        /// </summary>
        public IEnumerable<Worksheet> Sheets { get { return _sheets.AsEnumerable(); } }

        /// <summary>
        /// The title of the Workbook
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The author of the Workbook
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Add a worksheet to this workbook. Sheet names must be unique.
        /// </summary>
        /// <exception cref="ArgumentException">Thrown if a sheet with the same Name already exists</exception>
        /// <param name="sheet"></param>
        public void Add(Worksheet sheet)
        {
            // According to ECMA-376, sheet names must be unique
            if (_sheets.Any(s => s.Name == sheet.Name))
            {
                throw new ArgumentException("This workbook already contains a sheet named " + sheet.Name);
            }

            _sheets.Add(sheet);
        }

        /// <summary>
        /// Save this workbook to a file, overwriting the file if it exists
        /// </summary>
        public void Save(string filename)
        {
            Save(filename, CompressionLevel.Balanced);
        }

        /// <summary>
        /// Save this workbook to a file, overwriting the file if it exists
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="compressionLevel">How much should content in a workbook be compressed? (Higher CPU Usage)</param>
        public void Save(string filename, CompressionLevel compressionLevel)
        {
            using (var fs = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                Save(fs, compressionLevel);
            }
        }

        /// <summary>
        /// Save this workbook to the given stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="compressionLevel">How much should content in a workbook be compressed? (Higher CPU Usage)</param>
        public void Save(Stream stream, CompressionLevel compressionLevel)
        {
            using (var ms = XlsxWriter.Save(this, compressionLevel))
            {
                ms.CopyTo(stream);
            }
        }
    }
}