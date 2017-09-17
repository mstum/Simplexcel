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
        /// How many <see cref="Worksheet">sheets</see> are in the Workbook currently?
        /// </summary>
        public int SheetCount
        {
            get { return _sheets.Count; }
        }

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
        /// <param name="filename"></param>
        /// <param name="compress">Use compression? (Smaller files/higher CPU Usage)</param>
        /// <exception cref="InvalidOperationException">Thrown if there are no <see cref="Worksheet">sheets</see> in the workbook.</exception>
        public void Save(string filename, bool compress = true)
        {
            if(string.IsNullOrEmpty(filename))
            {
                throw new ArgumentException("Invalid Filename.", nameof(filename));
            }

            using (var fs = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                Save(fs, compress);
            }
        }

        /// <summary>
        /// Save this workbook to the given stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="compress">Use compression? (Smaller files/higher CPU Usage)</param>
        /// <exception cref="InvalidOperationException">Thrown if there are no <see cref="Worksheet">sheets</see> in the workbook.</exception>
        public void Save(Stream stream, bool compress = true)
        {
            if(stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if(!stream.CanWrite)
            {
                throw new InvalidOperationException("Stream to save to is not writeable.");
            }

            if (stream.CanSeek)
            {
                XlsxWriter.Save(this, stream, compress);
            }
            else
            {
                // ZipArchive needs a seekable stream. If a stream is not seekable (e.g., HttpContext.Response.OutputStream), wrap it in a MemoryStream instead.
                // TODO: Can we guess the required capacity?
                using(var ms = new MemoryStream())
                {
                    XlsxWriter.Save(this, ms, compress);
                    ms.CopyTo(stream);
                }
            }
        }
    }
}