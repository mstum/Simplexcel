using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// An XML File in the package
    /// </summary>
    internal class XmlFile
    {
        /// <summary>
        /// The path to the file within the package, without leading /
        /// </summary>
        internal string Path { get; set; }

        /// <summary>
        /// The Content Type of the file (default: application/xml)
        /// </summary>
        internal string ContentType { get; set; }

        /// <summary>
        /// The actual file content
        /// </summary>
        internal XDocument Content { get; set; }

        public XmlFile()
        {
            ContentType = "application/xml";
        }
    }
}
