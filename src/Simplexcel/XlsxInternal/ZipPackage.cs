using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    internal sealed class ZipPackage : IDisposable
    {
        private readonly ZipArchive _archive;
        private readonly IDictionary<string, string> _contentTypes;
        private readonly IList<Relationship> _relationships;
        private bool _closed;
        private readonly bool _compress;

        internal ZipPackage(Stream underlyingStream, bool useCompression)
        {
            if (underlyingStream == null)
            {
                throw new ArgumentNullException(nameof(underlyingStream));
            }

            _archive = new ZipArchive(underlyingStream, ZipArchiveMode.Create, true);
            _contentTypes = new Dictionary<string, string>();
            _relationships = new List<Relationship>();
            _compress = useCompression;
        }

        internal void AddRelationship(Relationship rel)
        {
            CheckClosed();
            _relationships.Add(rel);
        }

        internal void Close()
        {
            CheckClosed();
            WriteContentTypes();
            WriteRelationships();
            _archive.Dispose();
            _closed = true;
        }

        public void Dispose()
        {
            if (!_closed)
            {
                Close();
            }
        }


        private void CheckClosed()
        {
            if (_closed)
            {
                throw new InvalidOperationException("This " + nameof(ZipPackage) + " is already closed.");
            }
        }

        internal void WriteXmlFile(XmlFile file)
        {
            CheckClosed();
            _contentTypes["/" + file.Path] = file.ContentType;
            AddFile(file.Path, file.Content);
        }

        private void AddFile(string relativeFileName, XDocument xdoc)
        {
            CheckClosed();

            if (xdoc == null)
            {
                throw new ArgumentNullException(nameof(xdoc));
            }

            byte[] content = Encoding.UTF8.GetBytes(xdoc.ToString());
            AddFile(relativeFileName, content);
        }

        private void AddFile(string relativeFileName, byte[] content)
        {
            CheckClosed();
            if (string.IsNullOrEmpty(relativeFileName))
            {
                throw new ArgumentException("Invalid " + nameof(relativeFileName));
            }
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            var entry = _archive.CreateEntry(relativeFileName, _compress ? CompressionLevel.Optimal : CompressionLevel.NoCompression);

            using (var stream = entry.Open())
            {
                stream.Write(content, 0, content.Length);
            }
        }

        private void WriteContentTypes()
        {
            CheckClosed();
            const string defaultXmlContentType = "application/xml";
            var root = new XElement(Namespaces.contenttypes + "Types");
            root.Add(new XElement(Namespaces.contenttypes + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", defaultXmlContentType)));
            root.Add(new XElement(Namespaces.contenttypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")));

            foreach (var ct in _contentTypes)
            {
                if (string.Equals(Path.GetExtension(ct.Key), ".xml", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.Equals(ct.Value, defaultXmlContentType))
                    {
                        var ovr = new XElement(Namespaces.contenttypes + "Override");
                        ovr.Add(new XAttribute("PartName", ct.Key));
                        ovr.Add(new XAttribute("ContentType", ct.Value));
                        root.Add(ovr);
                    }
                }
            }

            AddFile("[Content_Types].xml", new XDocument(root));
        }

        private void WriteRelationships()
        {
            CheckClosed();
            var root = new XElement(Namespaces.relationship + "Relationships");
            foreach(var rel in _relationships)
            {
                var re = new XElement(Namespaces.relationship + "Relationship");
                re.Add(new XAttribute("Type", rel.Type));
                re.Add(new XAttribute("Target", "/" + rel.Target.Path));
                re.Add(new XAttribute("Id", rel.Id));
                root.Add(re);
            }

            AddFile("_rels/.rels", new XDocument(root));
        }
    }
}