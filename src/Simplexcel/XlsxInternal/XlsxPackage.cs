﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// A Wrapper for an .xlsx file, containing all the contents and relations and logic to create the file from that content
    /// </summary>
    internal class XlsxPackage
    {
        /// <summary>
        /// All XML Files in this package
        /// </summary>
        internal IList<XmlFile> XmlFiles { get; private set; }

        /// <summary>
        /// Package-level relationships (/_rels/.rels)
        /// </summary>
        internal IList<Relationship> Relationships { get; private set; }

        /// <summary>
        /// Workbook-level relationships (/xl/_rels/workbook.xml.rels)
        /// </summary>
        internal IList<Relationship> WorkbookRelationships { get; private set; }


        internal XlsxPackage()
        {
            Relationships = new List<Relationship>();
            WorkbookRelationships = new List<Relationship>();
            XmlFiles = new List<XmlFile>();
        }

        /// <summary>
        /// Save the Xlsx Package to a new Stream (that the caller owns and has to dispose)
        /// </summary>
        /// <returns></returns>
        internal void SaveToStream(Stream outputStream)
        {
            using (var pkg = new ZipPackage(outputStream))
            {
                WriteInfoXmlFile(pkg);

                foreach (var file in XmlFiles)
                {
                    pkg.WriteXmlFile(file);

                    var relations = Relationships.Where(r => r.Target == file);
                    foreach (var rel in relations)
                    {
                        pkg.AddRelationship(rel);
                    }
                }

                if (WorkbookRelationships.Count > 0)
                {
                    pkg.WriteXmlFile(WorkbookRelsXml());
                }
            }
            outputStream.Seek(0, SeekOrigin.Begin);
        }

        private void WriteInfoXmlFile(ZipPackage pkg)
        {
            var version = typeof(Workbook).GetTypeInfo().Assembly.GetName().Version;

            var infoXml = new XmlFile();
            infoXml.Path = "simplexcel.xml";
            infoXml.Content = new XDocument(new XElement(Namespaces.simplexcel + "docInfo", new XAttribute("xmlns", Namespaces.simplexcel)));

            infoXml.Content.Root.Add(new XElement(Namespaces.simplexcel + "version",
                new XAttribute("major", version.Major),
                new XAttribute("minor", version.Minor),
                new XAttribute("build", version.Build),
                new XAttribute("revision", version.Revision)
            ));

            infoXml.Content.Root.Add(new XElement(Namespaces.simplexcel + "created", DateTime.UtcNow));

            pkg.WriteXmlFile(infoXml);
        }

        /// <summary>
        /// Create the xl/_rels/workbook.xml.rels file
        /// </summary>
        /// <returns></returns>
        internal XmlFile WorkbookRelsXml()
        {
            var file = new XmlFile();
            file.ContentType = "application/vnd.openxmlformats-package.relationships+xml";
            file.Path = "xl/_rels/workbook.xml.rels";

            var content = new XDocument(new XElement(Namespaces.relationship + "Relationships", new XAttribute("xmlns", Namespaces.relationship)));
            foreach (var rel in WorkbookRelationships)
            {
                var elem = new XElement(Namespaces.relationship + "Relationship",
                                    new XAttribute("Target", "/" + rel.Target.Path),
                                    new XAttribute("Type", rel.Type),
                                    new XAttribute("Id", rel.Id));

                content.Root.Add(elem);
            }
            file.Content = content;

            return file;
        }
    }
}
