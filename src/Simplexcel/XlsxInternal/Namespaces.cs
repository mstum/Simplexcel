using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// Various xmlns namespaces
    /// </summary>
    internal static class Namespaces
    {
        // ReSharper disable InconsistentNaming
        internal static readonly XNamespace simplexcel = "http://stum.de/simplexcel/v1";

        internal static readonly XNamespace coreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        internal static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
        internal static readonly XNamespace dcterms = "http://purl.org/dc/terms/";
        internal static readonly XNamespace extendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        internal static readonly XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        internal static readonly XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal static readonly XNamespace relationship = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal static readonly XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        internal static readonly XNamespace workbook = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal static readonly XNamespace workbookRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal static readonly XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal static readonly XNamespace x14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
        internal static readonly XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        // ReSharper restore InconsistentNaming
    }
}