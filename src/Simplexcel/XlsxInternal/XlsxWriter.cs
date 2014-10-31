using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Simplexcel.XlsxInternal
{
    internal static class XlsxWriter
    {
        /// <summary>
        /// Save a workbook to a MemoryStream
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if there are no <see cref="Worksheet">sheets</see> in the workbook.</exception>
        /// <returns></returns>
        internal static void Save(Workbook workbook, CompressionLevel compressionLevel, Stream outputStream)
        {
            if (workbook.SheetCount == 0)
            {
                throw new InvalidOperationException("You are trying to save a Workbook that does not contain any Worksheets.");
            }

            CompressionOption option;
            switch (compressionLevel)
            {
                case CompressionLevel.Balanced:
                    option = CompressionOption.Normal;
                    break;
                case CompressionLevel.Maximum:
                    option = CompressionOption.Maximum;
                    break;
                case CompressionLevel.NoCompression:
                default:
                    option = CompressionOption.NotCompressed;
                    break;
            }

            var writer = new XlsxWriterInternal(workbook, option);
            writer.Save(outputStream);
        }

        /// <summary>
        /// This does the actual writing by manually creating the XML according to ECMA-376, 3rd Edition, Part 1's SpreadsheetML.
        /// Due to the way state is handled, Save() can only be called once, hence it's a private class around an internal static member that initialized it anew on every Save(Workbook)
        /// 
        /// Note that in many cases, the order of elements in an XML file matters!
        /// </summary>
        private sealed class XlsxWriterInternal
        {
            private readonly Workbook _workbook;
            private readonly RelationshipCounter _relationshipCounter;
            private readonly SharedStrings _sharedStrings;
            private readonly XlsxPackage _package;
            private readonly IList<XlsxCellStyle> _styles;

            // For some reason, Excel interprets column widths as the width minus this factor
            private const decimal ExcelColumnWidthDifference = 0.7109375m;

            public XlsxWriterInternal(Workbook workbook, CompressionOption compressionOption)
            {
                _workbook = workbook;
                _relationshipCounter = new RelationshipCounter();
                _sharedStrings = new SharedStrings();
                _package = new XlsxPackage { CompressionOption = compressionOption };
                _styles = GetXlsxStyles();
            }

            internal void Save(Stream outputStream)
            {
                // docProps/core.xml
                var cp = CreateCoreFileProperties();
                _package.XmlFiles.Add(cp.Target);
                _package.Relationships.Add(cp);

                // xl/styles.xml
                var styles = StyleWriter.CreateStyleXml(_styles);
                var stylesRel = new Relationship(_relationshipCounter)
                {
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                    Target = styles
                };
                _package.XmlFiles.Add(styles);
                _package.WorkbookRelationships.Add(stylesRel);

                // xl/worksheets/sheetX.xml
                var sheetinfos = new List<SheetPackageInfo>();
                int i = 0;
                foreach (var sheet in _workbook.Sheets)
                {
                    i++;
                    XmlFile sheetRels;
                    var rel = CreateSheetFile(sheet, i, out sheetRels);
                    if (sheetRels != null)
                    {
                        _package.XmlFiles.Add(sheetRels);
                    }

                    _package.XmlFiles.Add(rel.Target);
                    _package.WorkbookRelationships.Add(rel);

                    var sheetinfo = new SheetPackageInfo
                    {
                        RelationshipId = rel.Id,
                        SheetId = i,
                        SheetName = sheet.Name
                    };
                    if (sheet.PageSetup.PrintRepeatColumns > 0)
                    {
                        sheetinfo.RepeatCols = "'" + sheet.Name + "'!$A:$" + CellAddressHelper.ColToReference(sheet.PageSetup.PrintRepeatColumns-1);
                    }
                    if (sheet.PageSetup.PrintRepeatRows > 0)
                    {
                        sheetinfo.RepeatRows = "'" + sheet.Name + "'!$1:$" + sheet.PageSetup.PrintRepeatRows;
                    }

                    sheetinfos.Add(sheetinfo);
                }

                // xl/sharedStrings.xml
                if (_sharedStrings.Count > 0)
                {
                    var ssx = _sharedStrings.ToRelationship(_relationshipCounter);
                    _package.XmlFiles.Add(ssx.Target);
                    _package.WorkbookRelationships.Add(ssx);
                }

                // xl/workbook.xml
                var wb = CreateWorkbookFile(sheetinfos);
                _package.XmlFiles.Add(wb.Target);
                _package.Relationships.Add(wb);

                // xl/_rels/workbook.xml.rels
                _package.SaveToStream(outputStream);
            }

            /// <summary>
            /// Get all unique Styles throughout the workbook
            /// </summary>
            /// <returns>A dictionary where the style is the </returns>
            private IList<XlsxCellStyle> GetXlsxStyles()
            {
                var result = from sheet in _workbook.Sheets
                             from cells in sheet.Cells
                             select cells.Value.XlsxCellStyle;

                return result.Distinct().ToList();
            }

            /// <summary>
            /// Generated the docProps/core.xml which contains author, creation date etc.
            /// </summary>
            /// <returns></returns>
            private Relationship CreateCoreFileProperties()
            {
                var file = new XmlFile();
                file.ContentType = "application/vnd.openxmlformats-package.core-properties+xml";
                file.Path = "docProps/core.xml";

                var dc = Namespaces.dc;
                var dcterms = Namespaces.dcterms;
                var xsi = Namespaces.xsi;
                var cp = Namespaces.coreProperties;

                var doc = new XDocument();
                var root = new XElement(cp + "coreProperties",
                                new XAttribute(XNamespace.Xmlns + "cp", cp),
                                new XAttribute(XNamespace.Xmlns + "dc", dc),
                                new XAttribute(XNamespace.Xmlns + "dcterms", dcterms),
                                new XAttribute(XNamespace.Xmlns + "xsi", xsi)
                                );

                if (!string.IsNullOrEmpty(_workbook.Title))
                {
                    root.Add(new XElement(dc + "title", _workbook.Title));
                }
                if (!string.IsNullOrEmpty(_workbook.Author))
                {
                    root.Add(new XElement(dc + "creator", _workbook.Author));
                    root.Add(new XElement(cp + "lastModifiedBy", _workbook.Author));
                }

                root.Add(new XElement(dcterms + "created", DateTime.UtcNow, new XAttribute(xsi + "type", "dcterms:W3CDTF")));
                root.Add(new XElement(dcterms + "modified", DateTime.UtcNow, new XAttribute(xsi + "type", "dcterms:W3CDTF")));


                doc.Add(root);

                file.Content = doc;

                var rel = new Relationship(_relationshipCounter)
                {
                    Target = file,
                    Type = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                };

                return rel;
            }

            /// <summary>
            /// Create the xl/workbook.xml file and associated relationship
            /// </summary>
            /// <param name="sheetInfos"></param>
            /// <returns></returns>
            private Relationship CreateWorkbookFile(List<SheetPackageInfo> sheetInfos)
            {
                var file = new XmlFile();
                file.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
                file.Path = "xl/workbook.xml";

                var doc = new XDocument(new XElement(Namespaces.workbook + "workbook",
                        new XAttribute("xmlns", Namespaces.workbook),
                        new XAttribute(XNamespace.Xmlns + "r", Namespaces.workbookRelationship)
                    ));

                var sheets = new XElement(Namespaces.workbook + "sheets");
                foreach (var si in sheetInfos)
                {
                    sheets.Add(new XElement(Namespaces.workbook + "sheet",
                                        new XAttribute("name", si.SheetName),
                                        new XAttribute("sheetId", si.SheetId),
                                        new XAttribute(Namespaces.workbookRelationship + "id", si.RelationshipId)
                                        ));
                }

                doc.Root.Add(sheets);

                var repeatInfos = sheetInfos.Where(si => !string.IsNullOrEmpty(si.RepeatRows) || !string.IsNullOrEmpty(si.RepeatCols)).OrderBy(si => si.SheetId).ToList();
                if (repeatInfos.Count > 0)
                {
                    var dne = new XElement(Namespaces.workbook + "definedNames");
                    foreach (var re in repeatInfos)
                    {
                        var de = new XElement(Namespaces.workbook + "definedName",
                            new XAttribute("name", "_xlnm.Print_Titles"),
                            new XAttribute("localSheetId", re.SheetId - 1)
                        );

                        if (!string.IsNullOrEmpty(re.RepeatCols) && !string.IsNullOrEmpty(re.RepeatRows))
                        {
                            de.Add(new XText(re.RepeatCols + "," + re.RepeatRows));
                        }
                        else if (!string.IsNullOrEmpty(re.RepeatCols))
                        {
                            de.Add(new XText(re.RepeatCols));
                        }
                        else if (!string.IsNullOrEmpty(re.RepeatRows))
                        {
                            de.Add(new XText(re.RepeatRows));
                        }

                        dne.Add(de);
                    }
                    doc.Root.Add(dne);
                }

                file.Content = doc;

                var rel = new Relationship(_relationshipCounter)
                {
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    Target = file
                };

                return rel;
            }

            /// <summary>
            /// Create the xl/worksheets/sheetX.xml file
            /// </summary>
            /// <param name="sheet"></param>
            /// <param name="sheetIndex"></param>
            /// <param name="sheetRels">If this worksheet needs an xl/worksheets/_rels/sheetX.xml.rels file</param>
            /// <returns></returns>
            private Relationship CreateSheetFile(Worksheet sheet, int sheetIndex, out XmlFile sheetRels)
            {
                var rows = GetXlsxRows(sheet);

                var file = new XmlFile();
                file.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
                file.Path = "xl/worksheets/sheet" + sheetIndex + ".xml";

                var doc = new XDocument(new XElement(Namespaces.workbook + "worksheet",
                    new XAttribute("xmlns", Namespaces.workbook),
                    new XAttribute(XNamespace.Xmlns + "r", Namespaces.relationship),
                    new XAttribute(XNamespace.Xmlns + "mc", Namespaces.mc),
                    new XAttribute(XNamespace.Xmlns + "x14ac", Namespaces.x14ac),
                    new XAttribute(XNamespace.Xmlns + "or", Namespaces.officeRelationships),
                    new XAttribute(Namespaces.mc + "Ignorable", "x14ac")
                ));

                var sheetFormatPr = new XElement(Namespaces.workbook + "sheetFormatPr");
                sheetFormatPr.Add(new XAttribute("defaultRowHeight", 15));
                doc.Root.Add(sheetFormatPr);

                if (sheet.ColumnWidths.Any())
                {
                    var cols = new XElement(Namespaces.workbook + "cols");
                    foreach (var cw in sheet.ColumnWidths)
                    {
                        var rowId = cw.Key + 1;
                        var col = new XElement(Namespaces.workbook + "col",
                            new XAttribute("min", rowId),
                            new XAttribute("max", rowId),
                            new XAttribute("width", (decimal)cw.Value + ExcelColumnWidthDifference),
                            new XAttribute("customWidth", 1));
                        cols.Add(col);
                    }
                    doc.Root.Add(cols);
                }

                var sheetData = new XElement(Namespaces.workbook + "sheetData");
                foreach (var row in rows.OrderBy(rk => rk.Key))
                {
                    var re = new XElement(Namespaces.workbook + "row", new XAttribute("r", row.Value.RowIndex));
                    foreach (var cell in row.Value.Cells)
                    {
                        var ce = new XElement(Namespaces.workbook + "c",
                            new XAttribute("r", cell.Reference),
                            new XAttribute("t", cell.CellType),
                            new XAttribute("s", cell.StyleIndex),
                            new XElement(Namespaces.workbook + "v", cell.Value));

                        re.Add(ce);
                    }
                    sheetData.Add(re);
                }
                doc.Root.Add(sheetData);

                var hyperlinks = sheet.Cells.Where(c => c.Value != null && !string.IsNullOrEmpty(c.Value.Hyperlink)).ToList();
                if (hyperlinks.Count > 0)
                {
                    sheetRels = new XmlFile();
                    sheetRels.Path = "xl/worksheets/_rels/sheet" + sheetIndex + ".xml.rels";
                    sheetRels.ContentType = "application/vnd.openxmlformats-package.relationships+xml";

                    var hlRelsElem = new XElement(Namespaces.relationship + "Relationships");

                    var hlElem = new XElement(Namespaces.workbook + "hyperlinks");
                    for (int i = 0; i <= hyperlinks.Count - 1; i++)
                    {
                        string hyperLinkRelId = "rId" + (i + 1);

                        var link = hyperlinks[i];
                        var linkElem = new XElement(Namespaces.workbook + "hyperlink",
                                                    new XAttribute("ref", link.Key.ToString()),
                                                    new XAttribute(Namespaces.officeRelationships + "id", hyperLinkRelId)
                                );
                        hlElem.Add(linkElem);

                        hlRelsElem.Add(new XElement(Namespaces.relationship + "Relationship",
                            new XAttribute("Id", hyperLinkRelId),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                            new XAttribute("Target", link.Value.Hyperlink),
                            new XAttribute("TargetMode", "External")));
                    }
                    doc.Root.Add(hlElem);
                    sheetRels.Content = new XDocument();
                    sheetRels.Content.Add(hlRelsElem);
                }
                else
                {
                    sheetRels = null;
                }

                var pageSetup = new XElement(Namespaces.workbook + "pageSetup");
                pageSetup.Add(new XAttribute("orientation", sheet.PageSetup.Orientation == Orientation.Portrait ? "portrait" : "landscape"));
                doc.Root.Add(pageSetup);

                file.Content = doc;
                var rel = new Relationship(_relationshipCounter)
                {
                    Target = file,
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                };

                return rel;
            }

            private Dictionary<int, XlsxRow> GetXlsxRows(Worksheet sheet)
            {
                var rows = new Dictionary<int, XlsxRow>();

                // The order matters!
                foreach (var cell in sheet.Cells.OrderBy(c => c.Key.Row).ThenBy(c => c.Key.Column))
                {
                    if (!rows.ContainsKey(cell.Key.Row))
                    {
                        rows[cell.Key.Row] = new XlsxRow {RowIndex = cell.Key.Row + 1};
                    }

                    var styleIndex = _styles.IndexOf(cell.Value.XlsxCellStyle) + 1;

                    var xc = new XlsxCell
                    {
                            StyleIndex = styleIndex,
                            Reference = cell.Key.ToString()
                    };

                    switch (cell.Value.CellType)
                    {
                        case CellType.Text:
                            xc.CellType = XlsxCellTypes.SharedString;
                            xc.Value = _sharedStrings.GetStringIndex((string)cell.Value.Value);
                            break;
                        case CellType.Number:
                            xc.CellType = XlsxCellTypes.Number;
                            xc.Value = ((Decimal)cell.Value.Value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                            break;
                        default:
                            throw new ArgumentException("Unknown Cell Type: " + cell.Value.CellType + " in cell " + cell.Key.ToString() + " of " + sheet.Name);
                    }

                    rows[cell.Key.Row].Cells.Add(xc);
                }
                return rows;
            }
        }
    }
}
