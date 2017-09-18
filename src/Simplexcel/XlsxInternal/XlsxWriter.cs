using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    internal static class XlsxWriter
    {
        /// <summary>
        /// Save a workbook to a MemoryStream
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if there are no <see cref="Worksheet">sheets</see> in the workbook.</exception>
        /// <returns></returns>
        internal static void Save(Workbook workbook, Stream outputStream, bool compress)
        {
            if (workbook.SheetCount == 0)
            {
                throw new InvalidOperationException("You are trying to save a Workbook that does not contain any Worksheets.");
            }

            XlsxWriterInternal.Save(workbook, outputStream, compress);
        }

        /// <summary>
        /// This does the actual writing by manually creating the XML according to ECMA-376, 3rd Edition, Part 1's SpreadsheetML.
        /// 
        /// Note that in many cases, the order of elements in an XML file matters!
        /// </summary>
        private static class XlsxWriterInternal
        {
            // For some reason, Excel interprets column widths as the width minus this factor
            private const decimal ExcelColumnWidthDifference = 0.7109375m;

            internal static void Save(Workbook workbook, Stream outputStream, bool compress)
            {
                var relationshipCounter = new RelationshipCounter();
                var package = new XlsxPackage();
                var styles = GetXlsxStyles(workbook);
                var sharedStrings = new SharedStrings();

                // docProps/core.xml
                var cp = CreateCoreFileProperties(workbook, relationshipCounter);
                package.XmlFiles.Add(cp.Target);
                package.Relationships.Add(cp);

                // xl/styles.xml
                var stylesXml = StyleWriter.CreateStyleXml(styles);
                var stylesRel = new Relationship(relationshipCounter)
                {
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                    Target = stylesXml
                };
                package.XmlFiles.Add(stylesXml);
                package.WorkbookRelationships.Add(stylesRel);

                // xl/worksheets/sheetX.xml
                var sheetinfos = new List<SheetPackageInfo>();
                int i = 0;
                foreach (var sheet in workbook.Sheets)
                {
                    i++;
                    XmlFile sheetRels;
                    var rel = CreateSheetFile(sheet, i, relationshipCounter, styles, sharedStrings, out sheetRels);
                    if (sheetRels != null)
                    {
                        package.XmlFiles.Add(sheetRels);
                    }

                    package.XmlFiles.Add(rel.Target);
                    package.WorkbookRelationships.Add(rel);

                    var sheetinfo = new SheetPackageInfo
                    {
                        RelationshipId = rel.Id,
                        SheetId = i,
                        SheetName = sheet.Name
                    };
                    if (sheet.PageSetup.PrintRepeatColumns > 0)
                    {
                        sheetinfo.RepeatCols = "'" + sheet.Name + "'!$A:$" + CellAddressHelper.ColToReference(sheet.PageSetup.PrintRepeatColumns - 1);
                    }
                    if (sheet.PageSetup.PrintRepeatRows > 0)
                    {
                        sheetinfo.RepeatRows = "'" + sheet.Name + "'!$1:$" + sheet.PageSetup.PrintRepeatRows;
                    }

                    sheetinfos.Add(sheetinfo);
                }

                // xl/sharedStrings.xml
                if (sharedStrings.Count > 0)
                {
                    var ssx = sharedStrings.ToRelationship(relationshipCounter);
                    package.XmlFiles.Add(ssx.Target);
                    package.WorkbookRelationships.Add(ssx);
                }

                // xl/workbook.xml
                var wb = CreateWorkbookFile(sheetinfos, relationshipCounter);
                package.XmlFiles.Add(wb.Target);
                package.Relationships.Add(wb);

                // xl/_rels/workbook.xml.rels
                package.SaveToStream(outputStream, compress);
            }

            /// <summary>
            /// Get all unique Styles throughout the workbook
            /// </summary>
            /// <returns>A dictionary where the style is the </returns>
            private static IList<XlsxCellStyle> GetXlsxStyles(Workbook workbook)
            {
                var result = from sheet in workbook.Sheets
                             from cells in sheet.Cells
                             select cells.Value.XlsxCellStyle;

                return result.Distinct().ToList();
            }

            /// <summary>
            /// Generated the docProps/core.xml which contains author, creation date etc.
            /// </summary>
            /// <returns></returns>
            private static Relationship CreateCoreFileProperties(Workbook workbook, RelationshipCounter relationshipCounter)
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

                if (!string.IsNullOrEmpty(workbook.Title))
                {
                    root.Add(new XElement(dc + "title", workbook.Title));
                }
                if (!string.IsNullOrEmpty(workbook.Author))
                {
                    root.Add(new XElement(dc + "creator", workbook.Author));
                    root.Add(new XElement(cp + "lastModifiedBy", workbook.Author));
                }

                root.Add(new XElement(dcterms + "created", DateTime.UtcNow, new XAttribute(xsi + "type", "dcterms:W3CDTF")));
                root.Add(new XElement(dcterms + "modified", DateTime.UtcNow, new XAttribute(xsi + "type", "dcterms:W3CDTF")));


                doc.Add(root);

                file.Content = doc;

                var rel = new Relationship(relationshipCounter)
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
            private static Relationship CreateWorkbookFile(List<SheetPackageInfo> sheetInfos, RelationshipCounter relationshipCounter)
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

                var rel = new Relationship(relationshipCounter)
                {
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    Target = file
                };

                return rel;
            }

            private static void WriteSheetViews(Worksheet sheet, XDocument doc)
            {
                var sviews = sheet.GetSheetViews();
                if (sviews == null || sviews.Count == 0) { return; }

                var sheetViews = new XElement(Namespaces.workbook + "sheetViews");
                foreach (var sv in sviews)
                {
                    var sheetView = new XElement(Namespaces.workbook + "sheetView");
                    if (sv.ShowRuler.HasValue)
                    {
                        sheetView.Add(new XAttribute("showRuler", sv.ShowRuler.Value ? "1" : "0"));
                    }
                    if (sv.TabSelected.HasValue)
                    {
                        sheetView.Add(new XAttribute("tabSelected", sv.TabSelected.Value ? "1" : "0"));
                    }
                    // TODO: Consider adding a <bookViews> element to the workbook
                    sheetView.Add(new XAttribute("workbookViewId", sv.WorkbookViewId));

                    if (sv.Pane != null)
                    {
                        var p = sv.Pane;
                        var paneElem = new XElement(Namespaces.workbook + "pane");
                        if (p.XSplit.HasValue && p.XSplit.Value != Pane.DefaultXSplit)
                        {
                            paneElem.Add(new XAttribute("xSplit", p.XSplit.Value));
                        }
                        if (p.YSplit.HasValue && p.YSplit.Value != Pane.DefaultYSplit)
                        {
                            paneElem.Add(new XAttribute("ySplit", p.YSplit.Value));
                        }
                        if (p.ActivePane.HasValue)
                        {
                            paneElem.Add(new XAttribute("activePane", p.ActivePane.Value.GetXmlValue()));
                        }
                        if (!string.IsNullOrEmpty(p.TopLeftCell))
                        {
                            paneElem.Add(new XAttribute("topLeftCell", p.TopLeftCell));
                        }
                        if (p.State.HasValue)
                        {
                            paneElem.Add(new XAttribute("state", p.State.Value.GetXmlValue()));
                        }
                        sheetView.Add(paneElem);
                    }

                    foreach (var sel in sv.Selections ?? Enumerable.Empty<Selection>())
                    {
                        var selElem = new XElement(Namespaces.workbook + "selection");
                        if (!string.IsNullOrEmpty(sel.ActiveCell))
                        {
                            selElem.Add(new XAttribute("activeCell", sel.ActiveCell));
                        }
                        selElem.Add(new XAttribute("pane", sel.ActivePane.GetXmlValue()));
                        sheetView.Add(selElem);
                    }
                    sheetViews.Add(sheetView);
                }

                doc.Root.Add(sheetViews);
            }

            private static void WritePageBreaks(Worksheet sheet, XDocument doc)
            {
                XElement BreakToXml(PageBreak brk)
                {
                    var elem = new XElement(Namespaces.workbook + "brk");
                    elem.Add(new XAttribute("id", brk.Id));
                    if (brk.IsManualBreak)
                    {
                        elem.Add(new XAttribute("man", "1"));
                    }
                    if (brk.IsPivotCreatedPageBreak)
                    {
                        elem.Add(new XAttribute("pt", "1"));
                    }
                    if (brk.Min > 0)
                    {
                        elem.Add(new XAttribute("min", brk.Min));
                    }
                    if (brk.Max > 0)
                    {
                        elem.Add(new XAttribute("max", brk.Max));
                    }
                    return elem;
                }

                XElement PageBreakCollectionToXml(string elementName, ICollection<PageBreak> breaks)
                {
                    var elem = new XElement(Namespaces.workbook + elementName);
                    elem.Add(new XAttribute("count", breaks.Count));
                    elem.Add(new XAttribute("manualBreakCount", breaks.Count(r => r.IsManualBreak)));
                    foreach (var brk in breaks)
                    {
                        elem.Add(BreakToXml(brk));
                    }
                    return elem;
                }

                var rowBreaks = sheet.GetRowBreaks();
                if (rowBreaks != null && rowBreaks.Count > 0)
                {
                    var rowBreaksElem = PageBreakCollectionToXml("rowBreaks", rowBreaks);
                    doc.Root.Add(rowBreaksElem);
                }

                var colBreaks = sheet.GetColumnBreaks();
                if (colBreaks != null && colBreaks.Count > 0)
                {
                    var colBreaksElem = PageBreakCollectionToXml("colBreaks", colBreaks);
                    doc.Root.Add(colBreaksElem);
                }
            }

            /// <summary>
            /// Create the xl/worksheets/sheetX.xml file
            /// </summary>
            /// <param name="sheet"></param>
            /// <param name="sheetIndex"></param>
            /// <param name="sheetRels">If this worksheet needs an xl/worksheets/_rels/sheetX.xml.rels file</param>
            /// <returns></returns>
            private static Relationship CreateSheetFile(Worksheet sheet, int sheetIndex, RelationshipCounter relationshipCounter, IList<XlsxCellStyle> styles, SharedStrings sharedStrings, out XmlFile sheetRels)
            {
                var rows = GetXlsxRows(sheet, styles, sharedStrings);

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

                WriteSheetViews(sheet, doc);

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

                sheetRels = null;
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

                var pageSetup = new XElement(Namespaces.workbook + "pageSetup");
                pageSetup.Add(new XAttribute("orientation", sheet.PageSetup.Orientation == Orientation.Portrait ? "portrait" : "landscape"));
                doc.Root.Add(pageSetup);

                WritePageBreaks(sheet, doc);

                file.Content = doc;
                var rel = new Relationship(relationshipCounter)
                {
                    Target = file,
                    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                };

                return rel;
            }

            private static Dictionary<int, XlsxRow> GetXlsxRows(Worksheet sheet, IList<XlsxCellStyle> styles, SharedStrings sharedStrings)
            {
                var rows = new Dictionary<int, XlsxRow>();

                // The order matters!
                foreach (var cell in sheet.Cells.OrderBy(c => c.Key.Row).ThenBy(c => c.Key.Column))
                {
                    if (!rows.ContainsKey(cell.Key.Row))
                    {
                        rows[cell.Key.Row] = new XlsxRow { RowIndex = cell.Key.Row + 1 };
                    }

                    var styleIndex = styles.IndexOf(cell.Value.XlsxCellStyle) + 1;

                    var xc = new XlsxCell
                    {
                        StyleIndex = styleIndex,
                        Reference = cell.Key.ToString()
                    };

                    switch (cell.Value.CellType)
                    {
                        case CellType.Text:
                            xc.CellType = XlsxCellTypes.SharedString;
                            xc.Value = sharedStrings.GetStringIndex((string)cell.Value.Value);
                            break;
                        case CellType.Number:
                            xc.CellType = XlsxCellTypes.Number;
                            xc.Value = ((Decimal)cell.Value.Value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                            break;
                        case CellType.Date:
                            xc.CellType = XlsxCellTypes.Number;
                            if (cell.Value.Value != null)
                            {
                                xc.Value = ((DateTime)cell.Value.Value).ToOleAutDate();
                            }
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