using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    internal static class StyleWriter
    {
        // There are up to 164 built in number formats (0-163), all else are custom (164 and above)
        private const int CustomFormatIndex = 164;

        /// <summary>
        /// Create a styles.xml file
        /// </summary>
        /// <param name="styles"></param>
        /// <returns></returns>
        internal static XmlFile CreateStyleXml(IList<XlsxCellStyle> styles)
        {
            var uniqueBorders = styles.Select(s => s.Border).Where(s => s != CellBorder.None).Distinct().ToList();
            uniqueBorders.Insert(0, CellBorder.None);

            var numberFormats = styles.Select(s => s.Format).Distinct().ToList();
            var uniqueFonts = styles.Select(s => s.Font).Distinct().ToList();
            uniqueFonts.Insert(0, new XlsxFont());

            var file = new XmlFile
            {
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                Path = "xl/styles.xml"
            };

            var doc = new XDocument(new XElement(Namespaces.workbook + "styleSheet", new XAttribute("xmlns", Namespaces.workbook)));

            StyleAddNumFmtsElement(doc, numberFormats);
            StyleAddFontsElement(doc, uniqueFonts);
            StyleAddFillsElement(doc);
            StyleAddBordersElement(doc, uniqueBorders);
            StyleAddCellStyleXfsElement(doc);
            StyleAddCellXfsElement(doc, styles, uniqueBorders, numberFormats, uniqueFonts);

            file.Content = doc;

            return file;
        }

        private static void StyleAddCellXfsElement(XDocument doc, IList<XlsxCellStyle> styles, List<CellBorder> uniqueBorders, List<string> numberFormats, IList<XlsxFont> fontInfos)
        {
            var cellXfs = new XElement(Namespaces.workbook + "cellXfs", new XAttribute("count", styles.Count));

            // Default Cell Style, used to make sure the row numbers don't appear broken
            cellXfs.Add(new XElement(Namespaces.workbook + "xf",
                                     new XAttribute("numFmtId", 0),
                                     new XAttribute("fontId", 0),
                                     new XAttribute("fillId", 0),
                                     new XAttribute("borderId", 0),
                                     new XAttribute("xfId", 0)));

            foreach (var style in styles)
            {
                var numFmtId = numberFormats.IndexOf(style.Format) + CustomFormatIndex;
                var fontId = fontInfos.IndexOf(style.Font);
                var fillId = 0;
                var borderId = uniqueBorders.IndexOf(style.Border);
                var xfId = 0;

                var elem = new XElement(Namespaces.workbook + "xf",
                                        new XAttribute("numFmtId", numFmtId),
                                        new XAttribute("fontId", fontId),
                                        new XAttribute("fillId", fillId),
                                        new XAttribute("borderId", borderId),
                                        new XAttribute("xfId", xfId),
                                        new XAttribute("applyNumberFormat", 1));

                cellXfs.Add(elem);
            }
            doc.Root.Add(cellXfs);
        }

        private static void StyleAddCellStyleXfsElement(XDocument doc)
        {
            var cellStyleXfs = new XElement(Namespaces.workbook + "cellStyleXfs", new XAttribute("count", 1));
            cellStyleXfs.Add(new XElement(Namespaces.workbook + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0)));
            doc.Root.Add(cellStyleXfs);
        }

        private static void StyleAddNumFmtsElement(XDocument doc, List<string> numberFormats)
        {
            var numFmtsElem = new XElement(Namespaces.workbook + "numFmts", new XAttribute("count", numberFormats.Count + 1));

            for (int i = 0; i < numberFormats.Count; i++)
            {
                var fmtElem = new XElement(Namespaces.workbook + "numFmt",
                                           new XAttribute("numFmtId", i + CustomFormatIndex),
                                           new XAttribute("formatCode", numberFormats[i]));
                numFmtsElem.Add(fmtElem);
            }
            doc.Root.Add(numFmtsElem);
        }

        private static void StyleAddBordersElement(XDocument doc, List<CellBorder> uniqueBorders)
        {
            var borders = new XElement(Namespaces.workbook + "borders", new XAttribute("count", uniqueBorders.Count));
            foreach (var border in uniqueBorders)
            {
                var bex = new XElement(Namespaces.workbook + "border");

                var left = new XElement(Namespaces.workbook + "left");
                if (border.HasFlag(CellBorder.Left))
                {
                    left.Add(new XAttribute("style", "thin"));
                    left.Add(new XElement(Namespaces.workbook + "color", new XAttribute("auto", 1)));
                }
                bex.Add(left);

                var right = new XElement(Namespaces.workbook + "right");
                if (border.HasFlag(CellBorder.Right))
                {
                    right.Add(new XAttribute("style", "thin"));
                    right.Add(new XElement(Namespaces.workbook + "color", new XAttribute("auto", 1)));
                }
                bex.Add(right);

                var top = new XElement(Namespaces.workbook + "top");
                if (border.HasFlag(CellBorder.Top))
                {
                    top.Add(new XAttribute("style", "thin"));
                    top.Add(new XElement(Namespaces.workbook + "color", new XAttribute("auto", 1)));
                }
                bex.Add(top);

                var bottom = new XElement(Namespaces.workbook + "bottom");
                if (border.HasFlag(CellBorder.Bottom))
                {
                    bottom.Add(new XAttribute("style", "thin"));
                    bottom.Add(new XElement(Namespaces.workbook + "color", new XAttribute("auto", 1)));
                }
                bex.Add(bottom);

                bex.Add(new XElement(Namespaces.workbook + "diagonal"));

                borders.Add(bex);
            }
            doc.Root.Add(borders);
        }

        private static void StyleAddFillsElement(XDocument doc)
        {
            var fills = new XElement(Namespaces.workbook + "fills", new XAttribute("count", 2));
            fills.Add(new XElement(Namespaces.workbook + "fill",
                                   new XElement(Namespaces.workbook + "patternFill", new XAttribute("patternType", "none"))
                              ));
            fills.Add(new XElement(Namespaces.workbook + "fill",
                                   new XElement(Namespaces.workbook + "patternFill", new XAttribute("patternType", "gray125"))
                              ));
            doc.Root.Add(fills);
        }

        private static void StyleAddFontsElement(XDocument doc, IList<XlsxFont> fontInfos)
        {
            var fonts = new XElement(Namespaces.workbook + "fonts", new XAttribute("count", fontInfos.Count));

            foreach (var font in fontInfos)
            {
                var color = string.Format("{0:X2}{1:X2}{2:X2}{3:X2}", font.TextColor.A, font.TextColor.R, font.TextColor.G, font.TextColor.B);

                var elem = new XElement(Namespaces.workbook + "font",
                                       new XElement(Namespaces.workbook + "sz", new XAttribute("val", font.Size)),
                                       new XElement(Namespaces.workbook + "name", new XAttribute("val", font.Name)),
                                       new XElement(Namespaces.workbook + "color", new XAttribute("rgb", color))
                                  );

                if (font.Bold) { elem.Add(new XElement(Namespaces.workbook + "b")); }
                if (font.Italic) { elem.Add(new XElement(Namespaces.workbook + "i")); }
                if (font.Underline) { elem.Add(new XElement(Namespaces.workbook + "u")); }

                fonts.Add(elem);
            }
            doc.Root.Add(fonts);
        }
    }
}
