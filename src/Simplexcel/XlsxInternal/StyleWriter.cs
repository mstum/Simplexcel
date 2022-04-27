using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Simplexcel.XlsxInternal
{
    internal static class StyleWriter
    {
        // There are up to 164 built in number formats (0-163), all else are custom (164 and above)
        private const int CustomFormatIndex = 164;

        /// <summary>
        /// Standard format codes as defined in ECMA-376, 3rd Edition, Part 1, 18.8.30 numFmt (Number Format)
        /// </summary>
        private static Dictionary<string, int> StandardFormatIds = new Dictionary<string, int>
        {
            ["General"] = 0,
            ["0"] = 1,
            ["0.00"] = 2,
            ["#,##0"] = 3,
            ["#,##0.00"] = 4,
            ["0%"] = 9,
            ["0.00%"] = 10,
            ["0.00E+00"] = 11,
            ["# ?/?"] = 12,
            ["# ??/??"] = 13,
            ["mm-dd-yy"] = 14,
            ["d-mmm-yy"] = 15,
            ["d-mmm"] = 16,
            ["mmm-yy"] = 17,
            ["h:mm AM/PM"] = 18,
            ["h:mm:ss AM/PM"] = 19,
            ["h:mm"] = 20,
            ["h:mm:ss"] = 21,
            ["m/d/yy h:mm"] = 22,
            ["#,##0 ;(#,##0)"] = 37,
            ["#,##0 ;[Red](#,##0)"] = 38,
            ["#,##0.00;(#,##0.00)"] = 39,
            ["#,##0.00;[Red](#,##0.00)"] = 40,
            ["mm:ss"] = 45,
            ["[h]:mm:ss"] = 46,
            ["mmss.0"] = 47,
            ["##0.0E+0"] = 48,
            ["@"] = 49,
        };

        /// <summary>
        /// Create a styles.xml file
        /// </summary>
        /// <param name="styles"></param>
        /// <returns></returns>
        internal static XmlFile CreateStyleXml(IList<XlsxCellStyle> styles)
        {
            var numberFormats = new List<string>();
            var uniqueBorders = new List<CellBorder> { CellBorder.None };
            var uniqueFonts = new List<XlsxFont> { new XlsxFont() };
            // These two fills MUST exist as fill 0 (None) and 1 (Gray125)
            var uniqueFills = new List<PatternFill> { new PatternFill { PatternType = PatternType.None }, new PatternFill { PatternType = PatternType.Gray125 } };

            foreach (var style in styles)
            {
                if (style.Border != CellBorder.None && !uniqueBorders.Contains(style.Border))
                {
                    uniqueBorders.Add(style.Border);
                }

                if (!numberFormats.Contains(style.Format) && !StandardFormatIds.ContainsKey(style.Format))
                {
                    numberFormats.Add(style.Format);
                }

                if (style.Font != null && !uniqueFonts.Contains(style.Font))
                {
                    uniqueFonts.Add(style.Font);
                }

                if (style.Fill != null && !uniqueFills.Contains(style.Fill))
                {
                    uniqueFills.Add(style.Fill);
                }
            }

            var file = new XmlFile
            {
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                Path = "xl/styles.xml"
            };

            var doc = new XDocument(new XElement(Namespaces.workbook + "styleSheet", new XAttribute("xmlns", Namespaces.workbook)));

            StyleAddNumFmtsElement(doc, numberFormats);
            StyleAddFontsElement(doc, uniqueFonts);
            StyleAddFillsElement(doc, uniqueFills);
            StyleAddBordersElement(doc, uniqueBorders);
            StyleAddCellStyleXfsElement(doc);
            StyleAddCellXfsElement(doc, styles, uniqueBorders, numberFormats, uniqueFonts, uniqueFills);

            file.Content = doc;

            return file;
        }

        private static void StyleAddCellXfsElement(XDocument doc, IList<XlsxCellStyle> styles, IList<CellBorder> uniqueBorders, IList<string> numberFormats, IList<XlsxFont> fontInfos, IList<PatternFill> fills)
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
                var numFmtId = StandardFormatIds.TryGetValue(style.Format, out var standardNumFmtId) ? standardNumFmtId : numberFormats.IndexOf(style.Format) + CustomFormatIndex;
                var fontId = fontInfos.IndexOf(style.Font);
                var fillId = fills.IndexOf(style.Fill);
                var borderId = uniqueBorders.IndexOf(style.Border);
                var xfId = 0;

                var elem = new XElement(Namespaces.workbook + "xf",
                                        new XAttribute("numFmtId", numFmtId),
                                        new XAttribute("fontId", fontId),
                                        new XAttribute("fillId", fillId),
                                        new XAttribute("borderId", borderId),
                                        new XAttribute("xfId", xfId));

                if (numFmtId > 0)
                {
                    elem.Add(new XAttribute("applyNumberFormat", 1));
                }
                if (fillId > 0)
                {
                    elem.Add(new XAttribute("applyFill", 1));
                }
                if (fontId > 0)
                {
                    elem.Add(new XAttribute("applyFont", 1));
                }
                if (borderId > 0)
                {
                    elem.Add(new XAttribute("applyBorder", 1));
                }

                AddAlignmentElement(elem, style);

                cellXfs.Add(elem);
            }
            doc.Root.Add(cellXfs);
        }

        private static void AddAlignmentElement(XElement elem, XlsxCellStyle style)
        {
            if(elem == null)
            {
                throw new ArgumentNullException(nameof(elem));
            }

            if (style == null)
            {
                throw new ArgumentNullException(nameof(style));
            }

            if (style.HorizontalAlignment == HorizontalAlign.None && style.VerticalAlignment == VerticalAlign.None)
            {
                return;
            }

            var alignElem = new XElement(Namespaces.workbook + "alignment");
            if(style.HorizontalAlignment != HorizontalAlign.None)
            {
                var value = style.HorizontalAlignment switch
                {
                    HorizontalAlign.General => "general",
                    HorizontalAlign.Left => "left",
                    HorizontalAlign.Center => "center",
                    HorizontalAlign.Right => "right",
                    HorizontalAlign.Justify => "justify",
                    _ => throw new InvalidOperationException("Unhandled HorizontalAlignment in Cell Style: " + style.HorizontalAlignment),
                };
                alignElem.Add(new XAttribute("horizontal", value));
            }

            if (style.VerticalAlignment != VerticalAlign.None)
            {
                var value = style.VerticalAlignment switch
                {
                    VerticalAlign.Top => "top",
                    VerticalAlign.Middle => "center",
                    VerticalAlign.Bottom => "bottom",
                    VerticalAlign.Justify => "justify",
                    _ => throw new InvalidOperationException("Unhandled VerticalAlignment in Cell Style: " + style.VerticalAlignment),
                };
                alignElem.Add(new XAttribute("vertical", value));
            }

            elem.Add(alignElem);
        }

        private static void StyleAddCellStyleXfsElement(XDocument doc)
        {
            var cellStyleXfs = new XElement(Namespaces.workbook + "cellStyleXfs", new XAttribute("count", 1));
            cellStyleXfs.Add(new XElement(Namespaces.workbook + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0)));
            doc.Root.Add(cellStyleXfs);
        }

        private static void StyleAddNumFmtsElement(XDocument doc, List<string> numberFormats)
        {
            var numFmtsElem = new XElement(Namespaces.workbook + "numFmts", new XAttribute("count", numberFormats.Count));

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

        private static void StyleAddFillsElement(XDocument doc, List<PatternFill> uniqueFills)
        {
            var fills = new XElement(Namespaces.workbook + "fills", new XAttribute("count", uniqueFills.Count));

            foreach (var fill in uniqueFills)
            {
                var fe = new XElement(Namespaces.workbook + "fill");
                if (fill is PatternFill pf)
                {
                    var pfe = new XElement(Namespaces.workbook + "patternFill");
                    string patternType = "none";
                    switch (pf.PatternType)
                    {
                        case PatternType.Solid: patternType = "solid"; break;
                        case PatternType.Gray750: patternType = "darkGray"; break;
                        case PatternType.Gray500: patternType = "mediumGray"; break;
                        case PatternType.Gray250: patternType = "lightGray"; break;
                        case PatternType.Gray125: patternType = "gray125"; break;
                        case PatternType.Gray0625: patternType = "gray0625"; break;
                        case PatternType.HorizontalStripe: patternType = "darkHorizontal"; break;
                        case PatternType.VerticalStripe: patternType = "darkVertical"; break;
                        case PatternType.ReverseDiagonalStripe: patternType = "darkDown"; break;
                        case PatternType.DiagonalStripe: patternType = "darkUp"; break;
                        case PatternType.DiagonalCrosshatch: patternType = "darkGrid"; break;
                        case PatternType.ThickDiagonalCrosshatch: patternType = "darkTrellis"; break;
                        case PatternType.ThinHorizontalStripe: patternType = "lightHorizontal"; break;
                        case PatternType.ThinVerticalStripe: patternType = "lightVertical"; break;
                        case PatternType.ThinReverseDiagonalStripe: patternType = "lightDown"; break;
                        case PatternType.ThinDiagonalStripe: patternType = "lightUp"; break;
                        case PatternType.ThinHorizontalCrosshatch: patternType = "lightGrid"; break;
                        case PatternType.ThinDiagonalCrosshatch: patternType = "lightTrellis"; break;
                    }
                    pfe.Add(new XAttribute("patternType", patternType));

                    if (pf.BackgroundColor.HasValue)
                    {
                        // Not a typo. Excel calls the fgColor "Background Color" in the UI, and the bgColor "Pattern Color"
                        pfe.Add(new XElement(Namespaces.workbook + "fgColor", new XAttribute("rgb", pf.BackgroundColor.Value)));
                    }
                    if (pf.PatternColor.HasValue)
                    {
                        pfe.Add(new XElement(Namespaces.workbook + "bgColor", new XAttribute("rgb", pf.PatternColor.Value)));
                    }
                    else if (pf.BackgroundColor.HasValue)
                    {
                        // fgColor without bgColor causes Excel to show an error
                        pfe.Add(new XElement(Namespaces.workbook + "bgColor", new XAttribute("auto", 1)));
                    }

                    fe.Add(pfe);
                }
                //else if (fill is GradientFill gf)
                //{
                    // TODO: Not Yet Implemented
                //}
                fills.Add(fe);
            }
            doc.Root.Add(fills);
        }

        private static void StyleAddFontsElement(XDocument doc, IList<XlsxFont> fontInfos)
        {
            var fonts = new XElement(Namespaces.workbook + "fonts", new XAttribute("count", fontInfos.Count));

            foreach (var font in fontInfos)
            {
                var elem = new XElement(Namespaces.workbook + "font",
                                       new XElement(Namespaces.workbook + "sz", new XAttribute("val", font.Size)),
                                       new XElement(Namespaces.workbook + "name", new XAttribute("val", font.Name)),
                                       new XElement(Namespaces.workbook + "color", new XAttribute("rgb", font.TextColor))
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
