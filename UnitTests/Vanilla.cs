using FsExcel;
using CsExcel;
using static CsExcel.PositionFactory;
using static CsExcel.ItemFactory;
using static CsExcel.CellPropFactory;
using static CsExcel.BorderFactory;
using static CsExcel.FontEmphasisFactory;
using static CsExcel.HorizontalAlignmentFactory;
using static CsExcel.SizeFactory;
using System.Globalization;
using static ClosedXML.Excel.XLBorderStyleValues;
using static ClosedXML.Excel.XLFontUnderlineValues;
using ClosedXML.Excel;
using static CsExcel.CellLabelFactory;
using static CsExcel.StyleMergedCellFactory;
using static CsExcel.FreezePanesFactory;
using System.Runtime.InteropServices;

namespace UnitTests
{
    public class Vanilla
    {
        [Fact]
        public void HelloWorld()
        {
            var cells = new[]
                {
                    Cell([ String("Hello World") ])
                };

            CsExcel.Render.AsFile(cells, TestFiles.PathFor("helloWorld.xlsx"));

            using var wb = TestFiles.Open("helloWorld.xlsx");
            Assert.Equal("Hello World", wb.Worksheet(1).Cell(1, 1).GetString());
        }
        [Fact]
        public void MultipleCells()
        {
            var cells =
                from n in Enumerable.Range(1, 10)
                select Cell([Integer(n)]);
            CsExcel.Render.AsFile(cells, TestFiles.PathFor("MultipleCells.xlsx"));

            using var wb = TestFiles.Open("MultipleCells.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:J1", ws.RangeUsed().RangeAddress.ToString());
            for (var i = 1; i <= 10; i++)
            {
                Assert.Equal(i, ws.Cell(1, i).GetValue<int>());
            }
        }
        [Fact]
        public void VerticalMovement()
        {
            var cells =
                from m in Enumerable.Range(1, 12)
                let monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m)
                select Cell(
                    [
                        String(monthName),
                        Next(DownBy(1))
                    ]
                );
            CsExcel.Render.AsFile(cells, TestFiles.PathFor("VerticalMovement.xlsx"));

            using var wb = TestFiles.Open("VerticalMovement.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:A12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal("May", ws.Cell(5, 1).GetString());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
        }
        [Fact]
        public void VerticalMovement2()
        {
            // at this point using yield return is simpler than using a query expression
            IEnumerable<Item> Cells()
            {
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length), Next(PositionFactory.NewRow)]);
                }
            };
            CsExcel.Render.AsFile(Cells(), TestFiles.PathFor("VerticalMovement2.xlsx"));

            using var wb = TestFiles.Open("VerticalMovement2.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal(7, ws.Cell(1, 2).GetValue<int>());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
            Assert.Equal(8, ws.Cell(12, 2).GetValue<int>());
        }
        [Fact]
        public void VerticalMovement3()
        {
            IEnumerable<Item> Cells()
            {
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }
            };
            CsExcel.Render.AsFile(Cells(), TestFiles.PathFor("VerticalMovement3.xlsx"));

            using var wb = TestFiles.Open("VerticalMovement3.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal(7, ws.Cell(1, 2).GetValue<int>());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
            Assert.Equal(8, ws.Cell(12, 2).GetValue<int>());
        }
        [Fact]
        public void Indentation()
        {
            IEnumerable<Item> Cells()
            {
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Go(Indent(2));
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }
            };
            CsExcel.Render.AsFile(Cells(), TestFiles.PathFor("Indentation.xlsx"));

            using var wb = TestFiles.Open("Indentation.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("B1:C12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 2).GetString());
            Assert.Equal("December", ws.Cell(12, 2).GetString());
        }
        [Fact]
        public void BorderAndFontStyling()
        {
            IEnumerable<Item> Items()
            {
                foreach (var heading in new[] { "Month", "Letter Count" })
                {
                    yield return Cell(
                        [
                            String(heading),
                            Border(Bottom(XLBorderStyleValues.Medium)),
                            FontEmphasis(Bold),
                            FontEmphasis(Italic)
                        ]);
                }
                yield return Go(PositionFactory.NewRow);
                IEnumerable<CellProp> CellProps(string monthName)
                {
                    yield return String(monthName);
                    yield return FontEmphasis(Underline(DoubleAccounting));
                    if (monthName == "May")
                    {
                        yield return FontEmphasis(StrikeThrough);
                    }
                }
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(CellProps(monthName));
                    yield return Cell([
                            Integer(monthName.Length),
                        ]
                    );
                    yield return Go(PositionFactory.NewRow);
                }
            };

            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("BorderAndFontStyling.xlsx"));

            using var wb = TestFiles.Open("BorderAndFontStyling.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B13", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.Cell(1, 1).Style.Font.Bold);
            Assert.True(ws.Cell(1, 1).Style.Font.Italic);
            Assert.Equal("May", ws.Cell(6, 1).GetString());
            Assert.True(ws.Cell(6, 1).Style.Font.Strikethrough);
            Assert.False(ws.Cell(2, 1).Style.Font.Strikethrough);
            Assert.Equal(XLFontUnderlineValues.DoubleAccounting, ws.Cell(2, 1).Style.Font.Underline);
        }
        [Fact]
        public void BorderAndFontStyling2()
        {
            IEnumerable<Item> Items()
            {
                CellProp[] headingStyle = [
                    Border(Bottom(XLBorderStyleValues.Medium)),
                    FontEmphasis(Bold),
                    FontEmphasis(Italic)
                ];

                foreach (var heading in new[] { "Month", "Letter Count" })
                {
                    yield return Cell(
                        [
                            String(heading),
                            .. headingStyle
                        ]);
                }
                yield return Go(PositionFactory.NewRow);
                IEnumerable<CellProp> CellProps(string monthName)
                {
                    yield return String(monthName);
                    yield return FontEmphasis(Underline(DoubleAccounting));
                    if (monthName == "May")
                    {
                        yield return FontEmphasis(StrikeThrough);
                    }
                }
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(CellProps(monthName));
                    yield return Cell([
                            Integer(monthName.Length),
                        ]
                    );
                    yield return Go(PositionFactory.NewRow);
                }
            };

            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("BorderAndFontStyling2.xlsx"));

            using var wb = TestFiles.Open("BorderAndFontStyling2.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B13", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.Cell(1, 1).Style.Font.Bold);
            Assert.True(ws.Cell(1, 1).Style.Font.Italic);
            Assert.Equal("May", ws.Cell(6, 1).GetString());
            Assert.True(ws.Cell(6, 1).Style.Font.Strikethrough);
        }

        [Fact]
        public void FontAndNameSize()
        {
            // Materialized once (rather than left as a lazy query) so the render pass below and the
            // assertions afterward are guaranteed to see the exact same 20 fonts in the same order.
            var fontNames =
                SixLabors.Fonts.SystemFonts.Collection.Families.Select((fontFamily, i) => (fontFamily.Name, i)).OrderBy(f => f.Item1).Take(20).ToList();
            IEnumerable<Item> Items()
            {
                foreach (var (fontName, i) in fontNames)
                {
                    yield return
                            Cell([
                                String(fontName),
                                FontName(fontName),
                                FontSize(10 + (i * 2))]);

                }
                Go(PositionFactory.NewRow);
            };
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("FontAndNameSize.xlsx"));

            using var wb = TestFiles.Open("FontAndNameSize.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(20, fontNames.Count);
            for (var col = 0; col < fontNames.Count; col++)
            {
                var (name, i) = fontNames[col];
                var cell = ws.Cell(1, col + 1);
                Assert.Equal(name, cell.GetString());
                Assert.Equal(name, cell.Style.Font.FontName);
                Assert.Equal(10 + (i * 2), cell.Style.Font.FontSize);
            }
        }

        [Fact]
        public void WrapText()
        {
            Item[] items =
            [
                Cell([
                    String("Without wrap text:"),
                    HorizontalAlignment(Center),
                    VerticalAlignment(VerticalAlignmentFactory.Middle),
                    CellSize(ColWidth(16)),
                ]),
                Cell([
                    String("The quick brown fox jumps over the lazy dog."),
                    HorizontalAlignment(Center),
                    VerticalAlignment(VerticalAlignmentFactory.Middle),
                ]),
                Go(PositionFactory.NewRow),
                Cell([
                    String("Without wrap text:"),
                    HorizontalAlignment(Center),
                    VerticalAlignment(VerticalAlignmentFactory.Middle),
                    CellSize(ColWidth(16)),
                ]),
                Cell([
                    String("The quick brown fox jumps over the lazy dog."),
                    HorizontalAlignment(Center),
                    VerticalAlignment(VerticalAlignmentFactory.Middle),
                    CellPropFactory.WrapText(true),
                ]),
            ];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("WrapText.xlsx"));

            using var wb = TestFiles.Open("WrapText.xlsx");
            var ws = wb.Worksheet(1);
            Assert.False(ws.Cell(1, 2).Style.Alignment.WrapText);
            Assert.True(ws.Cell(2, 2).Style.Alignment.WrapText);
        }
        [Fact]
        public void TextRotation()
        {
            var (p, m, g) = ("⏺", "◑", "⭘");
            string[][] performances =
                [
                    [ p, m, g, g, p, p, g, p, p, g ],
                    [ g, m, g, m, g, p, g, p, p, g ],
                    [ g, m, m, g, g, p, g, g, p, g ],
                    [ m, m, m, p, p, p, g, m, p, g ],
                    [ p, p, p, p, g, g, m, m, p, g ],
                    [ p, g, p, g, g, g, p, g, m, m ],
                    [ g, p, g, p, m, p, m, p, p, g ],
                    [ p, p, m, g, p, p, p, m, p, m ],
                ];
            string GetPerformance(int categoryIndex, int supplierIndex) =>
                performances[supplierIndex - 1][categoryIndex - 1];

            IEnumerable<Item> Items()
            {
                yield return Go(RC(1, 2));
                foreach (var category in Enumerable.Range(1, 10))
                {
                    yield return Cell([
                        String($"Category {category}"),
                        CellPropFactory.TextRotation(45),
                        CellSize(RowHeight(45))]);
                }
                yield return Go(PositionFactory.NewRow);
                foreach (var supplier in Enumerable.Range(1, 8))
                {
                    yield return Cell([
                        String($"Supplier {supplier}"),
                        CellSize(ColWidth(10))]);
                    yield return Go(PositionFactory.NewRow);
                }
                yield return Go(RC(2, 2));
                yield return Go(Indent(2));
                foreach (var supplier in Enumerable.Range(1, 8))
                {
                    foreach (var category in Enumerable.Range(1, 10))
                    {
                        yield return Cell([String(GetPerformance(category, supplier)), HorizontalAlignment(Center)]);
                    }
                    yield return Go(PositionFactory.NewRow);
                }
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("TextRotation.xlsx"));

            using var wb = TestFiles.Open("TextRotation.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("Category 1", ws.Cell(1, 2).GetString());
            Assert.Equal(45, ws.Cell(1, 2).Style.Alignment.TextRotation);
            Assert.Equal("Supplier 1", ws.Cell(2, 1).GetString());
            Assert.Equal(GetPerformance(1, 1), ws.Cell(2, 2).GetString());
            Assert.Equal(GetPerformance(10, 8), ws.Cell(9, 11).GetString());
        }
        static class RandomGenerator
        {
            static uint state = 1; // Mutable state variable

            static ulong Mangle(ulong n)
            {
                return (n & 0x7FFFFFFF) + (n >> 31);
            }

            public static double NextDouble()
            {
                state = (uint)(Mangle(Mangle(state * 48271UL)));
                return (double)state / int.MaxValue;
            }
        }

        [Fact]
        public void NumberFormattingAndAlignment()
        {
            CellProp[] headingStyle = [
                Border(Bottom(XLBorderStyleValues.Medium)),
                FontEmphasis(Bold),
                FontEmphasis(Italic)
            ];
            IEnumerable<Item> Items()
            {
                foreach (var (heading, alignment) in new (string, FsExcel.HorizontalAlignment)[] {
                    ("Stock Item", HorizontalAlignmentFactory.Left),
                    ("Price", HorizontalAlignmentFactory.Right),
                    ("Count", HorizontalAlignmentFactory.Right) })
                {
                    yield return Cell([
                        String(heading),
                        .. headingStyle,
                        HorizontalAlignment(alignment)
                    ]);
                }
                yield return Go(PositionFactory.NewRow);
                foreach (var item in new[] { "Apples", "Oranges", "Pears" })
                {
                    yield return Cell([String(item)]);
                    yield return Cell([Float(RandomGenerator.NextDouble() * 1000.0), FormatCode("$0.00")]);
                    yield return Cell([Integer((int)(RandomGenerator.NextDouble() * 100.0)), FormatCode("#,##0")]);
                    yield return Go(PositionFactory.NewRow);
                }
            };
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("NumberFormattingAndAlignment.xlsx"));

            // Price/Count come from RandomGenerator, whose state is shared (and thus order-dependent)
            // across every test that uses it - assert structure/format/range rather than exact values.
            using var wb = TestFiles.Open("NumberFormattingAndAlignment.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:C4", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("Stock Item", ws.Cell(1, 1).GetString());
            Assert.Equal("Price", ws.Cell(1, 2).GetString());
            Assert.Equal("Count", ws.Cell(1, 3).GetString());
            var items = new[] { "Apples", "Oranges", "Pears" };
            for (var i = 0; i < items.Length; i++)
            {
                var row = i + 2;
                Assert.Equal(items[i], ws.Cell(row, 1).GetString());
                Assert.Equal("$0.00", ws.Cell(row, 2).Style.NumberFormat.Format);
                Assert.InRange(ws.Cell(row, 2).GetValue<double>(), 0.0, 1000.0);
                Assert.Equal("#,##0", ws.Cell(row, 3).Style.NumberFormat.Format);
                Assert.InRange(ws.Cell(row, 3).GetValue<int>(), 0, 99);
            }
        }
        [Fact]
        public void Formulae()
        {
            CellProp[] headingStyle = [
                Border(Bottom(XLBorderStyleValues.Medium)),
                FontEmphasis(Bold),
                FontEmphasis(Italic)
            ];
            IEnumerable<Item> Items()
            {
                foreach (var (heading, alignment) in new (string, FsExcel.HorizontalAlignment)[] {
                    ("Stock Item", HorizontalAlignmentFactory.Left),
                    ("Price", HorizontalAlignmentFactory.Right),
                    ("Count", HorizontalAlignmentFactory.Right),
                    ("Total", HorizontalAlignmentFactory.Right) })
                {
                    yield return Cell([
                        String(heading),
                        .. headingStyle,
                        HorizontalAlignment(alignment)
                    ]);
                }
                yield return Go(PositionFactory.NewRow);
                foreach (var (index, item) in new[] { "Apples", "Oranges", "Pears" }.Select((item, index) => (index, item)))
                {
                    yield return Cell([String(item)]);
                    yield return Cell([Float(RandomGenerator.NextDouble() * 1000.0), FormatCode("$0.00")]);
                    yield return Cell([Integer((int)(RandomGenerator.NextDouble() * 100.0)), FormatCode("#,##0")]);
                    yield return Cell([FormulaA1($"=B{index + 2}*C{index + 2}"), FormatCode("$#,##0.00")]);
                    yield return Go(PositionFactory.NewRow);
                }
            };
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("Formulae.xlsx"));

            using var wb = TestFiles.Open("Formulae.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:D4", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("Total", ws.Cell(1, 4).GetString());
            var items = new[] { "Apples", "Oranges", "Pears" };
            for (var i = 0; i < items.Length; i++)
            {
                var row = i + 2;
                Assert.Equal(items[i], ws.Cell(row, 1).GetString());
                Assert.Equal($"B{row}*C{row}", ws.Cell(row, 4).FormulaA1);
                Assert.Equal("$#,##0.00", ws.Cell(row, 4).Style.NumberFormat.Format);
                var expectedTotal = ws.Cell(row, 2).GetValue<double>() * ws.Cell(row, 3).GetValue<int>();
                Assert.Equal(expectedTotal, ws.Cell(row, 4).GetValue<double>(), precision: 6);
            }
        }
        [Fact]
        public void Color()
        {
            IEnumerable<Item> Items()
            {
                IEnumerable<int> values =
                    [
                        .. (Enumerable.Range(0, 8).Select(x => x * 32)),
                        255
                    ];
                foreach (var r in values)
                {
                    foreach (var g in values)
                    {
                        foreach (var b in values)
                        {
                            // N.B. the API refuses to fill a cell with black if its font is black
                            // so the very first cell won't be colored.
                            var backgroundColor = ClosedXML.Excel.XLColor.FromArgb(0, r, g, b);
                            var fontColor = ClosedXML.Excel.XLColor.FromArgb(0, b, r, g);
                            var borderColor = ClosedXML.Excel.XLColor.FromArgb(0, g, b, r);
                            yield return Cell([
                                String($"R={r};G={g};B={b}"),
                                FontColor(fontColor),
                                BackgroundColor(backgroundColor),
                                Border(Top(XLBorderStyleValues.Thick)),
                                Border(Right(XLBorderStyleValues.Thick)),
                                Border(Bottom(XLBorderStyleValues.Thick)),
                                Border(Left(XLBorderStyleValues.Thick)),
                                BorderColor(BorderColorFactory.Top(borderColor)),
                                BorderColor(BorderColorFactory.Right(borderColor)),
                                BorderColor(BorderColorFactory.Bottom(borderColor)),
                                BorderColor(BorderColorFactory.Left(borderColor))
                            ]);
                        }
                        yield return Go(PositionFactory.NewRow);
                    }
                    yield return Go(PositionFactory.NewRow);
                }
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("Color.xlsx"));

            using var wb = TestFiles.Open("Color.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:I89", ws.RangeUsed().RangeAddress.ToString());
            var a1 = ws.Cell(1, 1);
            Assert.Equal("R=0;G=0;B=0", a1.GetString());
            // the very first cell is left uncolored, since the API refuses to fill a cell with
            // black if its font is also black - so it keeps ClosedXML's default (indexed, unset) fill.
            Assert.Equal(XLColor.FromArgb(0, 0, 0, 0), a1.Style.Font.FontColor);
            Assert.Equal(XLColorType.Indexed, a1.Style.Fill.BackgroundColor.ColorType);
            var b1 = ws.Cell(1, 2);
            Assert.Equal("R=0;G=0;B=32", b1.GetString());
            Assert.Equal(XLColor.FromArgb(0, 32, 0, 0), b1.Style.Font.FontColor);
            Assert.Equal(XLColor.FromArgb(0, 0, 0, 32), b1.Style.Fill.BackgroundColor);
            Assert.Equal(XLBorderStyleValues.Thick, b1.Style.Border.TopBorder);
        }
        [Fact]
        public void RangeStyles()
        {
            IEnumerable<Item> Items()
            {
                yield return Style([
                        Border(Bottom(XLBorderStyleValues.Medium)),
                        FontEmphasis(Bold),
                        FontEmphasis(Italic)
                    ]);
                foreach (var heading in new[] { "Stock Item", "Price", "Count" })
                {
                    yield return Cell([String(heading)]);
                }
                yield return Go(PositionFactory.NewRow);
                foreach (var item in new[] { "Apples", "Oranges", "Pears" })
                {
                    yield return Cell([String(item)]);
                    yield return Style([CellPropFactory.FontEmphasis(Italic)]);
                    yield return Cell([Float(RandomGenerator.NextDouble() * 1000), FormatCode("$0.00")]);
                    yield return Cell([Integer((int)(RandomGenerator.NextDouble() * 100)), FormatCode("#,##0")]);
                    yield return Style([]);
                    yield return Go(PositionFactory.NewRow);
                }
            };
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("RangeStyles.xlsx"));

            using var wb = TestFiles.Open("RangeStyles.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:C4", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.Cell(1, 1).Style.Font.Bold);
            // Style sets an ambient "current style" applied going forward, not retroactively over a
            // range - so "Apples" (the first cell after the headings) still carries the heading's
            // bold+italic, since it's placed before the loop resets the style to italic-only.
            Assert.True(ws.Cell(2, 1).Style.Font.Bold);
            Assert.True(ws.Cell(2, 1).Style.Font.Italic);
            Assert.False(ws.Cell(2, 2).Style.Font.Bold);
            Assert.True(ws.Cell(2, 2).Style.Font.Italic);
            Assert.False(ws.Cell(3, 1).Style.Font.Bold);
            Assert.False(ws.Cell(3, 1).Style.Font.Italic);
        }
        [Fact]
        public void AddingABorderToMergedCells()
        {
            IEnumerable<Item> Items()
            {
                yield return Go(PositionFactory.NewRow);
                foreach (var (heading, colWidth) in new[] { ("ID", 3.22), ("Car Name", 10.33), ("Car Description", 49.33), ("Car Registration", 16.89) })
                {
                    yield return Cell([
                        String(heading),
                        FontEmphasis(Bold),
                        FontName("Calibri"),
                        FontSize(11),
                        HorizontalAlignment(Center),
                        FontColor(XLColor.FromArgb(0, 255, 255, 255)),
                        BackgroundColor(XLColor.FromArgb(0, 68, 114, 196)),
                        Border(All(Thin)),
                        CellSize(ColWidth(colWidth))
                    ]);
                }
                yield return Go(PositionFactory.NewRow);
                yield return Style([
                    HorizontalAlignment(Center),
                    VerticalAlignment(VerticalAlignmentFactory.Middle),
                    BackgroundColor(XLColor.FromArgb(0, 240, 240, 210))
                ]);
                yield return Cell([Integer(1), Name("ID")]);
                yield return Cell([String("Ford Fiesta")]);
                yield return Cell([String("Car Technical Details:"), Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 1"), Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 2"), Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 3"), Name("LastL")]);
                yield return Go(RC(3, 4));
                yield return Cell([String("AB12 CDE"), Name("Reg")]);
                yield return Go(RC(6, 4));
                yield return Cell([Name("RegEnd")]);
                yield return Go(RC(7, 3));
                yield return Cell([String("Another Technical Detail"), FontEmphasis(Italic), Name("TD"), Next(FsExcel.Position.Stay)]);
                yield return Go(DownBy(1));
                yield return Cell([Name("info")]);
                yield return MergeCells(ColRowLabel("B", 3), ColRowLabel("B", 6));
                yield return MergeCells(NamedCell("ID"), ColRowLabel("A", 6));
                yield return MergeCells(ColRowLabel("C", 7), NamedCell("info"));
                yield return MergeCells(NamedCell("Reg"), NamedCell("RegEnd"));
                yield return BorderMergedCell([
                    BorderType(All(Thin)),
                    ColorBorder(BorderColorFactory.All(XLColor.FromArgb(0, 68, 114, 196)))
                ]);
            };
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("AddingABorderToMergedCells.xlsx"));

            using var wb = TestFiles.Open("AddingABorderToMergedCells.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A2:D7", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("ID", ws.Cell(2, 1).GetString());
            Assert.Equal(1, ws.Cell(3, 1).GetValue<int>());
            Assert.Equal("Ford Fiesta", ws.Cell(3, 2).GetString());
            Assert.Equal("AB12 CDE", ws.Cell(3, 4).GetString());
            Assert.Equal("Another Technical Detail", ws.Cell(7, 3).GetString());
            var mergedRanges = ws.MergedRanges.Select(m => m.RangeAddress.ToString()).ToHashSet();
            Assert.Contains("A3:A6", mergedRanges);
            Assert.Contains("B3:B6", mergedRanges);
            Assert.Contains("D3:D6", mergedRanges);
            Assert.Contains("C7:C8", mergedRanges);
            Assert.Equal(XLBorderStyleValues.Thin, ws.Cell(3, 1).Style.Border.TopBorder);
        }
        [Fact]
        public void AbsolutePositioning()
        {
            var items =
                new[] {
                    Go(Col(3)),
                    Cell([String("Col 3")]),
                    Go(Row(4)),
                    Cell([String("Row 4")]),
                    Go(RC(6, 5)),
                    Cell([String("R6C5")]),
                    Cell([String("R6C6")])
                };
            CsExcel.Render.AsFile(items, TestFiles.PathFor("AbsolutePositioning.xlsx"));

            using var wb = TestFiles.Open("AbsolutePositioning.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("C1:F6", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("Col 3", ws.Cell(1, 3).GetString());
            Assert.Equal("Row 4", ws.Cell(4, 4).GetString());
            Assert.Equal("R6C5", ws.Cell(6, 5).GetString());
            Assert.Equal("R6C6", ws.Cell(6, 6).GetString());
        }
        [Fact]
        public void AbsolutePositionin2()
        {
            IEnumerable<Item> Items()
            {
                foreach (var i in Enumerable.Range(1, 5))
                {
                    yield return Cell([Integer(i), Next(PositionFactory.Stay)]);
                    yield return Go(DownBy(i));
                }
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("AbsolutePositioning2.xlsx"));

            using var wb = TestFiles.Open("AbsolutePositioning2.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(1, ws.Cell(1, 1).GetValue<int>());
            Assert.Equal(2, ws.Cell(2, 1).GetValue<int>());
            Assert.Equal(3, ws.Cell(4, 1).GetValue<int>());
            Assert.Equal(4, ws.Cell(7, 1).GetValue<int>());
            Assert.Equal(5, ws.Cell(11, 1).GetValue<int>());
        }
        [Fact]
        public void NamedCells()
        {
            var items = new[]
            {
                Cell([
                    String("JohnDoe"),
                    Name("Username"),
                ]),
                Cell([
                    String("john.doe@company.com"),
                    CellPropFactory.ScopedName("Email",NameScope.Workbook),
                ])
            };
            CsExcel.Render.AsFile(items, TestFiles.PathFor("NamedCells.xlsx"));

            using var wb = TestFiles.Open("NamedCells.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("JohnDoe", ws.Cell("A1").GetString());
            // A plain Name(...) is worksheet-scoped; ScopedName(..., NameScope.Workbook) is workbook-scoped.
            Assert.Equal("JohnDoe", ws.NamedRange("Username").Ranges.First().FirstCell().GetString());
            Assert.Equal("john.doe@company.com", wb.NamedRange("Email").Ranges.First().FirstCell().GetString());
        }

        static IEnumerable<Item> MakeWorksheetTabsItems()
        {
            var britishCultureNativeName = "English (United Kingdom)";
            var ukrainianCultureNativeName = "українська";
            var britishCultureDateTimeFormatGetMonthName =
                new[] { "January", "February", "March", "April", "May", "June", "July",
                        "August", "September", "October", "November", "December" };
            var britishCultureDateTimeFormatAbbreviatedMonthNames =
                new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
                        "Nov", "Dec" };
            var ukrainianCultureDateTimeFormatGetMonthName =
                new[] { "січень", "лютий", "березень", "квітень", "травень", "червень",
                        "липень", "серпень", "вересень", "жовтень", "листопад", "грудень" };
            var ukrainianCultureDateTimeFormatAbbreviatedMonthNames =
                new[] { "січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов",
                        "лис", "гру" };
            IEnumerable<Item> Items()
            {
                yield return Worksheet(britishCultureNativeName);
                foreach (var m in Enumerable.Range(0, 11))
                {
                    var monthName = britishCultureDateTimeFormatGetMonthName[m];
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }

                yield return Worksheet(ukrainianCultureNativeName);
                foreach (var m in Enumerable.Range(0, 11))
                {
                    var monthName = ukrainianCultureDateTimeFormatGetMonthName[m];
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }

                yield return Worksheet(britishCultureNativeName); // Switch back to the first worksheet
                yield return Go(RC(13, 1));
                foreach (var m in Enumerable.Range(0, 11))
                {
                    var monthAbbreviation = britishCultureDateTimeFormatAbbreviatedMonthNames[m];
                    yield return Cell([String(monthAbbreviation)]);
                    yield return Cell([Integer(monthAbbreviation.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }

                yield return Worksheet(ukrainianCultureNativeName); // Switch back to the second worksheet
                yield return Go(RC(13, 1));
                foreach (var m in Enumerable.Range(0, 11))
                {
                    var monthAbbreviation = ukrainianCultureDateTimeFormatAbbreviatedMonthNames[m];
                    yield return Cell([String(monthAbbreviation)]);
                    yield return Cell([Integer(monthAbbreviation.Length)]);
                    yield return Go(PositionFactory.NewRow);
                }
            };

            return Items();
        }

        [Fact]
        public void WorksheetsTabs()
        {
            CsExcel.Render.AsFile(MakeWorksheetTabsItems(), TestFiles.PathFor("WorksheetsTabs.xlsx"));

            using var wb = TestFiles.Open("WorksheetsTabs.xlsx");
            Assert.Equal(["English (United Kingdom)", "українська"], wb.Worksheets.Select(w => w.Name));
            var british = wb.Worksheet("English (United Kingdom)");
            Assert.Equal("January", british.Cell(1, 1).GetString());
            Assert.Equal("Nov", british.Cell(23, 1).GetString());
            var ukrainian = wb.Worksheet("українська");
            Assert.Equal("січень", ukrainian.Cell(1, 1).GetString());
        }
        [Fact]
        public void InsertingBlankRows()
        {
            // you can load an existing file
            // var workbook = new XLWorkbook(Path.Combine(savePath, "Worksheets.xlsx"))
            var workbook = CsExcel.Render.AsWorkBook(MakeWorksheetTabsItems());
            var britishCultureNativeName = "English (United Kingdom)";
            var ukrainianCultureNativeName = "українська";
            var altMonthNames = new[]
            {
                "Vintagearious", "Fogarious", "Frostarious", "Snowous", "Rainous",
                "Windous", "Buddal", "Floweral", "Meadowal", "Reapidor", "Heatidor", "Fruitidor"
            };
            IEnumerable<Item> Items()
            {
                yield return Workbook(workbook);
                yield return Worksheet(ukrainianCultureNativeName);
                yield return Go(RC(1, 3));
                yield return Cell([FormulaA1($"='{britishCultureNativeName}'!B1*2")]);
                yield return Worksheet(britishCultureNativeName);
                yield return InsertRowsAbove(12); // The cell reference in the formula above will be updated to B13
                for (var m = 0; m < 12; m++)
                {
                    yield return Cell([String(altMonthNames[m])]);
                    yield return Cell([Integer(altMonthNames[m].Length)]);
                    yield return Go(PositionFactory.NewRow);
                }
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("InsertingBlankRows.xlsx"));

            using var wb = TestFiles.Open("InsertingBlankRows.xlsx");
            var ukrainian = wb.Worksheet(ukrainianCultureNativeName);
            // The rows inserted above on the British sheet must shift the formula's row reference
            // from B1 to B13 - this is the whole point of the test.
            Assert.Equal($"'{britishCultureNativeName}'!B13*2", ukrainian.Cell(1, 3).FormulaA1);
            var british = wb.Worksheet(britishCultureNativeName);
            Assert.Equal("Vintagearious", british.Cell(1, 1).GetString());
            Assert.Equal("January", british.Cell(13, 1).GetString());
        }
        [Fact]
        public void ColumnWidthsAndRowHeightsForAllCells()
        {
            IEnumerable<Item> Items()
            {
                for (var x = 1; x <= 12; x++)
                {
                    for (var y = 0; y <= 12; y++)
                    {
                        yield return Cell([Integer(x * y)]);
                    }
                    yield return Go(PositionFactory.NewRow);
                }
                yield return SizeAll(ColWidth(5));
                yield return SizeAll(RowHeight(20));
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("ColumnWidthsAndRowHeightsForAllCells.xlsx"));

            using var wb = TestFiles.Open("ColumnWidthsAndRowHeightsForAllCells.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:M12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal(0, ws.Cell(1, 1).GetValue<int>());
            Assert.Equal(24, ws.Cell(2, 13).GetValue<int>());
            Assert.Equal(5.0, ws.Column(1).Width, precision: 2);
            Assert.Equal(20.0, ws.Row(1).Height, precision: 2);
        }
        [Fact]
        public void IndividualCellSizing()
        {
            IEnumerable<Item> Items()
            {
                yield return Go(PositionFactory.NewRow);
                foreach (var (heading, colWidth) in new (string, double)[]
                    {
                        ("ID", 3.22),
                        ("Car Name", 10.33),
                        ("Car Descriptions", 49.33),
                        ("Car Registration", 16.89),
                    })
                {
                    yield return Cell(
                        [
                            String(heading),
                            FontEmphasis(Bold),
                            FontName("Calibri"),
                            FontSize(11),
                            HorizontalAlignment(Center),
                            FontColor(XLColor.FromArgb(0,255,255,255)),
                            BackgroundColor(XLColor.FromArgb(0,68,114,196)),
                            Border(All(Thin)),
                            CellSize(ColWidth(colWidth)),
                        ]);
                }
                yield return Go(PositionFactory.NewRow);
                yield return Cell([Integer(1), HorizontalAlignment(Center)]);
                yield return Cell([String("Ford Fiesta")]);
                yield return Cell([String("Car Technical Details...")]);
                yield return Cell([String("AB12 CDE"), HorizontalAlignment(Center)]);
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("IndividualCellSizing.xlsx"));

            using var wb = TestFiles.Open("IndividualCellSizing.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A2:D3", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("ID", ws.Cell(2, 1).GetString());
            Assert.Equal(1, ws.Cell(3, 1).GetValue<int>());
            Assert.Equal("Ford Fiesta", ws.Cell(3, 2).GetString());
            Assert.Equal(3.22, ws.Column(1).Width, precision: 2);
            Assert.Equal(49.33, ws.Column(3).Width, precision: 2);
        }
        [Fact]
        public void Autofitting()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Liberation Sans");
            }
            CellProp[] headingStyle = [
                    Border(Bottom(XLBorderStyleValues.Medium)),
                    FontEmphasis(Bold),
                    FontEmphasis(Italic),
                ];
            IEnumerable<Item> Items()
            {
                foreach (var heading in new[] { "Month", "Letter Count" })
                {
                    yield return Cell([String(heading), .. headingStyle]);
                }
                yield return Go(NewRow);
                for (var m = 1; m <= 12; m++)
                {
                    var monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(NewRow);
                }
                yield return AutoFit(AutoFitFactory.AllCols);
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("Autofitting.xlsx"));

            using var wb = TestFiles.Open("Autofitting.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B13", ws.RangeUsed().RangeAddress.ToString());
            // AutoFit should have made both columns wider than ClosedXML's 8.43 default.
            Assert.True(ws.Column(1).Width > 8.43);
            Assert.True(ws.Column(2).Width > 8.43);
        }
        [Fact]
        void MergingCellsAndVerticalAlignment()
        {
            IEnumerable<Item> Items()
            {
                yield return Go(NewRow);
                foreach (var (heading, colWidth) in new (string, double)[]
                    {
                        ("ID", 3.22),
                        ("Car Name", 10.33),
                        ("Car Descriptions", 49.33),
                        ("Car Registration", 16.89),
                    })
                {
                    yield return Cell(
                        [
                            String(heading),
                            FontEmphasis(Bold),
                            FontName("Calibri"),
                            FontSize(11),
                            HorizontalAlignment(Center),
                            FontColor(XLColor.FromArgb(0,255,255,255)),
                            BackgroundColor(XLColor.FromArgb(0,68,114,196)),
                            Border(All(Thin)),
                            CellSize(ColWidth(colWidth)),
                        ]);
                }
                yield return Go(NewRow);
                yield return Cell([Integer(1),
                        HorizontalAlignment(HorizontalAlignmentFactory.Left),
                        VerticalAlignment(VerticalAlignmentFactory.TopMost),
                        Name("ID")]);
                yield return Cell([String("Ford Fiesta"),
                        HorizontalAlignment(Center),
                        VerticalAlignment(VerticalAlignmentFactory.Middle)]);
                yield return Cell([String("Car Technical Details:"),
                        Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 1"),
                        Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 2"),
                        Next(DownBy(1))]);
                yield return Cell([String("Technical Detail 3"),
                        Name("LastL")]);
                yield return Go(RC(3, 4));
                yield return Cell([  String("AB12 CDE"),
                            HorizontalAlignment(HorizontalAlignmentFactory.Right),
                            VerticalAlignment(VerticalAlignmentFactory.Base),
                            Name("Reg") ]);
                yield return Go(RC(6, 4));
                yield return Cell([Name("RegEnd")]);
                yield return Go(RC(7, 3));
                yield return Cell([String("Another Technical Detail"),
                        FontEmphasis(Italic),
                        VerticalAlignment(VerticalAlignmentFactory.Middle),
                        Name("TD"),
                        Next(Stay)]);
                yield return Go(DownBy(1));
                yield return Cell([Name("info")]);
                // Merging between named and specific cells
                yield return MergeCells(ColRowLabel("B", 3), ColRowLabel("B", 6));
                yield return MergeCells(NamedCell("ID"), ColRowLabel("A", 6));
                yield return MergeCells(ColRowLabel("C", 7), NamedCell("info"));
                yield return MergeCells(NamedCell("Reg"), NamedCell("RegEnd"));
                yield return Go(RC(10, 1));
                yield return Cell([String("Merging from a starting cell given a depth and span"),
                        BackgroundColor(XLColor.FromArgb(0, 80, 180, 220)),
                        FontEmphasis(Bold),
                        HorizontalAlignment(Center)]);
                yield return MergeCells(ColRowLabel("A", 10), ColRowLabel("D", 10));
                yield return Go(RC(12, 2));
                yield return Cell([String("The components that make up a car are: "),
                        Name("components"),
                        HorizontalAlignment(HorizontalAlignmentFactory.Left),
                        VerticalAlignment(VerticalAlignmentFactory.TopMost),
                        Border(BorderFactory.All(XLBorderStyleValues.MediumDashDot))]);
                yield return Go(RC(12, 4));
                yield return Cell([Border(BorderFactory.All(XLBorderStyleValues.MediumDashDot))]);
                yield return Go(RC(14, 4));
                yield return Cell([Border(BorderFactory.All(XLBorderStyleValues.MediumDashDot))]);
                yield return Go(RC(15, 2));
                yield return Cell([String("Road Tax"),
                        HorizontalAlignment(Center),
                        VerticalAlignment(VerticalAlignmentFactory.Middle),
                        Border(BorderFactory.All(XLBorderStyleValues.SlantDashDot))]);
                yield return Go(RC(16, 2));
                yield return Cell([Border(BorderFactory.All(XLBorderStyleValues.SlantDashDot))]);
                // Forward merging - cell name, cell contents, shading & top LH corner of border are retained
                yield return MergeCells(NamedCell("components"), SpanDepth(3, 3));
                yield return MergeCells(ColRowLabel("B", 15), SpanDepth(1, 2));
                yield return Go(RC(17, 4));
                yield return Cell([String("Insurance"),
                        Name("insurance"), // NamedCells cannot begin with a number
                        Border(BorderFactory.All(XLBorderStyleValues.Dashed))]);
                yield return Go(RC(17, 3));
                yield return Cell([Border(BorderFactory.All(XLBorderStyleValues.Dashed))]);
                yield return Go(RC(17, 2));
                yield return Cell([Border(BorderFactory.All(XLBorderStyleValues.Dashed))]);
                yield return Go(RC(16, 4));
                yield return Cell([String("Signature")]);
                // Reverse Merging - original cell contents, cell name and cell shading are lost
                // Only bottom RH corner of the border is retained
                yield return MergeCells(SpanDepth(3, 1), NamedCell("insurance"));
                yield return MergeCells(SpanDepth(2, 2), ColRowLabel("D", 16));
            }
            CsExcel.Render.AsFile(Items(), TestFiles.PathFor("MergeCellsWithVerticalAlignment.xlsx"));

            using var wb = TestFiles.Open("MergeCellsWithVerticalAlignment.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("ID", ws.Cell(2, 1).GetString());
            Assert.Equal(1, ws.Cell(3, 1).GetValue<int>());
            Assert.Equal("Ford Fiesta", ws.Cell(3, 2).GetString());
            Assert.Equal("AB12 CDE", ws.Cell(3, 4).GetString());
            Assert.Equal("Another Technical Detail", ws.Cell(7, 3).GetString());
            Assert.Equal("Merging from a starting cell given a depth and span", ws.Cell(10, 1).GetString());
            Assert.Equal("The components that make up a car are: ", ws.Cell(12, 2).GetString());
            Assert.Equal("Road Tax", ws.Cell(15, 2).GetString());
            var mergedRanges = ws.MergedRanges.Select(m => m.RangeAddress.ToString()).ToHashSet();
            foreach (var expected in new[] { "A3:A6", "B3:B6", "D3:D6", "C7:C8", "A10:D10", "B12:D14", "B15:B16", "C15:D16", "B17:D17" })
            {
                Assert.Contains(expected, mergedRanges);
            }
        }

        record JoiningInfo(string Name, int Age, decimal Fees, string DateJoined);

        // Tables from records/POCOs. FsExcel's own Table.fromInstance/fromSeq only work for F#
        // record types; CsExcel.Table reimplements the same layout logic with plain reflection so
        // it also works for C# classes, C# records (as used below) and anonymous types.
        [Fact]
        void TablesFromRecords()
        {
            var records = new[] {
                new JoiningInfo("Jane Smith", 32, 59.25m, "2022-03-12"), // Excel will treat these strings as dates
                new JoiningInfo("Michael Nguyễn", 23, 61.2m, "2022-03-13"),
                new JoiningInfo("Sofia Hernández", 58, 59.25m, "2022-03-15") };

            CellProp[] CellStyleVertical(int index, string name)
            {
                if (index == 0)
                {
                    return [ FontEmphasis(Bold) ];
                }
                else if (name == "Fees")
                {
                    return [ FormatCode("$0.00") ];
                }
                else
                {
                    return [];
                }
            }
            CellProp[] CellStyleHorizontal(int index, string name)
            {
                if (index == 0)
                {
                    return [ Border(BorderFactory.Bottom(XLBorderStyleValues.Medium)), FontEmphasis(Bold) ];
                }
                else if (name == "Fees")
                {
                    return [ FormatCode("$0.00") ];
                }
                else
                {
                    return [];
                }
            }
            var items = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Vertical,CellStyleVertical);
            CsExcel.Render.AsFile([.. items, AutoFit(AutoFitFactory.All)], TestFiles.PathFor("RecordSequenceVertical.xlsx"));

            var items2 = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Horizontal, CellStyleHorizontal);
            CsExcel.Render.AsFile([.. items2, AutoFit(AutoFitFactory.All)], TestFiles.PathFor("RecordSequenceHorizontal.xlsx"));

            foreach (var r in records.Take(1))
            {
                var cellsVertical = CsExcel.Table.fromInstance(r,CsExcel.Table.DirectionFactory.Vertical, CellStyleVertical);
                CsExcel.Render.AsFile([.. cellsVertical, AutoFit(AutoFitFactory.All)], TestFiles.PathFor("RecordInstanceVertical.xlsx"));

                var cellsHorizontal = CsExcel.Table.fromInstance(r,CsExcel.Table.DirectionFactory.Horizontal, CellStyleHorizontal);
                CsExcel.Render.AsFile([.. cellsHorizontal, AutoFit(AutoFitFactory.All)], TestFiles.PathFor("RecordInstanceHorizontal.xlsx"));
            }

            using (var wb = TestFiles.Open("RecordSequenceVertical.xlsx"))
            {
                var ws = wb.Worksheet(1);
                Assert.Equal("Name", ws.Cell(1, 1).GetString());
                Assert.True(ws.Cell(1, 1).Style.Font.Bold);
                Assert.Equal("Jane Smith", ws.Cell(1, 2).GetString());
                Assert.Equal("Michael Nguyễn", ws.Cell(1, 3).GetString());
                Assert.Equal("Fees", ws.Cell(3, 1).GetString());
                Assert.Equal(59.25, ws.Cell(3, 2).GetValue<double>());
                Assert.Equal("$0.00", ws.Cell(3, 2).Style.NumberFormat.Format);
                Assert.Equal("2022-03-12", ws.Cell(4, 2).GetString());
            }
            using (var wb = TestFiles.Open("RecordSequenceHorizontal.xlsx"))
            {
                var ws = wb.Worksheet(1);
                Assert.Equal(["Name", "Age", "Fees", "DateJoined"], new[] { 1, 2, 3, 4 }.Select(c => ws.Cell(1, c).GetString()));
                Assert.True(ws.Cell(1, 1).Style.Font.Bold);
                Assert.Equal("Jane Smith", ws.Cell(2, 1).GetString());
                Assert.Equal(32, ws.Cell(2, 2).GetValue<int>());
                Assert.Equal(59.25, ws.Cell(2, 3).GetValue<double>());
                Assert.Equal("$0.00", ws.Cell(2, 3).Style.NumberFormat.Format);
                Assert.Equal("Sofia Hernández", ws.Cell(4, 1).GetString());
            }
            using (var wb = TestFiles.Open("RecordInstanceVertical.xlsx"))
            {
                var ws = wb.Worksheet(1);
                Assert.Equal("Name", ws.Cell(1, 1).GetString());
                Assert.Equal("Jane Smith", ws.Cell(1, 2).GetString());
                Assert.Equal("Age", ws.Cell(2, 1).GetString());
                Assert.Equal(32, ws.Cell(2, 2).GetValue<int>());
            }
            using (var wb = TestFiles.Open("RecordInstanceHorizontal.xlsx"))
            {
                var ws = wb.Worksheet(1);
                Assert.Equal(["Name", "Age", "Fees", "DateJoined"], new[] { 1, 2, 3, 4 }.Select(c => ws.Cell(1, c).GetString()));
                Assert.Equal("Jane Smith", ws.Cell(2, 1).GetString());
                Assert.Equal(32, ws.Cell(2, 2).GetValue<int>());
            }
        }

        record Wide(string A, string B, string C, string D, string E);

        [Fact]
        void TableFromIEnumerableVerticalFieldCountDiffersFromRecordCount()
        {
            // Regression test: Table.fromIEnumerable's Vertical layout must return the cursor to
            // row 1 between columns based on the number of FIELDS per record (5, here), not the
            // number of records (2, here) - those two numbers used to be conflated. With more fields
            // than records+1, the old code returned the cursor too high, so every column after the
            // first started progressively below row 1 instead of at it.
            var rows = new[]
            {
                new Wide("A1", "A2", "A3", "A4", "A5"),
                new Wide("B1", "B2", "B3", "B4", "B5"),
            };
            CellProp[] NoStyle(int index, string name) => [];
            var items = CsExcel.Table.fromIEnumerable(rows, CsExcel.Table.DirectionFactory.Vertical, NoStyle);
            CsExcel.Render.AsFile(items, TestFiles.PathFor("TableVerticalFieldCountMismatch.xlsx"));

            using var wb = TestFiles.Open("TableVerticalFieldCountMismatch.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(["A", "B", "C", "D", "E"], new[] { 1, 2, 3, 4, 5 }.Select(r => ws.Cell(r, 1).GetString()));
            for (var i = 0; i < rows.Length; i++)
            {
                var col = i + 2; // column A is the header; each record gets the next column over
                var expected = new[] { rows[i].A, rows[i].B, rows[i].C, rows[i].D, rows[i].E };
                for (var f = 0; f < expected.Length; f++)
                {
                    Assert.Equal(expected[f], ws.Cell(f + 1, col).GetString());
                }
            }
        }

        interface IBase { string Name { get; } }
        record DerivedInfo(string Name, int Extra) : IBase;

        [Fact]
        void TableFromInstancePolymorphicTypeUsesDeclaredTypeFields()
        {
            // Regression test: the header and body field sets must both come from the declared
            // type parameter 'T (IBase, here), not the runtime type of the instance (DerivedInfo,
            // which has an extra "Extra" property) - otherwise they diverge and either throw
            // (Vertical: List.zip length mismatch) or silently misalign columns (Horizontal).
            IBase instance = new DerivedInfo("Widget", 42);
            CellProp[] NoStyle(int index, string name) => [];
            var items = CsExcel.Table.fromInstance(instance, CsExcel.Table.DirectionFactory.Vertical, NoStyle);
            CsExcel.Render.AsFile(items, TestFiles.PathFor("TablePolymorphicInstance.xlsx"));

            using var wb = TestFiles.Open("TablePolymorphicInstance.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B1", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("Name", ws.Cell(1, 1).GetString());
            Assert.Equal("Widget", ws.Cell(1, 2).GetString());
        }

        interface IHasId { string Id { get; } }
        interface IBaseWithInheritedProperty : IHasId { string Name { get; } }
        record DerivedWithInheritedProperty(string Id, string Name, int Extra) : IBaseWithInheritedProperty;

        [Fact]
        void TableFromInstanceMultiLevelInterfaceIncludesBaseInterfaceFields()
        {
            // Regression test: Type.GetProperties only returns properties declared directly on an
            // interface, not ones inherited from a base interface it extends (unlike classes, where
            // GetProperties does include inherited members) - so IBaseWithInheritedProperty's own
            // "Name" property must be combined with IHasId's "Id" property, not just "Name" alone.
            IBaseWithInheritedProperty instance = new DerivedWithInheritedProperty("id-1", "Widget", 42);
            CellProp[] NoStyle(int index, string name) => [];
            var items = CsExcel.Table.fromInstance(instance, CsExcel.Table.DirectionFactory.Vertical, NoStyle);
            CsExcel.Render.AsFile(items, TestFiles.PathFor("TableMultiLevelInterfaceInstance.xlsx"));

            using var wb = TestFiles.Open("TableMultiLevelInterfaceInstance.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B2", ws.RangeUsed().RangeAddress.ToString());
            var headers = new[] { ws.Cell(1, 1).GetString(), ws.Cell(2, 1).GetString() };
            Assert.Contains("Id", headers);
            Assert.Contains("Name", headers);
        }

        class NoFields { }

        [Fact]
        void TableFromIEnumerableVerticalZeroFieldTypeDoesNotThrow()
        {
            // Smoke test, not a true regression test: fromIEnumerable's Vertical branch ends with
            // Go(DownBy(depth-1)), where depth is the field count - a type with zero public
            // readable properties makes depth 0, which used to underflow to DownBy(-1). That
            // underflow turns out to be harmless in practice (Go's row math clamps elsewhere in
            // this module), so this test can't actually distinguish the guarded code from the
            // unguarded version - both pass it. Kept anyway as cheap insurance/documentation for the
            // degenerate zero-field case, and a leading marker cell forces a worksheet to exist (a
            // workbook with zero cells has no worksheet at all and can't be saved).
            var tableItems = CsExcel.Table.fromIEnumerable(new[] { new NoFields(), new NoFields() }, CsExcel.Table.DirectionFactory.Vertical, (int index, string name) => []);
            IEnumerable<Item> items = [Cell([String("marker")]), .. tableItems];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("TableVerticalZeroFieldType.xlsx"));

            using var wb = TestFiles.Open("TableVerticalZeroFieldType.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("marker", ws.Cell(1, 1).GetString());
        }
        [Fact]
        void RenderingInFableElmishOrSimilar()
        {
            var items = new Item[] { 
                Cell([String("Hello world!")])            
            };
            var bytes = CsExcel.Render.AsStreamBytes(items);
            Assert.Equal(6188,bytes.Length);
        }
        [Fact]
        void DataTypes()
        {
            var items = new Item[] {
                Cell([String("String")]),
                Cell([String("string")]),
                Go(NewRow),
                Cell([String("Integer")]),
                Cell([Integer(42)]),
                Go(NewRow),
                Cell([String("Number")]),
                Cell([Float(Math.PI)]),
                Go(NewRow),
                Cell([String("Boolean")]),
                Cell([Boolean(false)]),
                Go(NewRow),
                Cell([String("DateTime")]),
                Cell([DateTime(new System.DateTime(1903, 12, 17))]),
                Go(NewRow),
                Cell([String("TimeSpan")]),
                Cell([
                    TimeSpan(new System.TimeSpan(hours: 1, minutes: 2, seconds: 3)),
                    FormatCode("hh:mm:ss")
                ]),
            };
            CsExcel.Render.AsFile(items, TestFiles.PathFor("DataTypes.xlsx"));

            using var wb = TestFiles.Open("DataTypes.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("string", ws.Cell(1, 2).GetString());
            Assert.Equal(42, ws.Cell(2, 2).GetValue<int>());
            Assert.Equal(Math.PI, ws.Cell(3, 2).GetValue<double>(), precision: 10);
            Assert.False(ws.Cell(4, 2).GetValue<bool>());
            Assert.Equal(new System.DateTime(1903, 12, 17), ws.Cell(5, 2).GetValue<System.DateTime>());
            Assert.Equal("hh:mm:ss", ws.Cell(6, 2).Style.NumberFormat.Format);
            Assert.Equal(new System.TimeSpan(hours: 1, minutes: 2, seconds: 3).TotalDays, ws.Cell(6, 2).GetValue<double>(), precision: 6);
        }
        [Fact]
        void RenderingAsHtml()
        {
            bool IsHeader(int r, int c) => r == 0 || c == 0;
            var items = new Item[] {
                Worksheet("Worksheet 1"),
                Style([FontEmphasis(Bold)]),
                Cell([String("Item")]),
                Cell([String("Example")]),
                Style([]),
                Go(NewRow),
                Cell([String("String")]),
                Cell([String("string")]),
                Go(NewRow),
                Cell([String("Integer")]),
                Cell([Integer(42)]),
                Go(NewRow),
                Cell([String("Number")]),
                Cell([Float(Math.PI)]),
                Go(NewRow),
                Cell([String("Boolean")]),
                Cell([Boolean(false)]),
                Go(NewRow),
                Cell([String("DateTime")]),
                Cell([DateTime(new System.DateTime(1903, 12, 17))]),
                Go(NewRow),
                Cell([String("TimeSpan")]),
                Cell([
                    TimeSpan(new System.TimeSpan(hours: 1, minutes: 2, seconds: 3)),
                    FormatCode("hh:mm:ss")
                ]),
                Go(NewRow),
                Cell([String("Bold")]),
                Cell([
                    String("I am bold"),
                    FontEmphasis(Bold)
                ]),
                Go(NewRow),
                Cell([String("Italic")]),
                Cell([
                    String("I am Italic"),
                    FontEmphasis(Italic)
                ]),
                Go(NewRow),
                Cell([String("Underlined")]),
                Cell([
                    String("I am underlined"),
                    FontEmphasis(Underline(XLFontUnderlineValues.Single))
                ]),
                Go(NewRow),
                Worksheet("Worksheet 2"),
                Cell([String("I am another table")]),
            };

            var htmlString = CsExcel.Render.AsHtml(items, IsHeader);
            // HTML(htmlString) can be used in notebook

            Assert.Contains("<h3>Worksheet 1</h3>", htmlString);
            Assert.Contains("<h3>Worksheet 2</h3>", htmlString);
            Assert.Contains("I am another table", htmlString);
            Assert.Contains("font-weight: bold", htmlString);
            Assert.Contains("I am bold", htmlString);
        }
        [Fact]
        void AutoFilterEnableOnly()
        {
            var headings = new Item[]
            {
                Cell([String("StringCol"), HorizontalAlignment(Center)]),
                Cell([String("IntCol"), HorizontalAlignment(Center)]),
                Cell([String("FloatCol"), HorizontalAlignment(Center)]),
                Cell([String("DateTimeCol"), HorizontalAlignment(Center)]),
                Cell([String("BooleanCol"), HorizontalAlignment(Center)]),
                Go(NewRow)
            };
            var rows = (from i in Enumerable.Range(1, 5)
                       select new Item[]
                       {
                           Cell([String($"String{i}")]),
                           Cell([Integer(i)]),
                           Cell([Float((i + 0.1))]),
                           Cell([DateTime(new System.DateTime(2017, 7, 15, 5, 33, 0).AddMinutes(i))]),
                           Cell([Boolean(i % 2 == 0)]),
                           Go(NewRow)
                       }).SelectMany(x => x);
            IEnumerable<Item> items = [.. headings, .. rows, AutoFit(AutoFitFactory.All), AutoFilter([AutoFilterFactory.EnableOnly(AutoFilterRangeFactory.RangeUsed)])];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("AutoFilterEnableOnly.xlsx"));

            using var wb = TestFiles.Open("AutoFilterEnableOnly.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:E6", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.AutoFilter.IsEnabled);
            Assert.Equal("A1:E6", ws.AutoFilter.Range.RangeAddress.ToString());
            Assert.Equal("String3", ws.Cell(4, 1).GetString());
        }
        [Fact]
        void AutoFilterCompound()
        {
            var headings = new Item[]
            {
                Cell([String("StringCol"), HorizontalAlignment(Center)]),
                Cell([String("IntCol"), HorizontalAlignment(Center)]),
                Cell([String("FloatCol"), HorizontalAlignment(Center)]),
                Cell([String("DateTimeCol"), HorizontalAlignment(Center)]),
                Cell([String("BooleanCol"), HorizontalAlignment(Center)]),
                Go(NewRow)
            };
            var rows = (from i in Enumerable.Range(1, 5)
                        select new Item[]
                        {
                           Cell([String($"String{i}")]),
                           Cell([Integer(i)]),
                           Cell([Float((i + 0.1))]),
                           Cell([DateTime(new System.DateTime(2017, 7, 15, 5, 33, 0).AddMinutes(i))]),
                           Cell([Boolean(i % 2 == 0)]),
                           Go(NewRow)
                        }).SelectMany(x => x);
            IEnumerable<Item> items = 
                [   .. headings, 
                    .. rows, 
                    AutoFit(AutoFitFactory.All), 
                    AutoFilter(
                        [
                            AutoFilterFactory.GreaterThanInt(AutoFilterRangeFactory.RangeUsed,2,3),
                            AutoFilterFactory.EqualToBool(AutoFilterRangeFactory.RangeUsed,5,true)
                        ])
                ];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("AutoFilterCompound.xlsx"));

            using var wb = TestFiles.Open("AutoFilterCompound.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:E6", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.AutoFilter.IsEnabled);
            Assert.Equal("A1:E6", ws.AutoFilter.Range.RangeAddress.ToString());
        }

        static Item[] MakeFreezePanesHeadingsAndRows()
        {
            var headings = new Item[]
            {
                Cell([String("StringCol"), HorizontalAlignment(Center)]),
                Cell([String("IntCol"), HorizontalAlignment(Center)]),
                Cell([String("FloatCol"), HorizontalAlignment(Center)]),
                Cell([String("DateTimeCol"), HorizontalAlignment(Center)]),
                Cell([String("BooleanCol"), HorizontalAlignment(Center)]),
                Go(NewRow)
            };
            var rows = (from i in Enumerable.Range(1, 5)
                        select new Item[]
                        {
                           Cell([String($"String{i}")]),
                           Cell([Integer(i)]),
                           Cell([Float((i + 0.1))]),
                           Cell([DateTime(new System.DateTime(2017, 7, 15, 5, 33, 0).AddMinutes(i))]),
                           Cell([Boolean(i % 2 == 0)]),
                           Go(NewRow)
                        }).SelectMany(x => x);
            return [.. headings, .. rows];
        }

        [Fact]
        void FreezePanesRowAndColumn()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(Panes(1, 1))];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("FreezePanes.xlsx"));

            using var wb = TestFiles.Open("FreezePanes.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(1, ws.SheetView.SplitRow);
            Assert.Equal(1, ws.SheetView.SplitColumn);
        }
        [Fact]
        void FreezePanesTopRow()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(TopRow)];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("FreezePanesTopRow.xlsx"));

            using var wb = TestFiles.Open("FreezePanesTopRow.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(1, ws.SheetView.SplitRow);
            Assert.Equal(0, ws.SheetView.SplitColumn);
        }
        [Fact]
        void FreezePanesFirstColumn()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(FirstColumn)];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("FreezePanesFirstColumn.xlsx"));

            using var wb = TestFiles.Open("FreezePanesFirstColumn.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(0, ws.SheetView.SplitRow);
            Assert.Equal(1, ws.SheetView.SplitColumn);
        }
        [Fact]
        void FreezePanesUnfreezePanes()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(TopRow), FreezePanes(UnfreezePanes)];
            CsExcel.Render.AsFile(items, TestFiles.PathFor("FreezePanesUnfreezePanes.xlsx"));

            using var wb = TestFiles.Open("FreezePanesUnfreezePanes.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal(0, ws.SheetView.SplitRow);
            Assert.Equal(0, ws.SheetView.SplitColumn);
        }
    }
}
