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

            CsExcel.Render.AsFile(cells, """c:\temp\helloWorld.xlsx""");
        }
        [Fact]
        public void MultipleCells()
        {
            var cells =
                from n in Enumerable.Range(1, 10)
                select Cell([Integer(n)]);
            CsExcel.Render.AsFile(cells, """c:\temp\MultipleCells.xlsx""");
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
            CsExcel.Render.AsFile(cells, """c:\temp\VerticalMovement.xlsx""");
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
            CsExcel.Render.AsFile(Cells(), """c:\temp\VerticalMovement2.xlsx""");
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
            CsExcel.Render.AsFile(Cells(), """c:\temp\VerticalMovement3.xlsx""");
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
            CsExcel.Render.AsFile(Cells(), """c:\temp\Indentation.xlsx""");
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

            CsExcel.Render.AsFile(Items(), """c:\temp\BorderAndFontStyling.xlsx""");
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

            CsExcel.Render.AsFile(Items(), """c:\temp\BorderAndFontStyling.xlsx""");
        }

        [Fact]
        public void FontAndNameSize()
        {
            var fontNames =
                SixLabors.Fonts.SystemFonts.Collection.Families.Select((fontFamily, i) => (fontFamily.Name, i)).OrderBy(f => f.Item1).Take(20);
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
            CsExcel.Render.AsFile(Items(), """c:\temp\FontAndNameSize.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\WrapText.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\TextRotation.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\NumberFormattingAndAlignment.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\Formulae.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\Color.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\RangeStyles.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\AddingABorderToMergedCells.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\AbsolutePositioning.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\AbsolutePositioning2.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\NamedCells.xlsx""");
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
            CsExcel.Render.AsFile(MakeWorksheetTabsItems(), """c:\temp\WorksheetsTabs.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\InsertingBlankRows.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\ColumnWidthsAndRowHeightsForAllCells.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\IndividualCellSizing.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\Autofitting.xlsx""");
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
            CsExcel.Render.AsFile(Items(), """c:\temp\MergeCellsWithVerticalAlignment.xlsx""");
        }

        //record JoiningInfo(string Name, int Age, decimal Fees, string DateJoined);

        //Tables from Records
        [Fact]
        void TablesFromRecords()
        {
            // this functionality doesnt work because FsExcel assumes the types are F# record types

            //var records = new[] {
            //    new { Name = "Jane Smith", Age = 32, Fees= 59.25m, DateJoined = "2022-03-12" }, // Excel will treat these strings as dates
            //    new { Name = "Michael Nguyễn", Age = 23, Fees =61.2m, DateJoined = "2022-03-13" },
            //    new { Name = "Sofia Hernández",Age = 58, Fees = 59.25m, DateJoined = "2022-03-15" } };

            //CellProp[] CellStyleVertical(int index, string name)
            //{
            //    if (index == 0)
            //    {
            //        return [ FontEmphasis(Bold) ];
            //    }
            //    else if (name == "Fees")
            //    {
            //        return [ FormatCode("$0.00") ];
            //    }
            //    else
            //    {
            //        return [];
            //    }
            //}
            //CellProp[] CellStyleHorizontal(int index, string name)
            //{
            //    if (index == 0)
            //    {
            //        return [ Border(BorderFactory.Bottom(XLBorderStyleValues.Medium)), FontEmphasis(Bold) ];
            //    }
            //    else if (name == "Fees")
            //    {
            //        return [ FormatCode("$0.00") ];
            //    }
            //    else
            //    {
            //        return [];
            //    }
            //}
            //var items = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Vertical,CellStyleVertical);
            //CsExcel.Render.AsFile([.. items, AutoFit(AutoFitFactory.All)], """c:\temp\RecordSequenceVertical.xlsx""");

            //var items2 = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Horizontal, CellStyleHorizontal);
            //CsExcel.Render.AsFile([.. items2, AutoFit(AutoFitFactory.All)], """c:\temp\RecordSequenceHorizontal.xlsx""");

            //foreach (var r in records.Take(1))
            //{
            //    var cellsVertical = CsExcel.Table.fromInstance(r,CsExcel.Table.DirectionFactory.Vertical, CellStyleVertical);
            //    CsExcel.Render.AsFile([.. cellsVertical, AutoFit(AutoFitFactory.All)], """c:\temp\RecordInstanceVertical.xlsx""");

            //    var cellsHorizontal = CsExcel.Table.fromInstance(r,CsExcel.Table.DirectionFactory.Horizontal, CellStyleHorizontal);
            //    CsExcel.Render.AsFile([.. cellsHorizontal, AutoFit(AutoFitFactory.All)], """c:\temp\RecordInstanceHorizontal.xlsx""");
            //}
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
            CsExcel.Render.AsFile(items, """c:\temp\DataTypes.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\AutoFilterEnableOnly.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\AutoFilterCompound.xlsx""");
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
            CsExcel.Render.AsFile(items, """c:\temp\FreezePanes.xlsx""");
        }
        [Fact]
        void FreezePanesTopRow()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(TopRow)];
            CsExcel.Render.AsFile(items, """c:\temp\FreezePanesTopRow.xlsx""");
        }
        [Fact]
        void FreezePanesFirstColumn()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(FirstColumn)];
            CsExcel.Render.AsFile(items, """c:\temp\FreezePanesFirstColumn.xlsx""");
        }
        [Fact]
        void FreezePanesUnfreezePanes()
        {
            IEnumerable<Item> items = [.. MakeFreezePanesHeadingsAndRows(), AutoFit(AutoFitFactory.All), FreezePanes(TopRow), FreezePanes(UnfreezePanes)];
            CsExcel.Render.AsFile(items, """c:\temp\FreezePanesUnfreezePanes.xlsx""");
        }
    }
}
