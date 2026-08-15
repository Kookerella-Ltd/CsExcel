using static CsExcel.PositionFactory;
using static CsExcel.Fluent.ItemFactory;
using CsExcel.Fluent;
using System.Globalization;
using FsExcel;
using static CsExcel.BorderFactory;
using static CsExcel.FontEmphasisFactory;
using static ClosedXML.Excel.XLBorderStyleValues;
using static ClosedXML.Excel.XLFontUnderlineValues;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace UnitTests
{
    public class Fluent
    {
        [Fact]
        public void HelloWorld()
        {
            new[] {
                Cell(ps => ps.AddString("Hello World"))
            }.AsFile(TestFiles.PathFor("fhelloWorld.xlsx"));

            using var wb = TestFiles.Open("fhelloWorld.xlsx");
            Assert.Equal("Hello World", wb.Worksheet(1).Cell(1, 1).GetString());
        }
        [Fact]
        public void MultipleCells()
        {
            (from n in Enumerable.Range(1, 10)
             select Cell(ps => ps.AddInteger(n)))
                .AsFile(TestFiles.PathFor("fMultipleCells.xlsx"));

            using var wb = TestFiles.Open("fMultipleCells.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:J1", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal(1, ws.Cell(1, 1).GetValue<int>());
            Assert.Equal(10, ws.Cell(1, 10).GetValue<int>());
        }
        [Fact]
        public void VerticalMovement()
        {
            (from m in Enumerable.Range(1, 12)
             let monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m)
             select Cell(ps => ps.AddString(monthName).AddNext(DownBy(1))))
                .AsFile(TestFiles.PathFor("fVerticalMovement.xlsx"));

            using var wb = TestFiles.Open("fVerticalMovement.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:A12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
        }
        [Fact]
        public void VerticalMovement2()
        {
            IEnumerable<Item> Cells()
            {
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(ps => ps.AddString(monthName));
                    yield return Cell(ps => ps.AddInteger(monthName.Length).AddNext(NewRow));
                }
            };
            Cells().AsFile(TestFiles.PathFor("fVerticalMovement2.xlsx"));

            using var wb = TestFiles.Open("fVerticalMovement2.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal(7, ws.Cell(1, 2).GetValue<int>());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
        }
        [Fact]
        public void VerticalMovement3()
        {
            IEnumerable<Item> Cells()
            {
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(ps => ps.AddString(monthName));
                    yield return Cell(ps => ps.AddInteger(monthName.Length));
                    yield return Go(NewRow);
                }
            };

            Cells().AsFile(TestFiles.PathFor("fVerticalMovement3.xlsx"));

            using var wb = TestFiles.Open("fVerticalMovement3.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 1).GetString());
            Assert.Equal(7, ws.Cell(1, 2).GetValue<int>());
            Assert.Equal("December", ws.Cell(12, 1).GetString());
        }
        [Fact]
        public void Indentation()
        {
            IEnumerable<Item> Items()
            {
                yield return Go(IndentBy(2));
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(ps => ps.AddString(monthName));
                    yield return Cell(ps => ps.AddInteger(monthName.Length));
                    yield return Go(NewRow);
                }
            }

            Items().AsFile(TestFiles.PathFor("fIndentation.xlsx"));

            using var wb = TestFiles.Open("fIndentation.xlsx");
            var ws = wb.Worksheet(1);
            // IndentBy is relative (start col + 2 => column C), unlike Vanilla's Indent, which is
            // absolute (sets column to 2 => column B) - see Vanilla.Indentation for the contrast.
            Assert.Equal("C1:D12", ws.RangeUsed().RangeAddress.ToString());
            Assert.Equal("January", ws.Cell(1, 3).GetString());
            Assert.Equal("December", ws.Cell(12, 3).GetString());
        }
        [Fact]
        public void BorderAndFontStyling()
        {
            IEnumerable<Item> Items()
            {
                foreach (var heading in new[] { "Month", "Letter Count" })
                {
                    yield return Cell(empty =>
                        empty.AddString(heading)
                            .AddBorder(Bottom(Medium))
                            .AddFontEmphasis(Bold)
                            .AddFontEmphasis(Italic));
                }
                yield return Go(NewRow);
                foreach (var m in Enumerable.Range(1, 12))
                {
                    var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
                    yield return Cell(empty =>
                        empty.AddString(monthName)
                            .AddFontEmphasis(Underline(DoubleAccounting))
                            .Concat(empty.AddFontEmphasis(StrikeThrough)
                                .Where(_ => monthName == "May")));
                    yield return Cell(empty => empty.AddInteger(monthName.Length));
                    yield return Go(NewRow);
                }
            };

            Items().AsFile(TestFiles.PathFor("fBorderAndFontStyling.xlsx"));

            using var wb = TestFiles.Open("fBorderAndFontStyling.xlsx");
            var ws = wb.Worksheet(1);
            Assert.Equal("A1:B13", ws.RangeUsed().RangeAddress.ToString());
            Assert.True(ws.Cell(1, 1).Style.Font.Bold);
            Assert.True(ws.Cell(1, 1).Style.Font.Italic);
            Assert.Equal("May", ws.Cell(6, 1).GetString());
            Assert.True(ws.Cell(6, 1).Style.Font.Strikethrough);
            Assert.False(ws.Cell(2, 1).Style.Font.Strikethrough);
        }
    }
}