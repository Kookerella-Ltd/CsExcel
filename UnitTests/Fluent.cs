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

namespace UnitTests
{
    public class Fluent
    {
        [Fact]
        public void HelloWorld()
        {
            new[] {
                Cell(ps => ps.AddString("Hello World"))
            }.AsFile("""c:\temp\fhelloWorld.xlsx""");
        }
        [Fact]
        public void MultipleCells()
        {
            (from n in Enumerable.Range(1, 10)
             select Cell(ps => ps.AddInteger(n)))
                .AsFile("""c:\temp\fMultipleCells.xlsx""");
        }
        [Fact]
        public void VerticalMovement()
        {
            (from m in Enumerable.Range(1, 12)
             let monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m)
             select Cell(ps => ps.AddString(monthName).AddNext(DownBy(1))))
                .AsFile("""c:\temp\fVerticalMovement.xlsx""");
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
            Cells().AsFile("""c:\temp\fVerticalMovement2.xlsx""");
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

            Cells().AsFile("""c:\temp\fVerticalMovement3.xlsx""");
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

            Items().AsFile("""c:\temp\fIndentation.xlsx""");
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

            Items().AsFile("""c:\temp\fBorderAndFontStyling.xlsx""");
        }
    }
}