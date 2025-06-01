using FsExcel;
using CsExcel;
using static CsExcel.PositionFactory;
using static CsExcel.ItemFactory;
using static CsExcel.CellPropFactory;
using static CsExcel.BorderFactory;
using static CsExcel.FontEmphasisFactory;
using System.Globalization;
using static ClosedXML.Excel.XLBorderStyleValues;
using static ClosedXML.Excel.XLFontUnderlineValues;
using System.Runtime.CompilerServices;
using Microsoft.FSharp.Control;

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
                    yield return Cell([Integer(monthName.Length), Next(NewRow)]);
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
                    yield return Go(NewRow);
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
                    yield return Go(IndentBy(2));
                    yield return Cell([String(monthName)]);
                    yield return Cell([Integer(monthName.Length)]);
                    yield return Go(NewRow);
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
                            Border(Bottom(Medium)),
                            FontEmphasis(Bold),
                            FontEmphasis(Italic)
                        ]);
                }
                yield return Go(NewRow);
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
                    yield return Go(NewRow);
                }
            };

            CsExcel.Render.AsFile(Items(), """c:\temp\BorderAndFontStyling.xlsx""");
        }

        [Fact]
        public void HelloWorld2()
        {
            //Item[] makeExcel(ItemFactory factory)
            //{
            //    return new[]
            //    {
            //        factory.Cell(f => new[] { f.String("hello world") }),
            //        factory.Cell(f => new[] { f.String("This is a test") })
            //    };
            //}
            //ItemArrayExtensions.AsFile(makeExcel);
        }
        [Fact]
        public void HelloWorld3()
        {
            //Item[] makeExcel(ItemFactory factory)
            //{
            //    return new[]
            //    {
            //        factory.Cell(empty => empty.AddString("hello world")),
            //        factory.Cell(empty => empty.AddString("This is a test"))
            //    };
            //}
            //ItemArrayExtensions.AsFile(makeExcel);
        }
    }
}
