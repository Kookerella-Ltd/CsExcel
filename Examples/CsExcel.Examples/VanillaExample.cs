using FsExcel;
using CsExcel.Fluent;
using static CsExcel.CellPropFactory;
using static CsExcel.FontEmphasisFactory;
using static CsExcel.ItemFactory;
using static CsExcel.PositionFactory;

namespace CsExcel.Examples;

// The Vanilla calling style: static factory methods and an immutable CellProp list per cell -
// the closest match to FsExcel's own shape. Note there is no "workbook" object being created and
// mutated anywhere below - Worksheet("Report") is itself an Item in the sequence, and it's what
// creates (and switches to) the worksheet.
static class VanillaExample
{
    public static void Run()
    {
        IEnumerable<Item> items =
        [
            Worksheet("Report"),

            Cell([ String("Name"), FontEmphasis(Bold) ]),
            Cell([ String("Units Sold"), FontEmphasis(Bold) ]),
            Cell([ String("Revenue"), FontEmphasis(Bold), Next(NewRow) ]),

            Cell([ String("Widget") ]),
            Cell([ Integer(120) ]),
            Cell([ Float(1450.50), FormatCode("$0.00"), Next(NewRow) ]),

            Cell([ String("Gadget") ]),
            Cell([ Integer(75) ]),
            Cell([ Float(980.00), FormatCode("$0.00"), Next(NewRow) ]),

            AutoFit(AutoFitFactory.AllCols),
        ];

        var path = Path.Combine(AppContext.BaseDirectory, "vanilla-example.xlsx");
        items.AsFile(path);
        Console.WriteLine($"Wrote {path}");
    }
}
