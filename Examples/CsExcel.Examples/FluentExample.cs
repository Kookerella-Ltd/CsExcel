using FsExcel;
using CsExcel.Fluent;
using static CsExcel.Fluent.ItemFactory;
using static CsExcel.PositionFactory;

namespace CsExcel.Examples;

// The Fluent calling style: a chainable builder, closer to a typical C# builder API
// (Cell().String("x").Bold() rather than Cell([String("x"), FontEmphasis(Bold)])). Produces
// identical output to VanillaExample - the two styles can even be mixed in the same project
// (just not usefully in the same file, since both ItemFactory modules define a Cell/Worksheet/...
// that would otherwise collide).
static class FluentExample
{
    public static void Run()
    {
        List<Item> items =
        [
            Worksheet("Report"),

            Cell().String("Name").Bold(),
            Cell().String("Units Sold").Bold(),
            Cell().String("Revenue").Bold().Next(NewRow),

            Cell().String("Widget"),
            Cell().Integer(120),
            Cell().Float(1450.50).FormatCode("$0.00").Next(NewRow),

            Cell().String("Gadget"),
            Cell().Integer(75),
            Cell().Float(980.00).FormatCode("$0.00").Next(NewRow),

            AutoFit(AutoFitFactory.AllCols),
        ];

        var path = Path.Combine(AppContext.BaseDirectory, "fluent-example.xlsx");
        items.AsFile(path);
        Console.WriteLine($"Wrote {path}");
    }
}
