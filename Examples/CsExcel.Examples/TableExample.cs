using FsExcel;
using CsExcel.Fluent;
using static CsExcel.CellPropFactory;
using static CsExcel.FontEmphasisFactory;
using static CsExcel.ItemFactory;

namespace CsExcel.Examples;

// Builds table cells straight from a C# record's public properties via reflection, rather than
// listing out each Cell by hand - works for any C# class, record, or anonymous type. See
// CsExcel.Table.fromInstance for a single object, or fromIEnumerable for a sequence (as below).
static class TableExample
{
    record Product(string Name, int UnitsSold, double Revenue);

    public static void Run()
    {
        var products = new List<Product>
        {
            new("Widget", 120, 1450.50),
            new("Gadget", 75, 980.00),
        };

        IEnumerable<Item> items =
        [
            Worksheet("Products"),
            .. Table.fromIEnumerable(
                products,
                Table.DirectionFactory.Horizontal,
                (int index, string _) => index == 0 ? [FontEmphasis(Bold)] : Enumerable.Empty<CellProp>()),
        ];

        var path = Path.Combine(AppContext.BaseDirectory, "table-example.xlsx");
        items.AsFile(path);
        Console.WriteLine($"Wrote {path}");
    }
}
