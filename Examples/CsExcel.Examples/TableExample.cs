using FsExcel;
using CsExcel.Fluent;
using static CsExcel.CellPropFactory;
using static CsExcel.FontEmphasisFactory;
using static CsExcel.ItemFactory;

namespace CsExcel.Examples;

// Builds table cells straight from a plain C# class's public properties via reflection, rather
// than listing out each Cell by hand - works for any C# class, record, or anonymous type. See
// CsExcel.Table.fromInstance for a single object, or fromIEnumerable for a sequence (as below).
static class TableExample
{
    class Product
    {
        public string Name { get; set; } = "";
        public int UnitsSold { get; set; }
        public double Revenue { get; set; }
    }

    public static void Run()
    {
        var products = new List<Product>
        {
            new Product { Name = "Widget", UnitsSold = 120, Revenue = 1450.50 },
            new Product { Name = "Gadget", UnitsSold = 75, Revenue = 980.00 },
        };

        // Worksheet("Products") switches to (and creates) the worksheet the table is written to,
        // so it comes first in the list. fromIEnumerable then builds one Cell per property per
        // product, calling GetCellStyle once per cell so the header row (rowIndex 0) can be
        // styled differently from the data rows.
        IEnumerable<Item> items =
        [
            Worksheet("Products"),
            .. Table.fromIEnumerable(products, Table.DirectionFactory.Horizontal, GetCellStyle),
        ];

        var path = Path.Combine(AppContext.BaseDirectory, "table-example.xlsx");
        items.AsFile(path);
        Console.WriteLine($"Wrote {path}");
    }

    // rowIndex is 0 for the header row, 1 for the first product, 2 for the second, and so on.
    // propertyName is the C# property this cell came from (e.g. "Name", "UnitsSold").
    static IEnumerable<CellProp> GetCellStyle(int rowIndex, string propertyName)
    {
        if (rowIndex == 0)
        {
            return new List<CellProp> { FontEmphasis(Bold) };
        }

        return new List<CellProp>();
    }
}
