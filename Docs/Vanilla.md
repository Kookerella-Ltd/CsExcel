# CsExcel — Vanilla API Guide

CsExcel is a C#-friendly wrapper around [FsExcel](https://github.com/misterspeedy/FsExcel), an F# library for writing Excel workbooks (via [ClosedXML](https://github.com/ClosedXML/ClosedXML)) using a flat, declarative list of instructions instead of ClosedXML's own cell-by-cell object model.

CsExcel ships two calling styles over the same underlying model:

- **Vanilla** (this document) — static factory methods and collection literals, e.g. `Cell([String("x"), FontEmphasis(Bold)])`. Closest to FsExcel's own shape.
- **Fluent** (see [Fluent.md](Fluent.md)) — a chainable builder, e.g. `Cell().String("x").Bold()`. Reads more like a typical C# builder API.

Pick whichever your team prefers — both produce identical output, and a project can mix them freely.

## Setup

```csharp
using FsExcel;
using CsExcel;
using static CsExcel.PositionFactory;
using static CsExcel.ItemFactory;
using static CsExcel.CellPropFactory;
```

Later examples pull in a few more `using static` directives as needed (`BorderFactory`, `FontEmphasisFactory`, and so on) — each is called out where it's first used.

## The model: cells and a cursor

A workbook is built from a flat `IEnumerable<Item>` — mostly `Cell`s, plus a handful of other item kinds (`Go`, `Style`, `Worksheet`, `MergeCells`, ...). Rendering walks the sequence in order, maintaining an internal cursor: writing a cell moves the cursor one column to the right by default, much like typing a value into Excel and pressing Tab. `Item.Cell`, `Item.Go`, and friends are constructed through static factory classes (`ItemFactory`, `PositionFactory`, `CellPropFactory`, ...) rather than `new`, since the underlying types are F# discriminated unions.

## Hello World

```csharp
var cells = new[]
{
    Cell([ String("Hello World") ])
};

CsExcel.Render.AsFile(cells, @"c:\temp\HelloWorld.xlsx");
```

`Cell` takes a list of `CellProp` values — here just the cell's content (`String("Hello World")`). `Render.AsFile` writes the sequence to a real `.xlsx` file.

## Placing multiple cells

Since a cell without further instructions just advances the cursor rightward, a sequence of ten `Cell`s fills a single row:

```csharp
var cells =
    from n in Enumerable.Range(1, 10)
    select Cell([Integer(n)]);
CsExcel.Render.AsFile(cells, @"c:\temp\MultipleCells.xlsx");
```

## Moving the cursor explicitly

`Next(position)` inside a cell's prop list overrides where the cursor goes after that cell is written. Here every month name is followed by moving one row down instead of the default one column right:

```csharp
var cells =
    from m in Enumerable.Range(1, 12)
    let monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m)
    select Cell(
        [
            String(monthName),
            Next(DownBy(1))
        ]
    );
CsExcel.Render.AsFile(cells, @"c:\temp\VerticalMovement.xlsx");
```

The same "next row" move can also be issued as a standalone `Go` item between cells, rather than baked into the preceding cell's props — useful once a row has more than one cell in it:

```csharp
IEnumerable<Item> Cells()
{
    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        yield return Cell([String(monthName)]);
        yield return Cell([Integer(monthName.Length)]);
        yield return Go(PositionFactory.NewRow);
    }
}
CsExcel.Render.AsFile(Cells(), @"c:\temp\Rows.xlsx");
```

## Indentation

`Go(Indent(n))` sets the column the cursor returns to whenever `NewRow` fires — useful for indenting a whole block without repeating a column offset on every row:

```csharp
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
}
CsExcel.Render.AsFile(Cells(), @"c:\temp\Indentation.xlsx");
```

## Borders and font styling

Extra `CellProp` entries in a cell's list control formatting — here a bottom border and bold/italic on the headings, plus a conditional strikethrough on one particular row:

```csharp
using static CsExcel.BorderFactory;
using static CsExcel.FontEmphasisFactory;
using static ClosedXML.Excel.XLBorderStyleValues;
using static ClosedXML.Excel.XLFontUnderlineValues;

IEnumerable<Item> Items()
{
    foreach (var heading in new[] { "Month", "Letter Count" })
    {
        yield return Cell([
            String(heading),
            Border(Bottom(Medium)),
            FontEmphasis(Bold),
            FontEmphasis(Italic)
        ]);
    }
    yield return Go(PositionFactory.NewRow);

    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        IEnumerable<CellProp> props = [
            String(monthName),
            FontEmphasis(Underline(DoubleAccounting))
        ];
        if (monthName == "May") props = [.. props, FontEmphasis(StrikeThrough)];
        yield return Cell(props);
        yield return Cell([Integer(monthName.Length)]);
        yield return Go(PositionFactory.NewRow);
    }
}
CsExcel.Render.AsFile(Items(), @"c:\temp\BorderAndFontStyling.xlsx");
```

## Reusing a set of style properties

Since `CellProp` lists are ordinary collections, a shared style can be built once and spread into every cell that needs it, using C#'s `..` collection-expression spread:

```csharp
CellProp[] headingStyle =
[
    Border(Bottom(Medium)),
    FontEmphasis(Bold),
    FontEmphasis(Italic)
];

foreach (var heading in new[] { "Month", "Letter Count" })
{
    yield return Cell([ String(heading), .. headingStyle ]);
}
```

## Font name and size

```csharp
foreach (var (fontName, i) in fontNames)
{
    yield return Cell([
        String(fontName),
        FontName(fontName),
        FontSize(10 + (i * 2))
    ]);
}
```

## Wrap text

```csharp
using static CsExcel.HorizontalAlignmentFactory;
using static CsExcel.SizeFactory;

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
        CellPropFactory.WrapText(true),
    ]),
];
CsExcel.Render.AsFile(items, @"c:\temp\WrapText.xlsx");
```

## Text rotation

`TextRotation(degrees)` rotates a cell's content — handy for narrow column headers over a data grid:

```csharp
yield return Cell([
    String($"Category {category}"),
    CellPropFactory.TextRotation(45),
    CellSize(RowHeight(45))
]);
```

## Number formatting and alignment

`FormatCode` applies an Excel number format string; `HorizontalAlignment` controls left/center/right alignment per cell:

```csharp
foreach (var (heading, alignment) in new (string, FsExcel.HorizontalAlignment)[]
{
    ("Stock Item", HorizontalAlignmentFactory.Left),
    ("Price", HorizontalAlignmentFactory.Right),
    ("Count", HorizontalAlignmentFactory.Right)
})
{
    yield return Cell([String(heading), HorizontalAlignment(alignment)]);
}
yield return Go(PositionFactory.NewRow);

yield return Cell([String("Apples")]);
yield return Cell([Float(582.23), FormatCode("$0.00")]);
yield return Cell([Integer(80), FormatCode("#,##0")]);
```

## Formulas

`FormulaA1` writes an Excel formula using standard A1-style cell references:

```csharp
yield return Cell([String("Apples")]);
yield return Cell([Float(582.23), FormatCode("$0.00")]);
yield return Cell([Integer(80), FormatCode("#,##0")]);
yield return Cell([FormulaA1("=B2*C2"), FormatCode("$#,##0.00")]);
```

## Color

`FontColor`, `BackgroundColor`, and `BorderColor` all take a ClosedXML `XLColor`:

```csharp
var backgroundColor = XLColor.FromArgb(0, r, g, b);
yield return Cell([
    String($"R={r};G={g};B={b}"),
    FontColor(XLColor.FromArgb(0, b, r, g)),
    BackgroundColor(backgroundColor),
    Border(Top(Thick)),
    BorderColor(BorderColorFactory.Top(XLColor.FromArgb(0, g, b, r)))
]);
```

One quirk worth knowing: ClosedXML refuses to fill a cell with black background if its font color is also black, so the very first cell in an all-zero color sweep stays unstyled.

## Range styles: the "current style"

`Style([...])` doesn't retroactively style a range — it sets an *ambient* style that applies to every cell written *after* it, until the next `Style` call changes or clears it. This is the one item type where ordering relative to surrounding `Cell`s really matters:

```csharp
yield return Style([
    Border(Bottom(Medium)),
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
    yield return Cell([String(item)]);       // still bold+italic - the Style hasn't been reset yet
    yield return Style([FontEmphasis(Italic)]);
    yield return Cell([Float(582.23), FormatCode("$0.00")]);
    yield return Style([]);                  // clear the ambient style
    yield return Go(PositionFactory.NewRow);
}
```

## Merging cells and bordering the merge

```csharp
using static CsExcel.CellLabelFactory;
using static CsExcel.StyleMergedCellFactory;

yield return Cell([Integer(1), Name("ID")]);
yield return Cell([String("Ford Fiesta")]);
// ... more cells, some given Name(...) so they can be referenced by name below

yield return MergeCells(ColRowLabel("B", 3), ColRowLabel("B", 6));
yield return MergeCells(NamedCell("ID"), ColRowLabel("A", 6));
yield return BorderMergedCell([
    BorderType(BorderFactory.All(Thin)),
    ColorBorder(BorderColorFactory.All(XLColor.FromArgb(0, 68, 114, 196)))
]);
```

A cell can be addressed either by `ColRowLabel("B", 3)` (column letter + row number) or, if it was given a `Name(...)`, by `NamedCell("id")`. `BorderMergedCell` applies a border around the outside of the merged region — note that any border a cell had before merging is lost once it's merged.

## Absolute positioning

`Go` also accepts absolute coordinates rather than relative moves:

```csharp
var items = new[] {
    Go(Col(3)),
    Cell([String("Col 3")]),
    Go(Row(4)),
    Cell([String("Row 4")]),
    Go(RC(6, 5)),
    Cell([String("R6C5")]),
};
CsExcel.Render.AsFile(items, @"c:\temp\AbsolutePositioning.xlsx");
```

## Staying in place

`Next(Stay)` tells a cell not to move the cursor at all after being written — useful when the next instruction is an absolute `Go` anyway:

```csharp
foreach (var i in Enumerable.Range(1, 5))
{
    yield return Cell([Integer(i), Next(PositionFactory.Stay)]);
    yield return Go(DownBy(i));
}
```

## Named cells

`Name(...)` scopes a name to the current worksheet; `ScopedName(name, NameScope.Workbook)` makes it visible workbook-wide:

```csharp
var items = new[]
{
    Cell([String("JohnDoe"), Name("Username")]),
    Cell([String("john.doe@company.com"), CellPropFactory.ScopedName("Email", NameScope.Workbook)])
};
CsExcel.Render.AsFile(items, @"c:\temp\NamedCells.xlsx");
```

## Multiple worksheets

`Worksheet("name")` switches the active worksheet — creating it on first use, and re-selecting it (including its own cursor position) if it already exists:

```csharp
yield return Worksheet("English (United Kingdom)");
yield return Cell([String("January")]);
// ...

yield return Worksheet("українська");
yield return Cell([String("січень")]);
// ...

yield return Worksheet("English (United Kingdom)"); // switches back, cursor resumes where it left off
```

## Inserting rows above existing content

`Workbook(existingWorkbook)` lets you continue building on top of a workbook you already have (e.g. one loaded from disk, or produced by `Render.AsWorkBook` earlier). `InsertRowsAbove(n)` shifts existing rows down — and any formula elsewhere in the workbook that referenced those rows is automatically updated to point at their new location:

```csharp
var workbook = CsExcel.Render.AsWorkBook(existingItems);

IEnumerable<Item> Items()
{
    yield return Workbook(workbook);
    yield return Worksheet("Sheet2");
    yield return Cell([FormulaA1("='Sheet1'!B1*2")]);
    yield return Worksheet("Sheet1");
    yield return InsertRowsAbove(12); // the formula above now points at B13
    // ... write the 12 new rows
}
CsExcel.Render.AsFile(Items(), @"c:\temp\WorksheetsRevised.xlsx");
```

## Column widths and row heights for every cell

`SizeAll` applies a uniform width or height to the whole sheet:

```csharp
yield return SizeAll(ColWidth(5));
yield return SizeAll(RowHeight(20));
```

## Sizing individual cells

`CellSize(ColWidth(n))` on a specific cell overrides the column width for that cell's column:

```csharp
yield return Cell([
    String("Car Description"),
    CellSize(ColWidth(49.33))
]);
```

## Autofitting columns

`AutoFit(AutoFitFactory.AllCols)` sizes every column to its content automatically. On non-Windows platforms ClosedXML needs an explicit font to measure with:

```csharp
if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
{
    LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Liberation Sans");
}
// ... write cells ...
yield return AutoFit(AutoFitFactory.AllCols);
```

## Building tables from C# records or POCOs

`CsExcel.Table.fromInstance`/`fromIEnumerable` turn a single object, or a sequence of them, into a table — reading each public property via reflection, so any C# class, record, or anonymous type works (not just F# records):

```csharp
record JoiningInfo(string Name, int Age, decimal Fees, string DateJoined);

var records = new[] {
    new JoiningInfo("Jane Smith", 32, 59.25m, "2022-03-12"),
    new JoiningInfo("Michael Nguyễn", 23, 61.2m, "2022-03-13"),
};

CellProp[] CellStyle(int index, string name) =>
    index == 0 ? [FontEmphasis(Bold)]
    : name == "Fees" ? [FormatCode("$0.00")]
    : [];

var items = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Vertical, CellStyle);
CsExcel.Render.AsFile([.. items, AutoFit(AutoFitFactory.All)], @"c:\temp\RecordSequenceVertical.xlsx");
```

The style callback receives the column index and property name for each field, so headers or particular columns can be styled differently. `DirectionFactory.Horizontal` lays the same data out with properties as columns and records as rows instead.

## Rendering to a byte array

For scenarios that don't need a file on disk — a web download, an email attachment — `Render.AsStreamBytes` returns the workbook as a `byte[]` directly:

```csharp
var bytes = CsExcel.Render.AsStreamBytes(items);
```

`Render.AsStream` writes to a `Stream` you already have open, and `Render.AsWorkBook` returns the underlying ClosedXML `XLWorkbook` object for further manipulation before saving.

## The full range of supported data types

```csharp
var items = new Item[] {
    Cell([String("string")]),
    Cell([Integer(42)]),
    Cell([Float(Math.PI)]),
    Cell([Boolean(false)]),
    Cell([DateTime(new DateTime(1903, 12, 17))]),
    Cell([TimeSpan(new TimeSpan(hours: 1, minutes: 2, seconds: 3)), FormatCode("hh:mm:ss")]),
};
```

## Rendering as HTML

`Render.AsHtml` renders the same `Item` sequence as an HTML table instead of an xlsx file — useful for previewing in a notebook or embedding in a web page. The `isHeader` callback decides which row/column indices should render as `<th>` rather than `<td>`:

```csharp
bool IsHeader(int r, int c) => r == 0 || c == 0;
string html = CsExcel.Render.AsHtml(items, IsHeader);
```

## AutoFilter: a single condition

```csharp
using static CsExcel.AutoFilterFactory;
using static CsExcel.AutoFilterRangeFactory;

IEnumerable<Item> items = [
    .. headings, .. rows,
    AutoFit(AutoFitFactory.All),
    AutoFilter([EnableOnly(RangeUsed)])
];
```

## AutoFilter: compound conditions

Multiple conditions can be combined; `AutoFilterFactory` has one function per comparison kind (`GreaterThanInt`, `EqualToBool`, `ContainsString`, ...):

```csharp
IEnumerable<Item> items = [
    .. headings, .. rows,
    AutoFit(AutoFitFactory.All),
    AutoFilter([
        GreaterThanInt(RangeUsed, 2, 3),
        EqualToBool(RangeUsed, 5, true)
    ])
];
```

## Freeze panes

```csharp
using static CsExcel.FreezePanesFactory;

yield return FreezePanes(Panes(1, 1));  // freeze 1 row and 1 column
yield return FreezePanes(TopRow);       // freeze just the top row
yield return FreezePanes(FirstColumn);  // freeze just the first column
yield return FreezePanes(UnfreezePanes);
```

---

Every example above is a trimmed-down version of a working, tested scenario — see [UnitTests/Vanilla.cs](../UnitTests/Vanilla.cs) for the full, runnable versions (including the exact assertions that verify each one's output).
