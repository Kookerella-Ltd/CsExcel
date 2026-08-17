# CsExcel — Fluent API Guide

CsExcel brings a **declarative, functional** way to generate Excel workbooks to C#: a workbook is an immutable sequence of instructions, not an object you mutate cell by cell. It's a C# wrapper around [**FsExcel**](https://github.com/misterspeedy/FsExcel), an F# library that pioneered this model on top of [ClosedXML](https://github.com/ClosedXML/ClosedXML)'s own, more traditional cell-by-cell object model.

> **All of the design and functionality here is FsExcel's** — CsExcel just exposes it in a form C# code can call directly. This guide covers the same ground as [FsExcel's own README](https://github.com/misterspeedy/FsExcel#readme), translated into C#, but FsExcel's docs are the deeper, authoritative reference: check there first for anything not covered below, for the reasoning behind how a feature works, or for functionality added after this guide was written. (The fluent builder shown here is a CsExcel-specific addition on top of FsExcel's model, not something FsExcel itself has — see [Vanilla.md](Vanilla.md) for the style that maps most directly onto FsExcel's own API.)

CsExcel ships two calling styles over the same underlying model:

- **Fluent** (this document) — a chainable builder, e.g. `Cell().String("x").Bold()`. Reads more like a typical C# builder API.
- **Vanilla** (see [Vanilla.md](Vanilla.md)) — static factory methods and collection literals, e.g. `Cell([String("x"), FontEmphasis(Bold)])`. Closest to FsExcel's own shape.

Pick whichever your team prefers — both produce identical output, and a project can mix them freely.

> **Prefer to see it running first?** [Examples/CsExcel.Examples](../Examples/CsExcel.Examples) is a small console project covering this style, the vanilla style, and building a table from a C# type via reflection. Clone the repo and `dotnet run` it directly — see the [README](../README.md#example-project) for details.

## Setup

```csharp
using FsExcel;
using CsExcel;
using CsExcel.Fluent;
using static CsExcel.PositionFactory;
using static CsExcel.Fluent.ItemFactory;
```

Later examples pull in a few more `using static` directives as needed (`BorderFactory`, `HorizontalAlignmentFactory`, and so on) — each is called out where it's first used.

## The model: cells and a cursor

A workbook is built from a flat `IEnumerable<Item>` — mostly `Cell`s, plus a handful of other item kinds (`Go`, `Style`, `Worksheet`, `MergeCells`, ...). Rendering walks the sequence in order, maintaining an internal cursor: writing a cell moves the cursor one column to the right by default, much like typing a value into Excel and pressing Tab.

`Cell()` and `Style()` return an **immutable** `CellPropsBuilder` you chain calls onto — `.String(...)`, `.Bold()`, `.Border(...)`, and so on. Each call returns a *new* builder rather than changing the one you called it on, consistent with the rest of the API being built from immutable values. That has one concrete consequence: a chained call is an expression, not a mutating statement, so its result must be used —

```csharp
var cell = Cell().String("x");
cell.Bold();          // does nothing useful - the result is discarded, `cell` is unchanged
cell = cell.Bold();   // correct - captures the new, bold builder
```

The upside of immutability is that a partially-built chain can safely be kept in a variable and reused as the shared starting point for several different continuations, without them affecting each other or a later one silently picking up an earlier one's changes. The builder converts to `Item` automatically wherever an `Item` is expected (array elements, `yield return`, method arguments), so there's normally no explicit "build" step. Every other item kind (`Go`, `Worksheet`, `MergeCells`, `AutoFilter`, ...) is a plain static function, since those take simple values rather than a list of props to build up.

> **LINQ query syntax note:** a bare `select Cell()...` inside a `from ... select ...` query infers `IEnumerable<CellPropsBuilder>`, not `IEnumerable<Item>`, since C# doesn't apply implicit conversions inside query-expression type inference. Use `foreach`/`yield return` instead (as every example below does) — the iterator method's declared `IEnumerable<Item>` return type triggers the conversion correctly.

## Hello World

```csharp
Item[] cells = [ Cell().String("Hello World") ];

CsExcel.Render.AsFile(cells, @"c:\temp\HelloWorld.xlsx");
```

`Render.AsFile` writes the sequence to a real `.xlsx` file.

## Placing multiple cells

Since a cell without further instructions just advances the cursor rightward, a sequence of ten `Cell`s fills a single row:

```csharp
IEnumerable<Item> Cells()
{
    foreach (var n in Enumerable.Range(1, 10))
    {
        yield return Cell().Integer(n);
    }
}
CsExcel.Render.AsFile(Cells(), @"c:\temp\MultipleCells.xlsx");
```

## Moving the cursor explicitly

`.Next(position)` on a cell overrides where the cursor goes after that cell is written. Here every month name is followed by moving one row down instead of the default one column right:

```csharp
IEnumerable<Item> Cells()
{
    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        yield return Cell().String(monthName).Next(DownBy(1));
    }
}
CsExcel.Render.AsFile(Cells(), @"c:\temp\VerticalMovement.xlsx");
```

The same "next row" move can also be issued as a standalone `Go` item between cells, rather than baked into the preceding cell — useful once a row has more than one cell in it:

```csharp
IEnumerable<Item> Cells()
{
    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        yield return Cell().String(monthName);
        yield return Cell().Integer(monthName.Length);
        yield return Go(NewRow);
    }
}
CsExcel.Render.AsFile(Cells(), @"c:\temp\Rows.xlsx");
```

## Indentation

`Go(IndentBy(n))` moves the column the cursor returns to whenever `NewRow` fires, relative to where it currently is — useful for indenting a whole block without repeating a column offset on every row:

```csharp
IEnumerable<Item> Items()
{
    yield return Go(IndentBy(2));
    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        yield return Cell().String(monthName);
        yield return Cell().Integer(monthName.Length);
        yield return Go(NewRow);
    }
}
CsExcel.Render.AsFile(Items(), @"c:\temp\Indentation.xlsx");
```

(There's also an absolute `Indent(n)`, which sets the return column outright rather than shifting it — see [Vanilla.md](Vanilla.md#indentation).)

## Borders and font styling

Chained calls control formatting — here a bottom border and bold/italic on the headings, plus a conditional strikethrough on one particular row. `Bold()`, `Italic()`, and `StrikeThrough()` are plain zero-argument calls, so a plain `if` can decide whether to call one, rather than needing to build a conditional prop list up front. Remember the result still needs assigning back (see [above](#the-model-cells-and-a-cursor)), even for a conditional call:

```csharp
using static CsExcel.BorderFactory;
using static ClosedXML.Excel.XLBorderStyleValues;
using static ClosedXML.Excel.XLFontUnderlineValues;

IEnumerable<Item> Items()
{
    foreach (var heading in new[] { "Month", "Letter Count" })
    {
        yield return Cell().String(heading).Border(Bottom(Medium)).Bold().Italic();
    }
    yield return Go(NewRow);

    foreach (var m in Enumerable.Range(1, 12))
    {
        var monthName = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat.GetMonthName(m);
        var monthCell = Cell().String(monthName).Underline(DoubleAccounting);
        if (monthName == "May") monthCell = monthCell.StrikeThrough();
        yield return monthCell;
        yield return Cell().Integer(monthName.Length);
        yield return Go(NewRow);
    }
}
CsExcel.Render.AsFile(Items(), @"c:\temp\BorderAndFontStyling.xlsx");
```

## Reusing a set of style properties

Since the builder is just an object, a shared set of style calls can be factored into a small function and applied to any cell:

```csharp
CellPropsBuilder HeadingStyle(CellPropsBuilder cell) =>
    cell.Border(Bottom(Medium)).Bold().Italic();

foreach (var heading in new[] { "Month", "Letter Count" })
{
    yield return HeadingStyle(Cell()).String(heading);
}
```

## Font name and size

```csharp
foreach (var (fontName, i) in fontNames)
{
    yield return Cell().String(fontName).FontName(fontName).FontSize(10 + (i * 2));
}
```

## Wrap text

```csharp
using static CsExcel.HorizontalAlignmentFactory;
using static CsExcel.SizeFactory;

Item[] items =
[
    Cell().String("Without wrap text:")
        .HorizontalAlignment(Center)
        .VerticalAlignment(VerticalAlignmentFactory.Middle)
        .CellSize(ColWidth(16)),
    Cell().String("The quick brown fox jumps over the lazy dog.")
        .HorizontalAlignment(Center)
        .VerticalAlignment(VerticalAlignmentFactory.Middle)
        .WrapText(true),
];
CsExcel.Render.AsFile(items, @"c:\temp\WrapText.xlsx");
```

## Text rotation

`.TextRotation(degrees)` rotates a cell's content — handy for narrow column headers over a data grid:

```csharp
yield return Cell().String($"Category {category}").TextRotation(45).CellSize(RowHeight(45));
```

## Number formatting and alignment

`.FormatCode(...)` applies an Excel number format string; `.HorizontalAlignment(...)` controls left/center/right alignment per cell:

```csharp
foreach (var (heading, alignment) in new (string, FsExcel.HorizontalAlignment)[]
{
    ("Stock Item", HorizontalAlignmentFactory.Left),
    ("Price", HorizontalAlignmentFactory.Right),
    ("Count", HorizontalAlignmentFactory.Right)
})
{
    yield return Cell().String(heading).HorizontalAlignment(alignment);
}
yield return Go(NewRow);

yield return Cell().String("Apples");
yield return Cell().Float(582.23).FormatCode("$0.00");
yield return Cell().Integer(80).FormatCode("#,##0");
```

## Formulas

`.FormulaA1(...)` writes an Excel formula using standard A1-style cell references:

```csharp
yield return Cell().String("Apples");
yield return Cell().Float(582.23).FormatCode("$0.00");
yield return Cell().Integer(80).FormatCode("#,##0");
yield return Cell().FormulaA1("=B2*C2").FormatCode("$#,##0.00");
```

## Color

`.FontColor(...)`, `.BackgroundColor(...)`, and `.BorderColor(...)` all take a ClosedXML `XLColor`:

```csharp
using static CsExcel.BorderColorFactory;

var backgroundColor = XLColor.FromArgb(0, r, g, b);
yield return Cell().String($"R={r};G={g};B={b}")
    .FontColor(XLColor.FromArgb(0, b, r, g))
    .BackgroundColor(backgroundColor)
    .Border(Top(Thick))
    .BorderColor(BorderColorFactory.Top(XLColor.FromArgb(0, g, b, r)));
```

One quirk worth knowing: ClosedXML refuses to fill a cell with black background if its font color is also black, so the very first cell in an all-zero color sweep stays unstyled.

## Range styles: the "current style"

`Style()` doesn't retroactively style a range — it sets an *ambient* style that applies to every cell written *after* it, until the next `Style()` call changes or clears it (a bare `Style()` with no chained calls clears it). This is the one item type where ordering relative to surrounding `Cell()`s really matters:

```csharp
yield return Style().Border(Bottom(Medium)).Bold().Italic();
foreach (var heading in new[] { "Stock Item", "Price", "Count" })
{
    yield return Cell().String(heading);
}
yield return Go(NewRow);

foreach (var item in new[] { "Apples", "Oranges", "Pears" })
{
    yield return Cell().String(item);        // still bold+italic - the Style hasn't been reset yet
    yield return Style().Italic();
    yield return Cell().Float(582.23).FormatCode("$0.00");
    yield return Style();                    // clear the ambient style
    yield return Go(NewRow);
}
```

## Merging cells and bordering the merge

```csharp
using static CsExcel.CellLabelFactory;
using static CsExcel.StyleMergedCellFactory;

yield return Cell().Integer(1).Name("ID");
yield return Cell().String("Ford Fiesta");
// ... more cells, some given .Name(...) so they can be referenced by name below

yield return MergeCells(ColRowLabel("B", 3), ColRowLabel("B", 6));
yield return MergeCells(NamedCell("ID"), ColRowLabel("A", 6));
yield return BorderMergedCell([
    BorderType(BorderFactory.All(Thin)),
    ColorBorder(BorderColorFactory.All(XLColor.FromArgb(0, 68, 114, 196)))
]);
```

A cell can be addressed either by `ColRowLabel("B", 3)` (column letter + row number) or, if it was given a `.Name(...)`, by `NamedCell("id")`. `BorderMergedCell` applies a border around the outside of the merged region — note that any border a cell had before merging is lost once it's merged.

## Absolute positioning

`Go` also accepts absolute coordinates rather than relative moves:

```csharp
Item[] items = [
    Go(Col(3)),
    Cell().String("Col 3"),
    Go(Row(4)),
    Cell().String("Row 4"),
    Go(RC(6, 5)),
    Cell().String("R6C5"),
];
CsExcel.Render.AsFile(items, @"c:\temp\AbsolutePositioning.xlsx");
```

## Staying in place

`.Next(Stay)` tells a cell not to move the cursor at all after being written — useful when the next instruction is an absolute `Go` anyway:

```csharp
foreach (var i in Enumerable.Range(1, 5))
{
    yield return Cell().Integer(i).Next(Stay);
    yield return Go(DownBy(i));
}
```

## Named cells

`.Name(...)` scopes a name to the current worksheet; `.ScopedName(name, NameScope.Workbook)` makes it visible workbook-wide:

```csharp
Item[] items =
[
    Cell().String("JohnDoe").Name("Username"),
    Cell().String("john.doe@company.com").ScopedName("Email", NameScope.Workbook)
];
CsExcel.Render.AsFile(items, @"c:\temp\NamedCells.xlsx");
```

## Multiple worksheets

`Worksheet("name")` switches the active worksheet — creating it on first use, and re-selecting it (including its own cursor position) if it already exists:

```csharp
yield return Worksheet("English (United Kingdom)");
yield return Cell().String("January");
// ...

yield return Worksheet("українська");
yield return Cell().String("січень");
// ...

yield return Worksheet("English (United Kingdom)"); // switches back, cursor resumes where it left off
```

## Inserting rows above existing content

`Workbook(existingWorkbook)` lets you continue building on top of a workbook you already have (e.g. one loaded from disk, or produced by `.AsWorkBook()` earlier). `InsertRowsAbove(n)` shifts existing rows down — and any formula elsewhere in the workbook that referenced those rows is automatically updated to point at their new location:

```csharp
var workbook = existingItems.AsWorkBook();

IEnumerable<Item> Items()
{
    yield return Workbook(workbook);
    yield return Worksheet("Sheet2");
    yield return Cell().FormulaA1("='Sheet1'!B1*2");
    yield return Worksheet("Sheet1");
    yield return InsertRowsAbove(12); // the formula above now points at B13
    // ... write the 12 new rows
}
Items().AsFile(@"c:\temp\WorksheetsRevised.xlsx");
```

## Column widths and row heights for every cell

`SizeAll` applies a uniform width or height to the whole sheet:

```csharp
yield return SizeAll(ColWidth(5));
yield return SizeAll(RowHeight(20));
```

## Sizing individual cells

`.CellSize(ColWidth(n))` on a specific cell overrides the column width for that cell's column:

```csharp
yield return Cell().String("Car Description").CellSize(ColWidth(49.33));
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

`CsExcel.Table.fromInstance`/`fromIEnumerable` turn a single object, or a sequence of them, into a table — reading each public property via reflection, so any C# class, record, or anonymous type works (not just F# records). The style callback isn't building an `Item`, just a bare `CellProp` list, so it ends with `.ToCellProps()` instead of relying on the usual implicit conversion:

```csharp
record JoiningInfo(string Name, int Age, decimal Fees, string DateJoined);

var records = new[] {
    new JoiningInfo("Jane Smith", 32, 59.25m, "2022-03-12"),
    new JoiningInfo("Michael Nguyễn", 23, 61.2m, "2022-03-13"),
};

IEnumerable<CellProp> CellStyle(int index, string name) =>
    index == 0 ? Cell().Bold().ToCellProps()
    : name == "Fees" ? Cell().FormatCode("$0.00").ToCellProps()
    : [];

var items = CsExcel.Table.fromIEnumerable(records, CsExcel.Table.DirectionFactory.Vertical, CellStyle);
items.Append(AutoFit(AutoFitFactory.All)).AsFile(@"c:\temp\RecordSequenceVertical.xlsx");
```

The style callback receives the column index and property name for each field, so headers or particular columns can be styled differently. `DirectionFactory.Horizontal` lays the same data out with properties as columns and records as rows instead.

## Rendering to a byte array

For scenarios that don't need a file on disk — a web download, an email attachment — `.AsStreamBytes()` returns the workbook as a `byte[]` directly:

```csharp
var bytes = items.AsStreamBytes();
```

`.AsStream(stream)` writes to a `Stream` you already have open, and `.AsWorkBook()` returns the underlying ClosedXML `XLWorkbook` object for further manipulation before saving.

## The full range of supported data types

```csharp
Item[] items = [
    Cell().String("string"),
    Cell().Integer(42),
    Cell().Float(Math.PI),
    Cell().Boolean(false),
    Cell().DateTime(new DateTime(1903, 12, 17)),
    Cell().TimeSpan(new TimeSpan(hours: 1, minutes: 2, seconds: 3)).FormatCode("hh:mm:ss"),
];
```

## Rendering as HTML

`.AsHtml(isHeader)` renders the same `Item` sequence as an HTML table instead of an xlsx file — useful for previewing in a notebook or embedding in a web page. The callback decides which row/column indices should render as `<th>` rather than `<td>`:

```csharp
bool IsHeader(int r, int c) => r == 0 || c == 0;
string html = items.AsHtml(IsHeader);
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

Every example above is a trimmed-down version of a working, tested scenario — see [UnitTests/Fluent.cs](../UnitTests/Fluent.cs) for the full, runnable versions (including the exact assertions that verify each one's output).

For anything beyond this guide — a feature not shown here, the reasoning behind a design choice, or the latest additions — go to the source: [FsExcel on GitHub](https://github.com/misterspeedy/FsExcel).
