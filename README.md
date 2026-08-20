# CsExcel

CsExcel brings a **declarative, functional** way to generate Excel workbooks to C#: you build a workbook as an immutable sequence of instructions — not an object you mutate cell by cell — then hand that sequence to a renderer. It's a C# wrapper around [**FsExcel**](https://github.com/misterspeedy/FsExcel) — an F# library by [misterspeedy](https://github.com/misterspeedy) that pioneered this declarative model on top of [ClosedXML](https://github.com/ClosedXML/ClosedXML), which itself exposes a traditional, mutable cell-by-cell object model underneath.

**All credit for the underlying design and implementation belongs to FsExcel.** CsExcel doesn't reimplement any of that — it's a thin C#-facing layer over the real library, so that C# developers who aren't writing F# can still use it directly, with two calling styles to choose from. FsExcel's own [README](https://github.com/misterspeedy/FsExcel#readme) is the authoritative reference for the library's design, and is worth reading directly for anything not covered here — new features, edge cases, and the reasoning behind how the API is shaped will show up there first.

## Install

```bash
dotnet add package CsExcel
```

## Quick start

A workbook is an immutable sequence of `Item`s — not an object you call `.Save()` on after mutating it — handed to a renderer. The same output can be built either declaratively (**Vanilla**) or with a chainable builder (**Fluent**):

```csharp
// Vanilla — static factories and collection literals
using CsExcel;
using static CsExcel.ItemFactory;
using static CsExcel.CellPropFactory;

var cells = new[]
{
    Cell([ String("Hello, Excel!"), FontEmphasis(Bold) ])
};
CsExcel.Render.AsFile(cells, @"c:\temp\HelloWorld.xlsx");
```

```csharp
// Fluent — chainable, immutable builder
using CsExcel;
using CsExcel.Fluent;
using static CsExcel.Fluent.ItemFactory;

Item[] cells = [ Cell().String("Hello, Excel!").Bold() ];
CsExcel.Render.AsFile(cells, @"c:\temp\HelloWorld.xlsx");
```

Both produce the same `.xlsx` file. See [Docs/Vanilla.md](https://github.com/Kookerella-Ltd/CsExcel/blob/master/Docs/Vanilla.md) and [Docs/Fluent.md](https://github.com/Kookerella-Ltd/CsExcel/blob/master/Docs/Fluent.md) for the full guide — rows, columns, styles, borders, merged cells, and building tables from C# records via reflection.

## Documentation

- [Docs/Vanilla.md](https://github.com/Kookerella-Ltd/CsExcel/blob/master/Docs/Vanilla.md) — static factory methods and collection literals (`Cell([String("x"), FontEmphasis(Bold)])`), the closest match to FsExcel's own shape.
- [Docs/Fluent.md](https://github.com/Kookerella-Ltd/CsExcel/blob/master/Docs/Fluent.md) — a chainable builder (`Cell().String("x").Bold()`), closer to a typical C# builder API.

Both styles produce identical output and can be mixed freely in the same project. Each doc walks through the same set of features FsExcel's own tutorial covers, translated into C#, with runnable tests behind every example (see [UnitTests/](https://github.com/Kookerella-Ltd/CsExcel/tree/master/UnitTests)).

## Example project

[Examples/CsExcel.Examples](https://github.com/Kookerella-Ltd/CsExcel/tree/master/Examples/CsExcel.Examples) is a small runnable console project showing both calling styles side by side, plus building a table straight from a C# record via reflection. Clone the repo and run it directly:

```bash
git clone https://github.com/Kookerella-Ltd/CsExcel.git
cd CsExcel/Examples/CsExcel.Examples
dotnet run
```

It writes three `.xlsx` files to its output folder and prints their paths.

## Attribution

- **FsExcel**: https://github.com/misterspeedy/FsExcel — the library this project wraps.
- **ClosedXML**: https://github.com/ClosedXML/ClosedXML — the underlying Excel file library FsExcel itself builds on.

If something here seems to behave unexpectedly, check FsExcel's own documentation and issue tracker first — most of CsExcel's behavior is simply FsExcel's behavior, passed through.

## License

CsExcel's own code (this wrapper) is licensed under the [PolyForm Noncommercial License 1.0.0](https://github.com/Kookerella-Ltd/CsExcel/blob/master/LICENSE.txt) — free to use, modify, and distribute for noncommercial purposes; commercial use requires a separate arrangement with the licensor, [Kookerella Ltd](https://github.com/Kookerella-Ltd/Kookerella-Ltd).

This applies only to CsExcel's own code. FsExcel and ClosedXML, the libraries CsExcel depends on, are each licensed separately under the MIT License by their own authors — CsExcel's license doesn't change or restrict the terms those libraries are already available under as NuGet dependencies.
