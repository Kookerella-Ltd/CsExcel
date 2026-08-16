# CsExcel

CsExcel is a C# wrapper around [**FsExcel**](https://github.com/misterspeedy/FsExcel) — an F# library by [misterspeedy](https://github.com/misterspeedy) for writing Excel workbooks via [ClosedXML](https://github.com/ClosedXML/ClosedXML), using a flat, declarative list of instructions rather than ClosedXML's cell-by-cell object model.

**All credit for the underlying design and implementation belongs to FsExcel.** CsExcel doesn't reimplement any of that — it's a thin C#-facing layer over the real library, so that C# developers who aren't writing F# can still use it directly, with two calling styles to choose from. FsExcel's own [README](https://github.com/misterspeedy/FsExcel#readme) is the authoritative reference for the library's design, and is worth reading directly for anything not covered here — new features, edge cases, and the reasoning behind how the API is shaped will show up there first.

## Documentation

- [Docs/Vanilla.md](Docs/Vanilla.md) — static factory methods and collection literals (`Cell([String("x"), FontEmphasis(Bold)])`), the closest match to FsExcel's own shape.
- [Docs/Fluent.md](Docs/Fluent.md) — a chainable builder (`Cell().String("x").Bold()`), closer to a typical C# builder API.

Both styles produce identical output and can be mixed freely in the same project. Each doc walks through the same set of features FsExcel's own tutorial covers, translated into C#, with runnable tests behind every example (see [UnitTests/](UnitTests/)).

## Attribution

- **FsExcel**: https://github.com/misterspeedy/FsExcel — the library this project wraps.
- **ClosedXML**: https://github.com/ClosedXML/ClosedXML — the underlying Excel file library FsExcel itself builds on.

If something here seems to behave unexpectedly, check FsExcel's own documentation and issue tracker first — most of CsExcel's behavior is simply FsExcel's behavior, passed through.

## License

CsExcel's own code (this wrapper) is licensed under the [PolyForm Noncommercial License 1.0.0](LICENSE.txt) — free to use, modify, and distribute for noncommercial purposes; commercial use requires a separate arrangement with the licensor, [Kookerella Ltd](https://github.com/Kookerella-Ltd/Kookerella-Ltd).

This applies only to CsExcel's own code. FsExcel and ClosedXML, the libraries CsExcel depends on, are each licensed separately under the MIT License by their own authors — CsExcel's license doesn't change or restrict the terms those libraries are already available under as NuGet dependencies.
