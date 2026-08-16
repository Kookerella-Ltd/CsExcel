// CsExcel example: the same workbook built three ways - Vanilla, Fluent, and straight from a C#
// record via reflection. Run with `dotnet run` from this directory; each workbook is written to
// the current directory and its path is printed. See VanillaExample.cs, FluentExample.cs and
// TableExample.cs for the actual code - this file just runs them.
//
// The Vanilla and Fluent examples live in separate files because each one does
// `using static CsExcel.ItemFactory` / `using static CsExcel.Fluent.ItemFactory` respectively -
// importing both into the same file would make plain calls like `Cell(...)` ambiguous. A real
// project normally commits to one style rather than mixing both.

CsExcel.Examples.VanillaExample.Run();
CsExcel.Examples.FluentExample.Run();
CsExcel.Examples.TableExample.Run();
