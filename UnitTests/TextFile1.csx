#r "bin/Debug/net8.0/FsExcel.dll"
#r "bin/Debug/net8.0/CsExcel.dll"

using CsExcel;
using FsExcel;

var cells = new[]
    {
        Cell([ String("Hello World") ])
    };

CsExcel.Render.AsFile(cells, """c:\temp\helloWorld.xlsx""");
