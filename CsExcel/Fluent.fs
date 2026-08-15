namespace CsExcel.Fluent

open System.Runtime.CompilerServices
open FsExcel
open System.Linq
open System
open CsExcel
open ClosedXML.Excel

[<Extension>]
type CellPropSeqExtensions =
    [<Extension>]
// need to work out the efficiency here
    static member AddString(this : CellProp seq, s : string) = this.Prepend(CellProp.String s)
    [<Extension>]
    static member AddFloat(this : CellProp seq, f : float) = this.Prepend(CellProp.Float f)
    [<Extension>]
    static member AddInteger(this : CellProp seq, i : int) = this.Prepend(CellProp.Integer i)
    [<Extension>]
    static member AddBoolean(this : CellProp seq, b : bool) = CellProp.Boolean b |> this.Prepend
    [<Extension>]
    static member AddDateTime(this : CellProp seq, dt : DateTime) = CellProp.DateTime dt |> this.Prepend
    [<Extension>]
    static member AddTimeSpan(this : CellProp seq, ts) = CellProp.TimeSpan ts |> this.Prepend
    [<Extension>]
    static member AddFormulaA1(this : CellProp seq, formula : string) = CellProp.FormulaA1 formula |> this.Prepend
    [<Extension>]
    static member AddNext(this : CellProp seq, pos) = CellProp.Next pos |> this.Prepend
    [<Extension>]
    static member AddFontEmphasis(this : CellProp seq, emp) = CellProp.FontEmphasis emp |> this.Prepend
    [<Extension>]
    static member AddFontName(this : CellProp seq, name) = CellProp.FontName name |> this.Prepend
    [<Extension>]
    static member AddFontSize(this : CellProp seq, size) = CellProp.FontSize size |> this.Prepend
    [<Extension>]
    static member AddFontColor(this : CellProp seq, color) = CellProp.FontColor color |> this.Prepend
    [<Extension>]
    static member AddBorder(this : CellProp seq, border) = CellProp.Border border |> this.Prepend
    [<Extension>]
    static member AddBorderColor(this : CellProp seq, color) = CellProp.BorderColor color |> this.Prepend
    [<Extension>]
    static member AddBackgroundColor(this : CellProp seq, color) = CellProp.BackgroundColor color |> this.Prepend
    [<Extension>]
    static member AddHorizontalAlignment(this : CellProp seq, alignment) = CellProp.HorizontalAlignment alignment |> this.Prepend
    [<Extension>]
    static member AddVerticalAlignment(this : CellProp seq, alignment) = CellProp.VerticalAlignment alignment |> this.Prepend
    [<Extension>]
    static member AddTextRotation(this : CellProp seq, rotation) = CellProp.TextRotation rotation |> this.Prepend
    [<Extension>]
    static member AddWrapText(this : CellProp seq, wrap) = CellProp.WrapText wrap |> this.Prepend
    [<Extension>]
    static member AddFormatCode(this : CellProp seq, formatCode) = CellProp.FormatCode formatCode |> this.Prepend
    [<Extension>]
    static member AddName(this : CellProp seq, name) = CellProp.Name name |> this.Prepend
    [<Extension>]
    static member AddScopedName(this : CellProp seq, name) = CellProp.ScopedName name |> this.Prepend
    [<Extension>]
    static member AddCellSize(this : CellProp seq, size) = CellProp.CellSize size |> this.Prepend

[<Extension>]
type RenderExtensions =
    [<Extension>]
    static member AsFile(cells : Item seq,path : string) =
        Render.AsFile(cells,path)
    [<Extension>]
    static member AsHtml(cells : Item seq, isHeader : Render.IsHeader) =
        Render.AsHtml(cells, isHeader)
    [<Extension>]
    static member AsStream(cells : Item seq, stream) =
        Render.AsStream(cells, stream)
    [<Extension>]
    static member AsStreamBytes(cells : Item seq) =
        Render.AsStreamBytes(cells)
    [<Extension>]
    static member AsWorkBook(cells : Item seq) =
        Render.AsWorkBook(cells)

module ItemFactory =
    let Cell(fprops : Func<CellProp seq,CellProp seq>) =
        fprops.Invoke [||]
        |> ItemFactory.Cell
    let Style(fprops : Func<CellProp seq,CellProp seq>) =
        fprops.Invoke [||]
        |> ItemFactory.Style
    let AutoFilter(filters : AutoFilter seq) =
        filters
        |> ItemFactory.AutoFilter
    let BorderMergedCell(borderProps : StyleMergedCell seq) =
        borderProps
        |> ItemFactory.BorderMergedCell
    let Go(position) =
        Item.Go position
    let Worksheet(name) =
        Item.Worksheet name
    let AutoFit(autoFit) =
        Item.AutoFit autoFit
    let Workbook(xlWorkbook : XLWorkbook) =
        Item.Workbook xlWorkbook
    let InsertRowsAbove(rows) =
        Item.InsertRowsAbove rows
    let SizeAll(size) =
        Item.SizeAll size
    let MergeCells(c1,c2) =
        Item.MergeCells(c1,c2)
    let Table(props : TableProperty seq) =
        props
        |> ItemFactory.Table
    let FreezePanes(freezePanes) =
        Item.FreezePanes freezePanes
