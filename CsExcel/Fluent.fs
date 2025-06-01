namespace CsExcel.Fluent

open System.Runtime.CompilerServices
open FsExcel
open System.Linq
open System
open CsExcel

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


//    [<Extension>]
//    static member AddFloat(this : CellProp list, f : float) =
//        CellProp.Float f :: this

//    [<Extension>]
//    static member AddInteger(this : CellProp list, i : int) =
//        CellProp.Integer i :: this

//    [<Extension>]
//    static member AddBoolean(this : CellProp list, b : bool) =
//        CellProp.Boolean b :: this

//    [<Extension>]
//    static member AddDateTime(this : CellProp list, dt : DateTime) =
//        CellProp.DateTime dt :: this

//    [<Extension>]
//    static member AddTimeSpan(this : CellProp list, ts) =
//        CellProp.TimeSpan ts :: this
//module CellPropFactory = 
//    let String(s : string) = CellProp.String s
//    let Float(f : float) = CellProp.Float f
//    let Integer(i : int) = CellProp.Integer i
//    let Boolean(b : bool) = CellProp.Boolean b
//    let DateTime(dt : DateTime) = CellProp.DateTime dt
//    let TimeSpan(ts) = CellProp.TimeSpan ts
//    let FormulaA1(formula : string) = CellProp.FormulaA1 formula
//    let Next(pos) = CellProp.Next pos
//    let FontEmphasis(emp) = CellProp.FontEmphasis emp
//    let FontName(name) = CellProp.FontName name
//    let FontSize(size) = CellProp.FontSize size
//    let FontColor(color) = CellProp.FontColor color
//    let Border(border) = CellProp.Border border
//    let BorderColor(color) = CellProp.BorderColor color
//    let BackgroundColor(color) = CellProp.BackgroundColor color
//    let HorizontalAlignment(alignment) = CellProp.HorizontalAlignment alignment
//    let VerticalAlignment(alignment) = CellProp.VerticalAlignment alignment
//    let TextRotation(rotation) = CellProp.TextRotation rotation
//    let WrapText(wrap) = CellProp.WrapText wrap
//    let FormatCode(formatCode) = CellProp.FormatCode formatCode
//    let Name(name) = CellProp.Name name
//    let ScopedName(name) = CellProp.ScopedName name
//    let CellSize(size) = CellProp.CellSize size


module ItemFactory =
    let Cell(fprops : Func<CellProp seq,CellProp seq>) =
        fprops.Invoke [||]
        |> ItemFactory.Cell
    //let Style(props : CellProp seq) =
    //    props
    //    |> Seq.toList
    //    |> Item.Style
    //let AutoFilter(filters : AutoFilter seq) =
    //    filters
    //    |> Seq.toList
    //    |> Item.AutoFilter
    //let BorderMergedCell(borderProps : StyleMergedCell seq) =
    //    borderProps
    //    |> Seq.toList
    //    |> Item.BorderMergedCell
    let Go(position) = 
        Item.Go position
    //let Worksheet(name) = 
    //    Item.Worksheet name
    //let AutoFit(autoFit) = 
    //    Item.AutoFit autoFit
    //let Workbook(xlWorkbook) = 
    //    Item.Workbook xlWorkbook
    //let InsertRowsAbove(rows) = 
    //    Item.InsertRowsAbove rows
    //let SizeAll(size) = 
    //    Item.SizeAll size
    //let MergeCells(c1,c2) =
    //    Item.MergeCells(c1,c2)
    //let Table(props : TableProperty seq) =
    //    props
    //    |> Seq.toList
    //    |> Item.Table
    //let FreezePanes(freezePanes) = 
    //    Item.FreezePanes freezePanes


