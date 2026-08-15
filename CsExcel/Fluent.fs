namespace CsExcel.Fluent

open System.Runtime.CompilerServices
open FsExcel
open System
open CsExcel
open ClosedXML.Excel

// A mutable, chainable builder for the props of a single Cell or Style item - the C#-native
// alternative to composing an immutable CellProp seq by hand. `toItem` is which Item case the
// finished prop list becomes: Item.Cell for Cell(), Item.Style for Style(). Method names drop the
// "Add" prefix (plain .String(...) rather than .AddString(...)) and Bold/Italic/StrikeThrough are
// zero-arg calls rather than values imported from FontEmphasisFactory, matching how a C# fluent
// builder is normally shaped.
type CellPropsBuilder(toItem : CellProp list -> Item) =
    let props = ResizeArray<CellProp>()
    member private this.Add(p : CellProp) =
        props.Add p
        this
    member this.String(s : string) = this.Add(CellProp.String s)
    member this.Float(f : float) = this.Add(CellProp.Float f)
    member this.Integer(i : int) = this.Add(CellProp.Integer i)
    member this.Boolean(b : bool) = this.Add(CellProp.Boolean b)
    member this.DateTime(dt : DateTime) = this.Add(CellProp.DateTime dt)
    member this.TimeSpan(ts : TimeSpan) = this.Add(CellProp.TimeSpan ts)
    member this.FormulaA1(formula : string) = this.Add(CellProp.FormulaA1 formula)
    member this.Next(pos : Position) = this.Add(CellProp.Next pos)
    member this.Bold() = this.Add(CellProp.FontEmphasis FontEmphasis.Bold)
    member this.Italic() = this.Add(CellProp.FontEmphasis FontEmphasis.Italic)
    member this.StrikeThrough() = this.Add(CellProp.FontEmphasis FontEmphasis.StrikeThrough)
    member this.Underline(style : XLFontUnderlineValues) = this.Add(CellProp.FontEmphasis(FontEmphasis.Underline style))
    member this.FontName(name : string) = this.Add(CellProp.FontName name)
    member this.FontSize(size : float) = this.Add(CellProp.FontSize size)
    member this.FontColor(color : XLColor) = this.Add(CellProp.FontColor color)
    member this.Border(border : Border) = this.Add(CellProp.Border border)
    member this.BorderColor(color : BorderColor) = this.Add(CellProp.BorderColor color)
    member this.BackgroundColor(color : XLColor) = this.Add(CellProp.BackgroundColor color)
    member this.HorizontalAlignment(alignment : HorizontalAlignment) = this.Add(CellProp.HorizontalAlignment alignment)
    member this.VerticalAlignment(alignment : VerticalAlignment) = this.Add(CellProp.VerticalAlignment alignment)
    member this.TextRotation(rotation : int) = this.Add(CellProp.TextRotation rotation)
    member this.WrapText(wrap : bool) = this.Add(CellProp.WrapText wrap)
    member this.FormatCode(formatCode : string) = this.Add(CellProp.FormatCode formatCode)
    member this.Name(name : string) = this.Add(CellProp.Name name)
    member this.ScopedName(name : string, scope : NameScope) = this.Add(CellProp.ScopedName(name, scope))
    member this.CellSize(size : Size) = this.Add(CellProp.CellSize size)
    member this.ToItem() = props |> List.ofSeq |> toItem
    /// For the rarer case of building a bare CellProp list rather than a full Item - e.g. a
    /// per-column style callback passed to CsExcel.Table.fromInstance/fromIEnumerable.
    member this.ToCellProps() : CellProp seq = upcast props
    /// Lets a builder be used anywhere an Item is expected (array/collection-expression elements,
    /// `yield return`, method arguments) without an explicit terminal call.
    static member op_Implicit(builder : CellPropsBuilder) : Item = builder.ToItem()

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
    let Cell() = CellPropsBuilder(Item.Cell)
    let Style() = CellPropsBuilder(Item.Style)
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
