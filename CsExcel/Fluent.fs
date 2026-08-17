namespace CsExcel.Fluent

open System.Runtime.CompilerServices
open FsExcel
open System
open CsExcel
open ClosedXML.Excel

/// <summary>
/// An immutable, chainable builder for the props of a single Cell or Style item - the C#-native
/// alternative to composing an immutable CellProp seq by hand, e.g.
/// <c>Cell().String("x").Bold().FontColor(XLColor.Red)</c>.
/// <para>
/// <b>This type is immutable, like <see cref="string"/>: every method returns a new builder
/// instead of changing the one it was called on.</b> A call made without using its result (e.g.
/// <c>cell.Bold();</c> as its own statement) does nothing observable - assign the result back
/// (<c>cell = cell.Bold();</c>) or use it directly. The payoff is that a partially-built chain
/// can safely be kept in a variable and reused as the shared starting point for several different
/// continuations, since none of them can affect each other or the original.
/// </para>
/// Build one with <c>ItemFactory.Cell()</c> or <c>ItemFactory.Style()</c>, chain any of its
/// methods, then either call <c>ToItem()</c> explicitly or pass/yield the builder itself wherever
/// an <c>Item</c> is expected (it converts implicitly).
/// </summary>
type CellPropsBuilder private (toItem : CellProp list -> Item, props : CellProp list) =
    new(toItem : CellProp list -> Item) = CellPropsBuilder(toItem, [])
    member private this.Add(p : CellProp) = CellPropsBuilder(toItem, p :: props)
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
    /// <summary>Finishes the builder, producing the <c>Item</c> (Cell or Style) it was building.</summary>
    member this.ToItem() = props |> List.rev |> toItem
    /// For the rarer case of building a bare CellProp list rather than a full Item - e.g. a
    /// per-column style callback passed to CsExcel.Table.fromInstance/fromIEnumerable.
    member this.ToCellProps() : CellProp seq = props |> List.rev |> List.toSeq
    /// Lets a builder be used anywhere an Item is expected (array/collection-expression elements,
    /// `yield return`, method arguments) without an explicit terminal call.
    static member op_Implicit(builder : CellPropsBuilder) : Item = builder.ToItem()

/// <summary>
/// Extension methods that let an <c>IEnumerable&lt;Item&gt;</c> render itself directly (e.g.
/// <c>items.AsFile("out.xlsx")</c>) instead of calling <c>CsExcel.Render.AsFile(items, path)</c>.
/// Equivalent to the members of <c>CsExcel.Render</c> - see there for details on each.
/// </summary>
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

/// <summary>
/// Items other than a single styled cell - a workbook is still just a flat
/// <c>IEnumerable&lt;Item&gt;</c> (see <c>CsExcel.ItemFactory</c> for the underlying model), built
/// with a mix of <c>ItemFactory.Cell()</c>/<c>Style()</c> builders and these plain items.
/// </summary>
module ItemFactory =
    /// <summary>Starts building a cell at the current cursor position - chain props then call <c>ToItem()</c> or use the builder directly as an <c>Item</c>.</summary>
    let Cell() = CellPropsBuilder(Item.Cell)
    /// <summary>Starts building an ambient style applied to every cell written after this point, until the next <c>Style()</c> changes or clears it.</summary>
    let Style() = CellPropsBuilder(Item.Style)
    let AutoFilter(filters : AutoFilter seq) =
        filters
        |> ItemFactory.AutoFilter
    let BorderMergedCell(borderProps : StyleMergedCell seq) =
        borderProps
        |> ItemFactory.BorderMergedCell
    /// <summary>Moves the cursor without writing a cell - see <c>PositionFactory</c> for the available moves.</summary>
    let Go(position) =
        Item.Go position
    /// <summary>
    /// Creates a new worksheet with the given name if it doesn't already exist, and makes it the
    /// active sheet that subsequent items apply to. If it already exists, switches back to it
    /// instead. This is how worksheets are created and switched between - just include a
    /// <c>Worksheet</c> item in the sequence wherever you want one.
    /// </summary>
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
