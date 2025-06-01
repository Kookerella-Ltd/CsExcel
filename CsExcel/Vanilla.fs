namespace CsExcel

open FsExcel
open System
//open System.Runtime.CompilerServices
//open DiffList

//[<Extension>]
//type CellPropListExtensions =
//    [<Extension>]
//    static member AddString(this : CellProp list,s : string) =
//        CellProp.String s :: this
//    [<Extension>]
//    static member AddInteger(this : CellProp list,i) =
//        CellProp.Integer i :: this

//type PropFactory() = 
//    member _.String(s : string) = CellProp.String s
//    member _.Integer(i : int) = CellProp.Integer i

//type TablePropertyFactory() = 
//    member _.TableName(s) = 
//        TableProperty.TableName s
//    member _.Items(objs) = 
//        objs
//        |> Seq.toList
//        |> TableProperty.Items
//    member _.Theme(xlTableTheme) = 
//        TableProperty.Theme xlTableTheme
//    member _.ShowHeaderRow(show) = 
//        TableProperty.ShowHeaderRow show
//    member _.ShowRowStripes(show) = 
//        TableProperty.ShowRowStripes show
//    member _.ShowColumnStripes(show) =
//        TableProperty.ShowColumnStripes show
//    member _.EmphasizeFirstColumn(emphasize) =
//        TableProperty.EmphasizeFirstColumn emphasize
//    member _.EmphasizeLastColumn(emphasize) =
//        TableProperty.EmphasizeLastColumn emphasize
//    member _.ShowAutoFilter(show) =
//        TableProperty.ShowAutoFilter show
//    member _.Totals(fieldNames,rowItem) =
//        TableProperty.Totals(
//            fieldNames |> Seq.toList, 
//            rowItem)
//    member _.ColFormatCodes(fieldNames,formatCode) =
//        TableProperty.ColFormatCodes(
//            fieldNames |> Seq.toList,
//            formatCode)
//    member _.ColFormula(fieldName,formula) =
//        TableProperty.ColFormula(
//            fieldName,
//            formula)

//type ItemFactory() =
//    //member _.Cell(props : Func<PropFactory,CellProp array> ) = 
//    //    PropFactory()
//    //    |> props.Invoke
//    //    |> Array.toList
//    //    |> Cell
//    //member _.Cell(propsF : Func<CellProp list,CellProp list>) = 
//    //    propsF.Invoke []
//    //    |> Cell
//    member _.Cell(props : CellProp seq) =
//        props
//        |> Seq.toList
//        |> Item.Cell
//    member _.Style(props : CellProp seq) =
//        props
//        |> Seq.toList
//        |> Item.Style
//    member _.AutoFilter(filters : AutoFilter seq) =
//        filters
//        |> Seq.toList
//        |> Item.AutoFilter
//    member _.BorderMergedCell(borderProps : StyleMergedCell seq) =
//        borderProps
//        |> Seq.toList
//        |> Item.BorderMergedCell
//    member _.Go(position) = 
//        Item.Go position
//    member _.Worksheet(name) = 
//        Item.Worksheet name
//    member _.AutoFit(autoFit) = 
//        Item.AutoFit autoFit
//    member _.Workbook(xlWorkbook) = 
//        Item.Workbook xlWorkbook
//    member _.InsertRowsAbove(rows) = 
//        Item.InsertRowsAbove rows
//    member _.SizeAll(size) = 
//        Item.SizeAll size
//    member _.MergeCells(c1,c2) =
//        Item.MergeCells(c1,c2)
//    member _.Table(props : TableProperty seq) =
//        props
//        |> Seq.toList
//        |> Item.Table
//    member _.FreezePanes(freezePanes) = 
//        Item.FreezePanes freezePanes

module FontEmphasisFactory = 
    let Bold = FontEmphasis.Bold
    let Italic = FontEmphasis.Italic
    let Underline = FontEmphasis.Underline
    let StrikeThrough = FontEmphasis.StrikeThrough

module BorderFactory = 
    let All = Border.All
    let Top = Border.Top
    let Right = Border.Right
    let Left = Border.Left
    let Bottom = Border.Bottom

module PositionFactory = 
    let Row(i) = Position.Row i
    let Col(i) = Position.Col i
    let RC(row,col) = Position.RC(row,col)
    let RightBy(i) = Position.RightBy i
    let DownBy(i) = Position.DownBy i
    let LeftBy(i) = Position.LeftBy i
    let UpBy(i) = Position.UpBy i
    let IndentBy(i) = Position.IndentBy i
    let NewRow = Position.NewRow
    let Stay = Position.Stay

module CellPropFactory = 
    let String(s : string) = CellProp.String s
    let Float(f : float) = CellProp.Float f
    let Integer(i : int) = CellProp.Integer i
    let Boolean(b : bool) = CellProp.Boolean b
    let DateTime(dt : DateTime) = CellProp.DateTime dt
    let TimeSpan(ts) = CellProp.TimeSpan ts
    let FormulaA1(formula : string) = CellProp.FormulaA1 formula
    let Next(pos) = CellProp.Next pos
    let FontEmphasis(emp) = CellProp.FontEmphasis emp
    let FontName(name) = CellProp.FontName name
    let FontSize(size) = CellProp.FontSize size
    let FontColor(color) = CellProp.FontColor color
    let Border(border) = CellProp.Border border
    let BorderColor(color) = CellProp.BorderColor color
    let BackgroundColor(color) = CellProp.BackgroundColor color
    let HorizontalAlignment(alignment) = CellProp.HorizontalAlignment alignment
    let VerticalAlignment(alignment) = CellProp.VerticalAlignment alignment
    let TextRotation(rotation) = CellProp.TextRotation rotation
    let WrapText(wrap) = CellProp.WrapText wrap
    let FormatCode(formatCode) = CellProp.FormatCode formatCode
    let Name(name) = CellProp.Name name
    let ScopedName(name) = CellProp.ScopedName name
    let CellSize(size) = CellProp.CellSize size

module ItemFactory =
    let Cell(props : CellProp seq) =
        props
        |> Seq.toList
        |> Item.Cell
    let Style(props : CellProp seq) =
        props
        |> Seq.toList
        |> Item.Style
    let AutoFilter(filters : AutoFilter seq) =
        filters
        |> Seq.toList
        |> Item.AutoFilter
    let BorderMergedCell(borderProps : StyleMergedCell seq) =
        borderProps
        |> Seq.toList
        |> Item.BorderMergedCell
    let Go(position) = 
        Item.Go position
    let Worksheet(name) = 
        Item.Worksheet name
    let AutoFit(autoFit) = 
        Item.AutoFit autoFit
    let Workbook(xlWorkbook) = 
        Item.Workbook xlWorkbook
    let InsertRowsAbove(rows) = 
        Item.InsertRowsAbove rows
    let SizeAll(size) = 
        Item.SizeAll size
    let MergeCells(c1,c2) =
        Item.MergeCells(c1,c2)
    let Table(props : TableProperty seq) =
        props
        |> Seq.toList
        |> Item.Table
    let FreezePanes(freezePanes) = 
        Item.FreezePanes freezePanes

module Render = 
    let AsFile(cells : Item seq,path : string) = 
        cells
        |> Seq.toList
        |> FsExcel.Render.AsFile path
    
//[<Extension>]
//type ItemArrayExtensions =
//    [<Extension>]
//    static member AsFile(this : Func<ItemFactory, Item array>) =
//        this.Invoke (ItemFactory())
//        |> Array.toList
//        |> FsExcel.Render.AsFile """c:\temp\liam.xlsx"""

//[<Extension>]
//type ItemListExtensions =
//    [<Extension>]
//    static member AddCell(this : DiffList<Item>,props : Func<CellProp list,CellProp list>) =
//        snoc this (Cell (props.Invoke [])) 

//module Render =
//    let AsFileFluent : (Func<DiffList<Item>,DiffList<Item>>) -> unit = 
//        fun make ->
//            make.Invoke <| DiffList.empty ()
//            |> DiffList.toList
//            |> FsExcel.Render.AsFile """c:\temp\liam.xlsx"""



