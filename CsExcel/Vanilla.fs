namespace CsExcel

open FsExcel
open System
open ClosedXML.Excel
open FsExcel.Table
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

//type AutoFit =
//    | All
//    | ColRange of int * int
//    | RowRange of int * int
//    | AllCols
//    | AllRows

module AutoFitFactory = 
    let All = AutoFit.All
    let ColRange = AutoFit.ColRange
    let RowRange = AutoFit.RowRange
    let AllCols = AutoFit.AllCols
    let AllRows = AutoFit.AllRows

module BorderColorFactory =
    let Top(color : XLColor) = BorderColor.Top color
    let Right(color : XLColor) = BorderColor.Right color
    let Bottom(color : XLColor) = BorderColor.Bottom color
    let Left(color : XLColor) = BorderColor.Left color
    let All(color : XLColor) = BorderColor.All color

module StyleMergedCellFactory = 
    let BorderType(border : Border) = BorderType border
    let ColorBorder(color : BorderColor) = ColorBorder color


module CellLabelFactory =
    let ColRowLabel(col,row) = CellLabel.ColRowLabel(col,row)
    let NamedCell(name) = CellLabel.NamedCell name
    let SpanDepth(colSpan,rowDepth) = CellLabel.SpanDepth(colSpan,rowDepth)

module HorizontalAlignmentFactory = 
    let Left = HorizontalAlignment.Left
    let Center = HorizontalAlignment.Center
    let Right = HorizontalAlignment.Right

module FontEmphasisFactory = 
    let Bold = FontEmphasis.Bold
    let Italic = FontEmphasis.Italic
    let Underline = FontEmphasis.Underline
    let StrikeThrough = FontEmphasis.StrikeThrough

module SizeFactory = 
    let ColWidth = Size.ColWidth
    let RowHeight = Size.RowHeight

module VerticalAlignmentFactory =
    let Base = VerticalAlignment.Base
    let Middle = VerticalAlignment.Middle
    let TopMost = VerticalAlignment.TopMost

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
    let Indent(i) = Position.Indent i
    let IndentBy(i) = Position.IndentBy i
    let NewRow = Position.NewRow
    let Stay = Position.Stay

module Table =
    open System.Reflection
    open System.Collections.Concurrent

    module DirectionFactory =
        let Vertical = Table.Direction.Vertical
        let Horizontal = Table.Direction.Horizontal

    type CellStyleGetterDelegate = delegate of int * string -> CellProp seq

    // FsExcel's own Table.fromInstance/fromSeq only work for F# record types, because they're
    // implemented with FSharpType.GetRecordFields/FSharpValue.GetRecordField, which throw for any
    // other .NET type - and those helpers are private, so we can't reuse them. Reimplemented here
    // with plain PropertyInfo reflection instead, so it works for C# classes, C# records and
    // anonymous types too (as well as F# records, whose fields are ordinary public properties).
    module private Fields =
        let cache = ConcurrentDictionary<System.Type, PropertyInfo[]>()
        let ofType (t : System.Type) =
            cache.GetOrAdd(t, fun t ->
                t.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)
                |> Array.filter (fun p -> p.CanRead && p.GetIndexParameters().Length = 0))

    module private Cells =
        let toCellProp (value : obj) =
            match value with
            | null -> CellProp.String ""
            | :? string as s -> CellProp.String s
            | :? bool as b -> CellProp.Boolean b
            | :? System.DateTimeOffset as dto ->
                // TODO handle dates explicitly - currently rendered as text, matching FsExcel's own Table module
                CellProp.String (dto.ToString("u"))
            | :? System.DateTime as dt ->
                CellProp.String (dt.ToString("u"))
            | :? int as i -> CellProp.Integer i
            | :? float as f -> CellProp.Float f
            | :? float32 as f -> CellProp.Float (float f)
            | :? decimal as d -> CellProp.Float (float d)
            | value -> CellProp.String (string value)

        let body (getCellStyle : string -> CellProp list) (x : obj) =
            x.GetType()
            |> Fields.ofType
            |> Array.map (fun p -> p.Name, p.GetValue(x))
            |> Array.map (fun (name, value) ->
                let style = getCellStyle name
                let content = toCellProp value
                Cell [ content; yield! style; Next Stay ])
            |> List.ofArray

        let header<'T> (getCellStyle : string -> CellProp list) =
            typeof<'T>
            |> Fields.ofType
            |> Array.map (fun p ->
                let style = getCellStyle p.Name
                Cell [ CellProp.String p.Name; yield! style; Next Stay ])
            |> List.ofArray

    let fromInstance<'T>(x : 'T,direction : Direction,getCellStyle : CellStyleGetterDelegate) : Item seq =
        let getCellStyleF i s = getCellStyle.Invoke(i,s) |> Seq.toList
        let headerCells = Cells.header<'T> (getCellStyleF 0)
        let bodyCells = box x |> Cells.body (getCellStyleF 1)
        match direction with
        | Horizontal ->
            [
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                for bodyCell in bodyCells do
                    bodyCell
                    Go (RightBy 1)
                Go NewRow
            ]
        | Vertical ->
            [
                for heading, value in List.zip headerCells bodyCells do
                    heading
                    Go (RightBy 1)
                    value
                    Go (DownBy 1)
                    Go (LeftBy 1)
            ]
        |> List.toSeq

    type CellStyleGetterSeqDelegate = delegate of int * string -> CellProp seq

    let fromIEnumerable<'T>(xs : 'T seq,direction : Direction,getCellStyle : CellStyleGetterSeqDelegate) : Item seq =
        let getCellStyleF i s = getCellStyle.Invoke(i,s) |> Seq.toList
        let xs = xs |> Array.ofSeq
        let headerCells = Cells.header<'T> (getCellStyleF 0)
        match direction with
        | Vertical ->
            [
                let depth = xs.Length + 1
                for headerCell in headerCells do
                    headerCell
                    Go (DownBy 1)
                Go (UpBy depth)
                Go (RightBy 1)
                for i, x in xs |> Seq.indexed do
                    for bodyCell in box x |> Cells.body (getCellStyleF (i+1)) do
                        bodyCell
                        Go (DownBy 1)
                    Go (UpBy depth)
                    Go (RightBy 1)
                Go (DownBy (depth-1))
                Go NewRow
            ]
        | Horizontal ->
            [
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                for i, x in xs |> Seq.indexed do
                    for bodyCell in box x |> Cells.body (getCellStyleF (i+1)) do
                        bodyCell
                        Go (RightBy 1)
                    Go NewRow
            ]
        |> List.toSeq

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

module AutoFilterRangeFactory =
    let RangeUsed = AutoFilterRange.RangeUsed
    let CurrentRegion(cell : string) = AutoFilterRange.CurrentRegion cell
    let Range(range : string) = AutoFilterRange.Range range

module AutoFilterFactory = 
    let EnableOnly(range : AutoFilterRange) = AutoFilter.EnableOnly range
    let Clear(range : AutoFilterRange) = AutoFilter.Clear range
    let EqualToString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.EqualToString (range, column, value)
    let EqualToInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.EqualToInt (range, column, value)
    let EqualToFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.EqualToFloat (range, column, value)
    let EqualToDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.EqualToDateTime (range, column, value)
    let EqualToBool(range : AutoFilterRange, column : int, value : bool) =
        AutoFilter.EqualToBool (range, column, value)
    let NotEqualToString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.NotEqualToString (range, column, value)
    let NotEqualToInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.NotEqualToInt (range, column, value)
    let NotEqualToFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.NotEqualToFloat (range, column, value)
    let NotEqualToDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.NotEqualToDateTime (range, column, value)
    let NotEqualToBool(range : AutoFilterRange, column : int, value : bool) =
        AutoFilter.NotEqualToBool (range, column, value)
    let BetweenInt(range : AutoFilterRange, column : int, min : int, max : int) = 
        AutoFilter.BetweenInt (range, column, min, max)
    let BetweenFloat(range : AutoFilterRange, column : int, min : float, max : float) = 
        AutoFilter.BetweenFloat (range, column, min, max)
    let BetweenDateTime(range : AutoFilterRange, column : int, min : DateTime, max : DateTime) = 
        AutoFilter.BetweenDateTime (range, column, min, max)
    let NotBetweenInt(range : AutoFilterRange, column : int, min : int, max : int) = 
        AutoFilter.NotBetweenInt (range, column, min, max)
    let NotBetweenFloat(range : AutoFilterRange, column : int, min : float, max : float) = 
        AutoFilter.NotBetweenFloat (range, column, min, max)
    let NotBetweenDateTime(range : AutoFilterRange, column : int, min : DateTime, max : DateTime) = 
        AutoFilter.NotBetweenDateTime (range, column, min, max)
    let ContainsString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.ContainsString (range, column, value)
    let NotContainsString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.NotContainsString (range, column, value)
    let BeginsWithString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.BeginsWithString (range, column, value)
    let NotBeginsWithString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.NotBeginsWithString (range, column, value)
    let EndsWithString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.EndsWithString (range, column, value)
    let NotEndsWithString(range : AutoFilterRange, column : int, value : string) = 
        AutoFilter.NotEndsWithString (range, column, value)
    let Top(range : AutoFilterRange, column : int, value : int, topType : XLTopBottomType) = 
        AutoFilter.Top (range, column, value, topType)
    let Bottom(range : AutoFilterRange, column : int, value : int, bottomType : XLTopBottomType) = 
        AutoFilter.Bottom (range, column, value, bottomType)
    let GreaterThanInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.GreaterThanInt (range, column, value)
    let GreaterThanFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.GreaterThanFloat (range, column, value)
    let GreaterThanDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.GreaterThanDateTime (range, column, value)
    let LessThanInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.LessThanInt (range, column, value)
    let LessThanFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.LessThanFloat (range, column, value)
    let LessThanDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.LessThanDateTime (range, column, value)
    let EqualOrGreaterThanInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.EqualOrGreaterThanInt (range, column, value)
    let EqualOrGreaterThanFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.EqualOrGreaterThanFloat (range, column, value)
    let EqualOrGreaterThanDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.EqualOrGreaterThanDateTime (range, column, value)
    let EqualOrLessThanInt(range : AutoFilterRange, column : int, value : int) = 
        AutoFilter.EqualOrLessThanInt (range, column, value)
    let EqualOrLessThanFloat(range : AutoFilterRange, column : int, value : float) = 
        AutoFilter.EqualOrLessThanFloat (range, column, value)
    let EqualOrLessThanDateTime(range : AutoFilterRange, column : int, value : DateTime) = 
        AutoFilter.EqualOrLessThanDateTime (range, column, value)
    let AboveAverage(range : AutoFilterRange, column : int) = 
        AutoFilter.AboveAverage (range, column)
    let BelowAverage(range : AutoFilterRange, column : int) =
        AutoFilter.BelowAverage (range, column)

module FreezePanesFactory =
    let TopRow = FreezePanes.TopRow
    let FirstColumn = FreezePanes.FirstColumn
    let Panes(row : int, column : int) = FreezePanes.Panes(row, column)
    let UnfreezePanes = FreezePanes.UnfreezePanes

module Render =
    let AsFile(cells : Item seq,path : string) = 
        cells
        |> Seq.toList
        |> FsExcel.Render.AsFile path
    type IsHeader = delegate of int * int -> bool
    let AsHtml(cells : Item seq,isHeader : IsHeader) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsHtml (fun x y -> isHeader.Invoke(x,y));
    let AsStream(cells : Item seq,stream) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsStream stream
    let AsStreamBytes(cells : Item seq) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsStreamBytes
    let AsWorkBook(cells : Item seq) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsWorkBook

    
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



