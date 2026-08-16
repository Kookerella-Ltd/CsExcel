namespace CsExcel

open FsExcel
open System
open ClosedXML.Excel
open FsExcel.Table

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

/// <summary>
/// Cursor movements, used with <c>ItemFactory.Go</c> to move without writing a cell, or with
/// <c>CellPropFactory.Next</c> to override where the cursor goes after a specific cell.
/// </summary>
module PositionFactory =
    /// <summary>Moves to an absolute row, keeping the current column.</summary>
    let Row(i) = Position.Row i
    /// <summary>Moves to an absolute column, keeping the current row.</summary>
    let Col(i) = Position.Col i
    /// <summary>Moves to an absolute row and column.</summary>
    let RC(row,col) = Position.RC(row,col)
    /// <summary>Moves a relative number of columns to the right.</summary>
    let RightBy(i) = Position.RightBy i
    /// <summary>Moves a relative number of rows down.</summary>
    let DownBy(i) = Position.DownBy i
    /// <summary>Moves a relative number of columns to the left.</summary>
    let LeftBy(i) = Position.LeftBy i
    /// <summary>Moves a relative number of rows up.</summary>
    let UpBy(i) = Position.UpBy i
    /// <summary>
    /// Sets the column that <c>NewRow</c> returns to, as an absolute column number (not
    /// relative to the current position) - see <c>IndentBy</c> for a relative version.
    /// </summary>
    let Indent(i) = Position.Indent i
    /// <summary>
    /// Shifts the column that <c>NewRow</c> returns to by a relative amount - see
    /// <c>Indent</c> for an absolute version.
    /// </summary>
    let IndentBy(i) = Position.IndentBy i
    /// <summary>Moves down one row and back to the current indent column (see <c>Indent</c>/<c>IndentBy</c>).</summary>
    let NewRow = Position.NewRow
    /// <summary>
    /// Leaves the cursor exactly where it is after the current cell - useful when the next
    /// instruction is an absolute <c>Go</c> anyway, or to place several props-only cells (like
    /// a <c>Name</c>) at the same position.
    /// </summary>
    let Stay = Position.Stay

/// <summary>
/// Builds tables directly from C# objects' public properties, via reflection - this works for
/// any C# class, record, or anonymous type (not just F# records, which is all FsExcel's own
/// <c>Table.fromInstance</c>/<c>fromSeq</c> support).
/// </summary>
module Table =
    open System.Reflection
    open System.Collections.Concurrent

    module DirectionFactory =
        /// <summary>Lays a table out with headers down the first column and one further column per record.</summary>
        let Vertical = Table.Direction.Vertical
        /// <summary>Lays a table out with headers across the first row and one further row per record.</summary>
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
                // Type.GetProperties only returns members declared directly on an interface, not
                // ones inherited from a base interface it extends (unlike classes, where inherited
                // members are included automatically) - so for interfaces we also have to walk
                // GetInterfaces() ourselves and union in their properties too.
                let declared = t.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)
                let inherited =
                    if t.IsInterface then
                        t.GetInterfaces()
                        |> Array.collect (fun i -> i.GetProperties(BindingFlags.Public ||| BindingFlags.Instance))
                    else
                        [||]
                Array.append declared inherited
                |> Array.distinctBy (fun p -> p.Name)
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

        // fields is always derived from the declared type 'T, both here and in `header` below - never
        // from the runtime type of an individual instance - so a body value and its header always
        // line up, even when 'T is a base type/interface and some instances are a derived subtype.
        let body (fields : PropertyInfo[]) (getCellStyle : string -> CellProp list) (x : obj) =
            fields
            |> Array.map (fun p -> p.Name, p.GetValue(x))
            |> Array.map (fun (name, value) ->
                let style = getCellStyle name
                let content = toCellProp value
                Cell [ content; yield! style; Next Stay ])
            |> List.ofArray

        let header (fields : PropertyInfo[]) (getCellStyle : string -> CellProp list) =
            fields
            |> Array.map (fun p ->
                let style = getCellStyle p.Name
                Cell [ CellProp.String p.Name; yield! style; Next Stay ])
            |> List.ofArray

    /// <summary>
    /// Builds table cells from a single object's public properties. <c>getCellStyle(index,
    /// name)</c> is called once per field (index 0 for the header row/column, 1 for the value)
    /// and returns any extra <c>CellProp</c> styling/formatting to apply to that cell.
    /// </summary>
    let fromInstance<'T>(x : 'T,direction : Direction,getCellStyle : CellStyleGetterDelegate) : Item seq =
        let getCellStyleF i s = getCellStyle.Invoke(i,s) |> Seq.toList
        let fields = Fields.ofType typeof<'T>
        let headerCells = Cells.header fields (getCellStyleF 0)
        let bodyCells = box x |> Cells.body fields (getCellStyleF 1)
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

    /// <summary>
    /// Builds table cells from a sequence of objects' public properties - one column (Vertical)
    /// or row (Horizontal) per record. <c>getCellStyle(index, name)</c> is called once per field
    /// per record (index 0 for the header, 1+ for each record in order) and returns any extra
    /// <c>CellProp</c> styling/formatting to apply to that cell.
    /// </summary>
    let fromIEnumerable<'T>(xs : 'T seq,direction : Direction,getCellStyle : CellStyleGetterSeqDelegate) : Item seq =
        let getCellStyleF i s = getCellStyle.Invoke(i,s) |> Seq.toList
        let xs = xs |> Array.ofSeq
        let fields = Fields.ofType typeof<'T>
        let headerCells = Cells.header fields (getCellStyleF 0)
        match direction with
        | Vertical ->
            [
                // depth is how far each column has to travel back up to row 1 before moving right to
                // the next column - that's the number of fields per column, not the number of records.
                let depth = headerCells.Length
                for headerCell in headerCells do
                    headerCell
                    Go (DownBy 1)
                Go (UpBy depth)
                Go (RightBy 1)
                for i, x in xs |> Seq.indexed do
                    for bodyCell in box x |> Cells.body fields (getCellStyleF (i+1)) do
                        bodyCell
                        Go (DownBy 1)
                    Go (UpBy depth)
                    Go (RightBy 1)
                // depth-1 is normally >= 0, but guard the degenerate case of a type with no fields
                // (depth = 0), where DownBy(-1) would otherwise move the cursor up instead of down.
                Go (DownBy (max 0 (depth-1)))
                Go NewRow
            ]
        | Horizontal ->
            [
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                for i, x in xs |> Seq.indexed do
                    for bodyCell in box x |> Cells.body fields (getCellStyleF (i+1)) do
                        bodyCell
                        Go (RightBy 1)
                    Go NewRow
            ]
        |> List.toSeq

/// <summary>
/// Individual cell properties - a cell's content (<c>String</c>/<c>Integer</c>/<c>Float</c>/...),
/// formatting (<c>FontEmphasis</c>/<c>Border</c>/<c>BackgroundColor</c>/...), or cursor override
/// (<c>Next</c>). Pass a list of these to <c>ItemFactory.Cell</c> or <c>ItemFactory.Style</c>.
/// </summary>
module CellPropFactory =
    let String(s : string) = CellProp.String s
    let Float(f : float) = CellProp.Float f
    let Integer(i : int) = CellProp.Integer i
    let Boolean(b : bool) = CellProp.Boolean b
    let DateTime(dt : DateTime) = CellProp.DateTime dt
    let TimeSpan(ts) = CellProp.TimeSpan ts
    /// <summary>Writes an Excel formula, given in standard A1-reference syntax, e.g. <c>"=B2*C2"</c>.</summary>
    let FormulaA1(formula : string) = CellProp.FormulaA1 formula
    /// <summary>Overrides where the cursor moves after this cell - see <c>PositionFactory</c>.</summary>
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
    /// <summary>Rotates the cell's content by the given number of degrees.</summary>
    let TextRotation(rotation) = CellProp.TextRotation rotation
    let WrapText(wrap) = CellProp.WrapText wrap
    /// <summary>Applies an Excel number format string, e.g. <c>"$0.00"</c> or <c>"hh:mm:ss"</c>.</summary>
    let FormatCode(formatCode) = CellProp.FormatCode formatCode
    /// <summary>Names this cell, scoped to the current worksheet - see <c>CellLabelFactory.NamedCell</c> to address it later.</summary>
    let Name(name) = CellProp.Name name
    /// <summary>Names this cell with an explicit <c>NameScope</c> (e.g. <c>NameScope.Workbook</c> to make it visible workbook-wide, not just on the current sheet).</summary>
    let ScopedName(name) = CellProp.ScopedName name
    /// <summary>Sets this cell's column width or row height - see <c>SizeFactory</c>.</summary>
    let CellSize(size) = CellProp.CellSize size

/// <summary>
/// The core building blocks for a CsExcel workbook. A workbook is built by rendering a flat
/// sequence of <c>Item</c> values (see <c>CsExcel.Render</c>) - there is no "workbook" object to
/// create or mutate directly, and no separate call to add a worksheet: everything, including
/// switching worksheets, is an item in the same sequence. Rendering walks the sequence in order,
/// maintaining an internal cursor: writing a cell moves the cursor one column to the right by
/// default, similar to typing a value into Excel and pressing Tab.
/// </summary>
module ItemFactory =
    /// <summary>
    /// Writes a single cell at the current cursor position, described by a list of
    /// <c>CellProp</c> values (its content, plus optional styling/formatting/positioning props
    /// from <c>CellPropFactory</c>). After the cell is written, the cursor moves one column to
    /// the right, unless overridden by a <c>CellPropFactory.Next</c> prop in the list.
    /// </summary>
    let Cell(props : CellProp seq) =
        props
        |> Seq.toList
        |> Item.Cell
    /// <summary>
    /// Sets an ambient style applied to every <c>Cell</c> written after this point, until the
    /// next <c>Style</c> call changes or clears it (pass an empty list to clear it). This does
    /// NOT retroactively style cells already written before this call.
    /// </summary>
    let Style(props : CellProp seq) =
        props
        |> Seq.toList
        |> Item.Style
    /// <summary>Applies one or more AutoFilter conditions (see <c>AutoFilterFactory</c>) to the active worksheet.</summary>
    let AutoFilter(filters : AutoFilter seq) =
        filters
        |> Seq.toList
        |> Item.AutoFilter
    /// <summary>Applies a border around the outside of a previously merged cell region (see <c>MergeCells</c>).</summary>
    let BorderMergedCell(borderProps : StyleMergedCell seq) =
        borderProps
        |> Seq.toList
        |> Item.BorderMergedCell
    /// <summary>Moves the cursor without writing a cell - see <c>PositionFactory</c> for the available moves.</summary>
    let Go(position) =
        Item.Go position
    /// <summary>
    /// Creates a new worksheet with the given name if it doesn't already exist, and makes it the
    /// active sheet that subsequent items apply to. If a worksheet with this name already
    /// exists, switches back to it instead, resuming from wherever its cursor was left. This is
    /// how worksheets are created and switched between - there's no separate "add worksheet"
    /// call, just include a <c>Worksheet</c> item in the sequence wherever you want one.
    /// </summary>
    let Worksheet(name) =
        Item.Worksheet name
    /// <summary>Auto-sizes columns/rows to fit their content - see <c>AutoFitFactory</c> for which cells/range.</summary>
    let AutoFit(autoFit) =
        Item.AutoFit autoFit
    /// <summary>
    /// Continues building on top of an existing <c>XLWorkbook</c> (e.g. one returned by
    /// <c>Render.AsWorkBook</c>, or loaded from disk) instead of starting a new one.
    /// </summary>
    let Workbook(xlWorkbook) =
        Item.Workbook xlWorkbook
    /// <summary>
    /// Inserts blank rows above the current cursor row on the active worksheet, shifting
    /// existing rows down. Formulas elsewhere in the workbook that reference the shifted rows
    /// are automatically updated to point at their new location.
    /// </summary>
    let InsertRowsAbove(rows) =
        Item.InsertRowsAbove rows
    /// <summary>Applies a uniform column width or row height to every cell on the active worksheet - see <c>SizeFactory</c>.</summary>
    let SizeAll(size) =
        Item.SizeAll size
    /// <summary>
    /// Merges the cells spanning from <c>c1</c> to <c>c2</c> - see <c>CellLabelFactory</c> for
    /// how to address them, by column/row label or by a previously-assigned <c>Name</c>.
    /// </summary>
    let MergeCells(c1,c2) =
        Item.MergeCells(c1,c2)
    /// <summary>Formats the current range as a native Excel Table (ListObject).</summary>
    let Table(props : TableProperty seq) =
        props
        |> Seq.toList
        |> Item.Table
    /// <summary>Freezes or unfreezes panes on the active worksheet - see <c>FreezePanesFactory</c>.</summary>
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

/// <summary>
/// Turns a sequence of <c>Item</c> values into an actual Excel file, byte array, stream, or
/// in-memory <c>XLWorkbook</c>. This is the terminal step of every CsExcel program: build up an
/// <c>IEnumerable&lt;Item&gt;</c> (mostly via <c>ItemFactory.Cell</c>), then call one of these to
/// produce output.
/// </summary>
module Render =
    /// <summary>Renders the items and saves them directly to a file on disk at the given path.</summary>
    let AsFile(cells : Item seq,path : string) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsFile path
    type IsHeader = delegate of int * int -> bool
    /// <summary>
    /// Renders the items as an HTML table (one per worksheet) instead of an xlsx file - useful
    /// for previewing in a notebook or embedding in a web page. <c>isHeader(row, col)</c>
    /// decides which cells render as <c>&lt;th&gt;</c> instead of <c>&lt;td&gt;</c>.
    /// </summary>
    let AsHtml(cells : Item seq,isHeader : IsHeader) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsHtml (fun x y -> isHeader.Invoke(x,y));
    /// <summary>Renders the items and writes them to an already-open <c>Stream</c>.</summary>
    let AsStream(cells : Item seq,stream) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsStream stream
    /// <summary>
    /// Renders the items and returns the resulting xlsx file as a byte array - useful for
    /// scenarios that don't need a file on disk, e.g. a web download or email attachment.
    /// </summary>
    let AsStreamBytes(cells : Item seq) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsStreamBytes
    /// <summary>
    /// Renders the items and returns the underlying ClosedXML <c>XLWorkbook</c> object, without
    /// saving it anywhere - useful for further manipulation before saving, or as a base to keep
    /// building on via <c>ItemFactory.Workbook</c>.
    /// </summary>
    let AsWorkBook(cells : Item seq) =
        cells
        |> Seq.toList
        |> FsExcel.Render.AsWorkBook
