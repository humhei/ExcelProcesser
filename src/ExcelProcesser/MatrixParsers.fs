module ExcelProcesser.MatrixParsers

open System.Diagnostics

#nowarn "0104"
open OfficeOpenXml
open Extensions
open CellParsers
open FParsec.CharParsers
open CellScript.Core
open Shrimp.FSharp.Plus 
open System

type Direction =
    | Horizontal = 0
    | Vertical = 1

type Coordinate =
    { X: int 
      Y: int }

[<RequireQualifiedAccess>]
module Coordinate =
    let origin = { X = 0; Y = 0 } 

type Shift =
    | Start
    | Vertical of Coordinate * int
    | Horizontal of Coordinate * int
    | Compose of Shift list
with 
    member x.Last =
        let rec loop shift =
            match shift with 
            | Compose shifts -> 
                match shifts with
                | [] -> failwith "compose shifts cannot be empty after start"
                | h :: t ->
                    loop h
            | _ -> shift
        loop x

        



    member internal x.Folded =
        let rec loop shift  =
            match shift with 
            | Compose shifts -> 
                List.collect loop shifts
            | _ -> [shift]
        loop x 

[<RequireQualifiedAccess>]
module Shift =
    let rec length = function 
        | Start -> 1
        | Horizontal _ -> 1
        | Vertical _ -> 1
        | Compose shifts -> shifts |> List.sumBy length


    let isVertical = function
        | Vertical _ -> true
        | _ -> false

    let isHorizontal = function 
        | Horizontal _ -> true 
        | _ -> false

    let verticalOffsets shift = 
        let rec loop shift =
            match shift with 
            | Horizontal (coordinate, _) -> [coordinate.Y] 
            | Start _ -> [0]
            | Vertical (coordinate, offset) -> [coordinate.Y + offset]
            | Compose shifts -> 
                shifts
                |> List.collect loop

        loop shift

    let horizontalOffsets shift = 
        let rec loop shift =
            match shift with 
            | Horizontal (coordinate, offset) -> [coordinate.X + offset] 
            | Start _ -> [0]
            | Vertical (coordinate, _) -> [coordinate.X]
            | Compose shifts -> 
                shifts
                |> List.collect loop

        loop shift

    let isInDirectionOrStart (direction) shift = 
        match shift with 
        | Shift.Start -> true
        | _ -> 
            match direction with 
            | Direction.Horizontal -> isHorizontal shift
            | Direction.Vertical -> isVertical shift

    let private redirect (preCalculatedCoordinate: Coordinate) (direction: Direction) (shifts: Shift list) =
        if List.forall (fun shift -> true (*length shift = 1*)) shifts then 
            let indexedShifts = List.indexed shifts
            match direction with 
            | Direction.Horizontal ->
                let chooser (index: int, shift) =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.Y = preCalculatedCoordinate.Y 
                        then 
                            Some (
                                index,
                                Horizontal(
                                    coordinate,
                                    i + 1
                                )
                            )
                        else None

                    | Shift.Vertical (coordinate,i) ->
                        if coordinate.Y = preCalculatedCoordinate.Y 
                        then 
                            Some (
                                index,
                                Horizontal(
                                    { coordinate with X = coordinate.X + 1},
                                    0
                                )
                            )
                        else None
                    | _ -> None

                match List.tryPick chooser indexedShifts with 
                | Some (i, shift) ->
                    match i with 
                    | 0 -> Compose (shift :: shifts.Tail)
                    | _ -> Compose (shift :: shifts)
                | None ->
                    failwithf "not implemented preCalculatedCoordinate %A direction %A shifts %A" preCalculatedCoordinate direction indexedShifts

            | Direction.Vertical ->
                let chooser (index, shift) =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.X = preCalculatedCoordinate.X 
                        then 
                            Some (
                                index,
                                Vertical(
                                    { coordinate with Y = coordinate.Y + 1},
                                    0
                                )
                            )
                        else None
                    | Shift.Vertical (coordinate, i) ->
                        if coordinate.X = preCalculatedCoordinate.X 
                        then 
                            Some (
                                index,
                                Vertical(
                                    coordinate,
                                    i + 1
                                )
                            )
                        else None
                    | _ -> None

                match List.tryPick chooser indexedShifts with 
                | Some (i, shift) ->
                    match i with 
                    | 0 -> Compose (shift :: shifts.Tail)
                    | _ -> Compose (shift :: shifts)

                | None ->
                    failwithf "not implemented preCalculatedCoordinate %A direction %A shifts %A" preCalculatedCoordinate direction shifts
            | _ -> failwith "Invalid token"
        else
            failwith "Not implemented"

    let rec applyDirection (preShift: Shift) (direction: Direction) shift: Shift =
        match preShift, shift with 
        | Start, Start ->
            match direction with 
            | Direction.Horizontal -> Horizontal (Coordinate.origin, 1)
            | Direction.Vertical -> Vertical (Coordinate.origin, 1)
            | _ -> failwith "Invalid token"
        | Start, Horizontal (coordinate, i) ->
            match direction with 
            | Direction.Horizontal -> Horizontal (coordinate, i + 1)
            | Direction.Vertical -> 
                if coordinate = Coordinate.origin then 
                    //Vertical ({ X = 0; Y = 0},1)
                    Compose[Vertical ({ X = 0; Y = 0},1); shift]
                else failwith "not implemented"
            | _ -> failwith "Invalid token"

        | Start, Vertical (coordinate, i) ->
            match direction with 
            | Direction.Horizontal -> 
                if coordinate = Coordinate.origin then 
                    //Horizontal ({ X = 0; Y = 0},1)
                    Compose[Horizontal ({ X = 0; Y = 0},1); shift]
                else failwith ""
            | Direction.Vertical -> Vertical (coordinate, i + 1)
            | _ -> failwith "Invalid token"

        | Horizontal (coordinate1, i1), Horizontal (coordinate2, i2) ->
            match direction with 
            | Direction.Horizontal -> 
                if coordinate1 = coordinate2 then
                    Horizontal (coordinate1, i2 + 1)
                else failwith "not implemented"
            | Direction.Vertical ->
                if coordinate1 = coordinate2 then
                    Compose [Vertical ({X = coordinate1.X + i1; Y = coordinate1.Y},1);shift]
                else failwith ""
            | _ -> failwith "Invalid token"

        | Vertical (coordinate1, i1), Vertical (coordinate2, i2) ->
            match direction with 
            | Direction.Horizontal -> 
                if coordinate1 = coordinate2 then
                    Compose [Horizontal ({X = coordinate1.X; Y = coordinate1.Y + i1},1);shift]
                else failwith "not implemented"
            | Direction.Vertical ->
                if coordinate1 = coordinate2 then
                    Vertical (coordinate1, i2 + 1)
                else failwith "not implemented"

            | _ -> failwith "Invalid token"

        | Start, Compose shifts ->
            redirect {X = 0; Y =0} direction shifts

        | Compose (h1 :: t1), shift ->
            applyDirection h1 direction shift

        | _ , Compose (h :: t) when h = preShift ->
            match applyDirection h direction h  with
            | Compose shifts ->
                Compose (shifts @ t)
            | h -> Compose (h :: t)


        | Horizontal (coordiante1, i1), Compose (h :: t) ->
            let calculatedCoordinate1 = { coordiante1 with X = coordiante1.X + i1 }
            redirect calculatedCoordinate1 direction (h :: t)

        | Vertical (coordiante1, i1), Compose (h :: t) ->
            let calculatedCoordinate1 = { coordiante1 with Y = coordiante1.Y + i1 }
            redirect calculatedCoordinate1 direction (h :: t)

        | _ -> failwith "not implemented"

type Shift with 
    member x.ShiftHorizontally(columnOffset: int) =
    
        (x, [1..columnOffset])
        ||> List.fold(fun shift _ ->
            Shift.applyDirection shift Direction.Horizontal shift
        )

    member x.ShiftVertically(rowOffset: int) =
        (x, [1..rowOffset])
        ||> List.fold(fun shift _ ->
            Shift.applyDirection shift Direction.Vertical shift
        )
            

    member x.ShiftBy(columnOffset, rowOffset: int) =
        x.ShiftHorizontally(columnOffset).ShiftVertically(rowOffset)


    
[<RequireQualifiedAccess>]
module internal SingletonExcelRangeBaseUnion =
    let rec offset(shift: Shift) (range: SingletonExcelRangeBaseUnion) =
        match shift with 
        | Start -> range

        | Horizontal (coordinate, i) -> 
            range.Offset(0 + coordinate.Y, i + coordinate.X)

        | Vertical (coordinate, i) ->
            range.Offset(i + coordinate.Y, 0 + coordinate.X)

        | Compose shifts -> 
            match shifts with
            | [] -> failwith "compose shifts cannot be empty after start"

            | h :: _ -> offset h range


[<DebuggerDisplay("{ExcelAddress}")>]
type ParsingAddress = ParsingAddress of ComparableExcelAddress

with 
    member x.Value =
        let (ParsingAddress v) = x
        v

    member x.StartRow = x.Value.StartRow

    member x.StartColumn = x.Value.StartColumn

    member x.EndRow = x.Value.EndRow

    member x.EndColumn = x.Value.EndColumn

    static member OfRange(range: ExcelRangeUnion) =
        range.ComparableExcelAddress()
        |> ParsingAddress

    static member OfRange(range: ExcelRangeBase) =
        ComparableExcelAddress.OfRange range
        |> ParsingAddress

    member x.ExcelAddress = x.Value.ExcelAddress
        
    member x.Address = x.Value.ExcelAddress.Address

    member x.ComparableExcelAddress = x.Value


[<DebuggerDisplay("{Range.Address} {Range.Text}")>]
[<StructuredFormatDisplay("{Range.Address} {Range.Text}")>]
type InputMatrixStream = 
    { Range: SingletonExcelRangeBaseUnion
      Shift: Shift
      ParsingAddress: ParsingAddress
      Logger: Logger }
with 
    member internal stream.OffsetedRange =
        SingletonExcelRangeBaseUnion.offset stream.Shift stream.Range

    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded


type OutputMatrixStreamResult<'result> =
    { IsSkip: bool
      Value: 'result }

type RangeTransformer(range: ExcelRangeUnion) =
    
    member x.Range = range

    member x.SetRowCountTo(rowsCount) =
        let newRange =
            range.Offset(0, 0, rowsCount, range.Columns)
        RangeTransformer(
            range = newRange
        )
        

    member x.SetColumnCountTo(columnsCount) =
        let newRange = 
            range.Offset(0, 0, range.Rows, columnsCount)
        
        RangeTransformer(
            range = newRange
        )

    member x.SetStart(start: ComparableExcelCellAddress) =
        let endAddr = x.Range.End
        let addr = start.RangeTo(endAddr)
        x.Range.Rerange(addr)
        |> RangeTransformer

    member x.SetEnd(endAddr: ComparableExcelCellAddress) =
        let start = x.Range.Start
        let addr = start.RangeTo(endAddr)
        x.Range.Rerange(addr)
        |> RangeTransformer

    member x.SetEndRow(endRow: int) =
        let endAddr =
            { Column = x.Range.End.Column 
              Row = endRow}

        x.SetEnd(endAddr)

    member x.SetEndColumn(endColumn: int) =
        let endAddr =
            { Column = endColumn 
              Row = x.Range.End.Row }

        x.SetEnd(endAddr)

    member x.SetStartRow(startRow: int) =
        let start = 
            { 
                Column = x.Range.Start.Column
                Row = startRow
            }
        x.SetStart(start)
        
    member x.SetStartColumn(startColumn: int) =
        let start = 
            { 
                Column = startColumn
                Row = x.Range.Start.Row
            }
        x.SetStart(start)


    member x.RightOf(targetRange: SingletonExcelRangeBaseUnion) =
        let addr = 
            ExcelAddress(
                fromCol = targetRange.Column + 1,
                toColumn = x.Range.End.Column,
                fromRow = x.Range.Start.Row,
                toRow = x.Range.End.Row
            )

        let newRange = x.Range.Rerange(addr.Address)
        
        RangeTransformer(
            range = newRange
        )

    member x.LastRow() =
        let newRange = 
            let range = x.Range
            range.Offset(range.Rows-1, 0, 1, range.Columns)

        RangeTransformer(newRange)

type OutputMatrixStream<'result> =
    { Range: SingletonExcelRangeBaseUnion
      Shift: Shift
      Logger: Logger
      ParsingAddress: ParsingAddress
      Result: OutputMatrixStreamResult<'result> }

with 

    member stream.OffsetedRange =
        SingletonExcelRangeBaseUnion.offset stream.Shift stream.Range
    

    member stream.RangeToOffsetedRange_AllShifts =
        let shift = stream.Shift
        let vertialOffset = 
            shift
            |> Shift.verticalOffsets
            |> List.max

        let horizontalOffset =
            shift
            |> Shift.horizontalOffsets
            |> List.max
          
        stream.Range.RangeTo(
            stream.Range.Offset(
                vertialOffset, horizontalOffset
            )
        )


    member stream.RangeToOffsetedRange =
        stream.Range.RangeTo(stream.OffsetedRange)


    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift
          ParsingAddress = x.ParsingAddress
          Logger = x.Logger }


    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded

    member x.ShiftHorizontally(columnOffset: int) =
        { x with 
            Shift = x.Shift.ShiftHorizontally(columnOffset)
        }

    member x.ShiftVertically(rowOffset: int) =
        { x with 
            Shift = x.Shift.ShiftVertically(rowOffset)
        }

    member x.ShiftBy(columnOffset, rowOffset: int) =
        x.ShiftHorizontally(columnOffset).ShiftVertically(rowOffset)



[<RequireQualifiedAccess>]
module OutputMatrixStream =

    let redirectTo rowOffset columnOffset (inputStream: InputMatrixStream) (outputStream: OutputMatrixStream<_>) =
        let targetRange = inputStream.Range.Offset(rowOffset, columnOffset)
        
        let reranged = inputStream.Range.RangeTo(targetRange)

        { 
            Range = inputStream.Range
            Shift = 
                inputStream.Shift.ShiftBy(reranged.Columns-1, reranged.Rows-1)
            Logger = inputStream.Logger
            ParsingAddress = ParsingAddress.OfRange reranged
            Result = outputStream.Result
        }

    let reRangeByShift (stream: OutputMatrixStream<_>) = 
        let verticals = 
            let rec loop shift =
                match shift with 
                | Vertical (coordinate, i) -> [(coordinate, i)]
                | Compose shifts -> shifts |> List.collect loop
                | _ -> []
            loop stream.Shift

        let maxVertical = 
            verticals 
            |> List.map (fun (coordinate, i) -> coordinate.Y + i)
            |> function
                | [] -> 0
                | verticals -> List.max verticals

        let horizontals = 
            let rec loop shift =
                match shift with 
                | Horizontal (coordinate, i) -> [(coordinate, i)]
                | Compose shifts -> shifts |> List.collect loop
                | _ -> []
            loop stream.Shift

        let maxHorizontal = 
            horizontals 
            |> List.map (fun (coordinate, i) -> coordinate.X + i)
            |> function
                | [] -> 0
                | horizontals -> List.max horizontals

        let reranged = 
            stream.Range.Offset(0, 0, maxVertical + 1, maxHorizontal + 1)

        RangeTransformer(reranged)


    let reRangeToEnd (stream: OutputMatrixStream<_>) = 
        let addr = stream.Range.Address + ":" + stream.ParsingAddress.ComparableExcelAddress.End.Address
        let reranged = stream.Range.Rerange(addr)
        RangeTransformer(reranged)

    let reRangeTo (targetRange: SingletonExcelRangeBaseUnion) (stream: OutputMatrixStream<_>) = 
        let addr = stream.Range.Address + ":" + targetRange.Address
        let reranged = stream.Range.Rerange(addr)
        RangeTransformer(reranged)

    let reRangeToAsOutputStream (targetRange: SingletonExcelRangeBaseUnion) (stream: OutputMatrixStream<_>) = 
        let addr = stream.Range.Address + ":" + targetRange.Address
        let reranged = stream.Range.Rerange(addr)
        { 
            stream.ShiftBy(reranged.Columns-1, reranged.Rows-1) with 
                ParsingAddress = ParsingAddress.OfRange reranged }


    let topArea (stream: OutputMatrixStream<_>) =   
        match stream.ParsingAddress.StartRow < stream.Range.Row with 
        | true -> 
            let addr = 
                ExcelAddress(
                    fromCol = stream.Range.Column,
                    toColumn = stream.OffsetedRange.Column,
                    fromRow = stream.ParsingAddress.StartRow,
                    toRow = stream.Range.Row - 1)

            let reranged = stream.Range.Rerange(addr.Address)
            (reranged)
            |> Some

        | false -> None

    let leftArea (stream: OutputMatrixStream<_>) =
        match stream.ParsingAddress.StartColumn < stream.Range.Column with 
        | true ->

            let addr = 
                ExcelAddress(
                    fromCol = stream.ParsingAddress.StartColumn,
                    toColumn = stream.Range.Column - 1,
                    fromRow = stream.Range.Row,
                    toRow = stream.OffsetedRange.Row )

            let reranged = stream.Range.Rerange(addr.Address)
            (reranged)
            |> Some

        | false -> None

    let reRangeRowTo rowsCount (stream: OutputMatrixStream<_>) = 
        let column = stream.ParsingAddress.EndColumn - stream.Range.Column + 1
        let reranged = stream.Range.Offset(0, 0, rowsCount, column)
        RangeTransformer(reranged)

    let reRangeColumnTo columnsCount (stream: OutputMatrixStream<_>) = 
        let row = stream.ParsingAddress.EndRow - stream.Range.Row + 1
        let reranged = stream.Range.Offset(0, 0, row, columnsCount)
        RangeTransformer(reranged)

    let applyDirectionToShift direction (preInputstream: InputMatrixStream) (stream: OutputMatrixStream<_>) =
        
        let isInSomeDirectionOrStart = 
            Shift.isInDirectionOrStart direction stream.Shift.Last
        if stream.Result.IsSkip && isInSomeDirectionOrStart then stream
        else
            let newShift = Shift.applyDirection preInputstream.Shift direction stream.Shift
            { stream with 
                Shift = newShift }

    let mapResult mapping (stream: OutputMatrixStream<'result>) =
        
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = mapping stream.Result
          Logger = stream.Logger
          ParsingAddress = stream.ParsingAddress
          }

    let mapResultValue mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = mapping stream.Result.Value
              IsSkip = stream.Result.IsSkip }
          ParsingAddress = stream.ParsingAddress
          Logger = stream.Logger }

    let removeRedundants (streams: OutputMatrixStream<'result> list) =
        let streamWithAddressList =
            streams
            |> List.map (fun stream ->
                ComparableExcelAddress.OfAddress(stream.Range.RangeTo(stream.OffsetedRange).Address), stream
            )

        let length = streamWithAddressList.Length

        let rec loop i (streamWithAddressList: list<ComparableExcelAddress * OutputMatrixStream<'result>>) =
            match i >= length with 
            | true -> streamWithAddressList
            | false ->
                match streamWithAddressList with 
                | ((headAddr, _) & h) :: t ->
                    let redundants, filtered =
                        t |> List.partition (fun (otherAddr, _) ->
                            otherAddr.IsIncludedIn headAddr
                        )

                    loop (i + redundants.Length + 1) (filtered @ [h])

                | [] -> []

        loop 0 streamWithAddressList
        |> List.map snd


[<RequireQualifiedAccess>]
type MatrixStream<'result> =
    | Input of InputMatrixStream
    | Output of previousInputStream: InputMatrixStream * OutputMatrixStream<'result> list

type MatrixParser<'result>(invoke: InputMatrixStream -> OutputMatrixStream<'result> list) =
    member x.Invoke = invoke

    member x.InDebug = 
        fun inputStream ->
            let r = invoke inputStream
            r
        |> MatrixParser


    member x.InvokeToStreams (outputStream: OutputMatrixStream<_>, rangeTransformer: RangeTransformer) =
        let inputStream =
            let reranged = rangeTransformer.Range
            let newAddr: ParsingAddress =
                ParsingAddress.OfRange reranged

            let range = 
                reranged.Offset(0, 0, 1, 1)
                |> SingletonExcelRangeBaseUnion.Create

            { outputStream.AsInputStream with 
                Range = range 
                ParsingAddress = newAddr
                Shift = Shift.Start
            }

        x.Invoke inputStream 

    member x.InvokeToResults (outputStream: OutputMatrixStream<_>, rangeTransformer: RangeTransformer) =
        x.InvokeToStreams(outputStream, rangeTransformer)
        |> List.map (fun outputStream -> outputStream.Result.Value)

    member x.Map(mapping) = 
        mapping invoke
        |> MatrixParser

    static member (||>>) (p: MatrixParser<_>, f) =
        p.Map(fun streamMapping ->
            streamMapping >> List.map (OutputMatrixStream.mapResultValue f)
        )

type SingletonMatrixParser<'result>(invoke: InputMatrixStream -> OutputMatrixStream<'result> option) =
    inherit MatrixParser<'result>(invoke >> Option.toList)

    member x.Invoke = invoke

    member internal x.InDebug = 
        fun inputStream ->
            let r = invoke inputStream
            if r.IsSome 
            then
                ()

            r
        |> SingletonMatrixParser

    member internal x.Debug() = failwith ""

    member x.PrefixRange() =
        fun inputStream ->
            let r = invoke inputStream
            match r with 
            | Some r ->
                r |> OutputMatrixStream.mapResultValue(fun result ->
                    r.OffsetedRange, result
                )
                |> Some

            | None -> None
        |> SingletonMatrixParser



    member x.MapOp(mapping) =
        mapping invoke
        |> SingletonMatrixParser


    static member (||>>) (p: SingletonMatrixParser<_>, f) =
        p.MapOp(fun streamMapping ->
            streamMapping >> Option.map (OutputMatrixStream.mapResultValue f)
        )


[<RequireQualifiedAccess>]
module private List =
    let (|Some|None|) (list: 'a list) = 
        if list.Length = 0 then None
        else Some list

let private map (matrixParser: MatrixParser<_>) mapping =
    matrixParser.Map mapping

[<RequireQualifiedAccess>]
module MatrixParser =
    let mapOutputStream f p = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.map f outputStreams

    let mapResultValueWithOutputStream f p = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            
            outputStreams
            |> List.map (fun outputStream ->
                outputStream
                |> OutputMatrixStream.mapResultValue (fun _ -> f outputStream)
            )


    let mapOutputStreams f p = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            f outputStreams


    let collectOutputStream f p : MatrixParser<_> = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.collect f outputStreams

    let pickOutputStream f p : MatrixParser<_> = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.choose f outputStreams


    let filterOutputStreamByResultValue f p = map p <| fun p ->
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.filter (fun outputStream -> f outputStream.Result.Value) outputStreams

let private mxCellParserOp_common (cellParser: InputMatrixStream -> SingletonExcelRangeBaseUnion -> 'result option) =
    fun (stream: InputMatrixStream) ->
        let offsetedRange = stream.OffsetedRange
        let addr = offsetedRange.ExcelCellAddress
        let streamAddr = stream.ParsingAddress
        match addr.Row, addr.Column with 
        | Between(streamAddr.StartRow, streamAddr.EndRow),
            Between(streamAddr.StartColumn, streamAddr.EndColumn) ->
            match cellParser stream offsetedRange with 
            | Some result ->
                { Range = stream.Range 
                  Shift = stream.Shift 
                  ParsingAddress = stream.ParsingAddress
                  Result = 
                    { IsSkip = false
                      Value = result }
                  Logger = stream.Logger
                }
                |> Some
            | None -> None
        | _ -> None
    |> SingletonMatrixParser

let mxCellParserOp (cellParser: SingletonExcelRangeBaseUnion -> 'result option) =
    mxCellParserOp_common(fun _ range -> cellParser range)
//let mxCellParser_Result (cellParser: SingletonExcelRangeBaseUnion -> Result<'ok, string>) =
//    mxCellParserOp (fun range ->
//        match cellParser range with 
//        | Result.Ok v -> Some v
//        | Result.Error error -> 
//    )

let mxCellParser (cellParser: CellParser) getResult =
    fun range ->
        let b = cellParser range
        if b then 
            Some (getResult range)
        else None
    |> mxCellParserOp

let mxFParsec p =
    mxCellParserOp (pFParsec p)


let mxFParsecInt32 = mxFParsec (pint32) 
 
let mxText text =
    mxCellParser (pText text) SingletonExcelRangeBaseUnion.getText

let mxTextf f =
    mxCellParser (pTextf f) SingletonExcelRangeBaseUnion.getText

let mxNonEmpty = 
    mxTextf isTrimmedTextNotEmpty

let mxTextOp picker =
    mxCellParserOp (fun range ->
        picker range.Text
    )


let mxRegex pattern =
    mxCellParser (pTextf (fun text -> 
        match text with 
        | ParseRegex.Head pattern _ -> true
        | _ -> false
    )) SingletonExcelRangeBaseUnion.getText

let mxWord = mxRegex "\w" 

let mxSpace = mxCellParser pSpace ignore 

let mxEmpty = mxSpace


let mxStyleName styleName = mxCellParser (pStyleName styleName) SingletonExcelRangeBaseUnion.getText

let mxAnySkip = 
    mxCellParser pAny ignore 

let mxAnyOrigin = 
    mxCellParser pAny SingletonExcelRangeBaseUnion.getText 

let mxAddress address = 
    let targetAddress = new ExcelAddress(address)
    mxCellParserOp(fun range ->
        if ExcelAddress(range.Address).Address = targetAddress.Address
        then Some range.Text
        else None
    )

let mxAnyOriginObj = 
    mxCellParser pAny SingletonExcelRangeBaseUnion.getValue 


//let (||>>) p f = 
//    MatrixParser.mapOutputStream (fun outputStream -> OutputMatrixStream.mapResultValue f outputStream) p

[<AutoOpen>]
module LoggerExtensions =
    [<RequireQualifiedAccess>]
    module MatrixParser =
        let addLogger loggerLevel (name: string) (parser: MatrixParser<_>) = map parser <| fun parser ->
            fun (inputStream: InputMatrixStream) ->
                let logger = inputStream.Logger

                let outputStreams = parser inputStream
                //if not outputStreams.IsEmpty 
                //then logger.Log loggerLevel  (sprintf "BEGIN %s:" name)

                let name =  
                    (sprintf $"Parser: {name}").PadRight(30)

                for outputStream in outputStreams do
                    let message =   
                        let range = outputStream.OffsetedRange
                        let addr = range.Address
                        sprintf $"{name}\tResult: {addr}: {outputStream.Result.Value}"

                    logger.Log loggerLevel message

                //if not outputStreams.IsEmpty 
                //then logger.Log loggerLevel (sprintf "END %s:" name)

                outputStreams


    [<RequireQualifiedAccess>]
    module SingletonMatrixParser =
        let addLogger loggerLevel (name: string) (parser: SingletonMatrixParser<_>) = parser.MapOp <| fun parser ->
            fun (inputStream: InputMatrixStream) ->
                let logger = inputStream.Logger
                let outputStream = parser inputStream
                //if not outputStreams.IsEmpty 
                //then logger.Log loggerLevel  (sprintf "BEGIN %s:" name)
                let name =  
                    (sprintf $"Parser: {name}").PadRight(30)

                match outputStream with 
                | Some outputStream ->

                    let message =   
                        let range = outputStream.OffsetedRange
                        let addr = range.Address
                        sprintf $"{name}\tResult: {addr}: {outputStream.Result.Value}"

                    logger.Log loggerLevel message

                | None -> ()

                //if not outputStreams.IsEmpty 
                //then logger.Log loggerLevel (sprintf "END %s:" name)

                outputStream



let (|||>>) p f = 
    MatrixParser.mapOutputStream (fun outputStream ->
       OutputMatrixStream.mapResultValue (f (outputStream)) outputStream
    ) p

let mxEOF  =
    fun (stream: InputMatrixStream) ->
        let offsetedRange = stream.OffsetedRange
        let addr = offsetedRange.ExcelCellAddress
        let streamAddr = stream.ParsingAddress
        match addr.Row, addr.Column with 
        | Between(streamAddr.StartRow, streamAddr.EndRow),
            Between(streamAddr.StartColumn, streamAddr.EndColumn) ->
            None

        | _ -> 
            { Range = stream.Range 
              Shift = stream.Shift 
              ParsingAddress = stream.ParsingAddress
              Result = 
                { IsSkip = false
                  Value = () }
              Logger = stream.Logger
            }
            |> Some
    |> SingletonMatrixParser

let mxNotEOF()  =
    mxCellParserOp_common (fun inputStream range ->
        match inputStream.ParsingAddress.ComparableExcelAddress.Contains(range.ExcelAddress) with 
        | true -> Some range.Text
        | false -> None
    )

let mxOR (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>) =
    let p1 = 
        p1 ||>> Choice1Of2

    let p2 = 
        p2 ||>> Choice2Of2

    fun inputStream ->
        match p1.Invoke inputStream with
        | List.Some streams -> streams
        | _ -> p2.Invoke inputStream
    |> MatrixParser

let mxNot (p1: SingletonMatrixParser<'result1>) =
    fun inputStream ->
        match p1.Invoke inputStream with
        | Some outputStreams -> None
        | None -> mxAnyOrigin.Invoke inputStream
    |> SingletonMatrixParser

let (<||>) (p1: MatrixParser<'result>) (p2: MatrixParser<'result>) =
    mxOR p1 p2
    ||>> 
        function
            | Choice1Of2 v -> v
            | Choice2Of2 v -> v
 
/// p1 && not p2
let (<&!>) (p1: MatrixParser<'result>) (p2: MatrixParser<'exclude>) = 
    fun inputStream ->
        match p2.Invoke inputStream with 
        | List.Some _ -> []
        | _ -> p1.Invoke inputStream

    |> MatrixParser

let (<.&>) (p1: MatrixParser<'result>) (p2: MatrixParser<'predicate>) = 
    fun inputStream ->
        match p1.Invoke inputStream with 
        | List.Some outputStreams1 -> 
            match p2.Invoke inputStream with 
            | List.Some _ -> outputStreams1
            | List.None -> []
        | List.None -> []

    |> MatrixParser

let (<&>) (p1: MatrixParser<'predicate>) (p2: MatrixParser<'result>) = 
    fun inputStream ->
        match p1.Invoke inputStream with 
        | List.Some outputStreams1 -> 
            match p2.Invoke inputStream with 
            | List.Some outputStreams2 -> outputStreams2
            | List.None -> []
        | List.None -> []
    |> MatrixParser






let private pipe2RelativelyWithTupleStreamsReturn (direction: Direction) (p1: MatrixParser<'result1>) (buildP2: OutputMatrixStream<'result1> -> MatrixParser<'result2>) f =

    fun inputstream1 ->

        let newStreams1 = p1.Invoke inputstream1
        match newStreams1 with 
        | List.Some newStreams1 ->
            
            newStreams1
            |> List.collect (fun newStream1 ->
                let p2 = buildP2 newStream1
                let inputStream2 = (OutputMatrixStream.applyDirectionToShift direction inputstream1 newStream1).AsInputStream
                p2.Invoke inputStream2
                |> List.map (fun newStream2 -> 
                    let newStream2 = 
                        OutputMatrixStream.mapResultValue (fun result ->
                            f (newStream1.Result.Value, result)
                        ) newStream2
                    newStream1, newStream2
                )
            )
        | List.None -> []


let private pipe2Relatively (direction: Direction) (p1: MatrixParser<'result1>) (buildP2: OutputMatrixStream<'result1> -> MatrixParser<'result2>) f =
    
    pipe2RelativelyWithTupleStreamsReturn direction p1 buildP2 f
    >> List.map snd
    |> MatrixParser
    


let pipe2 direction p1 p2 f = 
    pipe2Relatively direction p1 (fun _ -> p2) f

let pipe3 direction p1 p2 p3 f =
    pipe2 direction (pipe2 direction p1 p2 id) p3 (fun ((a, b), c) ->
        f (a, b, c)
    )

let atLeastOne (p: MatrixParser<'a list>) =
    p
    |> MatrixParser.filterOutputStreamByResultValue (fun list ->
        not list.IsEmpty
    )

let atLeastTwo (p: MatrixParser<'a list>) =
    p
    |> MatrixParser.filterOutputStreamByResultValue (fun list ->
        list.Length > 1
    )

let mxManyWithMaxCount direction (maxCount: int option) (p: MatrixParser<'result>) = map p <| fun p ->
    let isSkip outputStream = 
        outputStream.Result.IsSkip

    fun inputStream ->
        let rec loop stream (accum: OutputMatrixStream<'result> list) = [
            let isReachMaxCount =
                match maxCount with 
                | Some maxCount -> 
                    accum.Length >= maxCount
                | None -> false

            if isReachMaxCount then yield accum
            else
                match stream with
                | MatrixStream.Input inputStream ->
                    match p inputStream with 
                    | List.Some outputStreams ->
                        let skip, outputStreams = List.partition isSkip outputStreams 

                        yield! List.replicate skip.Length []

                        yield! loop (MatrixStream.Output (inputStream,outputStreams)) (accum @ outputStreams) 

                    | List.None -> yield []

                | MatrixStream.Output (preInputStream,outputStreams1) ->
                    match outputStreams1 with 
                    | List.Some outputStreams1 ->

                        yield!
                            outputStreams1
                            |> List.collect (fun outputStream1 -> [
                                let inputStream = (OutputMatrixStream.applyDirectionToShift direction preInputStream outputStream1).AsInputStream
                                match p inputStream with 
                                | List.Some outputStreams2 ->
                                    let skip, outputStreams2 = List.partition isSkip outputStreams2

                                    yield! List.replicate skip.Length []

                                    yield! loop (MatrixStream.Output (inputStream, outputStreams2)) (accum @ outputStreams2)

                                | List.None ->  yield accum
                            ]
                            )

                    | List.None -> yield accum
        ]



        let outputStreamLists = loop (MatrixStream.Input inputStream) []
        outputStreamLists 
        |> List.map (fun outputStreams ->
            match outputStreams with 
            | _ :: _ ->
                let last = List.last outputStreams
                { Range = last.Range 
                  Shift = last.Shift 
                  Logger = last.Logger
                  ParsingAddress = last.ParsingAddress
                  Result = 
                    { IsSkip = false
                      Value = 
                        outputStreams 
                        |> List.map (fun outputStream ->
                              outputStream.Result.Value 
                        )
                    }
                }

            | [] -> 
                { Range = inputStream.Range
                  Shift = inputStream.Shift
                  Logger = inputStream.Logger
                  ParsingAddress = inputStream.ParsingAddress
                  Result = 
                    { IsSkip = true
                      Value = []
                    }
                }
        )

let mxMany_all_ForRowOrColumn direction p =
    fun (inputStream: InputMatrixStream) ->
        let maxCount = 
            match direction with 
            | Direction.Horizontal -> 
                let cols = 
                    inputStream.Shift
                    |> Shift.horizontalOffsets 

                (List.max cols) + 1
            
            | Direction.Vertical ->
                let rows = 
                    inputStream.Shift
                    |> Shift.verticalOffsets

                (List.max rows) + 1

        let r  = (mxManyWithMaxCount direction (Some maxCount) p).Invoke inputStream
        match r with 
        | [r] ->
            match r.Result.IsSkip with 
            | true -> []
            | false ->
                match r.Result.Value.Length with 
                | EqualTo maxCount -> [r]
                | _ -> []

        | _ -> failwith "Not implemented"
        //r
        //|> List.collect(fun r ->
        //    match r.Result.IsSkip with 
        //    | true -> []
        //    | false ->
        //        match r.Result.Value.Length with 
        //        | EqualTo maxCount -> [r]
        //        | _ -> []
        //)
    |> MatrixParser

let mxMany1WithMaxCount direction (maxCount: int option) (p: MatrixParser<'result>) =
    mxManyWithMaxCount direction maxCount p
    |> atLeastOne

let mxMany2WithMaxCount direction (maxCount: int option) (p: MatrixParser<'result>) =
    mxManyWithMaxCount direction maxCount p
    |> atLeastTwo

let mxMany direction p = mxManyWithMaxCount direction None p



let mxMany1 direction p =
    mxMany direction p
    |> atLeastOne

/// alaways backtrack
/// 
/// Result<_, _>
let mxManySkipRetain direction pSkip maxSkipCount p = 
    let skip = 
        mxManyWithMaxCount direction (Some maxSkipCount) pSkip 

    let many1 = mxMany1 direction p

    let piped = 
        fun (inputstream: InputMatrixStream) ->
            //None
            let outputStreams = (pipe2 direction skip many1 id).Invoke inputstream
            match outputStreams with 
            | List.Some outputStreams ->
                outputStreams
                |> List.filter (fun outputStream ->
                    let errors, values = outputStream.Result.Value
                    values.Length > 0
                )
            | List.None -> []
        |> MatrixParser

    ((mxMany direction piped))
    ||>> fun list ->
        ([], list) ||> List.fold (fun state (errors, values) ->
            state @ (List.map Result.Error errors) @ (List.map Result.Ok values)
        )

/// alaways backtrack
/// 
/// Result<_, _>
let mxMany1SkipRetain direction pSkip maxSkipCount p =
    let many1 = mxMany1 direction p
    pipe2RelativelyWithTupleStreamsReturn direction many1 (fun _ -> mxManySkipRetain direction pSkip maxSkipCount p) (fun (a,b) ->
        List.map Result.Ok a @ b
    ) >> List.map (fun (stream1, stream2) ->
        if stream2.Result.IsSkip then 
            stream1 |> OutputMatrixStream.mapResultValue (List.map Result.Ok)
        else stream2
    )
    |> MatrixParser


/// at least one skip
let mxMany1Skip1 direction pSkip maxSkipCount p =
    mxMany1SkipRetain direction pSkip maxSkipCount p
    |> MatrixParser.filterOutputStreamByResultValue (fun results ->
        let errors =
            results
            |> List.choose(function 
                | Result.Error error -> Some error
                | _ -> None
            )

        if errors.IsEmpty 
        then false
        else true
    )
    ||>> (List.choose (fun v ->
        match v with 
        | Result.Ok ok -> Some ok
        | Result.Error _ -> None
    ))

let mxMany1SkipRetain_BeginBy direction (pSkip:MatrixParser<_>) maxSkipCount p = 
    pipe2 direction 
        (mxManyWithMaxCount direction (Some maxSkipCount) pSkip)
        (mxMany1SkipRetain direction pSkip maxSkipCount p)
        (fun (skips, results) ->
            let skips =
                skips
                |> List.map Result.Error

            skips @ results
        )

let mxMany1Op direction maxSkipCount (p: SingletonMatrixParser<_>) =
    mxMany1SkipRetain_BeginBy direction (mxNot p) maxSkipCount p
    ||>> (List.map Result.toOption)


let mxMany1Skip direction pSkip maxSkipCount p =
    mxMany1SkipRetain direction pSkip maxSkipCount p
    ||>> (List.choose (fun v ->
        match v with 
        | Result.Ok ok -> Some ok
        | Result.Error _ -> None
    ))


let inDirection (p: Direction -> MatrixParser<'result>) =
    fun (inputStream: InputMatrixStream) ->
        let direction =
            let rec loop shift = 
                match shift with 
                | Start -> failwith "Cannot applied mxUntil in start position"
                | Horizontal _ -> Direction.Horizontal
                | Vertical _ -> Direction.Vertical
                | Compose shifts ->
                    match shifts with 
                    | h :: t -> 
                        loop h
                    | _ -> failwith "compose shifts cannot be empty after start"
            loop inputStream.Shift
        (p direction).Invoke inputStream
    |> MatrixParser


let mxUntil direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2 direction (mxManyWithMaxCount direction maxCount (pPrevious <&!> pLast)) pLast id
   

let mxUntilBacktrackLast direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2RelativelyWithTupleStreamsReturn direction (mxManyWithMaxCount direction maxCount (pPrevious <&!> pLast)) (fun _ -> pLast) id
    >> List.map fst

let mxUntil1 direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2 direction (mxMany1WithMaxCount direction maxCount (pPrevious <&!> pLast)) pLast id




let mxUntil1NoConfict direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2 direction (mxMany1WithMaxCount direction maxCount (pPrevious)) pLast id


let mxUntil1BacktrackLast direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2RelativelyWithTupleStreamsReturn direction (mxMany1WithMaxCount direction maxCount (pPrevious <&!> pLast)) (fun _ -> pLast) id
    >> List.map fst

let mxUntil2BacktrackLast direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2RelativelyWithTupleStreamsReturn direction (mxMany2WithMaxCount direction maxCount (pPrevious <&!> pLast)) (fun _ -> pLast) id
    >> List.map fst

/// IND = inDirection
let mxUntilIND maxCount pPrevious (pLast: MatrixParser<'result>) =
    inDirection (fun direction ->
        mxUntil direction maxCount pPrevious (pLast: MatrixParser<'result>)
    )

let mxUntilIND_EOF pPrevious =
    inDirection (fun direction ->
        mxUntil1 direction None (pPrevious) (mxEOF)
    )


/// Space >>. pUntil
let mxUntilS maxCount (p: MatrixParser<'result>) =
    inDirection (fun direction ->
        mxUntil direction maxCount mxSpace p
    )
    ||>> snd

/// AnySkip >>. pUntil
let mxUntilA maxCount (p: MatrixParser<'result>) =
    inDirection (fun direction ->
        mxUntil direction maxCount mxAnySkip p
    )
    ||>> snd




let mxUntilIND50 pPrevious (pLast: MatrixParser<'result>) = mxUntilIND (Some 50) pPrevious (pLast: MatrixParser<'result>)

let mxUntilA50 (p: MatrixParser<'result>) = mxUntilA (Some 50) p

let mxUntilA10 (p: MatrixParser<'result>) = mxUntilA (Some 10) p

let mxUntilA5 (p: MatrixParser<'result>) = mxUntilA (Some 5) p 
let mxUntilS5 (p: MatrixParser<'result>) = mxUntilS (Some 5) p 


let mxInt32 = mxTextOp(fun text ->
    match Int32.TryParse text with 
    | true, v -> Some v
    | false, _ -> None
)

let mxDouble = mxTextOp(fun text ->
    match Double.TryParse text with 
    | true, v -> Some v
    | false, _ -> None
)


type MergeStarterResult =
    { Address: string 
      Text: string }


let mxMergeStarter = 
    mxCellParser pMergeStarter (fun range -> { Address = range.Address; Text = range.Text}) 
    

let mxMergeWithAddresses direction =
    pipe2Relatively direction mxMergeStarter (fun outputStream ->
        let workSheet = outputStream.Range.WorksheetOrFail
        let mergeCellId = SingletonExcelRangeBaseUnion.Create(workSheet.Cells.[outputStream.Result.Value.Address]).GetMergeCellId()
        mxMany1 direction (mxCellParser (fun range -> range.GetMergeCellId() = mergeCellId) SingletonExcelRangeBaseUnion.getExcelCellAddress)
    ) id

let mxMerge direction =
    mxMergeWithAddresses direction
    ||>> (fun (start, addresses) ->
        start.Text
    )





let mxColMany p = mxMany Direction.Horizontal p

let mxColMany1 p = mxMany1 Direction.Horizontal p
let mxColMany1Op maxSkipCount p = mxMany1Op Direction.Horizontal maxSkipCount p
let mxColMany1WithMaxCount maxCount p = mxMany1WithMaxCount Direction.Horizontal maxCount p


let mxColMany1Skip pSkip maxSkipCount p = mxMany1Skip Direction.Horizontal pSkip maxSkipCount p

let mxColMany1SkipRetain pSkip maxSkipCount p = 
    mxMany1SkipRetain Direction.Horizontal pSkip maxSkipCount p
    
let mxColMany1SkipRetain_BeginBy pSkip maxSkipCount p = 
    mxMany1SkipRetain_BeginBy Direction.Horizontal pSkip maxSkipCount p

let mxRowMany1Skip pSkip maxSkipCount p = mxMany1Skip Direction.Vertical pSkip maxSkipCount p


let mxRowMany1SkipRetain pSkip maxSkipCount p = 
    mxMany1SkipRetain Direction.Vertical pSkip maxSkipCount p
    
let mxRowMany1SkipRetain_BeginBy pSkip maxSkipCount p =
    mxMany1SkipRetain_BeginBy Direction.Vertical maxSkipCount p

let mxRowMany1Op maxSkipCount p = mxMany1Op Direction.Vertical maxSkipCount p


let mxRowMany p = mxMany Direction.Vertical p
let mxRowMany1 p = mxMany1 Direction.Vertical p
let mxRowMany1WithMaxCount maxCount p = mxMany1WithMaxCount Direction.Vertical maxCount p

let mxEmptyRow: MatrixParser<_> =
    mxMany_all_ForRowOrColumn Direction.Horizontal mxEmpty
  

let mxEntityRow: MatrixParser<_> =
    MatrixParser(
        fun inputStream ->
            let r = mxEmptyRow.Invoke inputStream
            match r with 
            | [] -> 
                (mxMany_all_ForRowOrColumn Direction.Horizontal mxAnyOrigin).Invoke inputStream
            | _ -> []
    )



let c2 p1 p2 =
    pipe2 Direction.Horizontal p1 p2 id


/// R = Relatively
let c2R p1 buildP2 = 
    pipe2Relatively Direction.Horizontal p1 buildP2 id

let c3 p1 p2 p3 =
    pipe3 Direction.Horizontal p1 p2 p3 id

let c4 p1 p2 p3 p4 =
    c2 (c3 p1 p2 p3) p4
    ||>> (fun ((a, b, c), d) ->
        a, b, c, d
    )

let c5 p1 p2 p3 p4 p5 =
    c2  (c3 p1 p2 p3) (c2 p4 p5)
    ||>> (fun ((a, b, c), (d, e)) ->
        a, b, c, d, e
    )

let c6 p1 p2 p3 p4 p5 p6 =
    c2  (c3 p1 p2 p3) (c3 p4 p5 p6)
    ||>> (fun ((a, b, c), (d, e, f)) ->
        a, b, c, d, e, f
    )

let c7 p1 p2 p3 p4 p5 p6 p7 =
    c2  (c3 p1 p2 p3) (c4 p4 p5 p6 p7)
    ||>> (fun ((a, b, c), (d, e, f, g)) ->
        a, b, c, d, e, f, g
    )

let c8 p1 p2 p3 p4 p5 p6 p7 p8 =
    c2  (c3 p1 p2 p3) (c5 p4 p5 p6 p7 p8)
    ||>> (fun ((a, b, c), (d, e, f, g, h)) ->
        a, b, c, d, e, f, g, h
    )


let r2 p1 p2 =
    pipe2 Direction.Vertical p1 p2 id


/// R = Relatively
let r2R p1 buildP2 =
    pipe2Relatively Direction.Vertical p1 buildP2 id

let r3 p1 p2 p3 = 
    pipe3 Direction.Vertical p1 p2 p3 id

let r4 p1 p2 p3 p4 =
    r2 (r3 p1 p2 p3) p4
    ||>> (fun ((a, b, c), d) ->
        a, b, c, d
    )

let r5 p1 p2 p3 p4 p5 =
    r2  (r3 p1 p2 p3) (r2 p4 p5)
    ||>> (fun ((a, b, c), (d, e)) ->
        a, b, c, d, e
    )

let r6 p1 p2 p3 p4 p5 p6 =
    r2  (r3 p1 p2 p3) (r3 p4 p5 p6)
    ||>> (fun ((a, b, c), (d, e, f)) ->
        a, b, c, d, e, f
    )

let r7 p1 p2 p3 p4 p5 p6 p7 =
    r2  (r3 p1 p2 p3) (r4 p4 p5 p6 p7)
    ||>> (fun ((a, b, c), (d, e, f, g)) ->
        a, b, c, d, e, f, g
    )

let r8 p1 p2 p3 p4 p5 p6 p7 p8 =
    r2  (r3 p1 p2 p3) (r5 p4 p5 p6 p7 p8)
    ||>> (fun ((a, b, c), (d, e, f, g, h)) ->
        a, b, c, d, e, f, g, h
    )

let private runMatrixParserForRangesWithStreamsAsResult_Common addr logger (ranges : seq<SingletonExcelRangeBaseUnion>) (p : MatrixParser<_>) =
    let ranges =
        ranges 
        |> List.ofSeq


    let inputStreams = 
        ranges 
        |> List.map (fun range ->
            { Range = range 
              Shift = Shift.Start
              Logger = logger
              ParsingAddress = addr
              }
        )
    let r = 
        inputStreams 
        |> List.collect p.Invoke

    r

let private runMatrixParserForRangesWithStreamsAsResult addr (ranges : seq<SingletonExcelRangeBaseUnion>) (p : MatrixParser<_>) =
    runMatrixParserForRangesWithStreamsAsResult_Common addr (new Logger()) ranges p

let private runMatrixParserForRanges addr (ranges : seq<SingletonExcelRangeBaseUnion>) (p : MatrixParser<_>) =

    let mses = runMatrixParserForRangesWithStreamsAsResult addr ranges p
    mses |> List.map (fun ms -> ms.Result.Value)


let runMatrixParserForRangeWithStreamsAsResult (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeBase.asRangeList range
        |> List.map SingletonExcelRangeBaseUnion.Office


    let mses = runMatrixParserForRangesWithStreamsAsResult address ranges p
    mses

let runMatrixParserForRangeWithStreamsAsResultUnion (range : ExcelRangeUnion) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeUnion.asRangeList range

    let mses = runMatrixParserForRangesWithStreamsAsResult address ranges p
    mses

let runMatrixParserForRangeWithStreamsAsResult2 (logger: Logger) (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeBase.asRangeList range
        |> List.map SingletonExcelRangeBaseUnion.Office

    runMatrixParserForRangesWithStreamsAsResult_Common address logger ranges p

/// Including Empty Ranges
let runMatrixParserForRangeWithStreamsAsResult2_All (logger: Logger) (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeBase.asRangeList_All range
        |> List.map SingletonExcelRangeBaseUnion.Office

    runMatrixParserForRangesWithStreamsAsResult_Common address logger ranges p

/// Including Empty Ranges
let runMatrixParserForRangeWithStreamsAsResult2_All_Union (logger: Logger) (range : ExcelRangeUnion) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeUnion.asRangeList_All range

    runMatrixParserForRangesWithStreamsAsResult_Common address logger ranges p

let runMatrixParserForRange2 logger (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult2 logger range p
    |> List.map (fun m -> m.Result.Value)


/// Including Empty Ranges
let runMatrixParserForRange2_All logger (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult2_All logger range p
    |> List.map (fun m -> m.Result.Value)

let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult range p
    |> List.map (fun m -> m.Result.Value)

let runMatrixParserForRangeWithoutRedundent (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult range p
    |> OutputMatrixStream.removeRedundants
    |> List.map (fun m -> m.Result.Value)

/// Including Empty Ranges
let  private runMatrixParserWithStreamsAsResult_Common (worksheet: ValidExcelWorksheet) logger (p: MatrixParser<_>) =
  
    #if TestVirtual
    let userRange =
        let datas = worksheet.ReadDatas(RangeGettingOptions.UserRange)
        datas.Content
        |> VirtualExcelRange.OfData
        |> fun m -> m.AsCellRanges()
        |> List.map SingletonExcelRangeBaseUnion.Virtual
    #else
    let userRange = 
        worksheet.Value
        |> ExcelWorksheet.getUserRangeList
        |> List.map SingletonExcelRangeBaseUnion.Office
    #endif


    let addr =
        { 
            StartRow = 1
            StartColumn = 1
            EndRow = ExcelWorksheet.getMaxRowNumber worksheet.Value
            EndColumn = ExcelWorksheet.getMaxColNumber worksheet.Value
        }
        |> ParsingAddress

    let r = runMatrixParserForRangesWithStreamsAsResult_Common addr logger userRange p
    r

let runMatrixParser (worksheet: ValidExcelWorksheet) (p: MatrixParser<_>) =
    runMatrixParserWithStreamsAsResult_Common worksheet (new Logger()) p
    |> List.map (fun m -> m.Result.Value)

let runMatrixParserWithStreamsAsResult (worksheet: ValidExcelWorksheet) (p: MatrixParser<'result>) = 
    runMatrixParserWithStreamsAsResult_Common worksheet (new Logger()) p
    

let runMatrixParserWithStreamsAsResultSafe (worksheet: ValidExcelWorksheet) (p: MatrixParser<'result>) =

    let logger = new Logger()

    match runMatrixParserWithStreamsAsResult_Common worksheet (logger) p with 
    | [] -> 
        match logger.Messages().IsEmpty with 
        | true -> failwithf "SheetName: %s\n ParsingTarget: %s\nAll named parsed are parsed failured %A\nStackTrace:\n%s" worksheet.Name (typeof<'result>.Name) p System.Environment.StackTrace 
        | false ->
            failwithf "SheetName: %s\n ParsingTarget: %s\n%A\nStackTrace%s" worksheet.Name (typeof<'result>.Name) (logger.Messages()) System.Environment.StackTrace 
    | outputStreams -> (AtLeastOneList.Create outputStreams)


type MatrixParserSuccessfulResult<'result> = private MatrixParserSuccessfulResult of AtLeastOneList<'result>
with 
    member x.Value = 
        let (MatrixParserSuccessfulResult v) = x
        v

    member x.AsList = x.Value.AsList

let runMatrixParserSafe (worksheet: ValidExcelWorksheet) (p: MatrixParser<_>) =
    runMatrixParserWithStreamsAsResultSafe worksheet p
    |>AtLeastOneList.map (fun m -> m.Result.Value)
    |> MatrixParserSuccessfulResult
