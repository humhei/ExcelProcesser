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

    let private redirect (preCalculatedCoordinate: Coordinate) (direction: Direction) (shifts: Shift list) =
        if List.forall (fun shift -> true (*length shift = 1*)) shifts then 
            match direction with 
            | Direction.Horizontal ->
                let chooser shift =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.Y = preCalculatedCoordinate.Y 
                        then 
                            Some (
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
                                Horizontal(
                                    { coordinate with X = coordinate.X + 1},
                                    0
                                )
                            )
                        else None
                    | _ -> None

                match List.tryPick chooser shifts with 
                | Some shift ->
                    Compose (shift :: shifts)
                | None ->
                    failwithf "not implemented preCalculatedCoordinate %A direction %A shifts %A" preCalculatedCoordinate direction shifts

            | Direction.Vertical ->
                let chooser shift =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.X = preCalculatedCoordinate.X 
                        then 
                            Some (
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
                                Vertical(
                                    coordinate,
                                    i + 1
                                )
                            )
                        else None
                    | _ -> None

                match List.tryPick chooser shifts with 
                | Some shift ->
                    Compose (shift :: shifts)
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
                    Vertical ({ X = 0; Y = 0},1)
                else failwith "not implemented"
            | _ -> failwith "Invalid token"

        | Start, Vertical (coordinate, i) ->
            match direction with 
            | Direction.Horizontal -> 
                if coordinate = Coordinate.origin then 
                    Horizontal ({ X = 0; Y = 0},1)
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

[<RequireQualifiedAccess>]
module internal ExcelRangeBase =
    let rec offset (shift: Shift) (range: ExcelRangeBase) =
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
    
[<RequireQualifiedAccess>]
module internal SingletonExcelRangeBase =
    let offset(shift: Shift) (range: SingletonExcelRangeBase) =
        range.Value
        |> ExcelRangeBase.offset shift
        |> SingletonExcelRangeBase.Create

[<DebuggerDisplay("{ExcelAddress}")>]
type ParsingAddress =
    { StartRow: int 
      EndRow: int
      StartColumn: int 
      EndColumn: int 
      }
with 
    static member OfRange(range: ExcelRangeBase) =
        let startCell = range.Start

        let endCell = range.End
        {
            StartRow = startCell.Row
            EndRow = endCell.Row
            StartColumn = startCell.Column
            EndColumn = endCell.Column
        }

    member x.ExcelAddress =
        ExcelAddress(x.StartRow, x.StartColumn, x.EndRow, x.EndColumn)


[<DebuggerDisplay("{Range.Value.Address}")>]
type InputMatrixStream = 
    { Range: SingletonExcelRangeBase
      Shift: Shift
      ParsingAddress: ParsingAddress
      Logger: Logger }
with 
    member internal stream.OffsetedRange =
        SingletonExcelRangeBase.offset stream.Shift stream.Range
    

    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded

type OutputMatrixStreamResult<'result> =
    { IsSkip: bool
      Value: 'result }

type OutputMatrixStream<'result> =
    { Range: SingletonExcelRangeBase
      Shift: Shift
      Logger: Logger
      ParsingAddress: ParsingAddress
      Result: OutputMatrixStreamResult<'result> }

with 
    member internal stream.OffsetedRange =
        SingletonExcelRangeBase.offset stream.Shift stream.Range
    
    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift
          ParsingAddress = x.ParsingAddress
          Logger = x.Logger }

    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded

    member x.ShiftHorizontally(columns: int) =
        { x with 
            Shift = 
                (x.Shift, [1..columns])
                ||> List.fold(fun shift _ ->
                    Shift.applyDirection shift Direction.Horizontal shift
                )
            }

    member x.ShiftVertically(rows: int) =
        { x with 
            Shift = 
                (x.Shift, [1..rows])
                ||> List.fold(fun shift _ ->
                    Shift.applyDirection shift Direction.Vertical shift
                )
            }

    member x.ShiftBy(columns, rows: int) =
        x.ShiftHorizontally(columns).ShiftVertically(rows)


type ColumnRerangedResult(range: ExcelRangeBase, logger: Logger) =
    let inputStreamOfReranged logger (reranged: ExcelRangeBase) =
        let newAddr: ParsingAddress =
            ParsingAddress.OfRange reranged

        let range = 
            reranged.Offset(0, 0, 1, 1)
            |> SingletonExcelRangeBase.Create

        let resetedInputStream = 
            { Range = range
              Shift = Shift.Start
              ParsingAddress = newAddr
              Logger = logger }

        resetedInputStream

    
    member x.Range = range

    member val InputMatrixStream = inputStreamOfReranged logger range

    member x.SetRowCountTo(rowsCount) =
        let newRange =
            range.Offset(0, 0, rowsCount, range.Columns)
        ColumnRerangedResult(
            range = newRange,
            logger = logger
        )
        

    member x.SetColumnTo(columnsCount) =
        let newRange = 
            range.Offset(0, 0, range.Rows, columnsCount)
        
        ColumnRerangedResult(
            range = newRange,
            logger = logger
        )

    member x.RightOf(targetRange: SingletonExcelRangeBase) =
        let addr = 
            ExcelAddress(
                fromCol = targetRange.Column + 1,
                toColumn = x.Range.End.Column,
                fromRow = x.Range.Start.Row,
                toRow = x.Range.End.Row
            )

        let newRange = x.Range.Worksheet.Cells.[addr.Address]
        
        ColumnRerangedResult(
            range = newRange,
            logger = logger
        )

[<RequireQualifiedAccess>]
module OutputMatrixStream =


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

        ColumnRerangedResult(reranged, stream.Logger)


    let reRangeToEnd (stream: OutputMatrixStream<_>) = 
        let addr = stream.Range.Value.Address + ":" + stream.ParsingAddress.ExcelAddress.End.Address
        let reranged = stream.Range.Value.Worksheet.Cells.[addr]
        ColumnRerangedResult(reranged, stream.Logger)

    let reRangeTo (targetRange: SingletonExcelRangeBase) (stream: OutputMatrixStream<_>) = 
        let addr = stream.Range.Value.Address + ":" + targetRange.Address
        let reranged = stream.Range.Value.Worksheet.Cells.[addr]
        ColumnRerangedResult(reranged, stream.Logger)

    let topArea (stream: OutputMatrixStream<_>) =   
        match stream.ParsingAddress.StartRow < stream.Range.Row with 
        | true -> 
            let addr = 
                ExcelAddress(
                    fromCol = stream.Range.Column,
                    toColumn = stream.OffsetedRange.Column,
                    fromRow = stream.ParsingAddress.StartRow,
                    toRow = stream.Range.Row - 1)

            let reranged = stream.Range.Value.Worksheet.Cells.[addr.Address]
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

            let reranged = stream.Range.Value.Worksheet.Cells.[addr.Address]
            (reranged)
            |> Some

        | false -> None

    let reRangeRowTo rowsCount (stream: OutputMatrixStream<_>) = 
        let column = stream.ParsingAddress.EndColumn - stream.Range.Value.Start.Column + 1
        let reranged = stream.Range.Value.Offset(0, 0, rowsCount, column)
        ColumnRerangedResult(reranged, stream.Logger)

    let reRangeColumnTo columnsCount (stream: OutputMatrixStream<_>) = 
        let row = stream.ParsingAddress.EndRow - stream.Range.Value.Start.Row + 1
        let reranged = stream.Range.Value.Offset(0, 0, row, columnsCount)
        ColumnRerangedResult(reranged, stream.Logger)

    let applyDirectionToShift direction (preInputstream: InputMatrixStream) (stream: OutputMatrixStream<_>) =
        if stream.Result.IsSkip then stream
        else
            { stream with 
                Shift = Shift.applyDirection preInputstream.Shift direction stream.Shift }

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
                ExcelAddress(stream.Range.RangeTo(stream.OffsetedRange).Address), stream
            )

        let length = streamWithAddressList.Length

        let rec loop i (streamWithAddressList: list<ExcelAddress * OutputMatrixStream<'result>>) =
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

    member x.InvokeToResults = 
        x.Invoke >> List.map (fun outputStream -> outputStream.Result.Value)

    member x.Map(mapping) = 
        mapping invoke
        |> MatrixParser

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


let mxCellParserOp (cellParser: SingletonExcelRangeBase -> 'result option) =
    fun (stream: InputMatrixStream) ->
        let offsetedRange = stream.OffsetedRange
        let addr = offsetedRange.Value.Start
        let streamAddr = stream.ParsingAddress
        match addr.Row, addr.Column with 
        | Between(streamAddr.StartRow, streamAddr.EndRow),
            Between(streamAddr.StartColumn, streamAddr.EndColumn) ->
            match cellParser offsetedRange with 
            | Some result ->
                [
                    { Range = stream.Range 
                      Shift = stream.Shift 
                      ParsingAddress = stream.ParsingAddress
                      Result = 
                        { IsSkip = false
                          Value = result }
                      Logger = stream.Logger
                    }
                ]
            | None -> []
        | _ -> []
    |> MatrixParser
//let mxCellParser_Result (cellParser: SingletonExcelRangeBase -> Result<'ok, string>) =
//    mxCellParserOp (fun range ->
//        match cellParser range with 
//        | Result.Ok v -> Some v
//        | Result.Error error -> 
//    )

let mxCellParser (cellParser: CellParser) getResult =
    fun range ->
        let b = cellParser range
        if b then 
            Some (getResult range.Value)
        else None
    |> mxCellParserOp

let mxFParsec p =
    mxCellParserOp (pFParsec p)


let mxFParsecInt32 = mxFParsec (pint32) 

let mxText text =
    mxCellParser (pText text) ExcelRangeBase.getText

let mxTextf f =
    mxCellParser (pTextf f) ExcelRangeBase.getText

let mxNonEmpty = 
    mxTextf isTrimmedTextNotEmpty

let mxTextOp picker =
    mxCellParserOp (fun range ->
        picker range.Value.Text
    )


let mxRegex pattern =
    mxCellParser (pTextf (fun text -> 
        match text with 
        | ParseRegex.Head pattern _ -> true
        | _ -> false
    )) ExcelRangeBase.getText

let mxWord = mxRegex "\w" 

let mxSpace = mxCellParser pSpace ignore 

let mxStyleName styleName = mxCellParser (pStyleName styleName) ExcelRangeBase.getText

let mxAnySkip = 
    mxCellParser pAny ignore 

let mxAnyOrigin = 
    mxCellParser pAny ExcelRangeBase.getText 

let mxAddress address = 
    let targetAddress = new ExcelAddress(address)
    mxCellParserOp(fun range ->
        if ExcelAddress(range.Value.Address).Address = targetAddress.Address
        then Some range.Value.Text
        else None
    )

let mxAnyOriginObj = 
    mxCellParser pAny ExcelRangeBase.getValue 


let (||>>) p f = 
    MatrixParser.mapOutputStream (fun outputStream -> OutputMatrixStream.mapResultValue f outputStream) p

[<AutoOpen>]
module LoggerExtensions =
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
                        let addr = range.Value.Address
                        sprintf $"{name}\tResult: {addr}: {outputStream.Result.Value}"

                    logger.Log loggerLevel message

                //if not outputStreams.IsEmpty 
                //then logger.Log loggerLevel (sprintf "END %s:" name)

                outputStreams



let (|||>>) p f = 
    MatrixParser.mapOutputStream (fun outputStream ->
       OutputMatrixStream.mapResultValue (f (outputStream)) outputStream
    ) p


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


let internal inDebug (p: MatrixParser<_>) =
    fun inputStream -> 
        let m = p.Invoke inputStream
        match m with 
        | List.Some m -> m
        | List.None -> []
    |> MatrixParser
    |> MatrixParser.addLogger LoggerLevel.Important "Debug"

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

let private atLeastOne (p: MatrixParser<'a list>) =
    p
    |> MatrixParser.filterOutputStreamByResultValue (fun list ->
        not list.IsEmpty
    )

let private atLeastTwo (p: MatrixParser<'a list>) =
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

            | _ -> 
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


let mxManySkip direction pSkip maxSkipCount p =
    mxMany1SkipRetain direction pSkip maxSkipCount p
    ||>> (List.choose (fun v ->
        match v with 
        | Result.Ok ok -> Some ok
        | Result.Error _ -> None
    ))



/// at least one skip
let mxManySkip1 direction pSkip maxSkipCount p =
    mxManySkip direction pSkip maxSkipCount p
    |> atLeastOne



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
        let workSheet = outputStream.Range.Value.Worksheet
        let mergeCellId = ExcelWorksheet.getMergeCellIdOfRange workSheet.Cells.[outputStream.Result.Value.Address] workSheet
        mxMany1 direction (mxCellParser (fun range -> ExcelWorksheet.getMergeCellIdOfRange range.Value workSheet = mergeCellId) ExcelRangeBase.getAddress)
    ) id

let mxMerge direction =
    mxMergeWithAddresses direction
    ||>> (fun (start, addresses) ->
        start.Text
    )

let mxColMany p = mxMany Direction.Horizontal p

let mxColMany1 p = mxMany1 Direction.Horizontal p
let mxColMany1WithMaxCount maxCount p = mxMany1WithMaxCount Direction.Horizontal maxCount p



let mxColManySkip pSkip maxSkipCount p = mxManySkip Direction.Horizontal pSkip maxSkipCount p

let mxRowManySkip pSkip maxSkipCount p = mxManySkip Direction.Vertical pSkip maxSkipCount p

let mxRowMany1 p = mxMany1 Direction.Vertical p
let mxRowMany1WithMaxCount maxCount p = mxMany1WithMaxCount Direction.Vertical maxCount p

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

let private runMatrixParserForRangesWithStreamsAsResult_Common addr logger (ranges : seq<SingletonExcelRangeBase>) (p : MatrixParser<_>) =
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

    inputStreams 
    |> List.collect p.Invoke


let private runMatrixParserForRangesWithStreamsAsResult addr (ranges : seq<SingletonExcelRangeBase>) (p : MatrixParser<_>) =
    runMatrixParserForRangesWithStreamsAsResult_Common addr (new Logger()) ranges p

let private runMatrixParserForRanges addr (ranges : seq<SingletonExcelRangeBase>) (p : MatrixParser<_>) =

    let mses = runMatrixParserForRangesWithStreamsAsResult addr ranges p
    mses |> List.map (fun ms -> ms.Result.Value)


let runMatrixParserForRangeWithStreamsAsResult (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeBase.asRangeList range
    let mses = runMatrixParserForRangesWithStreamsAsResult address ranges p
    mses

let runMatrixParserForRangeWithStreamsAsResult2 (logger: Logger) (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let address = 
        ParsingAddress.OfRange range

    let ranges = 
        ExcelRangeBase.asRangeList range

    runMatrixParserForRangesWithStreamsAsResult_Common address logger ranges p

let runMatrixParserForRange2 logger (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult2 logger range p
    |> List.map (fun m -> m.Result.Value)


let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult range p
    |> List.map (fun m -> m.Result.Value)

let runMatrixParserForRangeWithoutRedundent (range : ExcelRangeBase) (p : MatrixParser<_>) =
    runMatrixParserForRangeWithStreamsAsResult range p
    |> OutputMatrixStream.removeRedundants
    |> List.map (fun m -> m.Result.Value)

let runMatrixParserWithStreamsAsResult (worksheet: ValidExcelWorksheet) (p: MatrixParser<_>) =
    let userRange = 
        worksheet.Value
        |> ExcelWorksheet.getUserRangeList

    let addr =
        { 
            StartRow = 1
            StartColumn = 1
            EndRow = ExcelWorksheet.getMaxRowNumber worksheet.Value
            EndColumn = ExcelWorksheet.getMaxColNumber worksheet.Value
        }

    runMatrixParserForRangesWithStreamsAsResult addr userRange p


let runMatrixParser (worksheet: ValidExcelWorksheet) (p: MatrixParser<_>) =
    runMatrixParserWithStreamsAsResult worksheet p
    |> List.map (fun m -> m.Result.Value)


let runMatrixParserWithStreamsAsResultSafe (worksheet: ValidExcelWorksheet) (p: MatrixParser<'result>) =

    let logger = new Logger()

    match runMatrixParserWithStreamsAsResult worksheet p with 
    | [] -> 
        match logger.Messages().IsEmpty with 
        | true -> failwithf "ParsingTarget: %s\nAll named parsed are parsed failured %A\nStackTrace:\n%s" (typeof<'result>.Name) p System.Environment.StackTrace 
        | false ->
            failwithf "ParsingTarget: %s\n%A\nStackTrace%s" (typeof<'result>.Name) (logger.Messages()) System.Environment.StackTrace 
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
