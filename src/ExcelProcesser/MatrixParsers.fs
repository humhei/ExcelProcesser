module ExcelProcesser.MatrixParsers

open OfficeOpenXml
open Extensions
open CellParsers
open FParsec.CharParsers

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
        if List.forall (fun shift -> length shift = 1) shifts then 
            match direction with 
            | Direction.Horizontal ->
                let chooser shift =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.Y = preCalculatedCoordinate.Y then Some { coordinate with X = coordinate.X + i }
                        else None

                    | Shift.Vertical (coordinate,i) ->
                        if coordinate.Y = preCalculatedCoordinate.Y then Some coordinate
                        else None
                    | _ -> None

                match List.tryPick chooser shifts with 
                | Some coordinate ->
                    Compose (Horizontal (coordinate,1) :: shifts)
                | None ->
                    failwith "not implemented"

            | Direction.Vertical ->
                let chooser shift =
                    match shift with 
                    | Shift.Horizontal (coordinate,i) ->
                        if coordinate.X = preCalculatedCoordinate.X then Some coordinate
                        else None
                    | Shift.Vertical (coordinate, i) ->
                        if coordinate.X = preCalculatedCoordinate.X then Some {coordinate with Y = coordinate.Y + i}
                        else None
                    | _ -> None

                match List.tryPick chooser shifts with 
                | Some coordinate ->
                    Compose (Vertical (coordinate, 1) :: shifts)
                | None ->
                    failwith "not implemented"
            | _ -> failwith "Invalid token"
        else
            failwith ""

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
            let coordinate2, i2 =
                match h with 
                | Horizontal (coordinate2, i2) | Vertical (coordinate2, i2) ->
                    coordinate2, i2
                | _ -> failwith "not implemented"

            if coordinate2 = calculatedCoordinate1 then
                match direction with 
                | Direction.Horizontal ->
                    redirect calculatedCoordinate1 direction (h :: t)
                        
                | Direction.Vertical ->
                    Compose (Vertical (coordinate2, i2 + 1) :: t)

                | _ -> failwith "Invalid token"
            else redirect calculatedCoordinate1 direction (h :: t)

        | Vertical (coordiante1, i1), Compose (h :: t) ->
            let calculatedCoordinate1 = { coordiante1 with Y = coordiante1.Y + i1 }
            let coordinate2, i2 =
                match h with 
                | Horizontal (coordinate2, i2) | Vertical (coordinate2, i2) ->
                    coordinate2, i2
                | _ -> failwith "not implemented"

            if coordinate2 = calculatedCoordinate1 then
                match direction with 
                | Direction.Horizontal ->
                    Compose (Horizontal (coordinate2, i2 + 1) :: t)
                        
                | Direction.Vertical ->
                    redirect calculatedCoordinate1 direction (h :: t)

                | _ -> failwith "Invalid token"
            else redirect calculatedCoordinate1 direction (h :: t)

        | _ -> failwith ""

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

            | h :: _ ->
                offset h range
    

type InputMatrixStream = 
    { Range: ExcelRangeBase
      Shift: Shift }
with 
    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded

type OutputMatrixStreamResult<'result> =
    { IsSkip: bool
      Value: 'result }

type OutputMatrixStream<'result> =
    { Range:  ExcelRangeBase
      Shift: Shift
      Result: OutputMatrixStreamResult<'result> }

with 
    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift }

    member private x.LastCellShift = x.Shift.Last

    member private x.FoldedShift = x.Shift.Folded

[<RequireQualifiedAccess>]
module OutputMatrixStream =

    let reRange (stream: OutputMatrixStream<_>) = 
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
            |> List.max

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
            |> List.max

        stream.Range.Offset(0, 0, maxVertical + 1, maxHorizontal + 1)

    let applyDirectionToShift direction (preInputstream: InputMatrixStream) (stream: OutputMatrixStream<_>) =
        if stream.Result.IsSkip then stream
        else
            { stream with 
                Shift = Shift.applyDirection preInputstream.Shift direction stream.Shift }

    let mapResult mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = mapping stream.Result }

    let mapResultValue mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = mapping stream.Result.Value
              IsSkip = stream.Result.IsSkip }}


[<RequireQualifiedAccess>]
type MatrixStream<'result> =
    | Input of InputMatrixStream
    | Output of previousInputStream: InputMatrixStream * OutputMatrixStream<'result> list



type MatrixParser<'result> = InputMatrixStream -> OutputMatrixStream<'result> list

[<RequireQualifiedAccess>]
module private List =
    let (|Some|None|) (list: 'a list) = 
        if list.Length = 0 then None
        else Some list

[<RequireQualifiedAccess>]
module MatrixParser =
    let mapOutputStream f p =
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.map f outputStreams

    let collectOutputStream f p : MatrixParser<_> =
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.collect f outputStreams

    let filterOutputStreamByResultValue f p =
        fun (inputStream: InputMatrixStream) ->
            let (outputStreams: OutputMatrixStream<'result> list) = p inputStream
            List.filter (fun outputStream -> f outputStream.Result.Value) outputStreams

let mxCellParserOp (cellParser: ExcelRangeBase -> 'result option) =
    fun (stream: InputMatrixStream) ->
        let offsetedRange = ExcelRangeBase.offset stream.Shift stream.Range
        match cellParser offsetedRange with 
        | Some result ->
            [
                { Range = stream.Range 
                  Shift = stream.Shift 
                  Result = 
                    { IsSkip = false
                      Value = result }
                }
            ]
        | None -> 
            //printfn "Parsing %O with %A failed" offsetedRange cellParser
            []

let mxCellParser (cellParser: CellParser) getResult =
    fun range ->
        let b = cellParser range
        if b then 
            Some (getResult range)
        else None
    |> mxCellParserOp

let mxFParsec p =
    mxCellParserOp (pFParsec p)

let mxFParsecInt32 inputStream = mxFParsec (pint32) inputStream

let mxText text =
    mxCellParser (pText text) ExcelRangeBase.getText

let mxTextf f =
    mxCellParser (pTextf f) ExcelRangeBase.getText

let mxSpace inputStream = mxCellParser pSpace ignore inputStream

let mxStyleName styleName = mxCellParser (pStyleName styleName) ExcelRangeBase.getText

let mxAnySkip inputStream = 
    mxCellParser pAny ignore inputStream

let mxAnyOrigin inputStream = 
    mxCellParser pAny ExcelRangeBase.getText inputStream

let (||>>) p f = 
    MatrixParser.mapOutputStream (fun outputStream -> OutputMatrixStream.mapResultValue f outputStream) p

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
        match p1 inputStream with
        | List.Some streams -> streams
        | _ -> p2 inputStream


/// p1 && not p2
let (<&!>) (p1: MatrixParser<'result>) (p2: MatrixParser<'exclude>) = 
    fun inputStream ->
        match p2 inputStream with 
        | List.Some _ -> []
        | _ -> p1 inputStream

let (<&>) (p1: MatrixParser<'result>) (p2: MatrixParser<'predicate>) = 
    fun inputStream ->
        match p1 inputStream with 
        | List.Some outputStreams1 -> 
            match p2 inputStream with 
            | List.Some _ -> outputStreams1
            | List.None -> []
        | List.None -> []


let internal inDebug p =
    fun inputStream -> 
        let m = p inputStream
        match m with 
        | List.Some m -> 
            m
        | List.None -> []

let private pipe2RelativelyWithTupleStreamsReturn (direction: Direction) (p1: MatrixParser<'result1>) (buildP2: OutputMatrixStream<'result1> -> MatrixParser<'result2>) f =

    fun inputstream1 ->

        let newStreams1 = p1 inputstream1
        match newStreams1 with 
        | List.Some newStreams1 ->
            
            newStreams1
            |> List.collect (fun newStream1 ->
                let p2 = buildP2 newStream1
                let inputStream2 = (OutputMatrixStream.applyDirectionToShift direction inputstream1 newStream1).AsInputStream
                p2 inputStream2
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

let mxManyWithMaxCount direction (maxCount: int option) (p: MatrixParser<'result>) = 
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
            let outputStreams = pipe2 direction skip many1 id inputstream
            match outputStreams with 
            | List.Some outputStreams ->
                outputStreams
                |> List.filter (fun outputStream ->
                    let errors, values = outputStream.Result.Value
                    values.Length > 0
                )
            | List.None -> []

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
    pipe2RelativelyWithTupleStreamsReturn direction (fun a -> many1 a) (fun _ -> mxManySkipRetain direction pSkip maxSkipCount p) (fun (a,b) ->
        List.map Result.Ok a @ b
    ) >> List.map (fun (stream1, stream2) ->
        if stream2.Result.IsSkip then 
            stream1 |> OutputMatrixStream.mapResultValue (List.map Result.Ok)
        else stream2
    )


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
        p direction inputStream

let mxUntil direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2 direction (mxManyWithMaxCount direction maxCount (pPrevious <&!> pLast)) pLast id
   

let mxUntilBacktrackLast direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2RelativelyWithTupleStreamsReturn direction (mxManyWithMaxCount direction maxCount (pPrevious <&!> pLast)) (fun _ -> pLast) id
    >> List.map fst

let mxUntil1 direction maxCount pPrevious (pLast: MatrixParser<'result>) =
    pipe2 direction (mxMany1WithMaxCount direction maxCount (pPrevious <&!> pLast)) pLast id

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


type MergeStarterResult =
    { Address: string 
      Text: string }


let mxMergeStarter inputStream = 
    mxCellParser pMergeStarter (fun range -> { Address = range.Address; Text = range.Text}) inputStream
    

let mxMerge direction =
    pipe2Relatively direction mxMergeStarter (fun outputStream ->
        let workSheet = outputStream.Range.Worksheet
        let mergeCellId = ExcelWorksheet.getMergeCellIdOfRange workSheet.Cells.[outputStream.Result.Value.Address] workSheet
        mxMany1 direction (mxCellParser (fun range -> ExcelWorksheet.getMergeCellIdOfRange range workSheet = mergeCellId) ExcelRangeBase.getAddress)
    ) id

let mxColMany1 p = mxMany1 Direction.Horizontal p

let mxColManySkip pSkip maxSkipCount p = mxManySkip Direction.Horizontal pSkip maxSkipCount p

let mxRowManySkip pSkip maxSkipCount p = mxManySkip Direction.Vertical pSkip maxSkipCount p

let mxRowMany1 p = mxMany1 Direction.Vertical p

let c2 p1 p2 =
    pipe2 Direction.Horizontal p1 p2 id

/// R = Relatively
let c2R p1 buildP2 = 
    pipe2Relatively Direction.Horizontal p1 buildP2 id

let c3 p1 p2 p3 =
    pipe3 Direction.Horizontal p1 p2 p3 id

let r2 p1 p2 =
    pipe2 Direction.Vertical p1 p2 id


/// R = Relatively
let r2R p1 buildP2 =
    pipe2Relatively Direction.Vertical p1 buildP2 id

let r3 p1 p2 p3 = 
    pipe3 Direction.Vertical p1 p2 p3 id

let runMatrixParserForRangesWithStreamsAsResult (ranges : seq<ExcelRangeBase>) (p : MatrixParser<_>) =
    let inputStreams = 
        ranges 
        |> List.ofSeq
        |> List.map (fun range ->
            { Range = range 
              Shift = Shift.Start }
        )

    inputStreams 
    |> List.collect p


let runMatrixParserForRanges (ranges : seq<ExcelRangeBase>) (p : MatrixParser<_>) =
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)


let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let ranges = ExcelRangeBase.asRangeList range
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParser (worksheet: ExcelWorksheet) (p: MatrixParser<_>) =
    match worksheet with 
    | null -> 
        failwithf "Work sheet is empty, please check xlsx file"
    | _ ->
        let userRange = 
            worksheet
            |> ExcelWorksheet.getUserRangeList

        runMatrixParserForRanges userRange p


