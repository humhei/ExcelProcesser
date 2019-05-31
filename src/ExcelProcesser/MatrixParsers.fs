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


[<RequireQualifiedAccess>]
type RelativeShift =
    | Skip
    | Start
    | Vertical of int
    | Horizontal of int

[<RequireQualifiedAccess>]
module RelativeShift =
    let getNumber = function 
        | RelativeShift.Skip -> 0
        | RelativeShift.Start -> 1
        | RelativeShift.Horizontal i -> i
        | RelativeShift.Vertical i -> i


    let plus direction shift1 shift2 =
        match shift1, shift2 with 
        | RelativeShift.Skip, _ -> shift2
        | _, RelativeShift.Skip -> shift1
        | RelativeShift.Start, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal 2
            | Direction.Vertical -> RelativeShift.Vertical 2
            | _ -> failwith "Invalid token"

        | RelativeShift.Start, RelativeShift.Horizontal i -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal (i + 1)
            | Direction.Vertical -> RelativeShift.Horizontal i
            | _ -> failwith "Invalid token"

        | RelativeShift.Start, RelativeShift.Vertical i ->
            match direction with 
            | Direction.Horizontal -> RelativeShift.Vertical(i)

            | Direction.Vertical ->
                RelativeShift.Vertical (i + 1)

            | _ -> failwith "Invalid token"


        | RelativeShift.Horizontal i, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal (i + 1)
            | Direction.Vertical -> RelativeShift.Vertical 1
            | _ -> failwith "Invalid token"

        | RelativeShift.Horizontal i, RelativeShift.Horizontal j -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal (i + j)
            | Direction.Vertical -> RelativeShift.Horizontal j
            | _ -> failwith "Invalid token"

        | RelativeShift.Horizontal i, RelativeShift.Vertical j -> RelativeShift.Vertical j

        | RelativeShift.Vertical i, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal 1

            | Direction.Vertical ->
                RelativeShift.Vertical (i + 1)

            | _ -> failwith "Invalid token"

        | RelativeShift.Vertical i, RelativeShift.Horizontal j -> RelativeShift.Horizontal j
        | RelativeShift.Vertical i, RelativeShift.Vertical j -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Vertical j
            | Direction.Vertical -> RelativeShift.Vertical (i + j)
            | _ -> failwith "Invalid token"

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

    let rec getCoordinate = function
        | Start -> Coordinate.origin
        | Horizontal (coordinate, _) -> coordinate
        | Vertical (coordinate, _) -> coordinate
        | Compose shifts ->
            match shifts with
            | [] -> failwith "compose shifts cannot be empty after start"
            | h :: t ->
                getCoordinate h

    let rec applyDirection (relativeShift: RelativeShift) (direction: Direction) shift = 
        
        match relativeShift with 
        | RelativeShift.Skip -> shift
        | _ ->
            let relativeShiftNumber = RelativeShift.getNumber relativeShift

            match shift with
            | Start ->
                match direction with 
                | Direction.Vertical ->
                    Vertical (Coordinate.origin, 1)

                | Direction.Horizontal ->
                    Horizontal (Coordinate.origin, 1)

                | _ -> failwith "Invalid token"

            | Vertical (coordinate, i) -> 
                match direction with 
                | Direction.Vertical ->
                    Vertical (coordinate, i + 1)

                | Direction.Horizontal ->
                    Compose([Horizontal({ coordinate with Y = coordinate.Y + i - relativeShiftNumber + 1 }, 1); Vertical(coordinate, i)])

                | _ -> failwith "Invalid token"

            | Horizontal (coordinate, i) ->
                match direction with 
                | Direction.Vertical ->
                    Compose([Vertical({ coordinate with X = coordinate.X + i - relativeShiftNumber + 1}, 1); Horizontal(coordinate, i)])

                | Direction.Horizontal ->
                    Horizontal (coordinate, i + 1)

                | _ -> failwith "Invalid token"


            | Compose (shifts) ->
                match shifts with
                | [] -> failwith "compose shifts cannot be empty after start"
                | h :: t ->
                    Compose (applyDirection relativeShift direction h :: t)


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
    member x.LastCellShift = x.Shift.Last


type OutputMatrixStreamResult<'result> =
    { RelativeShift: RelativeShift
      Value: 'result }

type OutputMatrixStream<'result> =
    { Range:  ExcelRangeBase
      Shift: Shift
      Result: OutputMatrixStreamResult<'result> }

with 
    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift }

    member x.LastCellShift = x.Shift.Last

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

    let applyDirectionToShift direction (stream: OutputMatrixStream<_>) =
        { stream with 
            Shift = Shift.applyDirection stream.Result.RelativeShift direction stream.Shift }

    let mapResult mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = mapping stream.Result }

    let mapResultValue mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = mapping stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}


[<RequireQualifiedAccess>]
type MatrixStream<'result> =
    | Input of InputMatrixStream
    | Output of OutputMatrixStream<'result> list



type MatrixParser<'result> = InputMatrixStream -> OutputMatrixStream<'result> list

[<RequireQualifiedAccess>]
module List =
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
                    { RelativeShift = RelativeShift.Start
                      Value = result }
                }
            ]
        | None -> []

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

let pipe2RelativelyWithTupleStreamsReturn (direction: Direction) (p1: MatrixParser<'result1>) (buildP2: OutputMatrixStream<'result1> -> MatrixParser<'result2>) f =

    fun inputstream1 ->

        let newStreams1 = p1 inputstream1
        match newStreams1 with 
        | List.Some newStreams1 ->
            
            newStreams1
            |> List.collect (fun newStream1 ->
                let p2 = buildP2 newStream1
                let inputStream2 = (OutputMatrixStream.applyDirectionToShift direction newStream1).AsInputStream
                match p2 inputStream2 with 
                | List.Some newStreams2 ->
                    List.map (
                        fun newStream2 ->
                            let stream2RelativeShift = newStream2.Result.RelativeShift
                            let stream2ResultValue = newStream2.Result.Value

                            let newStream2 = 
                                { Range = newStream2.Range
                                  Result = 
                                      { RelativeShift = RelativeShift.plus direction newStream1.Result.RelativeShift stream2RelativeShift
                                        Value = f (newStream1.Result.Value, stream2ResultValue) }
                                  Shift = 
                                      match stream2RelativeShift with 
                                      | RelativeShift.Skip -> newStream1.Shift 
                                      | _ -> newStream2.Shift }

                            newStreams1, newStream2

                    ) newStreams2

                | List.None -> []
            )
            
        | List.None -> []
        


let pipe2Relatively (direction: Direction) (p1: MatrixParser<'result1>) (buildP2: OutputMatrixStream<'result1> -> MatrixParser<'result2>) f =
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
        outputStream.Result.RelativeShift = RelativeShift.Skip

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

                        yield! loop (MatrixStream.Output outputStreams) (accum @ outputStreams) 

                    | List.None -> yield []

                | MatrixStream.Output outputStreams1 ->
                    if outputStreams1.Length = 0 then yield accum
                    else
                        yield!
                            outputStreams1
                            |> List.collect (fun outputStream1 -> [
                                let inputStream = (OutputMatrixStream.applyDirectionToShift direction outputStream1).AsInputStream
                                match p inputStream with 
                                | List.Some outputStreams2 ->
                                    let skip, outputStreams2 = List.partition isSkip outputStreams2

                                    yield! List.replicate skip.Length []

                                    let outputStreams2 = 
                                        outputStreams2
                                        |> List.map (fun outputStream2 ->
                                            OutputMatrixStream.mapResult (fun result2 -> 
                                                { RelativeShift = RelativeShift.plus direction outputStream1.Result.RelativeShift result2.RelativeShift
                                                  Value = result2.Value }
                                            ) outputStream2
                                        )

                                    yield! loop (MatrixStream.Output outputStreams2) (accum @ outputStreams2)

                                | List.None ->  yield accum
                            ]
                            )
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
                    { RelativeShift = last.Result.RelativeShift
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
                    { RelativeShift = RelativeShift.Skip
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
let private mxMany1SkipRetain direction pSkip maxSkipCount p =
    let many1 = mxMany1 direction p

    pipe2 direction many1 (mxManySkipRetain direction pSkip maxSkipCount p) (fun (a,b) ->
        List.map Result.Ok a @ b
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
        let mergeCellId = ExcelWorksheet.getMergeCellId workSheet.Cells.[outputStream.Result.Value.Address] workSheet
        mxMany1 direction (mxCellParser (fun range -> ExcelWorksheet.getMergeCellId range workSheet = mergeCellId) ExcelRangeBase.getAddressOfRange)
    ) id

let cm p = mxMany1 Direction.Horizontal p

let mxColManySkip pSkip maxSkipCount p = mxManySkip Direction.Horizontal pSkip maxSkipCount p

let mxRowManySkip pSkip maxSkipCount p = mxManySkip Direction.Vertical pSkip maxSkipCount p

let rm p = mxMany1 Direction.Vertical p

let c2 p1 p2 =
    pipe2 Direction.Horizontal p1 p2 id

let c2R p1 p2 = 
    pipe2 Direction.Horizontal p1 p2 id

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
    let ranges = ExcelRangeBase.asRanges range
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParser (worksheet: ExcelWorksheet) (p: MatrixParser<_>) =
    let userRange = 
        worksheet
        |> ExcelWorksheet.getUserRange

    runMatrixParserForRanges userRange p


