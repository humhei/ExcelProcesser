module ExcelProcesser.MatrixParsers

open OfficeOpenXml
open Extensions
open CellParsers

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
    | Start
    | Vertical of int
    | Horizontal of int

[<RequireQualifiedAccess>]
module RelativeShift =
    let getNumber = function 
        | RelativeShift.Start -> 1
        | RelativeShift.Horizontal i -> i
        | RelativeShift.Vertical i -> i


    let plus direction shift1 shift2 =
        match shift1, shift2 with 
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

        | RelativeShift.Horizontal i, RelativeShift.Horizontal j -> RelativeShift.Horizontal (i + j)

        | RelativeShift.Horizontal i, RelativeShift.Vertical j -> RelativeShift.Vertical j

        | RelativeShift.Vertical i, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal 1

            | Direction.Vertical ->
                RelativeShift.Vertical (i + 1)

            | _ -> failwith "Invalid token"

        | RelativeShift.Vertical i, RelativeShift.Horizontal j -> RelativeShift.Horizontal j
        | RelativeShift.Vertical i, RelativeShift.Vertical j -> RelativeShift.Vertical(i + j)

type Shift =
    | Start
    | Vertical of Coordinate * int
    | Horizontal of Coordinate * int
    | Compose of Shift list



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

    let rec applyDirection (archivement: RelativeShift) (direction: Direction) shift = 
        let archivementCount = RelativeShift.getNumber archivement

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
                Compose([Horizontal({ coordinate with Y = coordinate.Y + i - archivementCount + 1 }, 1); Vertical(coordinate, i)])

            | _ -> failwith "Invalid token"

        | Horizontal (coordinate, i) ->
            match direction with 
            | Direction.Vertical ->
                Compose([Vertical({ coordinate with X = coordinate.X + i - archivementCount + 1}, 1); Horizontal(coordinate, i)])

            | Direction.Horizontal ->
                Horizontal (coordinate, i + 1)

            | _ -> failwith "Invalid token"


        | Compose (shifts) ->
            match shifts with
            | [] -> failwith "compose shifts cannot be empty after start"
            | h :: t ->
                Compose (applyDirection archivement direction h :: t)


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


[<RequireQualifiedAccess>]
module OutputMatrixStream =

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
    | Output of OutputMatrixStream<'result>


type MatrixParser<'result> = InputMatrixStream -> OutputMatrixStream<'result> option

[<RequireQualifiedAccess>]
module MatrixParser =
    let mapOutputStream f p =
        fun (inputStream: InputMatrixStream) ->
            let (outputStream: OutputMatrixStream<'result> option) = p inputStream
            match outputStream with 
            | Some outputStream ->
                f outputStream
            | None -> None

let mxCellParserOp (cellParser: ExcelRangeBase -> 'result option) =
    fun (stream: InputMatrixStream) ->
        let offsetedRange = ExcelRangeBase.offset stream.Shift stream.Range
        match cellParser offsetedRange with 
        | Some result ->

            Some 
                { Range = stream.Range 
                  Shift = stream.Shift 
                  Result = 
                    { RelativeShift = RelativeShift.Start
                      Value = result }
                }
        | None -> None

let mxCellParser (cellParser: CellParser) getResult =
    fun range ->
        let b = cellParser range
        if b then 
            Some (getResult range)
        else None
    |> mxCellParserOp

let mxFParsec text =
    mxCellParserOp (pFParsec text)

let mxText text =
    mxCellParser (pText text) ExcelRangeBase.getText

let mxTextf f =
    mxCellParser (pTextf f) ExcelRangeBase.getText

let mxSpace inputStream = mxCellParser pSpace ignore inputStream

let mxAny inputStream = 
    mxCellParser pAny ignore inputStream


let (|||>) p f = 
    MatrixParser.mapOutputStream (fun outputStream ->
       Some (OutputMatrixStream.mapResultValue f outputStream) 
    ) p

let mxOR (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>) =
    let p1 = 
        p1 |||> Choice1Of2

    let p2 = 
        p2 |||> Choice2Of2

    fun inputStream ->
        match p1 inputStream with
        | Some outputStream ->
            Some outputStream
        | None -> p2 inputStream


let pipe2 (direction: Direction) (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>) f =

    fun inputstream1 ->

        let newStream1 = p1 inputstream1
        match newStream1 with 
        | Some newStream1 ->
            let inputStream2 = (OutputMatrixStream.applyDirectionToShift direction newStream1).AsInputStream
            
            match p2 inputStream2 with 
            | Some newStream2 ->
                OutputMatrixStream.mapResult (fun result2 -> 
                    { RelativeShift = RelativeShift.plus direction newStream1.Result.RelativeShift result2.RelativeShift
                      Value = f (newStream1.Result.Value, result2.Value) }
                ) newStream2
                |> Some

            | None -> None
        | None -> None




let pipe3 direction p1 p2 p3 f =
    pipe2 direction (pipe2 direction p1 p2 id) p3 (fun ((a, b), c) ->
        f (a, b, c)
    )


let private mxManyWithMaxCount direction (maxCount: int option) (p: MatrixParser<'result>) = 
    
    fun inputStream ->
        let rec loop stream (accum: OutputMatrixStream<'result> list) =
            let isReachMaxCount =
                match maxCount with 
                | Some maxCount -> 
                    accum.Length >= maxCount
                | None -> false

            if isReachMaxCount then accum
            else
                match stream with
                | MatrixStream.Input inputStream ->
                    match p inputStream with 
                    | Some outputStream ->
                        loop (MatrixStream.Output outputStream) (outputStream :: accum) 
                    | None -> accum

                | MatrixStream.Output outputStream ->
                    let inputStream = (OutputMatrixStream.applyDirectionToShift direction outputStream).AsInputStream

                    match p inputStream with 
                    | Some outputStream ->
                        let newOutputStream = 
                            OutputMatrixStream.mapResult (fun result2 -> 
                                { RelativeShift = RelativeShift.plus direction outputStream.Result.RelativeShift result2.RelativeShift
                                  Value = result2.Value }
                            ) outputStream
                        loop (MatrixStream.Output newOutputStream) (newOutputStream :: accum)

                    | None -> accum


        let outputStreams = loop (MatrixStream.Input inputStream) []
        match outputStreams with 
        | h :: t ->
            { Range = h.Range 
              Shift = h.Shift 
              Result = 
                { RelativeShift = h.Result.RelativeShift
                  Value = 
                    outputStreams 
                    |> List.map (fun outputStream ->
                          outputStream.Result.Value 
                    )
                    |> List.rev
                }
            }
            |> Some

        | _ -> 
            { Range = inputStream.Range
              Shift = inputStream.Shift
              Result = 
                { RelativeShift = RelativeShift.Start
                  Value = []
                }
            }
            |> Some

let mxMany direction p = mxManyWithMaxCount direction None p

let mxMany1 direction p =
        mxMany direction p
        |> MatrixParser.mapOutputStream (fun outputStream ->
            if outputStream.Result.Value.IsEmpty then None
            else Some outputStream
        )



let mxManySkip direction pSkip maxSkipCount p =
    let skip = 
        mxManyWithMaxCount direction (Some maxSkipCount) pSkip 

    let many1 = mxMany1 direction p

    let piped = 
        fun (inputstream: InputMatrixStream) ->
            let outputStream = pipe2 direction skip many1 snd inputstream
            outputStream

    pipe2 direction many1 (mxMany direction piped) (fun (a,b) ->
        a :: b
        |> List.concat
    )

let mxUntil maxCount (p: MatrixParser<'result>) =
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

        let rec greed accum stream =
            let isReachMaxCount = 
                match maxCount with 
                | Some maxCount -> accum > maxCount
                | None -> false

            if isReachMaxCount then None
            else
                match stream with 
                | MatrixStream.Input inputStream ->
                    match p inputStream with 
                    | Some outputStream ->
                        Some outputStream
                    | None ->
                        match mxAny inputStream with 
                        | Some outputStream ->
                            greed (accum + 1) (MatrixStream.Output outputStream) 
                        /// mxAny match everything
                        | None -> failwith "Invalid token"

                | MatrixStream.Output outputStream ->
                    let inputStream = (OutputMatrixStream.applyDirectionToShift direction outputStream).AsInputStream

                    match p inputStream with 
                    | Some stream -> Some stream

                    | None -> 
                        match mxAny inputStream with 
                        | Some outputStream ->
                            let newOutputStream = 
                                OutputMatrixStream.mapResult (fun result2 -> 
                                    { RelativeShift = RelativeShift.plus direction outputStream.Result.RelativeShift result2.RelativeShift
                                      Value = result2.Value }
                                ) outputStream
                            greed (accum + 1) (MatrixStream.Output newOutputStream)
                        /// mxAny match everything
                        | None -> failwith "Invalid token"

        greed 0 (MatrixStream.Input inputStream)



let cm p = mxMany1 Direction.Horizontal p

let mxManySkipCol pSkip maxSkipCount p = mxManySkip Direction.Horizontal pSkip maxSkipCount p

let mxManySkipRow pSkip maxSkipCount p = mxManySkip Direction.Vertical pSkip maxSkipCount p

let rm p = mxMany1 Direction.Vertical p

let c2 p1 p2 =
    pipe2 Direction.Horizontal p1 p2 id

let c3 p1 p2 p3 =
    pipe3 Direction.Horizontal p1 p2 p3 id

let r2 p1 p2 =
    pipe2 Direction.Vertical p1 p2 id

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
    |> List.choose p


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


