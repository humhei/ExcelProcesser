module ExcelProcesser.MatrixParsers

open OfficeOpenXml
open Extensions
open CellParsers
open Microsoft.FSharp.Reflection
open System.Collections.Concurrent
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

    let retype unboxEx (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = unboxEx stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}

    let untype (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = box stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}

[<RequireQualifiedAccess>]
type MatrixStream<'result> =
    | Input of InputMatrixStream
    | Output of OutputMatrixStream<'result>

type MatrixParserPort<'input, 'result> = 
    { Input: 'input
      ResultGetter: (ExcelRangeBase -> 'result)}

[<RequireQualifiedAccess>]
module MatrixParserPort =
    let mapResult mapping port =
        { Input = port.Input 
          ResultGetter = port.ResultGetter >> mapping }

type MatrixParserContent<'result> =
    | Text of MatrixParserPort<string, 'result>
    | Space of MatrixParserPort<unit, 'result>

[<RequireQualifiedAccess>]
module MatrixParserContent =
    let cellParser = function
        | Text port -> pText port.Input
        | Space _ -> pSpace

    let resultGetter = function 
        | Text port -> port.ResultGetter
        | Space port -> port.ResultGetter 

    let untype = function
        | Text port -> 
            Text (MatrixParserPort.mapResult box port)
        | Space port ->
            Space (MatrixParserPort.mapResult box port)

    let retype = function
        | Text port -> 
            Text (MatrixParserPort.mapResult unbox port)
        | Space port ->
            Space (MatrixParserPort.mapResult unbox port)

type MatrixParserOperator =
    | OR of MatrixParser<obj> * MatrixParser<obj>


and MatrixParser<'result> =
    | Content of MatrixParserContent<'result>
    | Operator of MatrixParserOperator


[<RequireQualifiedAccess>]
module MatrixParser =

    let untype = function
        | Content content -> 
            MatrixParserContent.untype content
            |> Content
        | Operator operator -> Operator operator

    let retype = function
        | Content content -> 
            MatrixParserContent.retype content
            |> Content
        | Operator operator -> Operator operator


    let private resultTpUciesCache = new ConcurrentDictionary<Type, UnionCaseInfo []>()

    let streamTransfer (p: MatrixParser<'result>) : InputMatrixStream -> option<OutputMatrixStream<'result>> =
        let p = untype p

        let rec loop (p: MatrixParser<obj>) =
            match p with
            | MatrixParser.Content content ->
                fun (inputStream: InputMatrixStream) ->
                    let toOpt cellParser getResult =
                        fun range ->
                            if cellParser range then Some (getResult range)
                            else None

                    let cellParserOpt = toOpt (MatrixParserContent.cellParser content) (MatrixParserContent.resultGetter content)
                    let offsetedRange = ExcelRangeBase.offset inputStream.Shift inputStream.Range
                    match cellParserOpt offsetedRange with 
                    | Some result ->
                        { Range = inputStream.Range 
                          Shift = inputStream.Shift 
                          Result = 
                            { RelativeShift = RelativeShift.Start
                              Value = result }
                        } 
                        |> Some
                    | None -> None

            | MatrixParser.Operator operator ->
                fun (inputStream: InputMatrixStream) ->
                    match operator with 
                    | OR (p1, p2) ->
                        match loop p1 inputStream with 
                        | Some outputStream ->
                            (OutputMatrixStream.mapResultValue (fun v -> Choice1Of2 v) outputStream)
                            |> OutputMatrixStream.untype
                            |> Some

                        | None ->
                            match loop (retype p2) inputStream with
                            | Some outputStream ->
                                (OutputMatrixStream.mapResultValue (fun v -> Choice2Of2 v) outputStream)
                                |> OutputMatrixStream.untype
                                |> Some
                            | None -> None
        
        fun inputStream ->       
            match loop p inputStream with 
            | Some outputStream ->

                let rec unboxEx (v: obj) =
                    match v with 
                    | :? Choice<obj, obj> as v ->
                        let resultTp = typeof<'result>
                        let ucies =
                            resultTpUciesCache.GetOrAdd(resultTp, fun _ ->
                                let choiceTp = 
                                    let generics = resultTp.GetGenericArguments()
                                    typedefof<Choice<_,_>>.MakeGenericType(generics)

                                FSharpType.GetUnionCases choiceTp
                            )

                        match v with 
                        | Choice1Of2 v1 ->  
                            FSharpValue.MakeUnion(ucies.[0], [|v1|])
                            |> unbox
                        | Choice2Of2 v2 ->
                            FSharpValue.MakeUnion(ucies.[1], [|v2|])
                            |> unbox
                    | _ -> unbox v

                Some (OutputMatrixStream.retype unboxEx outputStream)
            | None -> None


let mxText (text: string) =
    { Input = text 
      ResultGetter = ExcelRangeBase.getText }
    |> Text
    |> Content

let mxOR (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>): MatrixParser<Choice<'result1, 'result2>> =
    (MatrixParser.untype p1,MatrixParser.untype p2)
    |> OR
    |> Operator


let runMatrixParserForRangesWithStreamsAsResult (ranges : seq<ExcelRangeBase>) (p : MatrixParser<'result>) : OutputMatrixStream<'result> list =
    let inputStreams = 
        ranges 
        |> List.ofSeq
        |> List.map (fun range ->
            { Range = range 
              Shift = Shift.Start }
        )

    let streamTransfer = MatrixParser.streamTransfer p

    inputStreams 
    |> List.choose streamTransfer


let runMatrixParserForRanges (ranges : seq<ExcelRangeBase>) (p : MatrixParser<'result>) : 'result list =
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<'result>) : 'result list =
    let ranges = ExcelRangeBase.asRanges range
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParser (worksheet: ExcelWorksheet) (p : MatrixParser<'result>) : 'result list =
    let userRange = 
        worksheet
        |> ExcelWorksheet.getUserRange

    runMatrixParserForRanges userRange p

