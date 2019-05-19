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
            | [] -> failwith "shifts cannot be empty"
            | h :: t ->
                getCoordinate h

    let rec applyDirection (direction: Direction) = function 

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
                Compose([Horizontal({ coordinate with Y = coordinate.Y + i }, 1); Vertical(coordinate, i)])

            | _ -> failwith "Invalid token"

        | Horizontal (coordinate, i) ->
            match direction with 
            | Direction.Vertical ->
                Compose([Vertical({ coordinate with X = coordinate.X + i }, 1); Horizontal(coordinate, i)])

            | Direction.Horizontal ->
                Horizontal (coordinate, i + 1)

            | _ -> failwith "Invalid token"


        | Compose (shifts) ->
            match shifts with
            | [] -> failwith "shifts cannot be empty"
            | h :: t ->
                Compose (applyDirection direction h :: t)


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
            | [] -> failwith "shifts cannot be empty"
            //| a :: b :: _ ->
            //    match b with 
            //    | Horizontal (coordinate, i) ->
            //        failwith ""
            //    | Vertical (coordinate1, i) ->
            //        match a with
            //        | Horizontal (coordinate2, j) ->
            //            range
            //            |> offset a 
            //            |> offset b
            //        | Vertical _ -> failwith "Horizontal shift should be linked to Vertical shift"

            | h :: _ ->
                offset h range
    

type InputMatrixStream = 
    { Range: ExcelRangeBase
      Shift: Shift }

[<RequireQualifiedAccess>]
module InputMatrixStream =
    let applyDirectionToShift direction (stream: InputMatrixStream) =
        { stream with 
            Shift = Shift.applyDirection direction stream.Shift }


type OutputMatrixStream<'result> =
    { Range:  ExcelRangeBase
      Shift: Shift
      Result: 'result }

with 
    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift
        }

[<RequireQualifiedAccess>]
module OutputMatrixStream =
    let mapM mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = mapping stream.Result }

type MatrixParser<'result> = InputMatrixStream -> OutputMatrixStream<'result> option

let mxCellParser (cellParser: CellParser) getResult =

    fun (stream: InputMatrixStream) ->
        let offsetedRange = ExcelRangeBase.offset stream.Shift stream.Range
        if cellParser offsetedRange then 
            Some 
                { Range = stream.Range 
                  Shift = stream.Shift 
                  Result = getResult offsetedRange }
        else 
            None

let mxText text =
    mxCellParser (pText text) ExcelRangeBase.getText

let private pipe2 (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>) (direction: Direction) f =

    fun inputstream1 ->

        let newStream1 = p1 inputstream1
        match newStream1 with 
        | Some newStream1 ->
            let inputStream2 = InputMatrixStream.applyDirectionToShift direction newStream1.AsInputStream
            
            match p2 inputStream2 with 
            | Some newStream2 ->
                OutputMatrixStream.mapM (fun result2 -> 
                    let result = newStream1.Result, result2
                    f result
                ) newStream2
                |> Some

            | None -> None
        | None -> None

let private pipe3 p1 p2 p3 direction f =
    pipe2 (pipe2 p1 p2 direction id) p3 direction (fun ((a, b), c) ->
        f (a, b, c)
    )


let c2 p1 p2 =
    pipe2 p1 p2 Direction.Horizontal id

let c3 p1 p2 p3 =
    pipe3 p1 p2 p3 Direction.Horizontal id

let r2 p1 p2 =
    pipe2 p1 p2 Direction.Vertical id

let r3 p1 p2 p3 = 
    pipe3 p1 p2 p3 Direction.Vertical id

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
    mses |> List.map (fun ms -> ms.Result)

let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<_>) =
    let ranges = ExcelRangeBase.asRanges range
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result)

let runMatrixParser (worksheet: ExcelWorksheet) (p: MatrixParser<_>) =
    let userRange = 
        worksheet
        |> ExcelWorksheet.getUserRange

    runMatrixParserForRanges userRange p


