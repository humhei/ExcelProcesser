module ExcelProcesser.MatrixParsersAst

open OfficeOpenXml
open Extensions
open CellParsers
open MatrixParsers

type MatrixParser =
    | Text of string

module MatrixParser =
    let parse p =
        fun (inputStream: InputMatrixStream) ->
            match p with 
            | Text text ->


let runMatrixParserForRangesWithStreamsAsResult (ranges : seq<ExcelRangeBase>) (p : MatrixParser) =
    let inputStreams = 
        ranges 
        |> List.ofSeq
        |> List.map (fun range ->
            { Range = range 
              Shift = Shift.Start }
        )

    inputStreams 
    |> List.choose p

let runMatrixParser (worksheet: ExcelWorksheet) (p: MatrixParser<_>) =
    let userRange = 
        worksheet
        |> ExcelWorksheet.getUserRange

    runMatrixParserForRanges userRange p


