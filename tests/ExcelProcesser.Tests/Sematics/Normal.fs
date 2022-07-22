module Tests.SematicsParsers.Normal
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open ExcelProcesser.SematicsParsers
//open SematicsParsers.TwoHeadersPivotTable
open ExcelProcesser.Extensions
open Shrimp.FSharp.Plus
open CellScript.Core
open Deedle


let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = 
    excelPackage.Workbook.Worksheets.["Sematics_Normal"]
    |> CellScript.Core.Types.ValidExcelWorksheet.Create


let tests =
  testList "SematicsParsers Normal" [
    testCase "Multiple Rows Headers1" <| fun _ ->
        let stream = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (NormalColumnHeadersParser.Create(mxText "Multiple Rows Headers1", 5).Value)
            |> List.exactlyOne


        let result = stream.Result.Value

        match result.Columns with 
        | 11 -> pass()
        | _ -> fail()


    testCase "Normal Columns1" <| fun _ ->
        let stream = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (NormalColumnHeadersParser.Create(mxText "Normal Columns1", 2).Value)
            |> List.exactlyOne

        let result = stream.Result.Value

        let normalColumns =
            result.Value
            |> AtLeastOneList.map (fun columnHeader -> NormalColumn(columnHeader, 1, 7))

        let firstColumnContents = 
            normalColumns.Head.Contents.Contents.AsList
            |> List.map string

        match firstColumnContents  = List.replicate 7 "16501"  with 
        | true -> pass()
        | false -> fail()

        match result.Columns with 
        | 7 -> pass()
        | _ -> fail()

    testCase "OneColumn Contents Table" <| fun _ ->
        let stream = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (ContentsTable.Parser(mxText "OneColumnContentsTableHeader"))
            |> List.exactlyOne

        let table = stream.Result.Value.AsFrame
        match table.ColumnCount, table.RowCount with 
        | 1, 9 -> pass()
        | _ -> fail ()

    testCase "Multiple columns Contents Table" <| fun _ ->
        let stream = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (ContentsTable.Parser(mxText "MultipleContentsTableHeader"))
            |> List.exactlyOne

        let table = stream.Result.Value.AsFrame
        match table.ColumnCount, table.RowCount with 
        | 4, 11 -> pass()
        | _ -> fail ()

  ]