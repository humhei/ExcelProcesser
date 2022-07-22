module Tests.SematicsParsers.RangeInHeader
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
    excelPackage.Workbook.Worksheets.["Sematics_RangeInHeaders"]
    |> CellScript.Core.Types.ValidExcelWorksheet.Create


let tests =
  testList "SematicsParsers RangeInHeader" [

    testCase "Pivot Table Headers1" <| fun _ -> 

        let result = 
            let parser =
                NormalColumnHeadersParser
                    .Create(
                        start = mxText "Pivot Table Headers1",
                        rowsCount = 2
                    )
                    .RangeInHeader(headerName = GroupingColumnHeaderRowNameParser.TopOrNone)
                    .MultipleColumns(mxInt32).Parser

            runMatrixParser worksheet parser
            |> List.exactlyOne

        match result.NormalColumnHeaders.Columns, result.GroupingColumnHeaderRows.Value.ValuesLength with 
        | 15, 5 -> pass()
        | _ -> fail()

    testCase "Pivot Table" <| fun _ ->

        let result = 
            let parser =
                NormalColumnHeadersParser
                    .Create(
                        start = mxText "Pivot Table1",
                        rowsCount = 2
                    )
                    .RangeInHeader(headerName = GroupingColumnHeaderRowNameParser.TopOrNone)
                    .MultipleColumns(mxInt32)
                    .SelectColumn(mxInt32).Parser

            runMatrixParser worksheet parser
            |> List.exactlyOne

        match result.GroupingColumn.RowsCount, result.GroupingColumn.ValuesLength with 
        | 7, 8 -> pass()
        | _ -> fail ()

        match result.NormalColumns.Value.Length with 
        | 20 -> pass()
        | _ -> fail()

  ]