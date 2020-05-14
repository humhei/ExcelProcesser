module Tests.SematicsParsers
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open SematicsParsers
open ExcelProcesser.Extensions

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = excelPackage.Workbook.Worksheets.["Sematics"]

let sematicsParsers =
  testList "SematicsParsers" [
    testCase "groupingColumnsHeader" <| fun _ -> 
        let results = runMatrixParser worksheet (mxGroupingColumnsHeader (mxFParsec pint32))
        match results.[0].GroupedHeader, results.[0].ChildHeaders with 
        | ("Size", [28; 29; 30; 31; 32; 33; 34; 35])  -> pass()
        | _ -> fail()

    testCase "groupingColumn" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (mxGroupingColumn 
                    (GroupingColumnParserArg(mxFParsec pint32, Some mxSpace, mxFParsec pint32))
                )
        
        match results.[0].ElementsList |> List.map (List.choose id) with 
        | [1; 2; 2; 2; 2; 2; 1; 1] :: [1; 1; 1; 2; 2; 1; 1; 1] :: [1] :: [1] :: [1] :: [1] :: [1] :: _ -> pass()
        | _ -> fail()

    testCase "twoRowHeaderPivotTable" <| fun _ -> 
        let userRange = 
            worksheet
            |> ExcelWorksheet.getUserRangeList

        let results = 
            runMatrixParserForRangesWithStreamsAsResult
                userRange 
                (mxTwoHeadersPivotTable 
                    (mxStyleName "Border") 
                    (mxStyleName "Number") 
                    (mxStyleName "Border")
                    (GroupingColumnParserArg(mxFParsec pint32, Some mxSpace, mxFParsec pint32))
                )
        
        Expect.equal results.Length 1 "pass"
        
        let array2D = TwoHeadersPivotTable.toArray2D results.[0].Result.Value
        Expect.equal array2D.Length 484 "pass"

  ]