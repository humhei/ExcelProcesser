module Tests.SematicsParsers
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open SematicsParsers
open SematicsParsers.TwoHeadersPivotTable
open ExcelProcesser.Extensions
open Shrimp.FSharp.Plus
open CellScript.Core


let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = 
    excelPackage.Workbook.Worksheets.["Sematics"]
    |> CellScript.Core.Types.ValidExcelWorksheet


type GroupingHeaderStyle =
    | InOneCol = 0 
    | Horizontal = 1



let sematicsParsers =
  testList "SematicsParsers" [


    testCase "groupingColumnsHeader" <| fun _ -> 
        let results = runMatrixParser worksheet (mxGroupingColumnsHeader None (mxFParsec pint32))
        match results.[0].GroupedHeader, results.[0].ChildHeaders with 
        | ("Size", AtLeastOneList [28; 29; 30; 31; 32; 33; 34; 35])  -> pass()
        | _ -> fail()

    testCase "groupingColumn" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (mxGroupingColumn 
                    (GroupingColumnParserArg.Create(mxFParsec pint32, Some mxSpace, mxFParsec pint32))
                )
        let result = AtLeastOneList.toLists results.[0].ElementsList |> List.map (List.choose id)

        match result with 
        | [1; 2; 2; 2; 2; 2; 1; 1] :: [1; 1; 1; 2; 2; 1; 1; 1] :: [1] :: [1] :: [1] :: [1] :: [1] :: _ -> pass()
        | _ -> fail()

    testCase "one row one column table" <| fun _ ->
        let range = 
            worksheet.VisibleExcelWorksheet.GetRange(RangeGettingOptions.RangeIndexer "A83:E93")

        let results = 

            runMatrixParserForRange
                range
                (TwoHeadersPivotTable.Parser(
                    pLeftBorderHeader = (mxFParsec (pstringCI "ORDER NO.")),
                    pNumberHeader = (mxFParsec (pstringCI "ORDER NO.")) ,
                    pOriginRightBorderHeader = (Some (mxFParsec (pstringCI "Volume"))) ,
                    pGroupingColumn = 
                        (GroupingColumnParserArg.Create(
                            pChildHeader = mxFParsec pint32,
                            pElementSkip = Some mxSpace,
                            pElement = mxFParsec pint32, 
                            defaultGroupedHeaderText = "Size"))
                )
                )
        
        let rows = results.[0].Rows()

        Expect.equal rows.Length 7 "pass"

        Expect.equal results.Length 1 "pass"
        
        let array2D = TwoHeadersPivotTable.ToArray2D results.[0]
        Expect.equal array2D.Length 484 "pass"

    testCase "twoRowHeaderPivotTable" <| fun _ -> 
        let range = 
            worksheet.VisibleExcelWorksheet.GetRange(RangeGettingOptions.RangeIndexer "A37:AD50")

        let results = 

            runMatrixParserForRange
                range
                (TwoHeadersPivotTable.Parser(
                    pLeftBorderHeader = (mxFParsec (pstringCI "ORDER NO.")),
                    pNumberHeader = (mxFParsec (pstringCI "Pairs")) ,
                    pOriginRightBorderHeader = (Some (mxFParsec (pstringCI "Volume"))) ,
                    pGroupingColumn = 
                        (GroupingColumnParserArg.Create(
                            pChildHeader = mxFParsec pint32,
                            pElementSkip = Some mxSpace,
                            pElement = mxFParsec pint32, 
                            defaultGroupedHeaderText = "Size"))
                )
                )
        
        let rows = results.[0].Rows()

        Expect.equal rows.Length 7 "pass"

        Expect.equal results.Length 1 "pass"
        
        let array2D = TwoHeadersPivotTable.ToArray2D results.[0]
        Expect.equal array2D.Length 484 "pass"

  ]