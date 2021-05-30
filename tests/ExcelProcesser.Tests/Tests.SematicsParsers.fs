module Tests.SematicsParsers
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open SematicsParsers.PivotTable
//open SematicsParsers.TwoHeadersPivotTable
open ExcelProcesser.Extensions
open Shrimp.FSharp.Plus
open CellScript.Core


let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = 
    excelPackage.Workbook.Worksheets.["Sematics"]
    |> CellScript.Core.Types.ValidExcelWorksheet








let sematicsParsers =
  testList "SematicsParsers" [

  
    //testCase "groupingColumnsHeader" <| fun _ -> 
    //    let results = runMatrixParser worksheet (mxGroupingColumnsHeader None (mxFParsec pint32))
    //    match results.[0].GroupedHeader, results.[0].ChildHeaders with 
    //    | ("Size", AtLeastOneList [28; 29; 30; 31; 32; 33; 34; 35])  -> pass()
    //    | _ -> fail()

    //testCase "groupingColumn" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet
    //            (mxGroupingColumn 
    //                (GroupingColumnParserArg.Create(mxFParsec pint32, Some mxSpace, mxFParsec pint32))
    //            )
    //    let result = AtLeastOneList.toLists results.[0].ElementsList |> List.map (List.choose id)

    //    match result with 
    //    | [1; 2; 2; 2; 2; 2; 1; 1] :: [1; 1; 1; 2; 2; 1; 1; 1] :: [1] :: [1] :: [1] :: [1] :: [1] :: _ -> pass()
    //    | _ -> fail()

    //testCase "twoRowHeaderPivotTable" <| fun _ -> 

    //    let results = 

    //        runMatrixParser
    //            worksheet
    //            (TwoHeadersPivotTable.Parser(
    //                pLeftBorderHeader = (mxText ("ORDER NO.")),
    //                pNumberHeader = (mxFParsec (pstringCI "Pairs")) ,
    //                pOriginRightBorderHeader = (Some (mxFParsec (pstringCI "Volume"))) ,
    //                pGroupingColumn = 
    //                    (GroupingColumnParserArg.Create(
    //                        pChildHeader = mxFParsec pint32,
    //                        pElementSkip = Some mxSpace,
    //                        pElement = mxFParsec pint32, 
    //                        defaultGroupedHeaderText = "Size"))
    //            )
    //            )
        
    //    let rows = results.[0].Rows()

    //    Expect.equal rows.Length 7 "pass"

    //    Expect.equal results.Length 1 "pass"
        
    //    let array2D = TwoHeadersPivotTable.ToArray2D results.[0]
    //    Expect.equal array2D.Length 484 "pass"

    testCase "groupingColumnHeader" <| fun _ -> 

        let parser = 
            let start = 
                mxText "groupingColumnHeader1"

            let sizeRange = mxFParsec(sepEndBy1 pint32 spaces1)

            let groupingColumnParser =
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 1
                )

            c2 start (groupingColumnParser)
            ||>> snd

        let result = 
            runMatrixParser
                worksheet parser
            |> List.concatWithAtLeastOneList
            |> List.exactlyOne

        match result with 
        |  { Name = None
             Values = AtLeastOneList [35; 36; 37; 38; 39; 40] } -> pass()
        | _ -> fail()

    testCase "groupingColumnHeader2" <| fun _ -> 

        let parsingRange =
            let start = 
                mxText "groupingColumnHeader2"

            (runMatrixParserWithStreamsAsResult worksheet start)
            |> List.exactlyOne
            |> fun m -> m.Range.Offset(0, 1).Offset(0, 0, 1, 10)


        let parser = 

            let sizeRange = 
                mxColMany1 mxInt32

            GroupingColumnHeaderRow.Parser(
                groupingHeaders = sizeRange,
                headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                rowsCount = 1

            )

        let result = 
            runMatrixParserForRangeWithoutRedundent
                parsingRange
                parser
            |> List.concatWithAtLeastOneList
            |> List.exactlyOne

        match result with 
        |  { Name = None
             Values = AtLeastOneList [35; 36; 37; 38; 39; 40] } -> pass()
        | _ -> fail()

    testCase "groupingColumnHeader3" <| fun _ -> 

        let parsingRange =
            let start = 
                mxText "groupingColumnHeader3"

            (runMatrixParserWithStreamsAsResult worksheet start)
            |> List.exactlyOne
            |> fun m -> m.Range.Offset(0, 1).Offset(0, 0, 2, 10)


        let parser = 

            let sizeRange = 
                mxColMany1 mxInt32

            GroupingColumnHeaderRow.Parser(
                groupingHeaders = sizeRange,
                headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                rowsCount = 2
            )

        let result = 
            runMatrixParserForRangeWithoutRedundent
                parsingRange
                parser
            |> List.concatWithAtLeastOneList
            |> List.exactlyOne

        match result with 
        |  { Name = None
             Values = AtLeastOneList [35; 36; 37; 38; 39; 40] } -> pass()
        | _ -> fail()

    testCase "groupingColumnHeader4" <| fun _ -> 

        let parsingRange =
            let start = 
                mxText "groupingColumnHeader4"

            (runMatrixParserWithStreamsAsResult worksheet start)
            |> List.exactlyOne
            |> fun m -> m.Range.Offset(0, 1).Offset(0, 0, 2, 10)


        let parser = 

            let sizeRange = 
                mxColMany1 mxInt32

            GroupingColumnHeaderRow.Parser(
                groupingHeaders = sizeRange,
                headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                rowsCount = 2
            )

        let result = 
            runMatrixParserForRangeWithoutRedundent
                parsingRange
                parser
            |> List.concatWithAtLeastOneList
            |> List.exactlyOne

        match result with 
        |  { Name = Some (GroupingColumnHeaderRowName.Top "Size")
             Values = AtLeastOneList [35; 36; 37; 38; 39; 40] } -> pass()
        | _ -> fail()

    testCase "groupingColumnHeader5" <| fun _ -> 

        let rowsCount = 3

        let parsingRange =
            let start = 
                mxText "groupingColumnHeader5"

            (runMatrixParserWithStreamsAsResult worksheet start)
            |> List.exactlyOne
            |> fun m -> m.Range.Offset(0, 1).Offset(0, 0, rowsCount, 10)


        let parser = 

            let sizeRange = 
                mxColMany1 mxDouble

            GroupingColumnHeaderRow.Parser(
                groupingHeaders = sizeRange,
                headerName = GroupingColumnHeaderRowNameParser.Left mxNonEmpty,
                rowsCount = 2
            )

        let results = 
            runMatrixParserForRangeWithoutRedundent
                parsingRange
                parser
            |> List.exactlyOne

        match results.[0] with 
        |  { Name = Some (GroupingColumnHeaderRowName.Left "EUR")
             Values = AtLeastOneList [34.; 35.; 36.; 37.; 38.; 39.] } -> pass()
        | _ -> fail()

        match results.[1] with 
        |  { Name = Some (GroupingColumnHeaderRowName.Left "UK")
             Values = AtLeastOneList [2.; 2.5; 3.; 4.; 5.; 6.] } -> pass()
        | _ -> fail()



    testCase "one row table1" <| fun _ -> 

        let parser = 
            let start = 
                mxText "OneRowTable1_Art"

            //let color = mxNonEmpty

            let sizeRange = mxFParsec(sepEndBy1 pint32 spaces1)
                
            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 1
                ) 

            let fraction = mxFParsec(sepEndBy1 pint32 spaces1)
            PivotTable.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    )
            )


        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumns =  AtLeastOneList [_; _;] 
            GroupingColumn = { 
                HeaderRows = AtLeastOneList 
                    [ {Name = _
                       Values = AtLeastOneList [35; 36; 37; 38; 39; 40]} ]

                Elements = AtLeastOneLists [[1; 2; 3; 3; 2; 1]; _ ; _] 
                }
          } 
            -> 
            pass()
        | _ -> fail()

    testCase "one row table2" <| fun _ -> 

        let parser = 
            let start = mxText "OneRowTable2_Art"

            let sizeRange = mxFParsec(sepEndBy1 pint32 spaces1)
            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 1
                )
            let fraction = mxFParsec(sepEndBy1 pint32 spaces1)
            PivotTable.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    ),
                rightBorder = mxText "Weight"
            )


        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumns = AtLeastOneList [_; _; _; _] 
            GroupingColumn = { 
                HeaderRows = AtLeastOneList 
                    [ {Name = _
                       Values = AtLeastOneList [35; 36; 37; 38; 39; 40]} ]

                Elements = AtLeastOneLists [[1; 2; 3; 3; 2; 1]; _ ; _] 
                }} -> 
            pass()
        | _ -> fail()


    testCase "one row table3" <| fun _ -> 
        let parser = 
            let start = mxText "OneRowTable3_Art"
            let sizeRange = mxColMany1 (mxInt32)
            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 1
                )
            let fraction = mxColMany1 (mxInt32)
            PivotTable.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    ),
                rightBorder = mxText "Weight"
            )

        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumns = AtLeastOneList [_; _; _; _] 
            GroupingColumn = { 
                HeaderRows = AtLeastOneList 
                    [ {Name = _
                       Values = AtLeastOneList [35; 36; 37; 38; 39; 40]} ]

                Elements = AtLeastOneLists [[1; 2; 3; 3; 2; 1]; _ ; _] 
                }} -> pass()
        | _ -> fail()

    testCase "table4 normal Column header" <| fun _ -> 
        let parseByStart start = 
            let parser = 
                NormalColumnTreeHeader.Parser(
                    parser = start,
                    rowsCount = 5
                    //rightBorder = mxText "Weight"
                )

            runMatrixParserWithStreamsAsResult
                worksheet parser
            |> List.exactlyOne

        let result = parseByStart (mxText "Table4NormalHeaders_OneRow" )

        match result.Shift, result.Result.Value with 
        | Shift.Vertical ({X = 0; Y = 0}, 4), NormalColumnTreeHeader.Leaf "Table4NormalHeaders_OneRow"  -> pass()
        | _ -> fail()
        
        let result = parseByStart (mxText "Table4NH2")

        //let result = parseByStart (mxAddress "C109")

        match result.Shift, result.Result.Value with 
        | Shift.Compose([Shift.Vertical ({X = 2; Y = 0}, 4); Shift.Horizontal ({X = 0; Y = 0}, 2)]),
            NormalColumnTreeHeader.Node ("Table4NH2", _)  -> pass()
        | _ -> fail()

    testCase "table4 headers1" <| fun _ -> 
        
        let parser = 
            let start = mxText "Table4Headers_Art" 
            let sizeRange = mxColMany1 (mxInt32)
            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 2
                ) 

            let fraction = mxColMany1 (mxInt32)

            PivotTableHeadersParser.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    ),
                rowsCount = 2
            )

        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumnHeaders = AtLeastOneList [_; _] 
            GroupingColumnHeaderRows = AtLeastOneList [
                { Name = Some (GroupingColumnHeaderRowName.Top "Size")
                  Values = AtLeastOneList [35; 36; 37; 38; 39; 40] }
            ] 
            }  -> 
            pass()
        | _ -> fail()

    testCase "table4 headers2" <| fun _ -> 
        
        let parser = 
            let start = mxText "Table4Headers2_Art" 
            let sizeRange = mxColMany1 (mxInt32)
            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 2
                ) 

            let fraction = mxColMany1 (mxInt32)

            PivotTableHeadersParser.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    ),
                rowsCount = 2
            )

        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumnHeaders = AtLeastOneList [_; _; _; _] 
            GroupingColumnHeaderRows = AtLeastOneList [
                { Name = Some (GroupingColumnHeaderRowName.Top "Size")
                  Values = AtLeastOneList [35; 36; 37; 38; 39; 40] }
            ] 
            }  -> pass()
        | _ -> fail()

    testCase "one row table4" <| fun _ -> 
        let parser = 
            let start = mxText "OneRowTable4_Art"
            let sizeRange = mxColMany1 (mxInt32)
            let fraction = mxColMany1 (mxInt32)

            let groupingHeaderRows = 
                GroupingColumnHeaderRow.Parser(
                    groupingHeaders = sizeRange,
                    headerName = GroupingColumnHeaderRowNameParser.TopOrNone,
                    rowsCount = 2
                )

            PivotTable.Parser(
                start = start,
                groupingColumnParser = 
                    GroupingColumnParser(
                        groupingHeaderRows = groupingHeaderRows,
                        groupingElements = fraction
                    ),
                rightBorder = mxText "Weight"
            )

        let result = 
            runMatrixParser
                worksheet parser
            |> List.exactlyOne

        match result with 
        | { NormalColumns = AtLeastOneList [_; _; _; _] 
            GroupingColumn = { 
                HeaderRows = AtLeastOneList 
                    [ {Name = _
                       Values = AtLeastOneList [35; 36; 37; 38; 39; 40]} ]

                Elements = AtLeastOneLists [[1; 2; 3; 3; 2; 1]; _ ; _] 
                }} -> pass()
        | _ -> fail()

  ]