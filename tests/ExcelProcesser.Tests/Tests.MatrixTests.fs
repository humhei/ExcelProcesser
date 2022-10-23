module Tests.MatrixTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open ExcelProcesser.Extensions
open CellScript.Core

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = 
    excelPackage.Workbook.Worksheets.["Matrix"]
    |> ValidExcelWorksheet.Create

let shiftTests =
  testList "ShiftTests" [

    testCase "start + Horizontal ({X = 0; Y = 0;},1) + Direction.Horizontal = Horizontal {0,0},2" <| fun _ -> 
        let shift = 
            let shift = Horizontal ({X = 0; Y = 0;},1)
            Shift.applyDirection (Start) Direction.Horizontal shift
        match shift.Last with 
        | Horizontal({X = 0; Y = 0},2) -> pass()
        | _ -> fail()

    testCase "start + Vertical ({X = 0; Y = 0;},1) + Direction.Vertical = Vertical {0,0},2" <| fun _ -> 
        let shift = 
            let shift = Vertical ({X = 0; Y = 0;},1)
            Shift.applyDirection (Start) Direction.Vertical shift
        match shift.Last with 
        | Vertical({X = 0; Y = 0},2) -> pass()
        | _ -> fail()

    testCase "start + [Horizontal ({X = 0; Y = 0;},1);Vertical {1,0} 1;Horizontal {1,1} 1] => Horizontal {1,0},1" <| fun _ -> 
        let shift = 
            let shift = Compose [Horizontal ({X = 0; Y = 0;},1);Vertical({X = 1; Y =0},1) ;Horizontal({X=1; Y =1},1)]
            Shift.applyDirection (Start) Direction.Horizontal shift
        match shift.Last with 
        | Horizontal({X = 0; Y = 0},2) -> pass()
        | _ -> fail()

  ]

let matrixTests =
  testList "MatrixTests" [
    testCase "mxText" <| fun _ -> 
        let parser =
            (mxText "mxTextA")
            |> MatrixParser.addLogger LoggerLevel.Important "mxTextA"

        let results = runMatrixParserSafe worksheet parser
        match results.AsList with 
        | ["mxTextA"] -> pass()
        | _ -> fail()

    ftestCase "mxEOF" <| fun _ -> 
        let parser =
            c2 (mxText "mxEOF_STRAT") (mxUntil1 Direction.Horizontal None (mxAnyOrigin) (mxEOF))

        let results = runMatrixParser worksheet (parser)
        match results with 
        | ["mxEOF_STRAT", ([""; ""; ""; ""; ""; ""; ""; "mxEOF_R2"; "mxEOF_END"],_)] -> pass()
        | _ -> fail()

    testCase "mxOR" <| fun _ -> 
        let results = runMatrixParser worksheet (mxOR (mxText "mxOR_A") (mxText "mxOR_B"))
        match results with 
        | [Choice1Of2 "mxOR_A";Choice2Of2 "mxOR_B"] -> pass()
        | _ -> fail()

    testCase "c2" <| fun _ -> 
        let results = runMatrixParser worksheet (c2 (mxText "C2A") (mxText "C2B"))
        match results with 
        | ["C2A", "C2B"] -> pass()
        | _ -> fail()

    testCase "c3" <| fun _ -> 
        let results = runMatrixParser worksheet (c3 (mxText "C3A") (mxText "C3B") (mxText "C3C"))
        match results with 
        | ["C3A", "C3B", "C3C"] -> pass()
        | _ -> fail()

    testCase "r2" <| fun _ -> 
        let results = runMatrixParser worksheet (r2 (mxText "R2A") (mxText "R2B"))
        match results with 
        | ["R2A", "R2B"] -> pass()
        | _ -> fail()

    testCase "r3" <| fun _ -> 
        let results = runMatrixParser worksheet (r3 (mxText "R3A") (mxText "R3B") (mxText "R3C"))
        match results with 
        | ["R3A", "R3B", "R3C"] -> pass()
        | _ -> fail()

    testCase "compose c2 and c2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 (c2 (mxText "C2_C2A") (mxText "C2_C2B")) (mxText "C2_C2C"))

        match results with 
        | [("C2_C2A","C2_C2B"), "C2_C2C"] -> pass()
        | _ -> fail()

    testCase "compose c2 and r2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 (mxText "C2_R2A") (r2 (mxText "C2_R2B") (mxText "C2_R2C")))

        match results with 
        | [("C2_R2A"), ("C2_R2B", "C2_R2C")] -> pass()
        | _ -> fail()

    testCase "compose r3 and c2" <| fun _ -> 
        let results = runMatrixParser worksheet (r3 (c2 (mxText "R3_C2A") (mxText "R3_C2B")) (mxText "R3_C2C") (mxText "R3_C2D"))
        match results with 
        | [("R3_C2A", "R3_C2B"), "R3_C2C", "R3_C2D"] -> pass()
        | _ -> fail()

    testCase "compose c2 and r3" <| fun _ -> 
        let results = 
            runMatrixParser worksheet (c2 (r3 (mxText "C2_R3A") (mxText "C2_R3B") (mxText "C2_R3C")) (mxText "C2_R3D"))
        match results with 
        | [(("C2_R3A", "C2_R3B", "C2_R3C"),"C2_R3D")] -> pass()
        | _ -> fail()

    testCase "cross area1 - 1" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                    (r2 (mxText "Cross_1B") (c2 (mxText "Cross_1C") (mxText "Cross_1D")))
                
        //pass()
        match results with 
        | [(("Cross_1B", ("Cross_1C", "Cross_1D")))] -> pass()
        | _ -> fail()

    testCase "cross area1 - 2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                    (c3 
                        (mxText "Cross_1A") 
                        (r2 (mxText "Cross_1B") (c2 (mxText "Cross_1C") (mxText "Cross_1D")))
                        (mxText "Cross_1E")
                    )   
                
        //pass()
        match results with 
        | [(("Cross_1A",("Cross_1B", ("Cross_1C", "Cross_1D")),"Cross_1E"))] -> pass()
        | _ -> fail()

    testCase "cross area1 - 3" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                    (r3
                        (c3 
                            (mxText "Cross_1A") 
                            (r2 (mxText "Cross_1B") (c2 (mxText "Cross_1C") (mxText "Cross_1D")))
                            (mxText "Cross_1E")
                        )
                        (mxText "Cross_1F" )
                        (c2 (mxText "Cross_1G") (mxText "Cross_1H"))
                    )
                
        //pass()
        match results with 
        | [(("Cross_1A",("Cross_1B", ("Cross_1C", "Cross_1D")),"Cross_1E"),"Cross_1F",("Cross_1G","Cross_1H"))] -> pass()
        | _ -> fail()


    testCase "cross area1 - 4" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (r2
                    (r3
                        (c3 
                            (mxText "Cross_1A") 
                            (r2 (mxText "Cross_1B") (c2 (mxText "Cross_1C") (mxText "Cross_1D")))
                            (mxText "Cross_1E")
                        )
                        (mxText "Cross_1F")
                        (c2 (mxText "Cross_1G") (mxText "Cross_1H"))
                    )
                    (mxText "Cross_1I")
                )
        //pass()
        match results with 
        | [(("Cross_1A",("Cross_1B", ("Cross_1C", "Cross_1D")),"Cross_1E"),"Cross_1F",("Cross_1G","Cross_1H")),"Cross_1I"] -> pass()
        | _ -> fail()

    testCase "cross area2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 
                    (mxText "Cross_2A") 
                    (r3 
                        (mxText "Cross_2B") 
                        (mxText "Cross_2C") 
                        (c2 (mxText "Cross_2D") (mxText "Cross_2E"))
                    )
                )
        match results with 
        | [("Cross_2A",("Cross_2B", "Cross_2C", ("Cross_2D","Cross_2E")))] -> pass()
        | _ -> fail()

    testCase "cross area3-1" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                    (r2 
                        (mxText "Cross_3A") 
                        (c3 
                            (mxText "Cross_3B") 
                            (r2 
                                (mxText "Cross_3C") 
                                (c2 (mxText "Cross_3D" ) (mxText "Cross_3E"))
                            )
                            (mxText "Cross_3F")
                        ))

        match results with
        | [("Cross_3A", ("Cross_3B", ("Cross_3C", ("Cross_3D","Cross_3E")),"Cross_3F"))] -> pass()
        | _ -> fail()

    testCase "cross area3-2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                    (r2 
                        (mxText "Cross_3A") 
                        (c3 
                            (mxText "Cross_3B") 
                            (r2 
                                (mxText "Cross_3C") 
                                (c2 (mxText "Cross_3D" ) (mxText "Cross_3E"))
                            )
                            (mxText "Cross_3F")
                        )
                    )

        match results with
        | [("Cross_3A", ("Cross_3B", ("Cross_3C", ("Cross_3D","Cross_3E")),"Cross_3F"))] -> pass()
        | _ -> fail()


    testCase "cross area5" <| fun _ -> 
        let results = 
            let parser2 =
                let header = 
                    (mxTextf (fun text ->
                        [
                            "Cross_4B"
                            "Cross_4C"
                            "Cross_4D"
                        ] |> List.contains text
                    ))
                    |> mxRowMany1

                let last =
                    mxText "Cross_4E"

                r2 
                    header
                    (mxUntilA10 last)

            let parser3 =
                mxRowMany1(
                    c3
                        (mxTextf (fun m -> m.StartsWith "Cross_4"))
                        (mxTextf (fun m -> m.StartsWith "Cross_4"))
                        (mxTextf (fun m -> m.StartsWith "Cross_4"))
                )



            runMatrixParser 
                worksheet 
                (
                    c3
                        (mxText "Cross_4A")
                        parser2
                        parser3
                )

                   
        match results with
        | [("Cross_4A", (["Cross_4B"; "Cross_4C"; "Cross_4D"],"Cross_4E"),[ "Cross_4F", "Cross_4G", "Cross_4H" ; "Cross_4I", "Cross_4J", "Cross_4K"] )] -> pass()
        | _ -> fail()


    testCase "column many" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (mxColMany1 (mxTextf(fun text -> text.StartsWith "cm_n")))
        match results with 
        | ["cm_n1"; "cm_n2"; "cm_n3"; "cm_n4"] :: _  -> pass()
        | _ -> fail()

    testCase "column many without redundant" <| fun _ -> 
        let results = 
            runMatrixParserWithStreamsAsResult 
                worksheet 
                (mxColMany1 (mxTextf(fun text -> text.StartsWith "cm_n")))

            |> OutputMatrixStream.removeRedundants
            |> List.map (fun m -> m.Result.Value)

        match results with 
        | [["cm_n1"; "cm_n2"; "cm_n3"; "cm_n4"]]  -> pass()
        | _ -> fail()

    testCase "row many" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (mxRowMany1 (mxTextf(fun text -> text.StartsWith "rm_n")))
        match results with 
        | ["rm_n1"; "rm_n2"; "rm_n3"; "rm_n4"] :: _  -> pass()
        | _ -> fail()
        
    testCase "column many with skip" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (mxColMany1Skip mxSpace 1 (mxTextf(fun text -> text.StartsWith "cm_skip")))
        match results with 
        | ["cm_skip_1"; "cm_skip_2"; "cm_skip_3"] :: _  -> pass()
        | _ -> fail()

    testCase "column many with skip backtrack" <| fun _ -> 
        let results = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (mxColMany1Skip mxSpace 1 (mxTextf(fun text -> text.StartsWith "cm_skip_backTrack")))
        match results.[0].Shift.Last with 
        | Horizontal ({X = 0; Y = 0},2)  -> pass()
        | _ -> fail()

    testCase "mx until" <| fun _ -> 
        let streams = 
            runMatrixParserWithStreamsAsResult 
                worksheet 
                    (c2 (c2 (mxText "mx_until1") (mxUntilA50 (mxText "mx_until4"))) (mxText "mx_util5"))


        let results = 
            streams
            |> List.map (fun m -> m.Result.Value)

        match results with 
        | (("mx_until1", ("mx_until4")), "mx_util5") :: _  -> pass()
        | _ -> fail()

    testCase "mxMany1Op" <| fun _ -> 
        let streams = 
            let parser =
                c2 
                    ((mxText "mxMany1Op_Starter"))
                    (mxColMany1Op 1 (mxTextf(fun text -> text.StartsWith "cm")))
            runMatrixParserWithStreamsAsResult 
                worksheet 
                    parser


        let results = 
            streams
            |> List.map (fun m -> m.Result.Value)
            |> List.exactlyOne
            |> snd

        match results with 
        | [None; Some ("cm_op_1"); Some ("cm_op_2"); None; Some ("cm_op_3")] -> pass()
        | _ -> fail()


    testCase "cross area1 reRange" <| fun _ -> 
        let results = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (r2
                    (r3
                        (c3 
                            (mxText "Cross_1A") 
                            (r2 (mxText "Cross_1B") (c2 (mxText "Cross_1C") (mxText "Cross_1D")))
                            (mxText "Cross_1E")
                        )
                        (mxText "Cross_1F")
                        (c2 (mxText "Cross_1G") (mxText "Cross_1H"))
                    )
                    (mxText "Cross_1I")
                )
        results 
        |> List.map (OutputMatrixStream.reRangeByShift >> (fun (rerangedResult) -> 
            let ranges = ExcelRangeBase.asRangeList rerangedResult.Range
            List.map SingletonExcelRangeBase.getText ranges
            |> List.distinct
        ))
        |> function
            | ["Cross_1A"; "Cross_1B";"Cross_1E";"Cross_1F";"Cross_1C";"Cross_1D";"Cross_1G";"Cross_1H";"Cross_1I"] :: _ -> pass()
            | _ -> fail()

    testCase "cross area2 reRange" <| fun _ -> 
        let results = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (c2 
                    (mxText "Cross_2A") 
                    (r3 
                        (mxText "Cross_2B") 
                        (mxText "Cross_2C") 
                        (c2 (mxText "Cross_2D") (mxText "Cross_2E"))
                    )
                )
        results 
        |> List.map (OutputMatrixStream.reRangeByShift  >> (fun (rerangedResult) -> 
            let ranges = List.ofSeq rerangedResult.Range
            List.map ExcelRangeBase.getText ranges
            |> List.distinct
        ))
        |> function
            | ["Cross_2A"; "Cross_2B";"Cross_2C";"Cross_2D";"Cross_2E"] :: _ -> pass()
            | _ -> fail()



    testCase "mxMergeStarter" <| fun _ -> 
        let results = runMatrixParser worksheet (mxMergeStarter ||>> fun mergeStarter -> mergeStarter.Text)
        match results with 
        | "Merge1" :: "Merge2":: _ -> pass()
        | _ -> fail()

    testCase "mxMerge horizontal" <| fun _ -> 
        let results = 
            runMatrixParser worksheet ((mxMergeWithAddresses Direction.Horizontal) ||>> fun (mergeStarter,_) -> mergeStarter.Text)
        match results with 
        | ["Merge1"; "Merge2"] -> pass()
        | _ -> fail()
        


  ]

