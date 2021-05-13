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
    |> ValidExcelWorksheet

let shiftTests =
  ptestList "ShiftTests" [

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
        | Horizontal({X = 2; Y = 0},0) -> pass()
        | _ -> fail()

  ]

let matrixTests =
  testList "MatrixTests" [
    testCase "mxText" <| fun _ -> 
        let parser =
            (mxText "mxTextA")
            |> MatrixParser.addLogger "mxTextA"

        let results = runMatrixParser worksheet parser
        match results with 
        | ["mxTextA"] -> pass()
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
                (c2 (mxText "C2_R2A") (r2 (fun a -> mxText "C2_R2B" a) (mxText "C2_R2C")))

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
                        (fun a -> mxText "Cross_1E" a)
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
                        (fun a -> mxText "Cross_1F" a)
                        (fun a -> c2 (mxText "Cross_1G") (mxText "Cross_1H") a)
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
                                (fun a -> c2 (mxText "Cross_3D" ) (mxText "Cross_3E") a)
                            )
                            (fun a -> mxText "Cross_3F" a)
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
                                (fun a -> c2 (mxText "Cross_3D" ) (mxText "Cross_3E") a)
                            )
                            (fun a -> mxText "Cross_3F" a)
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

            let mxRowMany1Test x = 
                fun inputStream ->
                    mxRowMany1 x inputStream

            let parser3 =
                mxRowMany1Test(
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
                (mxColManySkip mxSpace 1 (mxTextf(fun text -> text.StartsWith "cm_skip")))
        match results with 
        | ["cm_skip_1"; "cm_skip_2"; "cm_skip_3"] :: _  -> pass()
        | _ -> fail()

    testCase "column many with skip backtrack" <| fun _ -> 
        let results = 
            runMatrixParserWithStreamsAsResult
                worksheet
                (mxColManySkip mxSpace 1 (mxTextf(fun text -> text.StartsWith "cm_skip_backTrack")))
        match results.[0].Shift.Last with 
        | Horizontal ({X = 0; Y = 0},2)  -> pass()
        | _ -> fail()

    testCase "mx until" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 (mxText "mx_until1") (mxUntilA50 (mxText "mx_until4")))

        match results with 
        | ("mx_until1", ("mx_until4")) :: _  -> pass()
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
        |> List.map (OutputMatrixStream.reRange >> (fun range -> 
            let ranges = ExcelRangeBase.asRangeList range
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
        |> List.map (OutputMatrixStream.reRange  >> (fun range -> 
            let ranges = List.ofSeq range
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
            runMatrixParser worksheet ((mxMerge Direction.Horizontal) ||>> fun (mergeStarter ,_) -> mergeStarter.Text)
        match results with 
        | ["Merge1"; "Merge2"] -> pass()
        | _ -> fail()

  ]

