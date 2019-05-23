module Tests.MatrixTreeTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParserTree
open FParsec
open OfficeOpenXml
open System.IO

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = excelPackage.Workbook.Worksheets.["Matrix"]

let matrixTreeTests =
  testList "MatrixTreeTests" [
    ftestCase "mxText" <| fun _ -> 
        let results = runMatrixParser worksheet (mxText "mxTextA")
        match results with 
        | ["mxTextA"] -> pass()
        | _ -> fail()

    //testCase "mxOR" <| fun _ -> 
    //    let results = runMatrixParser worksheet (mxOR (mxText "mxOR_A") (mxText "mxOR_B"))
    //    match results with 
    //    | [Choice1Of2 "mxOR_A";Choice2Of2 "mxOR_B"] -> pass()
    //    | _ -> fail()

    //testCase "c2" <| fun _ -> 
    //    let results = runMatrixParser worksheet (c2 (mxText "C2A") (mxText "C2B"))
    //    match results with 
    //    | ["C2A", "C2B"] -> pass()
    //    | _ -> fail()

    //testCase "c3" <| fun _ -> 
    //    let results = runMatrixParser worksheet (c3 (mxText "C3A") (mxText "C3B") (mxText "C3C"))
    //    match results with 
    //    | ["C3A", "C3B", "C3C"] -> pass()
    //    | _ -> fail()

    //testCase "r2" <| fun _ -> 
    //    let results = runMatrixParser worksheet (r2 (mxText "R2A") (mxText "R2B"))
    //    match results with 
    //    | ["R2A", "R2B"] -> pass()
    //    | _ -> fail()

    //testCase "r3" <| fun _ -> 
    //    let results = runMatrixParser worksheet (r3 (mxText "R3A") (mxText "R3B") (mxText "R3C"))
    //    match results with 
    //    | ["R3A", "R3B", "R3C"] -> pass()
    //    | _ -> fail()

    //testCase "compose c2 and r2" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (c2 (mxText "C2_R2A") (r2 (mxText "C2_R2B") (mxText "C2_R2C")))

    //    match results with 
    //    | [("C2_R2A"), ("C2_R2B", "C2_R2C")] -> pass()
    //    | _ -> fail()

    //testCase "compose r3 and c2" <| fun _ -> 
    //    let results = runMatrixParser worksheet (r3 (c2 (mxText "R3_C2A") (mxText "R3_C2B")) (mxText "R3_C2C") (mxText "R3_C2D"))
    //    match results with 
    //    | [("R3_C2A", "R3_C2B"), "R3_C2C", "R3_C2D"] -> pass()
    //    | _ -> fail()

    //testCase "compose c2 and r3" <| fun _ -> 
    //    let results = 
    //        runMatrixParser worksheet (c2 (r3 (mxText "C2_R3A") (mxText "C2_R3B") (mxText "C2_R3C")) (mxText "C2_R3D"))
    //    match results with 
    //    | [(("C2_R3A", "C2_R3B", "C2_R3C"),"C2_R3D")] -> pass()
    //    | _ -> fail()

    //testCase "cross area1" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (c3 
    //                (mxText "Cross_1A") 
    //                (r3 (mxText "Cross_1B") (mxText "Cross_1C") (mxText "Cross_1D"))
    //                (mxText "Cross_1E")
    //            )
    //    match results with 
    //    | [("Cross_1A",("Cross_1B", "Cross_1C", "Cross_1D"),"Cross_1E")] -> pass()
    //    | _ -> fail()

    //testCase "cross area2" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (c2 
    //                (mxText "Cross_2A") 
    //                (r3 
    //                    (mxText "Cross_2B") 
    //                    (mxText "Cross_2C") 
    //                    (c2 (mxText "Cross_2D") (mxText "Cross_2E"))
    //                )
    //            )
    //    match results with 
    //    | [("Cross_2A",("Cross_2B", "Cross_2C", ("Cross_2D","Cross_2E")))] -> pass()
    //    | _ -> fail()

    //testCase "column many" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (cm (mxTextf(fun text -> text.StartsWith "cm_n")))
    //    match results with 
    //    | ["cm_n1"; "cm_n2"; "cm_n3"; "cm_n4"] :: _  -> pass()
    //    | _ -> fail()

    //testCase "row many" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (rm (mxTextf(fun text -> text.StartsWith "rm_n")))
    //    match results with 
    //    | ["rm_n1"; "rm_n2"; "rm_n3"; "rm_n4"] :: _  -> pass()
    //    | _ -> fail()

    //testCase "column many with skip" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (mxColManySkip mxSpace 1 (mxTextf(fun text -> text.StartsWith "cm_skip")))
    //    match results with 
    //    | ["cm_skip_1"; "cm_skip_2"; "cm_skip_3"] :: _  -> pass()
    //    | _ -> fail()

    //testCase "mx until" <| fun _ -> 
    //    let results = 
    //        runMatrixParser 
    //            worksheet 
    //            (c2 (mxText "mx_until1") (mxUntilA50 (mxText "mx_until4")))

    //    match results with 
    //    | ("mx_until1", ("mx_until4")) :: _  -> pass()
    //    | _ -> fail()

  ]