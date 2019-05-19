module Tests.MyTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let _, worksheet = excelPackageAndWorksheet 0 XLPath.matrix

let MyTests =
  testList "MyTests" [
    testCase "mxText" <| fun _ -> 
        let results = runMatrixParser worksheet (mxText "mxTextA")
        match results with 
        | ["mxTextA"] -> pass()
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

    testCase "compose c2 and r2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 (mxText "C2_R2A") (r2 (mxText "C2_R2B") (mxText "C2_R2C")))

        match results with 
        | [("C2_R2A"), ("C2_R2B", "C2_R2C")] -> pass()
        | _ -> fail()

    ftestCase "compose r3 and c2" <| fun _ -> 
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

    testCase "cross area1" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c3 
                    (mxText "Cross1_A") 
                    (r3 (mxText "Cross1_B") (mxText "Cross1_C") (mxText "Cross1_D"))
                    (mxText "Cross1_E"))
        match results with 
        | [("Cross1_A",("Cross1_B", "Cross1_C", "Cross1_D"),"Cross1_E")] -> pass()
        | _ -> fail()

    testCase "cross area2" <| fun _ -> 
        let results = 
            runMatrixParser 
                worksheet 
                (c2 
                    (mxText "Cross2_A") 
                    (r3 
                        (mxText "Cross2_B") 
                        (mxText "Cross2_C") 
                        (c2 (mxText "Cross2_D") 
                        (mxText "Cross2_E")))
                )
        match results with 
        | [("Cross2_A",("Cross2_B", "Cross2_C", ("Cross2_D","Cross2_E")))] -> pass()
        | _ -> fail()

  ]