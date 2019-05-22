module Tests.MatrixAstTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsersAst
open FParsec
open OfficeOpenXml
open System.IO

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = excelPackage.Workbook.Worksheets.["Matrix"]

let matrixAstTests =
  testList "MatrixAstTests" [
    testCase "mxText" <| fun _ -> 
        let results = runMatrixParser worksheet (mxText "mxTextA")
        match results with 
        | ["mxTextA"] -> pass()
        | _ -> fail()

  ]