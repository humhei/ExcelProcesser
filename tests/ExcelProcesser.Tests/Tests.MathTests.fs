module Tests.MathTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open MathParsers

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = excelPackage.Workbook.Worksheets.["Math"]

let mathTests =
  testList "MathTests" [
    testCase "sum" <| fun _ -> 
        let results = runMatrixParser worksheet (mxSum Direction.Vertical)
        match results with 
        | ([1;2], 3) :: _ -> pass()
        | _ -> fail()

  ]