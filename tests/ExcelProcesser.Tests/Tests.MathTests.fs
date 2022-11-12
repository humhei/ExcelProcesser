module Tests.MathTests
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open MathParsers
open CellScript.Core

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"

let excelPackage = new ExcelPackage(FileInfo(XLPath.testData))

let worksheet = 
    excelPackage.Workbook.Worksheets.["Math"]
    |> ValidExcelWorksheet.Create


let mathTests =
  testList "MathTests" [
    testCase "mxSumContinuously" <| fun _ -> 
        let results = runMatrixParser worksheet (mxSumContinuously Direction.Vertical)
        #if TestVirtual
        match results with 
        | [] -> pass()
        | _ -> fail()
        #else
        match results with 
        | ([1;2], 3) :: _ -> pass()
        | _ -> fail()
        #endif
  ]