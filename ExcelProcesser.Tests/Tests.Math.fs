module Tests.Math

open ExcelProcess
open CellParsers
open Expecto
open System.Drawing
open ArrayParsers
open FParsec
open Tests.Types
open System.IO
open MatrixParsers
open Math
let workSheet = XLPath.mathTest |> Excel.getWorksheetByIndex 0

let MathTests =
  testList "MathExTests" [
    testCase "Sum" <| fun _ ->
        runMatrixParserBack (mxRowSum) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | _ -> pass()
            | _ -> fail()

  ]