module Tests.MyTests
open CellParsers
open Expecto
open System.Drawing
let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"
let fileName= "test3.xlsx"
let MyTests =
  testList "MyTests" [
    testCase "FontColorTest" <| fun _ -> 
          fileName
          |>Excel.getWorksheetByIndex 1
          |>fun c->c.Cells.["E39"]
          |>pFontColor Color.Red
          |>function |true->pass()
                     |false->fail()
          
  ]