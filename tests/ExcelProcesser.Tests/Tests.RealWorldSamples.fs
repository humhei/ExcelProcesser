module Tests.RealWorldSamples
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open Deedle

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"



let realWorldSamples =
  testList "Real world samples" [
    ftestCase "19SPX16" <| fun _ -> 

        use excelPackage = new ExcelPackage(FileInfo(XLPath.RealWorldSamples.``19SPX16合同附件``))
        
        let worksheet = excelPackage.Workbook.Worksheets.["Sheet1"]

        let record = Types.XLPath.RealWorldSamples.Module_19SPX16合同附件.Record.Parse(worksheet)
        let m = record.ToTable()
        pass()
  ]