module Tests.RealWorldSamples
open Expecto
open ExcelProcesser
open Tests.Types
open MatrixParsers
open FParsec
open OfficeOpenXml
open System.IO
open Deedle
open CellScript.Core

let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"



let mergeColumnOn columnKey (infoFrame: Frame<_, _>) (primaryFrame: Frame<_, _>) =
    let infoFrame = infoFrame.IndexRows<string>(columnKey)
    primaryFrame.Rows
    |> Series.map (fun _ row ->
        infoFrame.Rows.TryGet(row.GetAs<string>(columnKey)).ValueOrDefault
        |> Series.merge row
    )
    |> Frame.ofRows

let realWorldSamples =
  testList "Real world samples" [
    ftestCase "19SPX16" <| fun _ -> 
        use excelPackage = new ExcelPackage(FileInfo(XLPath.RealWorldSamples.``19SPX16合同附件``))
        
        let worksheet = ValidExcelWorksheet(excelPackage.Workbook.Worksheets.["Sheet1"])

        let record = Types.XLPath.RealWorldSamples.Module_嘴唇.Record.Parse(worksheet)

        let frame: Frame<_, string> = 
            (record.ToTable())

        let mergeFrame = 
            (Frame.ReadCsv(XLPath.RealWorldSamples.``19SPX16Merge``))

        let cc =
            [ 1.. 1000 ]
            |> List.map (fun _ ->
                mergeColumnOn "Art" mergeFrame frame
            )


        pass()
  ]