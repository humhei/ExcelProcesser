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
    testCase "19SPX16" <| fun _ -> 
        use excelPackage = new ExcelPackage(FileInfo(XLPath.RealWorldSamples.``19SPX16合同附件``))
        
        let worksheet = ValidExcelWorksheet.Create(excelPackage.Workbook.Worksheets.["Sheet1"])

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


    testCase "DIS26677装箱单" <| fun _ ->
        use excelPackage = new ExcelPackage(FileInfo(XLPath.RealWorldSamples.DIS26677装箱单))
        let worksheet = ValidExcelWorksheet.Create(excelPackage.Workbook.Worksheets.["PACKING LIST"])
        let parser =
            
            r2
                (
                    c2 
                        (mxRegex "Discription") 
                        (mxUntilA50 (mxRegex "VOLUME"))
                )
                (
                    mxUntilA50 (mxRegex "Total")
                )
            
        let parser = parser.InDebug

        let r = runMatrixParserWithStreamsAsResult worksheet parser
        pass()

  ]