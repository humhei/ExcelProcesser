[<RequireQualifiedAccess>]
module Excel
open OfficeOpenXml
open System.IO
open System.Drawing
open CellParsers
open RangeParsers
let getWorksheets filename = seq {
    let file = FileInfo(filename) 
    let xlPackage = new ExcelPackage(file)
    for i in 1..xlPackage.Workbook.Worksheets.Count do
        yield xlPackage.Workbook.Worksheets.[i]
    }
let getWorksheetByIndex (index:int) filename = 
    let file = FileInfo(filename) 
    let xlPackage = new ExcelPackage(file)
    xlPackage.Workbook.Worksheets.[index]
let getMaxColNumber (worksheet:ExcelWorksheet) = 
    worksheet.Dimension.End.Column
let getMaxRowNumber (worksheet:ExcelWorksheet) = 
    worksheet.Dimension.End.Row     
let getContent worksheet = seq {        
    let maxRow = getMaxRowNumber worksheet
    let maxCol = getMaxColNumber worksheet
    for i in 1..maxRow do
        for j in 1..maxCol do
            let content = worksheet.Cells.[i,j].Value
            yield content
}
let getUserRange  worksheet:seq<ExcelRangeBase> = seq {        
    let maxRow = getMaxRowNumber worksheet
    let maxCol = getMaxColNumber worksheet
    for i in 1..maxRow do
        for j in 1..maxCol do
            let content = worksheet.Cells.[i,j]
            yield content:>ExcelRangeBase
          
}
let runParser (parser:RangeParser<'a>)  worksheet=
    let t= ref 0
    worksheet
    |>getUserRange
    |>Seq.cache
    |>fun c->{position=t;userRange=c}  
    |>parser
