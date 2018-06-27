namespace ExcelProcess
//Below code adpated from igorkulman's ExcelPackageF
//https://github.com/igorkulman/ExcelPackageF
[<RequireQualifiedAccess>]
module Excel=
    open OfficeOpenXml
    open System.IO
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

    let translate address (xOffset:int) (yOffset:int) =
        ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(address, -yOffset, -xOffset), 0, 0)