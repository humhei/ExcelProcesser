namespace ExcelProcess
//Below code adpated from igorkulman's ExcelPackageF
//https://github.com/igorkulman/ExcelPackageF
open OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
[<RequireQualifiedAccess>]
module Address =
    let isCell (add:string) =
        not (add.Contains ":") 
    let isRange (add:string) =
        isCell add |> not

[<RequireQualifiedAccess>]
module Excel=
    open OfficeOpenXml
    open System.IO
    open FParsec
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
    /// r2 include r1
    let contain (r1: ExcelRangeBase) (r2: ExcelRangeBase) =
        let runWithValueBack s =
            let p = (asciiUpper .>>. pint64) 
            run p s 
            |> function
                | ParserResult.Success (s,_,_) -> s 
                | _ -> failwithf "failed parsed with %A" s

        let add1 = r1.Address
        let add2 = r2.Address
        let inMiddle l r s = 
            s >= l && s <= r
        if Address.isCell add1 && Address.isRange add2 then
            let c00,r00 = runWithValueBack add1
            let a1 = add2.Split(':')
            let c10,r10 = runWithValueBack a1.[0]
            let c11,r11 = runWithValueBack a1.[1]
            inMiddle c10 c11 c00 && inMiddle r00 r10 r11

        elif Address.isRange add1 && Address.isRange add2 then
            let a0 =  add1.Split(':')
            let c00,r00 = runWithValueBack a0.[0]
            let c01,r01 = runWithValueBack a0.[1]
            let a1 = add2.Split(':')
            let c10,r10 = runWithValueBack a1.[0]
            let c11,r11 = runWithValueBack a1.[1]
            c00 |> inMiddle c10 c11
            && c01 |> inMiddle c10 c11
            && r00 |> inMiddle r10 r11
            && r01 |> inMiddle r10 r11
        else 
            false

    let distinctRanges (ranges: seq<ExcelRangeBase>) =
        let r = 
            ranges |> Seq.fold (fun accum range ->
                let others = ranges |> Seq.filter (fun r -> r.Address <> range.Address)
                if others |> Seq.exists (contain range) then
                    accum                       
                else 
                    accum @ [range]     
            ) []
        r        