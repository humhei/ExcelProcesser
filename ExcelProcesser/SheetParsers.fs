module ExcelProcess.SheetParsers
open OfficeOpenXml
open MatrixParsers
type SheetParser<'a> = ExcelWorksheet -> 'a
    
let (>=>) (p : MatrixParser<'a>) (f: 'a -> 'b) =
    fun sheet ->
        runMatrixParser p sheet
let runSheetParser sheet (p: SheetParser<'a>) =
    p sheet
let spf2 (p1: SheetParser<'a>) (p2: SheetParser<'b>) f =
    fun sheet ->
        let s1 = p1 sheet
        let s2 = p2 sheet
        f s1 s2        
let sp2 p1 p2 = 
    fun sheet -> 
        spf2 p1 p2 (fun a b -> a,b) sheet

let spf3 (p1: SheetParser<'a>) (p2: SheetParser<'b>) (p3: SheetParser<'c>) f =
    spf2 (sp2 p1 p2) p3 (fun (s1,s2) s3 -> f s1 s2 s3)
let sp3 p1 p2 p3 = spf3 p1 p2 p3 (fun a b c -> a,b,c)

let spf4 p1 p2 p3 p4 f =
    spf2 (sp3 p1 p2 p3) p4 (fun (s1,s2,s3) s4 -> f s1 s2 s3 s4)

let sp4 p1 p2 p3 p4 f =
    spf4 p1 p2 p3 p4 (fun s1 s2 s3 s4 -> s1,s2,s3,s4)