[<RequireQualifiedAccess>]
module Excel
open OfficeOpenXml
open System.IO
open LinearParsers
open MatrixParsers
open ArrayParsers
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
let runLinearParser (parser:LinearParser<'a>)  worksheet=
    let t= ref 0
    worksheet
    |>getUserRange
    |>Seq.cache
    |>fun c->{position=t;userRange=c}  
    |>parser
let runMatrixParser (parser:MatrixParser)  worksheet=
    let rowAny (n:int) (cell:ExcelRangeBase)=
        seq {
            for i=0 to n do 
            yield cell.Offset(0,i)
        }
    let extend (range:Range list) (userRange:ExcelRangeBase seq) =
        if Seq.isEmpty range then seq{yield userRange}
        else 
            let rec loop (accum:ExcelRangeBase seq) (n:int) range =
                match range with
                |h::t->match h with 
                        |AnyCell m->  loop accum (m+n) t
                        |CellParser c-> 
                            accum
                            |>Seq.where(fun m->
                              m.Offset(0,n+1)|>c)
                            |>fun c->loop c (n+1) t
                        |_ ->failwithf "Not implemented"
                |[]->
                    accum
                    |>Seq.map(rowAny n)  
            loop userRange 0 range
    let (h,t)=
        match Seq.head parser with
          |PLRow row->
            match Seq.head row with
            |CellParser p->p,[]
            |Range r-> match r with
                        |h::t->match h with
                               |CellParser p->p,t
                               |_->failwithf "Never excute"
                        |_-> failwithf "Never excute"
            |_->failwithf "cell (1,1) must be CellParser"
          |_->failwithf "R1 must be PLRow"
    worksheet
    |>getUserRange
    |>Seq.cache
    |>Seq.where(h)
    |>extend t
    |>Seq.map(fun c->
        c
        |>Seq.map(fun n->n.Text))
let runArraryParser (parser:ArrayParser)  worksheet=
    worksheet
    |>getUserRange
    |>Seq.cache
    |>fun c->{userRange=c;shift=[0]}
    |>parser     
let runParser= runArraryParser