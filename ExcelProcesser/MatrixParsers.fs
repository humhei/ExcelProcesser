module MatrixParsers
open OfficeOpenXml
open CellParsers
type Range=
     |CellParser of CellParser
     |Auto
     |AnyCell of int
     |Range of Range list
type PLRow=
    Range list
type Row=
    |PLRow of PLRow
    |Auto
    |AnyRow of int
type MatrixParser=Row list
let (>>.) (p1:CellParser) (p2:CellParser)=
    fun (cell:ExcelRangeBase)->
        let nextCell=cell.Offset(0,1)
        p1 cell&&p2 nextCell        
let (.>>) (p1:CellParser) (p2:CellParser)=
    fun (cell:ExcelRangeBase)->
        let preCell=cell.Offset(0,-1)
        p1 cell&&p2 preCell
let (.>>.) (p1:Range) (p2:Range):Range=
    match p1 with
    |Range r1->
        match p2 with
        |Range r2-> Range<| ([r1;r2]|>List.concat)
        |_-> Range<| ([r1;[p2]]|>List.concat)
    |_-> match p2 with
          |Range r2-> Range (p1::r2)
          |_->Range [p1;p2]