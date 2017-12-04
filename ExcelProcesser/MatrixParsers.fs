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
    Range [p1;p2]
    