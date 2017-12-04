module CellParsers
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
type CellParser=ExcelRangeBase -> bool
let pColor (color:Color):CellParser=
    fun (cell:ExcelRangeBase)->
        let toHex (color:Color)=sprintf "#%02X%02X%02X%02x" color.A color.R color.G color.B
        let bkColor=cell.Style.Fill.BackgroundColor.LookupColor()
        let targetColor=  toHex color
        bkColor=targetColor
let pRegex pattern:CellParser=
    fun (cell:ExcelRangeBase)->
        let m=Regex.Match(cell.Text, pattern)
        m.Success
let pAny :CellParser=fun _->true
let (<&>) (p1:CellParser) (p2:CellParser)=
    fun (cell:ExcelRangeBase)->
        p1 cell&&p2 cell
   