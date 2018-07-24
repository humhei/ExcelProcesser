module ExcelProcess.CellParsers

open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
type CellParser=ExcelRangeBase -> bool
let getColor (color:ExcelColor)=
    if color.Indexed >0 then color.LookupColor()
    else "#"+color.Rgb
let pBkColor (color:Color):CellParser=
    fun (cell:ExcelRangeBase)->
        let toHex (color:Color)=sprintf "#%02X%02X%02X%02X" color.A color.R color.G color.B
        let bkColor=cell.Style.Fill.BackgroundColor|>getColor
        let targetColor=  toHex color
        bkColor=targetColor
let pFontColor (color:Color):CellParser=
    fun (cell:ExcelRangeBase)->
        let toHex (color:Color)=sprintf "#%02X%02X%02X%02X" color.A color.R color.G color.B
        let fontColor=cell.Style.Font.Color|>getColor
        let targetColor=  toHex color
        fontColor=targetColor        

let pText (f: string -> bool) =
    fun (cell:ExcelRangeBase)->
        f cell.Text

let pRegex pattern:CellParser =
    pText (fun text ->
        let m=Regex.Match(text, pattern)
        m.Success
    )

let pFParsec (p: Parser<_,_>) =
    fun (cell:ExcelRangeBase) ->
        let text = cell.Text
        match run p text with 
        | ParserResult.Success _ -> true
        | _ -> false

let pFParsecWith (p: Parser<_,_>) f =
    fun (cell:ExcelRangeBase) ->
        let text = cell.Text
        match run p text with 
        | ParserResult.Success (r,_,_) -> f r
        | _ -> false

let pAny :CellParser=fun _->true
let (<&>) (p1:CellParser) (p2:CellParser)=
    fun (cell:ExcelRangeBase)->
        p1 cell&&p2 cell
