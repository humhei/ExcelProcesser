module ExcelProcesser.CellParsers

open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
open System
open Extensions

type CellParser = ExcelRangeBase -> bool


let pBkColor (color : Color) : CellParser =
    fun (cell : ExcelRangeBase) ->
        let toHex (color : Color) =
            sprintf "#%02X%02X%02X%02X" color.A color.R color.G color.B
        let bkColor = cell.Style.Fill.BackgroundColor |> ExcelColor.hex
        let targetColor = toHex color
        bkColor = targetColor

let pFontColor (color : Color) : CellParser =
    fun (cell : ExcelRangeBase) ->
        let toHex (color : Color) =
            sprintf "#%02X%02X%02X%02X" color.A color.R color.G color.B
        let fontColor = cell.Style.Font.Color |> ExcelColor.hex
        let targetColor = toHex color
        fontColor = targetColor

let pTextf (f : string -> bool) = fun (cell : ExcelRangeBase) -> f cell.Text

let pTextContain text = pTextf (fun cellText -> cellText.Contains text)

let pText text = pTextf (fun cellText -> cellText = text)

let pStyleName (styleName : string) : CellParser =
    fun (cell : ExcelRangeBase) -> cell.StyleName = styleName

let pRegex pattern : CellParser =
    pTextf (fun text ->
        let m = Regex.Match(text, pattern)
        m.Success)

let pFormula (firstFormula : Formula) =
    fun (cell : ExcelRangeBase) ->
        let formula = cell.Formula
        //if cell.Address = "B13" then printf ""
        formula.StartsWith(Enum.GetName(typeof<Formula>, firstFormula) + "(")


let pFParsec (p : Parser<_, _>) =
    fun (cell : ExcelRangeBase) ->
        let text = cell.Text
        match run p text with
        | ParserResult.Success (result, _, _) -> Some result
        | _ -> None


let pSpace range = pTextf (fun text -> text.Trim() = "") range

let pAny : CellParser = fun _ -> true

let pMerge = fun (range: ExcelRangeBase) -> range.Merge

let pMergeStarter = fun (range: ExcelRangeBase) -> 
    if range.Merge then 
        let worksheet = range.Worksheet
        let index = ExcelWorksheet.getMergeCellIdOfRange range worksheet
        let addr = worksheet.MergedCells.[index-1]
        worksheet.Cells.[addr].Start.Address = range.Address
    else false
