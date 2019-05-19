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
        let bkColor = cell.Style.Fill.BackgroundColor |> ExcelColor.getColorHex
        let targetColor = toHex color
        bkColor = targetColor

let pFontColor (color : Color) : CellParser =
    fun (cell : ExcelRangeBase) ->
        let toHex (color : Color) =
            sprintf "#%02X%02X%02X%02X" color.A color.R color.G color.B
        let fontColor = cell.Style.Font.Color |> ExcelColor.getColorHex
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
        | ParserResult.Success _ -> true
        | _ -> false

let pFParsecWithMappingResult (p : Parser<_, _>) mapping =
    fun (cell : ExcelRangeBase) ->
        let text = cell.Text
        match run p text with
        | ParserResult.Success(r, _, _) -> mapping r
        | _ -> false

let pAny : CellParser = fun _ -> true
