module Parser
open System.Text.RegularExpressions
open OfficeOpenXml
open CellParsers
open System.Diagnostics
open System.Drawing
open ArrayParsers
let matchRegex pattern (cell:ExcelRangeBase)=
    let m=Regex.Match(cell.Text, pattern)
    m.Success
let IRunPaser (fileName:string)=
    let a=Stopwatch()
    a.Start()
    let parser:ArrayParser=
        filter[
            !@(pRegex("GF.*双装")<&>pBkColor Color.Yellow) .>>. xShift 3 .>>. !@(pFontColor Color.Blue)
            yShift 2
            !@(pFontColor Color.Red) .>>. !@(pBkColor Color.Yellow) .>>. xShift 3
            ]
    let numberData=
        fileName
        |>Excel.getWorksheetByIndex 1
        |>Excel.runParser parser
    let t=numberData
    let w=t
    printf "%A" a.Elapsed
    ()    
