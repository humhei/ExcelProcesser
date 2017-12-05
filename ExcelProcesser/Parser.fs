module Parser
open System.Text.RegularExpressions
open OfficeOpenXml
open CellParsers
open System.Diagnostics
open MatrixParsers
open System.Drawing

let matchRegex pattern (cell:ExcelRangeBase)=
    let m=Regex.Match(cell.Text, pattern)
    m.Success
let IRunPaser (fileName:string)=
    let a=Stopwatch()
    a.Start()
    let parser:MatrixParser=
       [PLRow [
           CellParser (pRegex("GF.*双装")<&>pBkColor Color.Yellow) 
           .>>. CellParser (pFontColor Color.Blue) 
           .>>. AnyCell 2]
        AnyRow 1 ]
    let numberData=
        fileName
        |>Excel.getWorksheetByIndex 1
        |>Excel.runParser parser
    let t=numberData
    let w=t
    let t =numberData
    printf "%A" a.Elapsed
    ()    
