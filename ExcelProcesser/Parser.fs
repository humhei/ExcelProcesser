module Parser
open System.Drawing
open System.Text.RegularExpressions
open OfficeOpenXml
open CellParsers
open RangeParsers
open System.Diagnostics

let matchRegex pattern (cell:ExcelRangeBase)=
    let m=Regex.Match(cell.Text, pattern)
    m.Success
let IRunPaser (fileName:string)=
    let a=Stopwatch()
    a.Start()
    let parser=
       many <| (manyTill <| !@(pRegex("GF.*双装")<&>pColor Color.Yellow))
    let numberData=
        fileName
        |>Excel.getWorksheetByIndex 1
        |>Excel.runParser parser
    let t=numberData|>Option.get
    let w=t
    let t =numberData
    printf "%A" a.Elapsed
    // printfn "%A" numberData
    ()    
    // let numberData=
    //     fileName
    //     |>Excel.getWorksheetByIndex 1
    //     |>Excel.getUserRange
    //     |>Seq.where(pRegex("GF.*双装"))
    //     |>List.ofSeq
    // let t =numberData
    // printf "%A" a.Elapsed
    // // printfn "%A" numberData
    // ()
