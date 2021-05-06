// Learn more about F# at http://fsharp.org

namespace ExcelProcesser

open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
open System
open System.IO
open CellScript.Core
open Shrimp.FSharp.Plus

type LoggerLevel = 
    | Trace_Red = 0
    | Slient = 1

type SingletonExcelRangeBase = private SingletonExcelRangeBase of ExcelRangeBase
with 
    static member Create (excelRangeBase: ExcelRangeBase) =
        match excelRangeBase.Columns, excelRangeBase.Rows with 
        | 1, 1 -> SingletonExcelRangeBase excelRangeBase
        | _ -> failwithf "Cannot create SingletonExcelRangeBase when columns is %d and rows is %d" excelRangeBase.Columns excelRangeBase.Rows


    member x.Value =
        let (SingletonExcelRangeBase value) = x
        value

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) = x.Value.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)

[<RequireQualifiedAccess>]
module SingletonExcelRangeBase =
    let getValue (range: SingletonExcelRangeBase) =
        range.Value

    let getText (range: SingletonExcelRangeBase) = 
        range.Value.Text

[<AutoOpen>]
module AutoOpen_Extensions =
    
    [<RequireQualifiedAccess>]
    module String =
        let contains pattern (text: string) =
            text.Contains pattern

module Extensions =

    [<RequireQualifiedAccess>]
    module internal Array2D =

        let joinByRows (a1: 'a[,]) (a2: 'a[,]) =
            let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
            if a1l2 <> a2l2 then failwith "arrays have different column sizes"
            let result = Array2D.zeroCreate (a1l1 + a2l1) a1l2
            Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
            Array2D.blit a2 0 0 result a1l1 0 a2l1 a2l2
            result

        let concatByRows (array2DList: 'a[,] seq) =
            if Seq.length array2DList = 0 then Array2D.zeroCreate 0 0
            else
                array2DList
                |> Seq.reduce joinByRows

        let joinByCols (a1: 'a[,]) (a2: 'a[,]) =
            let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
            if a1l1 <> a2l1 then failwith "arrays have different row sizes"
            let result = Array2D.zeroCreate a1l1 (a1l2 + a2l2)
            Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
            Array2D.blit a2 0 0 result 0 a1l2 a2l1 a2l2
            result


    [<RequireQualifiedAccess>]
    module ExcelColor =
        let hex (color: ExcelColor)=
            if color.Indexed > 0 then color.LookupColor()
            else "#" + color.Rgb


    [<RequireQualifiedAccess>]
    module ExcelWorksheet =

        let private getMaxColNumber (worksheet:ExcelWorksheet) =
            worksheet.Dimension.End.Column

        let private getMaxRowNumber (worksheet:ExcelWorksheet) =
            worksheet.Dimension.End.Row

        let getUserRangeList (worksheet: ExcelWorksheet) =
            [ let maxRow = getMaxRowNumber worksheet
              let maxCol = getMaxColNumber worksheet
              for i in 1..maxRow do
                  for j in 1..maxCol do
                      let content = worksheet.Cells.[i, j]
                      yield SingletonExcelRangeBase.Create(content :> ExcelRangeBase) ]

        let getMergeCellIdOfRange (range: ExcelRangeBase) (worksheet: ExcelWorksheet) =
            worksheet.GetMergeCellId (range.Start.Row, range.Start.Column)

    [<RequireQualifiedAccess>]
    module ExcelRangeBase =

        let asRangeList (range: ExcelRangeBase) =
            range :> seq<ExcelRangeBase>
            |> List.ofSeq
            |> List.map SingletonExcelRangeBase.Create

        let getText (range: ExcelRangeBase) = range.Text

        let getAddress (range: ExcelRangeBase) = range.Address

