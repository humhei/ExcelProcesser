// Learn more about F# at http://fsharp.org

namespace ExcelProcesser

open OfficeOpenXml
open OfficeOpenXml.Style
open Shrimp.FSharp.Plus
open System.Diagnostics

type LoggerLevel = 
    | Info = 0
    | Important = 1
    | Slient = 2




[<DebuggerDisplay("{Value.Address}")>]
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
    
    member x.Offset(rowOffset, columnOffset) = 
        x.Value.Offset(rowOffset, columnOffset)
        |> SingletonExcelRangeBase.Create

    member x.Address = x.Value.Address

    member x.Text = x.Value.Text

    member x.Row = x.Value.Start.Row

    member x.Column = x.Value.Start.Column

    member x.RangeTo(targetRange: SingletonExcelRangeBase) =
        let addr = x.Value.Address + ":" + targetRange.Address
        x.Value.Worksheet.Cells.[addr]



[<RequireQualifiedAccess>]
module SingletonExcelRangeBase =
    let getValue (range: SingletonExcelRangeBase) =
        range.Value

    let getText (range: SingletonExcelRangeBase) = 
        range.Value.Text

    let tryGetMergedRange(range: SingletonExcelRangeBase) =
        let range = range.Value
        let worksheet = range.Worksheet
        match range.Merge with 
        | true ->
            let id = worksheet.GetMergeCellId (range.Start.Row, range.Start.Column)

            let addr = 
                worksheet.MergedCells.[id-1]
                |> ExcelAddress
                |> Some

            addr

        | false -> None

[<AutoOpen>]
module AutoOpen_Extensions =
    

    type ExcelAddress with 
        member x.Contains(y: ExcelAddress) = 


            match x.Start.Column, x.Start.Row, x.End.Column, x.End.Row with 
            | SmallerOrEqual y.Start.Column, SmallerOrEqual y.Start.Row, BiggerOrEqual y.End.Column, BiggerOrEqual y.End.Row ->
                true
            | _ -> false

        member x.IsIncludedIn(y: ExcelAddress) = y.Contains(x)

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

        let internal getMaxColNumber (worksheet:ExcelWorksheet) =
            worksheet.Dimension.End.Column

        let internal getMaxRowNumber (worksheet:ExcelWorksheet) =
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
        let getValue (range: ExcelRangeBase) = range.Value

        let getAddress (range: ExcelRangeBase) = range.Address

