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




[<DebuggerDisplay("{ExcelCellAddress}")>]
type ComparableExcelCellAddress =
    { Row: int 
      Column: int }
with 
    static member OfExcelCellAddress(address: ExcelCellAddress) =
        { Row = address.Row 
          Column = address.Column }

    static member OfAddress(address: string) =
        ComparableExcelCellAddress.OfExcelCellAddress(ExcelCellAddress(address))

    member x.ExcelCellAddress =
        ExcelCellAddress(x.Row, x.Column)

    member x.Address = x.ExcelCellAddress.Address

[<DebuggerDisplay("{ExcelAddress}")>]
type ComparableExcelAddress =
    { StartRow: int 
      EndRow: int
      StartColumn: int 
      EndColumn: int 
      }
with 
    member x.Start: ComparableExcelCellAddress =
        { Row = x.StartRow 
          Column = x.StartColumn }

    member x.End: ComparableExcelCellAddress =
        { Row = x.EndRow
          Column = x.EndColumn }

    member x.Rows = x.EndRow - x.StartRow + 1

    member x.Columns = x.EndColumn - x.StartColumn + 1

    static member OfAddress(excelAddress: ExcelAddress) =
        let startCell = excelAddress.Start

        let endCell = excelAddress.End
        {
            StartRow = startCell.Row
            EndRow = endCell.Row
            StartColumn = startCell.Column
            EndColumn = endCell.Column
        }

    static member OfAddress(address: string) =
        ComparableExcelAddress.OfAddress(ExcelAddress(address))

    static member OfRange(range: ExcelRangeBase) =
        let startCell = range.Start

        let endCell = range.End
        {
            StartRow = startCell.Row
            EndRow = endCell.Row
            StartColumn = startCell.Column
            EndColumn = endCell.Column
        }



    member x.ExcelAddress =
        ExcelAddress(x.StartRow, x.StartColumn, x.EndRow, x.EndColumn)
    
    member x.Address = x.ExcelAddress.Address

    member x.Contains(y: ComparableExcelAddress) = 
        match x.Start.Column, x.Start.Row, x.End.Column, x.End.Row with 
        | SmallerOrEqual y.Start.Column, SmallerOrEqual y.Start.Row, BiggerOrEqual y.End.Column, BiggerOrEqual y.End.Row ->
            true
        | _ -> false

    member x.IsIncludedIn(y: ComparableExcelAddress) = y.Contains(x)


type ComparableExcelCellAddress with 
    member x.RangeTo(y: ComparableExcelCellAddress) =
        x.Address + ":" + y.Address

[<DebuggerDisplay("{Address} {Text}")>]
[<StructuredFormatDisplay("{Address} {Text}")>]
type SingletonExcelRangeBase private (range: ExcelRangeBase) =
    
    member x.Value = range.Value

    member x.Style = range.Style

    member x.StyleName = range.StyleName

    member x.StyleID = range.StyleID

    member x.Formula = range.Formula

    member x.Worksheet = range.Worksheet

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) = range.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)
    
    member x.Offset(rowOffset, columnOffset) = 
        range.Offset(rowOffset, columnOffset)
        |> SingletonExcelRangeBase.Create

    member x.Address = range.Address

    member val ExcelCellAddress = ComparableExcelCellAddress.OfExcelCellAddress(ExcelCellAddress(range.Address))

    member x.Text = range.Text

    member x.Row = range.Start.Row

    member x.Column = range.Start.Column

    member x.RangeTo(targetRange: SingletonExcelRangeBase) =
        let addr = x.ExcelCellAddress.RangeTo(targetRange.ExcelCellAddress)
        range.Worksheet.Cells.[addr]

    member x.Merge = range.Merge
    
    member x.TryGetMergeCellId() =
        match x.Merge with 
        | true ->
            x.Worksheet.GetMergeCellId (x.Row, x.Column)
            |> Some

        | false -> None

    member x.GetMergeCellId() =
        x.Worksheet.GetMergeCellId (x.Row, x.Column)

    member range.TryGetMergedRangeAddress() =
        match range.TryGetMergeCellId() with 
        | Some id ->
            let addr = 
                range.Worksheet.MergedCells.[id-1]
                |> ComparableExcelAddress.OfAddress
                |> Some

            addr

        | None -> None

    static member Create (excelRangeBase: ExcelRangeBase) =
        match excelRangeBase.Columns, excelRangeBase.Rows with 
        | 1, 1 -> SingletonExcelRangeBase excelRangeBase
        | _ -> failwithf "Cannot create SingletonExcelRangeBase when columns is %d and rows is %d" excelRangeBase.Columns excelRangeBase.Rows



[<RequireQualifiedAccess>]
module SingletonExcelRangeBase =
    let getValue (range: SingletonExcelRangeBase) =
        range.Value

    let getText (range: SingletonExcelRangeBase) = 
        range.Text

    let getExcelCellAddress (range: SingletonExcelRangeBase) = 
        range.ExcelCellAddress


    let tryGetMergedRangeAddress(range: SingletonExcelRangeBase) =
        range.TryGetMergedRangeAddress()

[<AutoOpen>]
module AutoOpen_Extensions =
    [<RequireQualifiedAccess>]
    module String =
        let contains pattern (text: string) =
            text.Contains pattern

module Extensions =

    //[<RequireQualifiedAccess>]
    //module internal Array2D =

    //    let joinByRows (a1: 'a[,]) (a2: 'a[,]) =
    //        let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
    //        if a1l2 <> a2l2 then failwith "arrays have different column sizes"
    //        let result = Array2D.zeroCreate (a1l1 + a2l1) a1l2
    //        Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
    //        Array2D.blit a2 0 0 result a1l1 0 a2l1 a2l2
    //        result

    //    let concatByRows (array2DList: 'a[,] seq) =
    //        if Seq.length array2DList = 0 then Array2D.zeroCreate 0 0
    //        else
    //            array2DList
    //            |> Seq.reduce joinByRows

    //    let joinByCols (a1: 'a[,]) (a2: 'a[,]) =
    //        let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
    //        if a1l1 <> a2l1 then failwith "arrays have different row sizes"
    //        let result = Array2D.zeroCreate a1l1 (a1l2 + a2l2)
    //        Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
    //        Array2D.blit a2 0 0 result 0 a1l2 a2l1 a2l2
    //        result


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





    [<RequireQualifiedAccess>]
    module ExcelRangeBase =

        let asRangeList (range: ExcelRangeBase) =
            range :> seq<ExcelRangeBase>
            |> List.ofSeq
            |> List.map SingletonExcelRangeBase.Create


        /// Including Empty Ranges
        let asRangeList_All  (range: ExcelRangeBase) =
            let originColumns = range.Columns
            let rec loop rows columns accum (range: SingletonExcelRangeBase) =
                match rows, columns with 
                | 1, 1 -> (range :: accum) |> List.rev
                | _, BiggerThan 1 ->
                    loop rows (columns - 1) (range :: accum) (range.Offset(0, 1))
                 
                | BiggerThan 1, 1 ->
                    loop (rows - 1) (originColumns) (range :: accum) (range.Offset(1, 0))
                | _ -> failwith "Invalid token"

            loop (range.Rows) originColumns [] (SingletonExcelRangeBase.Create(range.Offset(0, 0, 1, 1)))


        let getText (range: ExcelRangeBase) = range.Text
        let getValue (range: ExcelRangeBase) = range.Value

        let getComparableAddress (range: ExcelRangeBase) = ComparableExcelAddress.OfRange(range)

