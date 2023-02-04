// Learn more about F# at http://fsharp.org

namespace ExcelProcesser
open CellScript.Core

open OfficeOpenXml
open OfficeOpenXml.Style
open Shrimp.FSharp.Plus
open CellScript.Core.Extensions
open System.Diagnostics


type LoggerLevel = 
    | Info = 0
    | Important = 1
    | Slient = 2





[<AutoOpen>]
module AutoOpen_Extensions =
    [<RequireQualifiedAccess>]
    module String =
        let contains pattern (text: string) =
            text.Contains pattern


[<DebuggerDisplay("{Address} {Text}")>]
[<StructuredFormatDisplay("{Address} {Text}")>]
type SingletonExcelRangeBase private (range: ExcelRangeBase) =
    let addr = ComparableExcelAddress.OfRange range

    let cellAddr = ComparableExcelCellAddress.OfExcelCellAddress(ExcelCellAddress(range.Address))
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

    member x.ExcelCellAddress = cellAddr

    member x.ExcelAddress = addr

    member x.Text = range.Text

    member x.Row = range.Start.Row

    member x.Column = range.Start.Column

    member x.RangeTo(targetRange: SingletonExcelRangeBase) =
        let addr = x.ExcelCellAddress.RangeTo(targetRange.ExcelCellAddress)
        range.Worksheet.Cells.[addr] :> ExcelRangeBase

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
        | _ -> failwithf "Cannot create SingletonExcelRangeBaseUnion when columns is %d and rows is %d" excelRangeBase.Columns excelRangeBase.Rows



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

        let getUserRangeListWith maximumEmptyColumns (worksheet: ExcelWorksheet) =
            let maxRow = getMaxRowNumber worksheet
            let maxCol = getMaxColNumber worksheet

            match maximumEmptyColumns with 
            | None ->
                let r = 
                    [ for row in 1..maxRow do
                          for col in 1..maxCol do
                              let content = worksheet.Cells.[row, col]
                              yield SingletonExcelRangeBase.Create(content :> ExcelRangeBase) ]

                {|MaxCol = maxCol; MaxRow = maxRow; UserRange = r |}

            | Some maximumEmptyColumns ->
                let rec loop columns (emptyColumns: _ list) colNum =
                    match emptyColumns.Length with 
                    | BiggerThan maximumEmptyColumns -> columns
                    | _ ->
                        match colNum with 
                        | BiggerThan maxCol -> columns
                        | _ ->
                            let column = 
                                [1..maxRow]
                                |> List.map(fun row ->
                                    let content = worksheet.Cells.[row, colNum]
                                    SingletonExcelRangeBase.Create(content :> ExcelRangeBase) 
                                )

                            let isColumnEmpty = 
                                column
                                |> List.forall(fun (m: SingletonExcelRangeBase) -> isNull m.Value)

                            match isColumnEmpty with 
                            | true ->
                                loop (columns) (column :: emptyColumns) (colNum + 1)

                            | false ->
                                loop (column :: emptyColumns @ columns) [] (colNum + 1)
                                

                let r = 
                    loop [] [] 1
                    |> List.rev

                let r = 
                    r
                    |> array2D
                    |> Array2D.transpose

                let r2 =
                    r
                    |> Array2D.toLists
                    |> List.concat

                {|MaxCol = Array2D.length2 r; MaxRow = Array2D.length1 r; UserRange = r2 |}

        let getUserRangeList (worksheet: ExcelWorksheet) =
            getUserRangeListWith None worksheet


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

