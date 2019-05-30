// Learn more about F# at http://fsharp.org

namespace ExcelProcesser

open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
open System
open System.IO

module Extensions =

    [<RequireQualifiedAccess>]
    module ExcelColor =
        let getColorHex (color: ExcelColor)=
            if color.Indexed > 0 then color.LookupColor()
            else "#" + color.Rgb


    [<RequireQualifiedAccess>]
    module ExcelWorksheet =

        let private getMaxColNumber (worksheet:ExcelWorksheet) =
            worksheet.Dimension.End.Column

        let private getMaxRowNumber (worksheet:ExcelWorksheet) =
            worksheet.Dimension.End.Row

        let getUserRange worksheet =
            [ let maxRow = getMaxRowNumber worksheet
              let maxCol = getMaxColNumber worksheet
              for i in 1..maxRow do
                  for j in 1..maxCol do
                      let content = worksheet.Cells.[i, j]
                      yield content :> ExcelRangeBase ]

        let getMergeCellId (range: ExcelRangeBase) (worksheet: ExcelWorksheet) =
            worksheet.GetMergeCellId (range.Start.Row, range.Start.Column)

    [<RequireQualifiedAccess>]
    module ExcelRangeBase =

        let asRanges (range: ExcelRangeBase) =
            range :> seq<ExcelRangeBase>
            |> List.ofSeq

        let getText (range: ExcelRangeBase) = range.Text

        let getAddressOfRange (range: ExcelRangeBase) = range.Address


