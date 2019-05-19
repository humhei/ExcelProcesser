namespace ExcelProcesser

open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
open System
open System.IO

type Formula =
    | SUM = 0

[<AutoOpen>]
module Operators =
    let excelPackageAndWorksheet (index: int) filename =
        if not (File.Exists filename) then
            failwithf "file %s is not existed" filename

        let file = FileInfo(filename)

        let xlPackage = new ExcelPackage(file)

        let worksheet = xlPackage.Workbook.Worksheets.[index]
        xlPackage,worksheet