#r "nuget: ExcelProcesser"
open FParsec
open ExcelProcesser
open MatrixParsers
open OfficeOpenXml
open System.IO
open ExcelProcesser.Extensions
open CellScript.Core
open Shrimp.FSharp.Plus 

let file = XlsxFile @"resources/real world samples/19SPX16合同生产细节.xlsx"
let excelPackage = new ExcelPackage(FileInfo file.Path)
let worksheet = excelPackage.GetValidWorksheet(SheetGettingOptions.DefaultValue)

ExcelProcesserLoggerLevel <- LoggerLevel.Trace_Red

let parser = 
    mxTextf (fun text -> text.StartsWith "KOI")
    |> MatrixParser.addLogger "artParser"

let results = runMatrixParser worksheet parser

excelPackage.Dispose()