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

