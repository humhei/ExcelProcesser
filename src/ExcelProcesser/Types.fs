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

    let mutable ExcelProcesserLoggerLevel = LoggerLevel.Slient

    let ensureFParsecValid text parser  =
        match run parser text with 
        | Success _ -> parser
        | Failure (error, _, _) -> failwith error