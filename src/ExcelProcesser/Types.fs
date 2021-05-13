namespace ExcelProcesser
#nowarn "0104"
open FParsec
open System.Drawing
open OfficeOpenXml
open System.Text.RegularExpressions
open OfficeOpenXml.Style
open System
open System.IO
open System.Collections.Generic

type Formula =
    | SUM = 0

[<AutoOpen>]
module Operators =

    [<RequireQualifiedAccess>]
    type private LoggerMsg =
        | Info of AsyncReplyChannel<unit> * string

    type Logger(loggerLevel: LoggerLevel) =
        let nlog = NLog.LogManager.GetCurrentClassLogger()

        let messages = new List<string>()

        member x.Info (message: string) = 
            match loggerLevel with 
            | LoggerLevel.Trace_Red  -> 
                nlog.Error message
                messages.Add(message)
            | LoggerLevel.Slient -> ()


        member x.Messages() = List.ofSeq messages

    let ensureFParsecValid text parser  =
        match run parser text with 
        | Success _ -> parser
        | Failure (error, _, _) -> failwith error