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
        | Info of string

    type Logger(loggerLevel: LoggerLevel) =
        let mailbox = MailboxProcessor.Start(fun inbox ->
            let nlog = NLog.LogManager.GetCurrentClassLogger()
            
            let rec loop (traces: List<string>) = async {
                let! msg = inbox.Receive()
                match msg with 
                | LoggerMsg.Info message -> 
                    match loggerLevel with 
                    | LoggerLevel.Trace_Red  -> 
                        nlog.Error message
                        traces.Add(message)
                        return! loop traces
                    | LoggerLevel.Slient -> ()

            }
            loop (new List<_>())
        
        )

        member x.Info line = mailbox.Post (LoggerMsg.Info line)

