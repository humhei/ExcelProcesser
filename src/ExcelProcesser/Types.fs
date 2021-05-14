

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
open NLog

type Formula =
    | SUM = 0

[<AutoOpen>]
module Operators =


    let nlog = NLog.LogManager.GetCurrentClassLogger()

    type Messages =
        { Infos: string list 
          Importants: string list
          AllMessages: string list }
    with 
        member x.IsEmpty = x.Infos.IsEmpty && x.Importants.IsEmpty



    type Logger() =
        do
            GlobalDiagnosticsContext.Set("Application", "My cool app");
        

        let infos = new List<string>()
        
        let imports = new List<string>()
        let allMessages = new List<string>()

        member x.Log loggerLevel (message: string) = 
            match loggerLevel with 
            | LoggerLevel.Info  -> 
                nlog.Info message
                infos.Add(message)
                allMessages.Add(message)

            | LoggerLevel.Important ->
                nlog.Error message
                imports.Add(message)
                allMessages.Add(message)

            | LoggerLevel.Slient -> ()


        member x.Messages() = 
            { Infos = List.ofSeq infos
              Importants = List.ofSeq imports
              AllMessages = List.ofSeq allMessages }


    let ensureFParsecValid text parser  =
        match run parser text with 
        | Success _ -> parser
        | Failure (error, _, _) -> failwith error