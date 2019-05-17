namespace ExcelProcess
open Fake.Core
open NLog



[<AutoOpen>]
module Global =
    let logger = LogManager.GetCurrentClassLogger()
