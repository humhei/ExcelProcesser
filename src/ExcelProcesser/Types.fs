

namespace ExcelProcesser
#nowarn "0104"
open FParsec
open System.Collections.Generic
open NLog
open OfficeOpenXml
open CellScript.Core
open Extensions
open Shrimp.FSharp.Plus

type Formula =
    | SUM = 0

exception VirtualSingletonExcelRangeNotSupportedPropsException = exn

type VirtualSingletonExcelRange =
    { CellRelativeAddress: ComparableExcelCellAddress 
      RootData: ConvertibleUnion [,]
      Starter: ComparableExcelCellAddress }
with 
    
    member x.CellAddress = 
        x.CellRelativeAddress.Offset(x.Starter.Row, x.Starter.Column)

    member x.Address = x.CellAddress.ExcelCellAddress.Address

    member x.Offset(rowOffset, columnOffset) =
        { x with 
            CellRelativeAddress = x.CellRelativeAddress.Offset(rowOffset, columnOffset)
        }

    member x.Column = x.CellAddress.Column

    member x.Row = x.CellAddress.Row

    member x.Value = 
        x.RootData.[x.CellAddress.Row-1, x.CellAddress.Column-1]

    member x.Text = x.Value.Text


type VirtualExcelRange =
    { RelativeAddress: ComparableExcelAddress 
      RootData: ConvertibleUnion [,]
      Starter: ComparableExcelCellAddress }
with 
    member x.AsCellRanges() =
        x.RelativeAddress.AsCellAddresses()
        |> List.map(fun addr ->
            { CellRelativeAddress = addr 
              RootData = x.RootData
              Starter = x.Starter }
        )

    member x.ComparableExcelAddress: ComparableExcelAddress = 
        x.RelativeAddress.Offset(x.Starter.Row, x.Starter.Column)
     

    member x.End =
        { Row = x.ComparableExcelAddress.EndRow 
          Column = x.ComparableExcelAddress.EndColumn }

    member x.Start =
        { Row = x.ComparableExcelAddress.StartRow
          Column = x.ComparableExcelAddress.StartColumn }

    member x.Columns = x.ComparableExcelAddress.Columns

    member x.Address = x.ComparableExcelAddress.Address

    member x.Rows = x.ComparableExcelAddress.Rows

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        let rowStart = x.RelativeAddress.StartRow + rowOffset
        let columnStart = x.RelativeAddress.StartColumn + columnOffset
        let rowEnd = rowStart + numberOfRows
        let columnEnd = columnStart + numberOfColumns

        let addr = 
            { StartRow = rowStart
              StartColumn = columnStart
              EndRow  = rowEnd
              EndColumn = columnEnd }

        { RelativeAddress = addr  
          RootData = x.RootData
          Starter = x.Starter }

    member x.Rerange(address: string) =
        let address = ComparableExcelAddress.OfAddress address
        let rootData =
            x.RootData.[address.StartRow-1..address.EndRow-1, address.StartColumn-1..address.EndColumn-1]

        let starter = x.Starter.Offset(address.Start.Row-1, address.Start.Column-1)

        { RelativeAddress = 
            {StartColumn = 1
             StartRow = 1
             EndRow = address.Rows
             EndColumn = address.EndColumn }
          Starter  = starter
          RootData = rootData
        }

type VirtualSingletonExcelRange with 
    member x.AsRange =
        { RootData = x.RootData 
          Starter = x.Starter
          RelativeAddress = 
            { StartRow = x.CellRelativeAddress.Row
              EndRow = x.CellRelativeAddress.Row
              StartColumn = x.CellRelativeAddress.Column
              EndColumn = x.CellRelativeAddress.Column }
          }

    member x.RangeTo(target: VirtualSingletonExcelRange) =
        let starter =
            [x.Starter; target.Starter]
            |> List.exactlyOne_DetailFailingText

        let addr =
            x.CellRelativeAddress.RangeTo(target.CellRelativeAddress)
            |> ComparableExcelAddress.OfAddress

        { RelativeAddress  = addr 
          RootData = x.RootData
          Starter = starter }

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        x.AsRange.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)

    static member Create(range: VirtualExcelRange) =
        let address =
            try 
                { Row = 
                    [ range.RelativeAddress.StartRow 
                      range.RelativeAddress.EndRow ]
                    |> List.exactlyOne_DetailFailingText

                  Column =
                    [ range.RelativeAddress.StartColumn
                      range.RelativeAddress.EndColumn ]
                    |> List.exactlyOne_DetailFailingText
                }
            with ex ->
                failwithf "Cannot create VirtualSingletonExcelRange from %A" range.RelativeAddress


        {
            Starter = range.Starter
            RootData = range.RootData
            CellRelativeAddress = address
        }


[<RequireQualifiedAccess>]
type ExcelRangeUnion =
    | Office of ExcelRangeBase
    | Virtual of VirtualExcelRange
with 
    member x.Rerange(address: string) =
        match x with 
        | Office v -> v.Worksheet.Cells.[address]  :> ExcelRangeBase  |> ExcelRangeUnion.Office
        | Virtual v -> 
            v.Rerange(address)
            |> ExcelRangeUnion.Virtual

    member x.Address =
        match x with 
        | Office v -> v.Address
        | Virtual v -> v.Address

    member x.Columns =
        match x with 
        | Office v -> v.Columns
        | Virtual v -> v.Columns

    member x.Rows =
        match x with 
        | Office v  -> v.Rows
        | Virtual v -> v.Rows

    member x.ComparableExcelAddress() =
        match x with 
        | Office v -> ComparableExcelAddress.OfRange v
        | Virtual v -> v.ComparableExcelAddress

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        match x with 
        | Office  v -> v.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) |> Office 
        | Virtual v -> v.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) |> Virtual

    member x.End =
        match x with 
        | Office  v -> ComparableExcelCellAddress.OfExcelCellAddress v.End
        | Virtual v -> v.End

    member x.Start =
        match x with 
        | Office  v -> ComparableExcelCellAddress.OfExcelCellAddress v.Start
        | Virtual v -> v.Start



[<RequireQualifiedAccess>]
type SingletonExcelRangeBaseUnion =
    | Office of SingletonExcelRangeBase
    | Virtual of VirtualSingletonExcelRange
with 
        
    member x.Value =
        match x with 
        | Office v -> v.Value
        | Virtual v -> v.Value.Value

    member x.Rerange(address: string) =
        match x with 
        | Office v -> v.Worksheet.Cells.[address]  :> ExcelRangeBase  |> ExcelRangeUnion.Office
        | Virtual v -> 
            v.AsRange.Rerange(address)
            |> ExcelRangeUnion.Virtual
   

    static member Create(range: ExcelRangeUnion) =
        match range with 
        | ExcelRangeUnion.Office v -> SingletonExcelRangeBase.Create v  |> SingletonExcelRangeBaseUnion.Office
        | ExcelRangeUnion.Virtual v -> 
            VirtualSingletonExcelRange.Create(v)
            |> SingletonExcelRangeBaseUnion.Virtual

    static member Create(range: ExcelRangeBase) =
        SingletonExcelRangeBaseUnion.Create(ExcelRangeUnion.Office range)

    member x.RangeTo(target) =
        match x, target with 
        | Office  v, Office target -> v.RangeTo(target)      |> ExcelRangeUnion.Office
        | Virtual v, Virtual target -> v.RangeTo(target)     |> ExcelRangeUnion.Virtual
        | _ -> failwithf "Invalid token, %A and %A are not the same type" x target

    member x.Column = 
        match x with 
        | Office  v -> v.Column 
        | Virtual v -> v.Column
            
    member x.Row = 
        match x with 
        | Office  v -> v.Row
        | Virtual v -> v.Row

    member x.Offset(rowOffset, columnOffset) =
        match x with 
        | Office  v -> v.Offset(rowOffset, columnOffset)  |> Office 
        | Virtual v -> v.Offset(rowOffset, columnOffset)  |> Virtual
            
    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        match x with 
        | Office  v -> v.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)   |> ExcelRangeUnion.Office
        | Virtual v -> v.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)   |> ExcelRangeUnion.Virtual
            

    member x.Style =
        match x with 
        | Office v -> v.Style
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException

    member x.StyleName =
        match x with 
        | Office v -> v.StyleName
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException

    member x.Formula = 
        match x with 
        | Office v -> v.Formula
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException

    member x.GetMergeCellId() =
        match x with 
        | Office v -> v.GetMergeCellId()
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException

    member x.WorksheetOrFail =
        match x with 
        | Office v -> v.Worksheet
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException


    member x.Merge = 
        match x with 
        | Office v -> v.Merge
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException

    member x.TryGetMergedRangeAddress() =
        match x with 
        | Office v -> v.TryGetMergedRangeAddress()
        | Virtual _ -> raise VirtualSingletonExcelRangeNotSupportedPropsException


    member x.Text = 
        match x with 
        | Office v -> v.Text
        | Virtual v -> v.Text

    member x.Address =  
        match x with 
        | Office v -> v.Address
        | Virtual v -> v.Address
        

    member x.ExcelCellAddress =
        match x with 
        | Office v -> v.ExcelCellAddress
        | Virtual v -> v.CellAddress

    member x.ExcelAddress =
        match x with 
        | Office v -> v.ExcelAddress
        | Virtual v -> 
            v.AsRange.ComparableExcelAddress
        

[<RequireQualifiedAccess>]
module SingletonExcelRangeBaseUnion =
    let getText (range: SingletonExcelRangeBaseUnion) =
        range.Text

    let getExcelCellAddress (range: SingletonExcelRangeBaseUnion) = 
        range.ExcelCellAddress


    let tryGetMergedRangeAddress(range: SingletonExcelRangeBaseUnion) =
        range.TryGetMergedRangeAddress()

    let getValue(range: SingletonExcelRangeBaseUnion) =
        match range with 
        | SingletonExcelRangeBaseUnion.Office v -> SingletonExcelRangeBase.getValue v
        | SingletonExcelRangeBaseUnion.Virtual v -> v.Value.Value 

[<RequireQualifiedAccess>]
module ExcelRangeUnion =
    let asRangeList (range: ExcelRangeUnion) =
        match range with 
        | ExcelRangeUnion.Office v -> 
            ExcelRangeBase.asRangeList v
            |> List.map SingletonExcelRangeBaseUnion.Office

        | ExcelRangeUnion.Virtual v ->
            v.AsCellRanges()
            |> List.map SingletonExcelRangeBaseUnion.Virtual


    let asRangeList_All (range: ExcelRangeUnion) =
        match range with 
        | ExcelRangeUnion.Office v -> 
            ExcelRangeBase.asRangeList_All v
            |> List.map SingletonExcelRangeBaseUnion.Office
        | ExcelRangeUnion.Virtual v ->
            v.AsCellRanges()
            |> List.map SingletonExcelRangeBaseUnion.Virtual

[<AutoOpen>]
module Operators =
    let internal isTrimmedTextEmpty (text: string) = text.Trim() = ""
    let internal isTrimmedTextNotEmpty (text: string) = text.Trim() <> ""


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