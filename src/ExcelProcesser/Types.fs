

namespace ExcelProcesser

open System.Diagnostics

#nowarn "0104"
open FParsec
open CellScript.Core.Extensions
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
    { ExcelCellAddress: ComparableExcelCellAddress 
      RootData: ConvertibleUnion [,] }
with 
        

    member x.Address = x.ExcelCellAddress.ExcelCellAddress.Address

    member x.Offset(rowOffset, columnOffset) =
        { x with 
            ExcelCellAddress = x.ExcelCellAddress.Offset(rowOffset, columnOffset)
        }

    member x.Column = x.ExcelCellAddress.Column

    member x.Row = x.ExcelCellAddress.Row

    member x.Value = 
        let l1 = Array2D.length1 x.RootData
        let l2 = Array2D.length2 x.RootData
        let row = x.ExcelCellAddress.Row
        let column = x.ExcelCellAddress.Column

        if row > l1 || column > l2 
        then ConvertibleUnion.Missing
        else x.RootData.[row-1, column-1]

    member x.Text = x.Value.Text


type VirtualExcelRange =
    { ExcelAddress: ComparableExcelAddress 
      RootData: ConvertibleUnion [,]  }
with 
    static member OfData(data: ConvertibleUnion [,]) =
        { ExcelAddress =
            { StartRow    = 1
              StartColumn = 1
              EndRow      = Array2D.length1 data
              EndColumn   = Array2D.length2 data 
            }
          RootData = data
        }

    member x.AsCellRanges() =
        x.ExcelAddress.AsCellAddresses()
        |> List.map(fun addr ->
            { ExcelCellAddress = addr 
              RootData = x.RootData  }
        )
        |> List.filter(fun m -> 
            match m.Value with 
            | ConvertibleUnion.Missing _ -> false
            | _ -> true
        )

    member x.AsCellRanges_All() =
        x.ExcelAddress.AsCellAddresses()
        |> List.map(fun addr ->
            { ExcelCellAddress = addr 
              RootData = x.RootData  }
        )


    member x.ComparableExcelAddress: ComparableExcelAddress = x.ExcelAddress
     

    member x.End =
        { Row = x.ComparableExcelAddress.EndRow 
          Column = x.ComparableExcelAddress.EndColumn }

    member x.Start =
        { Row = x.ComparableExcelAddress.StartRow
          Column = x.ComparableExcelAddress.StartColumn }

    member internal x.StarterValue = 
        let start = x.Start
        x.RootData.[start.Row-1, start.Column-1]


    member x.Columns = x.ComparableExcelAddress.Columns

    member x.Address = x.ComparableExcelAddress.Address

    member x.Rows = x.ComparableExcelAddress.Rows

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        let addr = x.ExcelAddress.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) 

        { ExcelAddress = addr  
          RootData = x.RootData }

    member x.Rerange(address: string) =
        let address = ComparableExcelAddress.OfAddress address

        { ExcelAddress = address
          RootData = x.RootData
        }

type VirtualSingletonExcelRange with 
    member x.AsRange =
        { RootData = x.RootData 
          ExcelAddress = 
            { StartRow = x.ExcelCellAddress.Row
              EndRow = x.ExcelCellAddress.Row
              StartColumn = x.ExcelCellAddress.Column
              EndColumn = x.ExcelCellAddress.Column }
          }

    member x.RangeTo(target: VirtualSingletonExcelRange) =
        let addr =
            x.ExcelCellAddress.RangeTo(target.ExcelCellAddress)
            |> ComparableExcelAddress.OfAddress

        { ExcelAddress  = addr 
          RootData = x.RootData }

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        x.AsRange.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns)

    static member Create(range: VirtualExcelRange) =
        let address =
            try 
                { Row = 
                    [ range.ExcelAddress.StartRow 
                      range.ExcelAddress.EndRow ]
                    |> List.exactlyOne_DetailFailingText

                  Column =
                    [ range.ExcelAddress.StartColumn
                      range.ExcelAddress.EndColumn ]
                    |> List.exactlyOne_DetailFailingText
                }
            with ex ->
                failwithf "Cannot create VirtualSingletonExcelRange from %A" range.ExcelAddress


        {
            RootData = range.RootData
            ExcelCellAddress = address
        }

[<DebuggerDisplay("{Address} {Text}")>]
[<StructuredFormatDisplay("{Address} {Text}")>]
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

    member private x.Text = 
        match x with 
        | Office v -> v.Text
        | Virtual v -> v.StarterValue.Text

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


[<DebuggerDisplay("{Address} {Text}")>]
[<StructuredFormatDisplay("{Address} {Text}")>]

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
        | Virtual _ -> ""

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
        | Virtual _ -> None


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
        | Virtual v -> v.ExcelCellAddress

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
            #if TestVirtual
            let datas = v.ReadDatas()
            datas
            |> VirtualExcelRange.OfData
            |> fun m -> m.AsCellRanges()
            |> List.map SingletonExcelRangeBaseUnion.Virtual
            #else
            ExcelRangeBase.asRangeList v
            |> List.map SingletonExcelRangeBaseUnion.Office
            #endif
        | ExcelRangeUnion.Virtual v ->
            v.AsCellRanges()
            |> List.filter(fun m -> 
                match m.Value with 
                | ConvertibleUnion.Missing _ -> false
                | _ -> true
            )
            |> List.map SingletonExcelRangeBaseUnion.Virtual


    let asRangeList_All (range: ExcelRangeUnion) =
        match range with 
        | ExcelRangeUnion.Office v -> 
            #if TestVirtual
            let datas = v.ReadDatas()
            datas
            |> VirtualExcelRange.OfData
            |> fun m -> m.AsCellRanges_All()
            |> List.map SingletonExcelRangeBaseUnion.Virtual
            #else
            ExcelRangeBase.asRangeList_All v
            |> List.map SingletonExcelRangeBaseUnion.Office
            #endif

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