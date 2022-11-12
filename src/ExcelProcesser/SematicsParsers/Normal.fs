namespace ExcelProcesser.SematicsParsers
#nowarn "0104"
open ExcelProcesser
open ExcelProcesser.MatrixParsers
open Shrimp.FSharp.Plus
open CellScript.Core
open OfficeOpenXml
open System.Diagnostics
open ExcelProcesser.Extensions


[<AutoOpen>]
module _Normal =
    [<AutoOpen>]
    module _Headers =
        [<RequireQualifiedAccess>]
        type HeaderTree = 
            | Node of current: SingletonExcelRangeBaseUnion * HeaderTree
            | Leaf of current: SingletonExcelRangeBaseUnion
            | LeftTailMerged of current: SingletonExcelRangeBaseUnion * merged: HeaderTree
            | BottomEmpty of cureent: SingletonExcelRangeBaseUnion * tops: HeaderTree
            | RightTailMerged of current: SingletonExcelRangeBaseUnion * merged: HeaderTree
        with 
            member x.Current =
                match x with 
                | HeaderTree.Node (v, _) 
                | HeaderTree.Leaf v
                | HeaderTree.BottomEmpty (v, _) 
                | HeaderTree.LeftTailMerged (v, _) 
                | HeaderTree.RightTailMerged (v, _) -> v

            member x.FirstValidRange =
                let rec loop (tree: HeaderTree) = 
                    match tree with 
                    | Leaf range -> range
                    | Node (range, _) -> range
                    | BottomEmpty (_, tree) 
                    | LeftTailMerged (_ , tree) 
                    | RightTailMerged (_ , tree) -> loop tree

                loop x


        type private MergedTailTag =
            | Left = 0
            | Right = 1
            | None = 2

        [<DebuggerDisplay("NormalColumnHeader ({Tree}) ({Tree.FirstValidRange}) ")>]
        type NormalColumnHeader internal (originRange: SingletonExcelRangeBaseUnion, rowsCount: int) =
            let tree =
                let rec loop i (range: SingletonExcelRangeBaseUnion) = 
                    match i with 
                    | 1 -> HeaderTree.Leaf range
                    | i when i < 1 -> failwith "Invalid token"
                    | _ ->
                        let mergedTailTag, seedingRange, rowsCost = 
                            match range.TryGetMergedRangeAddress() with 
                            | Some mergedAdress ->
                                let currentRangeAddr = range.ExcelCellAddress
                    
                                match mergedAdress.Start = currentRangeAddr with 
                                | true -> 
                                    MergedTailTag.None, range, (mergedAdress.Rows)
                                | false -> 
                                    let seedingRange = 
                                        range.WorksheetOrFail.Cells.[mergedAdress.Start.Address]
                                        |> SingletonExcelRangeBaseUnion.Create
                        
                                    let mergedTailTag =
                                        match mergedAdress.Start.Column = currentRangeAddr.Column with 
                                        | true -> MergedTailTag.Left
                                        | false -> MergedTailTag.Right


                                    mergedTailTag, seedingRange, (mergedAdress.Rows)

                            | None -> MergedTailTag.None, range, 1

                        let nextTree = 
                            match i - rowsCost with 
                            | 0 -> HeaderTree.Leaf (seedingRange)
                            | i when i < 0 -> failwith "Invalid token"
                            | _ ->
                                match seedingRange.Text.Trim() = "" with 
                                | true -> 
                                    match seedingRange.Row = originRange.Row with 
                                    | true -> HeaderTree.BottomEmpty(seedingRange, loop (i - rowsCost) (seedingRange.Offset(-1, 0)))
                                    | false ->
                                        loop (i - rowsCost) (seedingRange.Offset(-1, 0))
                                | false -> HeaderTree.Node(seedingRange, loop (i - rowsCost) (seedingRange.Offset(-1, 0)))


                        match mergedTailTag with 
                        | MergedTailTag.Left -> HeaderTree.LeftTailMerged(range, nextTree)
                        | MergedTailTag.Right -> HeaderTree.RightTailMerged(range, nextTree)
                        | MergedTailTag.None -> nextTree

                loop rowsCount originRange

            member x.Tree = tree

            member x.WorksheetOrFail = x.Tree.Current.WorksheetOrFail

            member x.LastRow = x.Tree.Current.Row

        [<RequireQualifiedAccess>]
        module NormalColumnHeader =

            let (|NonEmpty|Empty|) (normalColumnHeader: NormalColumnHeader) =
                match normalColumnHeader.Tree.FirstValidRange.Text.Trim() with 
                | "" -> Empty(normalColumnHeader)
                | _ -> NonEmpty(normalColumnHeader)
    

            let emptyParser rowsCount = 
                mxCellParserOp(fun range ->
                    let header = NormalColumnHeader(range, rowsCount)
                    match header with 
                    | Empty v -> Some v
                    | NonEmpty _ -> None
                )

            let nonEmptyParser rowsCount = 
                mxCellParserOp(fun range ->
                    let header = NormalColumnHeader(range, rowsCount)
                    match header with 
                    | Empty _ -> None
                    | NonEmpty v-> Some v
                )

        type NormalColumnHeaders internal (normalColumnHeaders: NormalColumnHeader al1List) =

            member x.Value = normalColumnHeaders

            member x.Length =
                normalColumnHeaders.Length

            member x.Columns = x.Length

            member headers1.Add(headers2: NormalColumnHeaders) =
                headers1.Value.Add(headers2.Value)
                |> NormalColumnHeaders

            member x.LastColumn = x.Value.Last.Tree.Current.Column

            member val LastRow =
                normalColumnHeaders.AsList
                |> List.map (fun m -> m.LastRow)
                |> List.distinct
                |> List.exactlyOne

        type NormalColumnHeadersParser private (start: MatrixParser<unit>, rowsCount: int, ?maxEmptySkipCount) =

            member x.RowsCount = rowsCount

            member x.Start = start

            member x.MaxEmptySkipCount = maxEmptySkipCount

            member x.Value =
                start
                |> MatrixParser.pickOutputStream(fun outputStream ->
                    let lastRow = (OutputMatrixStream.reRangeRowTo rowsCount outputStream).LastRow().Range
                    let normalColumnHeaderLists = 
                        let maxSkipCount = defaultArg maxEmptySkipCount 5

                        let parser = 
                            mxColMany1SkipRetain_BeginBy 
                                (NormalColumnHeader.emptyParser rowsCount)
                                maxSkipCount
                                (NormalColumnHeader.nonEmptyParser rowsCount)
                            ||>> fun results -> 
                                let index =
                                    results 
                                    |> List.tryFindIndexBack(Result.isOk)

                                let results_end_trimmed =
                                    match index with 
                                    | Some index ->
                                        results.[0..index]
                                    | None -> results

                                results_end_trimmed
                                |> List.map (function
                                    | Result.Ok v -> v
                                    | Result.Error v -> v
                                )
                   
                        runMatrixParserForRangeWithStreamsAsResult2_All_Union outputStream.Logger lastRow parser
                        |> OutputMatrixStream.removeRedundants
                        |> List.map (fun m -> m.Result.Value)
            

                    match normalColumnHeaderLists with 
                    | [] -> None
                    | [normalColumnHeaders] -> 
                        let normalColumnHeaders_typed = 
                            AtLeastOneList.Create normalColumnHeaders
                            |> NormalColumnHeaders

                        let endRange = outputStream.Range.Offset(rowsCount - 1, normalColumnHeaders_typed.Columns - 1)

                        let newOutputStream = 
                            outputStream
                            |> OutputMatrixStream.reRangeToAsOutputStream endRange
                            |> OutputMatrixStream.mapResultValue (fun _ -> normalColumnHeaders_typed)
            
                        newOutputStream
                        |> Some

                    | _ -> failwith "Invalid token"
                )

            static member Create(start: MatrixParser<'start>, rowsCount, ?maxEmptySkipCount) =
                NormalColumnHeadersParser(start ||>> ignore, rowsCount, ?maxEmptySkipCount = maxEmptySkipCount)
    [<AutoOpen>]
    module _Content =
        
        type NormalColumnContents private (ranges: SingletonExcelRangeBaseUnion al1List, contents: obj al1List) =
            let isEmpty (v: obj) =
                match v with 
                | null -> true
                | _ -> v.ToString().Trim() = ""

            member x.Contents = contents

            member x.Ranges = ranges

            member x.FillEmptyUp() = 
                let newContents =
                    contents.AsList
                    |> List.mapi (fun i v ->
                        match isEmpty v with 
                        | true -> 
                            contents.[0..i].AsList
                            |> List.tryFindBack(isEmpty >> not)
                            |> function
                                | Some notEmpty -> notEmpty
                                | None -> null
                        | false -> v
                    )

                NormalColumnContents(ranges, AtLeastOneList.Create newContents)

            member x.FillEmptyDown() =
                
                let newContents =
                    contents.AsList
                    |> List.mapi (fun i v ->
                        match isEmpty v with 
                        | true -> 
                            contents.[i+1..].AsList
                            |> List.tryFind(isEmpty >> not)
                            |> function
                                | Some notEmpty -> notEmpty
                                | None -> null
                        | false -> v
                    )

                NormalColumnContents(ranges, AtLeastOneList.Create newContents)



            new (ranges: SingletonExcelRangeBaseUnion al1List, ?fillEmpty: bool) =
                let fillEmpty = defaultArg fillEmpty true 

                let contents_mergedExpanded =
                    ranges.AsList 
                    |> List.map (fun range ->
                        match SingletonExcelRangeBaseUnion.tryGetMergedRangeAddress range with 
                        | Some addr -> 
                            let range = range.WorksheetOrFail.Cells.[addr.Start.Address]
                            range.Value
                        | None -> range.Value
                    )

                let contents_emptyFilled =
                    match fillEmpty with 
                    | true -> 
                        let notEmptys =
                            contents_mergedExpanded
                            |> List.choose (fun v ->
                                match v with 
                                | null -> None
                                | v -> 
                                    match v.ToString() = "" with 
                                    | true -> None
                                    | false -> Some v
                            )

                        match notEmptys with 
                        | [ notEmpty ] -> 
                            List.replicate contents_mergedExpanded.Length notEmpty
                        | [] ->  contents_mergedExpanded
                        | notEmptys when notEmptys.Length = contents_mergedExpanded.Length -> notEmptys
                        | notEmptys -> contents_mergedExpanded

                    | false -> contents_mergedExpanded

                NormalColumnContents(ranges = ranges, contents = AtLeastOneList.Create contents_emptyFilled)


    [<AutoOpen>]
    module _Column =    

        type NormalColumn(header: NormalColumnHeader, rowIndexes: int al1List) =
            let worksheet = header.WorksheetOrFail
            let contents = 
                let ranges = 
                    let column = header.Tree.Current.Column
                    let addrs =
                        rowIndexes.AsList
                        |> List.map (fun row ->
                            {
                                Column = column
                                Row = row
                            }
                        )

                    addrs
                    |> List.map (fun addr ->
                        worksheet.Cells.[addr.Address]
                        |> SingletonExcelRangeBaseUnion.Create
                    )
                    |> AtLeastOneList.Create

                NormalColumnContents(ranges)

            member x.Contents = contents

            member x.Header = header

            member x.RowIndexes = rowIndexes

            new (header: NormalColumnHeader, startRowOffset: int, rowsCount: int) =
                let range = header.Tree.Current.Offset(startRowOffset, 0, rowsCount, 1)
                let rowIndexes  = 
                    ExcelRangeUnion.asRangeList_All range
                    |> List.map (fun m -> m.Row)
                    |> AtLeastOneList.Create

                NormalColumn(header = header, rowIndexes = rowIndexes)

        type NormalColumns(normalColumns: NormalColumn al1List) =
            member val RowIndexes =   
                normalColumns.AsList
                |> List.map (fun m -> m.RowIndexes)
                |> List.distinct
                |> List.exactlyOne

            member x.Value = normalColumns


    [<AutoOpen>]
    module _DSL =

        type NormalColumnHeader with 
            member x.SelectColumn(startRowOffset, rowsCount) =
                NormalColumn(x, startRowOffset, rowsCount)

            member x.SelectColumn(rowIndexes) =
                NormalColumn(x, rowIndexes)

        type NormalColumnHeaders with 
            member x.SelectColumns(startRowOffset, rowsCount) =
                x.Value
                |> AtLeastOneList.map (fun m -> m.SelectColumn(startRowOffset, rowsCount))
                |> NormalColumns

            member x.SelectColumns(rowIndexes: int al1List) =
                x.Value
                |> AtLeastOneList.map (fun m -> m.SelectColumn(rowIndexes))
                |> NormalColumns