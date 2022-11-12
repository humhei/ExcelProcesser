namespace ExcelProcesser.SematicsParsers
#nowarn "0104"
open ExcelProcesser
open ExcelProcesser.MatrixParsers
open ExcelProcesser.Extensions
open Shrimp.FSharp.Plus

[<AutoOpen>]
module  _RangeInHeader =

    [<RequireQualifiedAccess>]
    type GroupingColumnHeaderRowNameParser =
        | Left of MatrixParser<string>
        | TopOrNone

    module RangeInHeader =


        module Calc =
            [<AutoOpen>]
            module Header =

                [<AutoOpen>]
                module GroupingColumnHeader =
                    type GroupingHeaderNamePosition =
                        | TopOrNone = 0
                        | Left = 1
    
                    [<RequireQualifiedAccess>]
                    type GroupingColumnHeaderRowName =
                        | Left of SingletonExcelRangeBaseUnion * string
                        | Top of string
                    with 
                        member x.Value =
                            match x with 
                            | GroupingColumnHeaderRowName.Left (_, v) -> v
                            | GroupingColumnHeaderRowName.Top v -> v
    
    
                    type GroupingColumnHeaderRow<'groupingHeader> =
                        { Name: GroupingColumnHeaderRowName option
                          Values: al1List<SingletonExcelRangeBaseUnion * 'groupingHeader> }
    
    
                    type GroupingColumnHeaderRows<'groupingHeader>
                        (rows: GroupingColumnHeaderRow<'groupingHeader> al1List) =

                        member val ValuesLength = 
                            rows.AsList
                            |> List.map (fun m -> m.Values.Length)
                            |> List.distinct
                            |> List.exactlyOne

                        member val ValuesStartColumn = 
                            rows.AsList
                            |> List.map (fun row -> 
                                let range = fst row.Values.Head
                                range.Column
                            )
                            |> List.distinct
                            |> List.exactlyOne

                          

                        member val ValuesEndColumn =
                            rows.AsList
                            |> List.map (fun row -> 
                                let range = fst row.Values.Last
                                range.Column
                            )
                            |> List.distinct
                            |> List.exactlyOne


                        member internal x.Value = rows
        
    
    
                    type MultipleColumnsGroupingColumnHeaderRows<'groupingHeader> = 
                        private MultipleColumnsGroupingColumnHeaderRows of GroupingColumnHeaderRows<'groupingHeader>
                    with 
                        member x.Value =
                            let (MultipleColumnsGroupingColumnHeaderRows v) = x
                            v
    
    
                    type OneColumnGroupingColumnHeaderRows<'groupingHeader> = 
                        private OneColumnGroupingColumnHeaderRows of GroupingColumnHeaderRows<'groupingHeader>
                    with 
                        member x.Value =
                            let (OneColumnGroupingColumnHeaderRows v) = x
                            v
    
    
                    [<RequireQualifiedAccess>]
                    type GroupingColumnHeaderRowsChoice<'groupingHeader> =
                        | MultipleColumn of MultipleColumnsGroupingColumnHeaderRows<'groupingHeader>
                        | OneColumn of OneColumnGroupingColumnHeaderRows<'groupingHeader>
                    with 
                        member x.Value =
                            match x with 
                            | GroupingColumnHeaderRowsChoice.MultipleColumn v -> v.Value
                            | GroupingColumnHeaderRowsChoice.OneColumn v -> v.Value
    
                    

                    type GroupingColumnHeaderRows =
    
                        static member private Parser_Common(groupingHeaders: MatrixParser<list<SingletonExcelRangeBaseUnion * 'groupingHeader>>, headerName: GroupingColumnHeaderRowNameParser, rowsCount: int) =
                            let groupingHeaders =
                                groupingHeaders
                                |> atLeastOne
                
            
                            MatrixParser(fun inputStream ->
                                let inputStream =
                                    let addr =
                                        let addr = inputStream.ParsingAddress.Value
                                        { addr with EndRow = min(addr.StartRow + rowsCount) addr.EndRow }
                                        |> ParsingAddress

                                    { inputStream with ParsingAddress = addr }
    
                                let newParser =
                                    match headerName with 
                                    | GroupingColumnHeaderRowNameParser.TopOrNone ->
                                        groupingHeaders
                                        |> MatrixParser.mapOutputStreams OutputMatrixStream.removeRedundants
                                        |> MatrixParser.mapOutputStream (fun outputStream ->
                                            let newOutputStream = 
                                                outputStream
                                                |> OutputMatrixStream.mapResultValue(fun groupingHeaders ->
                                                    let nameCellOfArea area =
                                                        let exactlyOneAreaText =
                                                            ExcelRangeUnion.asRangeList area
                                                            |> List.filter(fun m -> m.Text.Trim() <> "")
                                                            |> List.tryExactlyOne
    
                                                        exactlyOneAreaText
    
                                                    let topNameCell = 
                                                        OutputMatrixStream.topArea outputStream
                                                        |> Option.map nameCellOfArea
                                                        |> Option.flatten
    
                                                    { Name = 
                                                        topNameCell
                                                        |> Option.map (fun m -> GroupingColumnHeaderRowName.Top m.Text )
                                                      Values = 
                                                        groupingHeaders
                                                        |> AtLeastOneList.Create  }
                                                    |> List.singleton
                                                )
    
    
                                            OutputMatrixStream.redirectTo 
                                                (rowsCount-1) 
                                                (newOutputStream.OffsetedRange.Column - newOutputStream.Range.Column) 
                                                inputStream 
                                                newOutputStream
                                        )
    
    
                                    | GroupingColumnHeaderRowNameParser.Left leftName ->
                                        let leftName =
                                            leftName
                                            |> MatrixParser.mapOutputStream(fun outputStream ->
                                                outputStream
                                                |> OutputMatrixStream.mapResultValue(fun v ->
                                                    outputStream.OffsetedRange, v
                                                )
                                            )
    
                                        (
                                            mxRowMany1 (c2 leftName groupingHeaders)
                                            |> MatrixParser.mapOutputStreams OutputMatrixStream.removeRedundants
                                            |> MatrixParser.mapOutputStream (fun outputStream ->
                                                outputStream
                                                |> OutputMatrixStream.mapResultValue 
                                                    (List.map (fun (leftName, groupingHeaders) ->
                                                        { Name = Some (GroupingColumnHeaderRowName.Left leftName)
                                                          Values = 
                                                            groupingHeaders
                                                            |> AtLeastOneList.Create
                                                        }
                                                    ))
                                            )
                                        )
                                    |> fun a -> a ||>> AtLeastOneList.Create
            
                                newParser.Invoke(inputStream)
    
                            )
                            ||>> GroupingColumnHeaderRows
    
                        static member MultipleColumns(groupingHeader: SingletonMatrixParser<'groupingHeader>, headerName: GroupingColumnHeaderRowNameParser, rowsCount: int) =
                            let groupingHeaders =
                                groupingHeader
                                |> MatrixParser.mapResultValueWithOutputStream(fun outputStream  ->
                                    (outputStream.OffsetedRange, outputStream.Result.Value)
                                )
                                |> mxColMany1 
    
                            GroupingColumnHeaderRows.Parser_Common(groupingHeaders, headerName, rowsCount = rowsCount)
                            ||>> MultipleColumnsGroupingColumnHeaderRows
                            ||>> GroupingColumnHeaderRowsChoice.MultipleColumn
    
                        static member OneColumn(groupingHeaders: MatrixParser<'groupingHeader list>, headerName: GroupingColumnHeaderRowNameParser, rowsCount: int) =
                            let groupingHeaders =
                                groupingHeaders
                                |> MatrixParser.mapResultValueWithOutputStream(fun outputStream  ->
                                    outputStream.Result.Value
                                    |> List.map (fun groupingHeader ->
                                        outputStream.OffsetedRange, groupingHeader
                                    )
                                )
    
                            GroupingColumnHeaderRows.Parser_Common(groupingHeaders, headerName, rowsCount = rowsCount)
                            ||>> OneColumnGroupingColumnHeaderRows
                            ||>> GroupingColumnHeaderRowsChoice.OneColumn
    
                [<AutoOpen>]
                module PivotTableHeader = 
                
                    type PivotTableHeaders<'groupingHeader> =
                        { NormalColumnHeaders: NormalColumnHeaders 
                          GroupingColumnHeaderRows: GroupingColumnHeaderRowsChoice<'groupingHeader> }
                
                    type PivotTableHeadersParser<'groupingHeader> internal (start: MatrixParser<unit>, groupingHeaderRows: MatrixParser<GroupingColumnHeaderRowsChoice<'groupingHeader>>, rowsCount, ?maxEmptySkipCount) =
                        let parser = 
                            start
                            |> MatrixParser.collectOutputStream(fun outputStream ->
                                let rerangedResult = OutputMatrixStream.reRangeRowTo rowsCount outputStream
                
                                runMatrixParserForRangeWithStreamsAsResultUnion rerangedResult.Range (groupingHeaderRows)
                                |> OutputMatrixStream.removeRedundants
                                |> List.tryExactlyOne
                                |> function
                                    | Some groupingHeadersOutputStream ->
                                        let leftNormalColumns =
                                            let leftNormalColumnsReranged = 
                                                let columnsCount = 
                                                    groupingHeadersOutputStream.Range.Column -
                                                        outputStream.Range.Column 
                                                
                                                rerangedResult.SetColumnCountTo(columnsCount) 
                
                                            let normalColumnTreeHeaderParser =
                                                NormalColumnHeadersParser.Create(mxNonEmpty, rowsCount, ?maxEmptySkipCount = maxEmptySkipCount).Value
                
                                            normalColumnTreeHeaderParser.InvokeToResults(outputStream, leftNormalColumnsReranged)
                                            |> List.exactlyOne
                
                                        let rightNormalColumns =
                                            let rightNormalColumnsReranged = 
                                                rerangedResult.RightOf(groupingHeadersOutputStream.OffsetedRange)
                
                                            let normalColumnTreeHeaderParser =
                                        
                                                NormalColumnHeadersParser.Create(mxAnyOrigin, rowsCount, ?maxEmptySkipCount = maxEmptySkipCount).Value
                
                                            normalColumnTreeHeaderParser.InvokeToResults(outputStream, rightNormalColumnsReranged)
                                            |> function
                                                | [ v ] -> Some v
                                                | [] -> None
                                                | _ -> failwith "Invalid token"
                
                                        let normalColumns =
                                            match rightNormalColumns with 
                                            | Some rightNormalColumns -> leftNormalColumns.Add(rightNormalColumns)
                                            | None -> leftNormalColumns
                
                                        let columnsCount = 
                                            let groupingColumnsCount =
                                                groupingHeadersOutputStream.OffsetedRange.Column 
                                                - groupingHeadersOutputStream.Range.Column 
                                                + 1
                                    
                                            normalColumns.Columns + groupingColumnsCount
                
                                        let newOutputStream =
                                            outputStream.ShiftBy(columnsCount-1, rowsCount-1)
                                            |> OutputMatrixStream.mapResultValue (fun _ ->
                                                { NormalColumnHeaders = normalColumns
                                                  GroupingColumnHeaderRows = groupingHeadersOutputStream.Result.Value }
                                            )
                
                                        [newOutputStream]
                                    | None -> []
                            )
        
                        member x.Value = parser

            [<AutoOpen>]
            module Table =
                
                type GroupingColumnElementLists<'groupingElement>(elementLists: (SingletonExcelRangeBaseUnion * 'groupingElement) option al1List al1List) =
                    member x.ElementLists = elementLists

                    member x.AsList = x.ElementLists.AsList


                    member x.RowIndexes = 
                        elementLists
                        |> AtLeastOneList.map (fun elements ->
                            let (range, _) =
                                elements.AsList
                                |> List.pick id
                            range.Row
                        )

                type GroupingColumn<'groupingHeader, 'groupingElement>
                    (headers: GroupingColumnHeaderRowsChoice<'groupingHeader>,
                     elementLists: GroupingColumnElementLists<'groupingElement>,
                     startRowIndex: int,
                     endRowIndex: int) =
                    
                    member val ValuesLength =
                        let l1 =
                            elementLists.AsList
                            |> List.map (fun m -> m.Length)
                            |> List.distinct
                            |> List.exactlyOne

                        let l2 =
                            headers.Value.ValuesLength

                        [ l1; l2 ]
                        |> List.distinct
                        |> List.exactlyOne
                    
                    member x.RowsCount = endRowIndex - startRowIndex + 1

                    member x.ElementLists = elementLists

                    member x.Headers = headers
                

                type PivotTable<'groupingHeader, 'groupingElement> =
                    { NormalColumns: NormalColumns
                      GroupingColumn: GroupingColumn<'groupingHeader, 'groupingElement> }

                type PivotTableParser<'groupingHeader, 'groupingElement> internal (headers: PivotTableHeadersParser<'groupingHeader>, elements: MatrixParser<list<(SingletonExcelRangeBaseUnion * 'groupingElement) option>>) =
                    let elementLists =
                        headers.Value
                        |> MatrixParser.mapOutputStream(fun outputStream ->
                            let headers = outputStream.Result.Value

                            let groupingColumn = 
                                let groupingColumnHeaders = headers.GroupingColumnHeaderRows

                                let stream = 
                                    let rangeTransformer =
                                        RangeTransformer(outputStream.RangeToOffsetedRange)
                                            .SetEndRow(outputStream.ParsingAddress.EndRow)
                                            .SetEndColumn(groupingColumnHeaders.Value.ValuesEndColumn)
                                            .SetStart(
                                                { Row = headers.NormalColumnHeaders.LastRow + 1
                                                  Column = groupingColumnHeaders.Value.ValuesStartColumn }
                                            )

                                    (mxRowMany1 elements).InvokeToStreams(outputStream, rangeTransformer)
                                    |> List.exactlyOne


                                let elementLists =
                                    let maxLength = groupingColumnHeaders.Value.ValuesLength
                                    stream.Result.Value
                                    |> List.map (fun m ->
                                        m @ List.replicate (maxLength - m.Length) None
                                    )

                                GroupingColumn(
                                    headers = groupingColumnHeaders,
                                    elementLists = GroupingColumnElementLists (AtLeastOneList.ofLists elementLists),
                                    startRowIndex = headers.NormalColumnHeaders.LastRow + 1,
                                    endRowIndex = stream.OffsetedRange.Row
                                )

                            let normalColumns =
                                headers.NormalColumnHeaders.SelectColumns(groupingColumn.ElementLists.RowIndexes)

                            let newStream = 
                                outputStream.ShiftBy(0, groupingColumn.RowsCount)
                                |> OutputMatrixStream.mapResultValue (fun _ -> 
                                    {
                                        NormalColumns = normalColumns
                                        GroupingColumn = groupingColumn
                                    }
                                )

                            newStream
                        )

                    member x.Parser = elementLists

        module DSL =
            open Calc
            [<AutoOpen>]
            module Header =
                type MultipleColumnsPivotTableHeaders<'groupingHeader> = private MultipleColumnsPivotTableHeaders of PivotTableHeadersParser<'groupingHeader>
                with 
                    member x.Value =
                        let (MultipleColumnsPivotTableHeaders v) = x
                        v

                    member x.Parser = x.Value.Value

                    member x.SelectColumn(parser: SingletonMatrixParser<'groupingElement>) =
                        let elementsParser = mxColMany1Op (System.Int32.MaxValue) (parser.PrefixRange())
                        PivotTableParser(headers= x.Value, elements= elementsParser)

                type PivotTableHeaders(normalColumnHeadersParser: NormalColumnHeadersParser, headerName) =
                    member internal x.NormalColumnHeadersParser = normalColumnHeadersParser

                    member x.OneColumn(groupingHeaders) = 
                        PivotTableHeadersParser(
                            start = normalColumnHeadersParser.Start,
                            groupingHeaderRows = 
                                GroupingColumnHeaderRows.OneColumn(
                                    groupingHeaders = groupingHeaders,
                                    headerName = headerName,
                                    rowsCount = normalColumnHeadersParser.RowsCount),
                            rowsCount = normalColumnHeadersParser.RowsCount,
                            ?maxEmptySkipCount = normalColumnHeadersParser.MaxEmptySkipCount
                        )
                        |> MultipleColumnsPivotTableHeaders

                    member x.MultipleColumns(groupingHeader) = 
                        PivotTableHeadersParser(
                            start = normalColumnHeadersParser.Start,
                            groupingHeaderRows = 
                                GroupingColumnHeaderRows.MultipleColumns(
                                    groupingHeader = groupingHeader,
                                    headerName = headerName,
                                    rowsCount = normalColumnHeadersParser.RowsCount),
                            rowsCount = normalColumnHeadersParser.RowsCount,
                            ?maxEmptySkipCount = normalColumnHeadersParser.MaxEmptySkipCount
                        )
                        |> MultipleColumnsPivotTableHeaders
                    

    open RangeInHeader
    type NormalColumnHeadersParser with 
        member x.RangeInHeader(headerName) =
            DSL.Header.PivotTableHeaders(normalColumnHeadersParser = x, headerName = headerName)
            
