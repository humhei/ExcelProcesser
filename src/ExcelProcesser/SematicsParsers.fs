module ExcelProcesser.SematicsParsers
open ExcelProcesser.MatrixParsers
open FParsec.CharParsers
open ExcelProcesser.CellParsers
open System
open ExcelProcesser.MathParsers
open Extensions
open Deedle
open Shrimp.FSharp.Plus

[<AutoOpen>]
module internal InternalExtensions =
    [<RequireQualifiedAccess>]
    module Array2D =

        let private rebasingMap mapping (array2D: 'a [,]) =
            let lower = Array2D.base1 array2D
            let upper = Array2D.base2 array2D
            if lower = 0 && upper = 0 
            then mapping array2D
            else mapping (Array2D.rebase array2D)


        let pickHeaderTailRowsNotInclude headerIndex tailIndex (array2D: 'a [,]) =
            rebasingMap (fun array2D ->
                let headers = array2D.[0 .. headerIndex - 1, *]
                let tails = array2D.[tailIndex + 1 .. (Array2D.length1 array2D) - 1, *]
                Array2D.joinByRows headers tails
            ) array2D

        let removeSencondRow (array2D: 'a [,]) =
            pickHeaderTailRowsNotInclude 1 1 array2D

        let removeLastRow (array2D: 'a [,]) =
            array2D.[0.. (Array2D.length1 array2D - 2), *]

        let pickHeaderTailColumnsNotInclude headerIndex tailIndex (array2D: 'a [,]) =
            rebasingMap (fun array2D ->
                let headers = array2D.[*, 0 .. headerIndex - 1]
                let tails = array2D.[*, tailIndex + 1 .. (Array2D.length2 array2D) - 1]
                Array2D.joinByCols headers tails
            ) array2D

        let pickHeaderTailColumnsNotIncludeByIndexer (coordinate: Coordinate, shift) (array2D: 'a [,]) =
            let headerIndex = coordinate.X
            let tailIndex = headerIndex + shift
            pickHeaderTailColumnsNotInclude headerIndex tailIndex array2D

        let transpose (input: 'a[,]) =
            let l1 = input.GetLowerBound(1)
            let u1 = input.GetUpperBound(1)
            seq {
                for i = l1 to u1 do
                    yield input.[*,i]
            }
            |> array2D

    [<RequireQualifiedAccess>]
    module internal Frame =
        let ofArray2DWithHeader (array: obj[,]) =

            let array = Array2D.rebase array

            let headers = array.[0,*] |> Seq.map (fun header -> header.ToString())

            let contents = array.[1..,*] |> Array2D.map box |> Frame.ofArray2D

            Frame.indexColsWith headers contents

        let splitRowToMany addtionalHeaders (mapping : 'R -> ObjectSeries<_> -> seq<seq<obj>>)  (frame: Frame<'R,'C>) =
            let headers = 
                let keys = frame.ColumnKeys
                Seq.append keys addtionalHeaders

            let values = 
                frame.Rows.Values
                |> Seq.mapi (fun i value -> mapping (Seq.item i frame.RowKeys) value)
                |> Seq.concat
                |> array2D

            Frame.ofArray2D values
            |> Frame.indexColsWith headers

        let toArray2DWithHeader (frame: Frame<_,_>) =

            let header = array2D [Seq.map box frame.ColumnKeys]

            let contents = frame.ToArray2D(null) 
            Array2D.joinByRows header contents


module TwoHeadersPivotTable =

    type GroupingColumnHeader<'childHeader> =
        { GroupedHeader: string 
          ChildHeaders: AtLeastOneList<'childHeader>
          Shift: Shift }
    with 
        member x.Indexer = 
            match x.Shift.Last with 
            | Horizontal (coordinate, i) -> coordinate, i
            | _ -> failwith "Last shift should be horizontal"

    let mxGroupingColumnsHeader (defaultGroupedHeaderText: string option) pChild =
        r2
            (mxMerge Direction.Horizontal <||> (mxMany1 Direction.Horizontal mxSpace ||>> fun _ -> ""))
            (mxMany1 Direction.Horizontal pChild)

        |||>> fun outputStream ((groupedHeader), childHeaders) -> 
            { GroupedHeader = 
                match defaultGroupedHeaderText with 
                | None -> 
                    match groupedHeader.Trim() with 
                    | "" -> "GroupedHeader"
                    | groupedHeader -> groupedHeader

                | Some text -> text
              ChildHeaders = AtLeastOneList.Create childHeaders
              Shift = outputStream.Shift }



    type GroupingColumn<'childHeader, 'element> =
        { Header: GroupingColumnHeader<'childHeader>
          ElementsList: ('element option) al1List al1List }


    type GroupingColumnParserArg<'childHeader,'elementSkip, 'element> =
        GroupingColumnParserArg_ of 
            defaultGroupedHeaderText: string option
            * pChildHeader: MatrixParser<'childHeader> 
            * pElementSkip: MatrixParser<'elementSkip> option
            * pElement: MatrixParser<'element>

    type GroupingColumnParserArg =
        static member Create(pChildHeader, pElementSkip, pElement, ?defaultGroupedHeaderText) =
            GroupingColumnParserArg_(defaultGroupedHeaderText, pChildHeader, pElementSkip, pElement)

    let mxGroupingColumn (GroupingColumnParserArg_(defaultGroupedHeaderText, pChildHeader, pElementSkip, pElement)) =
        let pChildHeader = 
            MatrixParser.addLogger LoggerLevel.Info "pChildHeader" pChildHeader
    
        let pElement = 
            MatrixParser.addLogger LoggerLevel.Info "pElement" pElement
    

        (r2R 
            (mxGroupingColumnsHeader defaultGroupedHeaderText pChildHeader)
            (fun outputStream ->

                let maxColNum, columns = 
                    let reranged, addr = OutputMatrixStream.reRange outputStream
                    reranged.End.Column, reranged.Columns

                let pElementInRange =
                    pElement <.&> (mxCellParser (fun range -> range.Value.Start.Column <= maxColNum) ignore)

                match pElementSkip with 
                | Some pElementSkip ->
                    mxRowMany1 
                        ((mxManySkipRetain Direction.Horizontal pElementSkip columns pElementInRange) 
                        ||>> (List.mapi (fun i result ->
                            match result with 
                            | Result.Ok ok -> Some (ok)
                            | Result.Error _ -> None
                        )))
                | None -> mxRowMany1 ((mxColMany1 pElementInRange) ||>> (List.map Some))
            )
        ) 
        ||>> fun (header, elementsList) ->
            { Header = header
              ElementsList = AtLeastOneList.ofLists elementsList }


    type TwoHeadersPivotTableBorder<'leftBorderHeader,'numberHeader,'rightBorderHeader> =
        { LeftBorderHeader: 'leftBorderHeader
          NumberHeader: 'numberHeader
          SumElements: int list
          SumResult: int
          RightBorderHeader: 'rightBorderHeader }

    let private mxTwoHeadersPivotTableBorder pLeftBorderHeader pNumberHeader pRightBorderHeader =
        c3 
            (pLeftBorderHeader <.&> mxMergeStarter)
            (mxUntilA50
                (r2 
                    (pNumberHeader <&> mxMergeStarter)
                    (mxUntilA50
                        (mxSumContinuously Direction.Vertical) 
                    )
                ) 
            )
            (mxUntilA50 (pRightBorderHeader <&> mxMergeStarter))
        ||>> (fun (leftBorderHeader,(numberHeader,(sumElements, sumResult)),rightBorderHeader) ->
            { LeftBorderHeader = leftBorderHeader
              NumberHeader = numberHeader 
              SumElements = sumElements
              SumResult = sumResult
              RightBorderHeader = rightBorderHeader }
        )


    type NormalColumn =
        { Header: string 
          Contents: obj list }

    //[<RequireQualifiedAccess>]
    //module NormalColumn =
    //    let internal fixEmptyUp column =
    //        { column with 
    //            Contents = 
    //                column.Contents 
    //                |> List.mapi (fun i content ->
    //                    let isNullOrEmpty = 
    //                        match content with 
    //                        | null -> true
    //                        | :? string as text -> text.Trim() = ""
    //                        | _ -> false

    //                    if isNullOrEmpty then 
    //                        column.Contents.[0 .. i - 1]
    //                        |> List.tryFindBack (fun content -> not (isNull content))
    //                        |> function 
    //                            | Some v -> v
    //                            | None -> null
    //                    else content
    //                )
    //        }



    type GroupingElement<'groupingColumnChildHeader, 'groupingColumnElement> =
        {
            IndexInHeaders: int
            Header: 'groupingColumnChildHeader
            Element: 'groupingColumnElement
        }

    type TwoHeadersPivotTableRow<'groupingColumnChildHeader, 'groupingColumnElement> =
        { GroupingElments: AtLeastOneList<GroupingElement<'groupingColumnChildHeader, 'groupingColumnElement>>
          NormalValueObservations: ObjectSeries<StringIC>
        }

    type TwoHeadersPivotTable<'groupingColumnChildHeader, 'groupingColumnElement> =
        { GroupingColumn: GroupingColumn<'groupingColumnChildHeader, 'groupingColumnElement>
          NormalColumns: NormalColumn list
          SumNumber: int }

    with 
        member this.Rows(?fixEmptyUp: bool) =
            let fixEmptyUp = defaultArg fixEmptyUp true
            let basicRows = 
                this.GroupingColumn.ElementsList.AsList
                |> List.indexed
                |> List.choose (fun (i, elements) ->
                    let groupingElements =
                        let childHeaders = this.GroupingColumn.Header.ChildHeaders
                        elements.AsList
                        |> List.mapi(fun j element ->
                            match element with
                            | Some element -> 
                                { IndexInHeaders = j
                                  Header = childHeaders.[j]
                                  Element = element }
                                |> Some
                            | None -> None
                        )
                        |> List.choose id

                    match groupingElements with 
                    | [] -> None
                    | _ -> 
                        {
                            NormalValueObservations = 
                                this.NormalColumns
                                |> List.map(fun normalColumn -> 
                                    StringIC normalColumn.Header => normalColumn.Contents.[i]
                                )
                                |> Series.ofObservations
                                |> ObjectSeries
                            GroupingElments = AtLeastOneList.Create groupingElements
                        }
                        |> Some
                )
        
            match fixEmptyUp with 
            | true ->
                let exactlyOneItemObservations =
                    match basicRows with 
                    | basicRow :: _ ->  
                        let keys =
                            List.ofSeq basicRow.NormalValueObservations.Keys

                        keys
                        |> List.map (fun key ->
                            let exactlyValue = 
                                basicRows
                                |> List.choose (fun m -> 
                                    m.NormalValueObservations.TryGet(key)  
                                    |> OptionalValue.asOption
                                )
                                |> List.tryExactlyOne

                            (key => exactlyValue)
                        )
                        |> dict
                    | _ -> failwith "Invalid token"

                (None, basicRows)
                ||> List.mapFold (fun previousRow (basicRow) ->
                    let basicRow =
                        { 
                            basicRow with 
                                NormalValueObservations =
                                    basicRow.NormalValueObservations
                                    |> Series.mapAll(fun key value ->
                                        match value with 
                                        | None -> exactlyOneItemObservations.[key]
                                        | Some value ->  Some (value)
                                    )
                                    |> ObjectSeries
                        }
                    
                    match previousRow with 
                    | None -> basicRow, Some basicRow
                    | Some previousRow -> 
                        let previousRowNormalValueObsevations = previousRow.NormalValueObservations
                        let newBasicRow = 
                            { 
                                basicRow with 
                                    NormalValueObservations =
                                        basicRow.NormalValueObservations
                                        |> Series.mapAll(fun key value ->
                                            match value with 
                                            | None -> 
                                                match previousRowNormalValueObsevations.TryGet(key) with 
                                                | OptionalValue.Missing -> None
                                                | OptionalValue.Present v -> Some v
                                            | Some value ->  Some (value)
                                        )
                                        |> ObjectSeries
                            }
                        newBasicRow, Some newBasicRow
                )|> fst

            | false -> basicRows

        

    type TwoHeadersPivotTable =
        static member ToFrame (twoHeadersPivotTable: TwoHeadersPivotTable<_, _>, ?fixEmptyUp) =
            let rows = twoHeadersPivotTable.Rows(?fixEmptyUp = fixEmptyUp)
            let groupedHeader = twoHeadersPivotTable.GroupingColumn.Header.GroupedHeader

            rows
            |> List.collect (fun row ->
                row.GroupingElments.AsList
                |> List.map (fun groupingElement ->
                    let addtionalSeries =
                        series
                            [
                                groupedHeader => box groupingElement.Header
                                groupedHeader + "_Value" => box groupingElement.Element
                            ]
                        |> Series.mapKeys StringIC
                    Series.mergeUsing UnionBehavior.Exclusive row.NormalValueObservations addtionalSeries 
                )
            )
            |> Frame.ofRowsOrdinal
            |> Frame.mapRowKeys int

        static member ToArray2D (twoHeadersPivotTable, ?fixEmptyUp) =
            TwoHeadersPivotTable.ToFrame(twoHeadersPivotTable, ?fixEmptyUp = fixEmptyUp)
            |> Frame.toArray2DWithHeader

        static member Parser(pLeftBorderHeader, pNumberHeader, (pOriginRightBorderHeader: MatrixParser<_> option), (pGroupingColumn:GroupingColumnParserArg<_, _, _>)) =
            
            let pLeftBorderHeader = 
                MatrixParser.addLogger LoggerLevel.Info "pLeftBorderHeader" pLeftBorderHeader

            let pNumberHeader = 
                MatrixParser.addLogger LoggerLevel.Info "pNumberHeader" pNumberHeader

            let pRightBorderHeader = 
                pOriginRightBorderHeader
                |> Option.map (
                    MatrixParser.addLogger LoggerLevel.Info "pRightBorderHeader" 
                )
                |> function
                    | Some pRightBorderHeader -> pRightBorderHeader ||>> ignore
                    | None ->
                        mxGroupingColumn pGroupingColumn
                        ||>> ignore


            mxTwoHeadersPivotTableBorder pLeftBorderHeader pNumberHeader pRightBorderHeader
            |> MatrixParser.collectOutputStream (fun outputStream ->
                let reranged, addr = OutputMatrixStream.reRange outputStream
                let range = 
                    reranged
                    |> Seq.head
                    |> SingletonExcelRangeBase.Create

                let resetedInputStream = 
                    { Range = range
                      Shift = Shift.Start
                      ParsingAddress = addr
                      Logger = outputStream.Logger }

                let p = 
                    let p =
                        match pOriginRightBorderHeader with 
                        | Some _ ->
                            c3 
                                (inDebug(pLeftBorderHeader <&> mxMergeStarter) ||>> ignore)
                                ((mxUntilA50
                                    ((mxGroupingColumn pGroupingColumn))))
                                ((mxUntilA50 (pRightBorderHeader <&> mxMergeStarter)) ||>> ignore)
                            ||>> (fun (_, b, _) -> b)
                        | None ->
                            c2 
                                ((pLeftBorderHeader <&> mxMergeStarter) ||>> ignore)
                                (mxUntilA50
                                    ((mxGroupingColumn pGroupingColumn)))
                            ||>> (fun (_, b) -> b)
                    p
                    ||>> (fun groupingColumn ->
                        let normalColumns =
                            let array2D = 
                                (reranged).Value :?> obj[,]

                            let exceptGroupingColumn array2D = 
                                Array2D.pickHeaderTailColumnsNotIncludeByIndexer groupingColumn.Header.Indexer array2D
            
                            let newArray2D = 
                                array2D
                                |> exceptGroupingColumn
                                |> Array2D.removeSencondRow
                                |> Array2D.removeLastRow

                            [
                                for i = 0 to Array2D.length2 newArray2D - 1 do 
                                    let column = newArray2D.[*, i]
                                    yield
                                        { Header = column.[0].ToString()
                                          Contents = List.ofArray column.[1..]}
                            ]
                        { GroupingColumn = groupingColumn 
                          NormalColumns = normalColumns
                          SumNumber = outputStream.Result.Value.SumResult }
                    )
                let r = p.Invoke resetedInputStream
                r
            )



    