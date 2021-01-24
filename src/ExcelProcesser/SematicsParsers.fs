module ExcelProcesser.SematicsParsers
open ExcelProcesser.MatrixParsers
open FParsec.CharParsers
open ExcelProcesser.CellParsers
open System
open ExcelProcesser.MathParsers
open Extensions
open Deedle

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

type GroupingColumnHeader<'childHeader> =
    { GroupedHeader: string 
      ChildHeaders: 'childHeader list
      Shift: Shift }
with 
    member x.Indexer = 
        match x.Shift.Last with 
        | Horizontal (coordinate, i) -> coordinate, i
        | _ -> failwith "Last shift should be horizontal"

let mxGroupingColumnsHeader (defaultGroupedHeaderText: string option) pChild =
    r2
        (mxMerge Direction.Horizontal)
        (mxMany1 Direction.Horizontal pChild)
    |> MatrixParser.filterOutputStreamByResultValue (fun ((groupedHeader, emptys), childs) ->
        match defaultGroupedHeaderText with 
        | None -> groupedHeader.Text.Trim() <> "" 
        | Some _ -> true
        && emptys.Length = childs.Length - 1
    )
    |||>> fun outputStream ((groupedHeader, _), childHeaders) -> 
        { GroupedHeader = 
            match defaultGroupedHeaderText with 
            | None -> groupedHeader.Text
            | Some text -> text
          ChildHeaders = childHeaders
          Shift = outputStream.Shift }



type GroupingColumn<'childHeader, 'element> =
    { Header: GroupingColumnHeader<'childHeader>
      ElementsList: ('element option) list list }


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
    (r2R 
        (mxGroupingColumnsHeader defaultGroupedHeaderText pChildHeader)
        (fun outputStream ->

            let maxColNum, columns = 
                let reranged = OutputMatrixStream.reRange outputStream
                reranged.End.Column, reranged.Columns

            let pElementInRange =
                pElement <&> (mxCellParser (fun range -> range.Value.Start.Column <= maxColNum) ignore)


            match pElementSkip with 
            | Some pElementSkip ->
                mxRowMany1 ((fun a -> mxManySkipRetain Direction.Horizontal pElementSkip columns pElementInRange a) ||>> (List.mapi (fun i result ->
                    match result with 
                    | Result.Ok ok -> Some (ok)
                    | Result.Error _ -> None
                )))
            | None -> mxRowMany1 ((mxColMany1 pElementInRange) ||>> (List.map Some))
        )
    ) 
    ||>> fun (header, elementsList) ->
        { Header = header
          ElementsList = elementsList }


type TwoHeadersPivotTableBorder<'leftBorderHeader,'numberHeader,'rightBorderHeader> =
    { LeftBorderHeader: 'leftBorderHeader
      NumberHeader: 'numberHeader
      SumElements: int list
      SumResult: int
      RightBorderHeader: 'rightBorderHeader }

let private mxTwoHeadersPivotTableBorder pLeftBorderHeader pNumberHeader pRightBorderHeader =
    c3 
        (pLeftBorderHeader <&> mxMergeStarter)
        (mxUntilA50
            (r2 
                (pNumberHeader <&> mxMergeStarter)
                (mxUntilA50 (mxSumContinuously Direction.Vertical))
            )
        )
        (mxUntilA50 (pRightBorderHeader <&> mxMergeStarter))
    ||>> (fun (leftBorderHeader,(numberHeader,(sumElements,sumResult)),rightBorderHeader) ->
        { LeftBorderHeader = leftBorderHeader
          NumberHeader = numberHeader 
          SumElements = sumElements
          SumResult = sumResult 
          RightBorderHeader = rightBorderHeader }
    )


type NormalColumn =
    { Header: string 
      Contents: obj list }

[<RequireQualifiedAccess>]
module NormalColumn =
    let internal fixEmptyUp column =
        { column with 
            Contents = 
                column.Contents 
                |> List.mapi (fun i content ->
                    let isNullOrEmpty = 
                        match content with 
                        | null -> true
                        | :? string as text -> text.Trim() = ""
                        | _ -> false

                    if isNullOrEmpty then 
                        column.Contents.[0 .. i - 1]
                        |> List.tryFindBack (fun content -> not (isNull content))
                        |> function 
                            | Some v -> v
                            | None -> null
                    else content
                )
        }

type TwoHeadersPivotTable<'groupingColumnChildHeader, 'groupingColumnElement> =
    { GroupingColumn: GroupingColumn<'groupingColumnChildHeader, 'groupingColumnElement>
      NormalColumns: NormalColumn list
      SumNumber: int }

type TwoHeadersPivotTable =
    static member private FixEmptyUp twoHeadersPivotTable =
        { twoHeadersPivotTable with 
            NormalColumns = 
                twoHeadersPivotTable.NormalColumns |> List.map NormalColumn.fixEmptyUp
        }

    static member ToFrame (twoHeadersPivotTable, ?fixEmptyUp) =
        let twoHeadersPivotTable = 
            let isFixEmptyUp = defaultArg fixEmptyUp true
            match isFixEmptyUp with 
            | true ->
                TwoHeadersPivotTable.FixEmptyUp twoHeadersPivotTable
            | false -> twoHeadersPivotTable

        let groupingColumn = twoHeadersPivotTable.GroupingColumn

        let groupingColumnHeader = groupingColumn.Header
        
        let baseTable = 
            let normalColumnsHeaders = 
                List.map (fun column -> column.Header) twoHeadersPivotTable.NormalColumns 
            
            let contentFrame = 
                twoHeadersPivotTable.NormalColumns 
                |> List.mapi(fun i column -> normalColumnsHeaders.[i], Series.ofValues column.Contents)
                |> Frame.ofColumns

            contentFrame.IndexColumnsWith normalColumnsHeaders

        baseTable
        |> Frame.splitRowToMany [groupingColumnHeader.GroupedHeader; groupingColumnHeader.GroupedHeader + "_Value"] (fun rowKey row ->
            let elements = groupingColumn.ElementsList.[rowKey]
            elements 
            |> Seq.indexed
            |> Seq.choose (fun (i,element) ->
                match element with 
                | Some element -> 
                    let addtionalValues = [box groupingColumnHeader.ChildHeaders.[i]; box element]
                    Some (Seq.append row.ValuesAll addtionalValues)
                | None -> None
            )
        )

    static member ToArray2D (twoHeadersPivotTable, ?fixEmptyUp) =
        TwoHeadersPivotTable.ToFrame(twoHeadersPivotTable, ?fixEmptyUp = fixEmptyUp)
        |> Frame.toArray2DWithHeader



let mxTwoHeadersPivotTable pLeftBorderHeader pNumberHeader pRightBorderHeader pGroupingColumn =
    
    mxTwoHeadersPivotTableBorder pLeftBorderHeader pNumberHeader pRightBorderHeader
    |> MatrixParser.collectOutputStream (fun outputStream ->
        let reranged = OutputMatrixStream.reRange outputStream
        let range = 
            reranged
            |> Seq.head
            |> SingletonExcelRangeBase.Create

        let resetedInputStream = 
            { Range = range
              Shift = Shift.Start }

        let p = 
            c3 
                ((pLeftBorderHeader <&> mxMergeStarter) ||>> ignore)
                (mxUntilA50
                    ((mxGroupingColumn pGroupingColumn)))
                (inDebug(mxUntilA50 (pRightBorderHeader <&> mxMergeStarter)) ||>> ignore)
            ||>> ((fun (_, b, _) -> b) >> fun groupingColumn ->
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
        p resetedInputStream

     

    )

