module ExcelProcesser.SematicsParsers
open ExcelProcesser.MatrixParsers
open FParsec.CharParsers
open ExcelProcesser.CellParsers
open System
open ExcelProcesser.MathParsers

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

        let private joinByRows (a1: 'a[,]) (a2: 'a[,]) =
            let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
            if a1l2 <> a2l2 then failwith "arrays have different column sizes"
            let result = Array2D.zeroCreate (a1l1 + a2l1) a1l2
            Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
            Array2D.blit a2 0 0 result a1l1 0 a2l1 a2l2
            result

        let private joinByCols (a1: 'a[,]) (a2: 'a[,]) =
            let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
            if a1l1 <> a2l1 then failwith "arrays have different row sizes"
            let result = Array2D.zeroCreate a1l1 (a1l2 + a2l2)
            Array2D.blit a1 0 0 result 0 0 a1l1 a1l2
            Array2D.blit a2 0 0 result 0 a1l2 a2l1 a2l2
            result

        let pickHeaderTailRowsNotInclude headerIndex tailIndex (array2D: 'a [,]) =
            rebasingMap (fun array2D ->
                let headers = array2D.[0 .. headerIndex - 1, *]
                let tails = array2D.[tailIndex + 1 .. (Array2D.length1 array2D) - 1, *]
                joinByRows headers tails
            ) array2D

        let removeSencondRow (array2D: 'a [,]) =
            pickHeaderTailRowsNotInclude 1 1 array2D

        let removeLastRow (array2D: 'a [,]) =
            array2D.[0.. (Array2D.length1 array2D - 2), *]

        let pickHeaderTailColumnsNotInclude headerIndex tailIndex (array2D: 'a [,]) =
            rebasingMap (fun array2D ->
                let headers = array2D.[*, 0 .. headerIndex - 1]
                let tails = array2D.[*, tailIndex + 1 .. (Array2D.length2 array2D) - 1]
                joinByCols headers tails
            ) array2D

        let pickHeaderTailColumnsNotIncludeByIndexer (coordinate: Coordinate, shift) (array2D: 'a [,]) =
            let headerIndex = coordinate.X
            let tailIndex = headerIndex + shift
            pickHeaderTailColumnsNotInclude headerIndex tailIndex array2D


type GroupingColumnHeader<'childHeader> =
    { GroupedHeader: string 
      ChildHeaders: 'childHeader list
      Shift: Shift }
with 
    member x.Indexer = 
        match x.Shift.Last with 
        | Horizontal (coordinate, i) -> coordinate, i
        | _ -> failwith "Last shift should be horizontal"

let mxGroupingColumnsHeader pChild =
    r2
        (mxMerge Direction.Horizontal)
        (mxMany1 Direction.Horizontal pChild)
    |> MatrixParser.filterOutputStreamByResultValue (fun ((groupedHeader, emptys), childs) ->
        groupedHeader.Text.Trim() <> "" 
        && emptys.Length = childs.Length - 1
    )
    |||>> fun outputStream ((groupedHeader, _), childHeaders) -> 
        { GroupedHeader = groupedHeader.Text
          ChildHeaders = childHeaders
          Shift = outputStream.Shift }



type GroupingColumn<'childHeader, 'element> =
    { Header: GroupingColumnHeader<'childHeader>
      ElementsList: (int * 'element) list list }

[<RequireQualifiedAccess>]
module GroupingColumn =
    let untype groupingColumn =
        { Header = 
            { GroupedHeader = groupingColumn.Header.GroupedHeader
              ChildHeaders = groupingColumn.Header.ChildHeaders |> List.map box
              Shift = groupingColumn.Header.Shift }
          ElementsList = 
            groupingColumn.ElementsList |> List.map (List.map (fun (a,b) -> a,box b))
      }


type GroupingColumnParserArg<'childHeader,'elementSkip, 'element> =
    GroupingColumnParserArg of 
        pChildHeader: MatrixParser<'childHeader> 
        * pElementSkip: MatrixParser<'elementSkip> option
        * pElement: MatrixParser<'element>

let mxGroupingColumn (GroupingColumnParserArg(pChildHeader, pElementSkip, pElement)) =
    (r2R 
        (mxGroupingColumnsHeader pChildHeader)
        (fun outputStream ->

            let maxColNum, columns = 
                let reranged = OutputMatrixStream.reRange outputStream
                reranged.End.Column,reranged.Columns

            let pElementInRange =
                pElement <&> (mxCellParser (fun range -> range.Start.Column <= maxColNum) ignore)


            match pElementSkip with 
            | Some pElementSkip ->
                rm ((fun a -> mxManySkipRetain Direction.Horizontal pElementSkip columns pElementInRange a) ||>> (List.mapi (fun i result ->
                    match result with 
                    | Result.Ok ok -> Some (i,ok)
                    | Result.Error _ -> None
                ) >> List.choose id))
            | None -> rm ((cm pElementInRange) ||>> List.indexed )
        )
    ) 
    ||>> fun (header, elementsList) ->
        { Header = header
          ElementsList = elementsList
        }

type NormalColumn =
    { Header: string 
      Contents: obj list }

type TwoHeadersPivotTable<'groupingColumnChildHeader, 'groupingColumnElement> =
    { GroupingColumn: GroupingColumn<'groupingColumnChildHeader, 'groupingColumnElement>
      NormalColumns: NormalColumn list }

[<RequireQualifiedAccess>]
module TwoHeadersPivotTable =

    let mxFrame pLeftBorderHeader pNumberHeader pRightBorderHeader =
        c3 
            (pLeftBorderHeader <&> mxMergeStarter)
            (mxUntilA50
                (r2 
                    (pNumberHeader <&> mxMergeStarter)
                    (mxUntilA5 (mxSum Direction.Vertical))
                )
            )
            (mxUntilA50 (pRightBorderHeader <&> mxMergeStarter))

    let parser pLeftBorderHeader pNumberHeader pRightBorderHeader pGroupingColumn =
        
        mxFrame pLeftBorderHeader pNumberHeader pRightBorderHeader
        |> MatrixParser.collectOutputStream (fun outputStream ->
            let reranged = OutputMatrixStream.reRange outputStream
            reranged
            |> List.ofSeq
            |> List.collect (fun range ->
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
                          NormalColumns = normalColumns }
                    )
                p resetedInputStream

            )

        )

