module ExcelProcesser.SematicsParsers
open ExcelProcesser.MatrixParsers
open FParsec.CharParsers
open ExcelProcesser.CellParsers
open System

type GroupingColumnHeader<'childHeader> =
    { GroupedHeader: string 
      ChildHeaders: 'childHeader list }

let mxGroupingColumnsHeader pChild =
    r2
        (mxMerge Direction.Horizontal)
        (mxMany1 Direction.Horizontal pChild)
    |> MatrixParser.filterOutputStreamByResultValue (fun ((groupedHeader, emptys), childs) ->
        groupedHeader.Trim() <> "" 
        && emptys.Length = childs.Length - 1
    )
    |||> fun ((groupedHeader, _), childHeaders) -> { GroupedHeader = groupedHeader; ChildHeaders = childHeaders}


type GroupColumns<'childHeader, 'element> =
    { Header: GroupingColumnHeader<'childHeader>
      ElementsList: (int * 'element) list list }

let mxGroupingColumns pChildHeader pElementSkip pElement =
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
                rm (inDebug (mxManySkipKeep Direction.Horizontal pElementSkip columns pElementInRange) |||> (List.mapi (fun i result ->
                    match result with 
                    | Result.Ok ok -> Some (i,ok)
                    | Result.Error _ -> None
                ) >> List.choose id))
            | None -> rm ((cm pElementInRange) |||> List.indexed )
        )
    ) 
    |||> fun (header, elementsList) ->
        { Header = header
          ElementsList = elementsList
        }