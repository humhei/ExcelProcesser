namespace ExcelProcesser.SematicsParsers
open CellScript.Core
open ExcelProcesser.MatrixParsers

type ContentsTable = private ContentsTable of Table
with 
    static member Parser(start: SingletonMatrixParser<_>) =
        
        let headers = 
            c2 start (mxColMany (mxNonEmpty))
            ||>> (fun (a, b) -> a :: b)

        let parser = 
            let rows = 
                (mxRowMany1 (mxColMany1 mxNonEmpty))
            (r2 
                headers
                rows)

        parser
        ||>> (fun (headers, rows) ->
            let rows = 
                rows
                |> List.map (fun row -> List.take headers.Length row)

            headers :: rows
            |> array2D
            |> Array2D.map box
            |> Table.OfArray2D
        )
