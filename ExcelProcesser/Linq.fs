namespace ExcelProcess
[<RequireQualifiedAccessAttribute>]
module List=
    let mapTail f list=
        list
        |>List.mapi(fun i o->if i = list.Length - 1 then f o else o)        