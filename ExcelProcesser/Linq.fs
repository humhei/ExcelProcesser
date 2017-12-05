[<RequireQualifiedAccessAttribute>]
module List
let mapHead f list=
    list
    |>List.mapi(fun i o->if i=0 then f o else o)        