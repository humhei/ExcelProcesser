// Learn more about F# at http://fsharp.org

open Parser

[<EntryPoint>]
let main _ =
    IRunPaser "test3.xlsx"|>ignore
    printfn "Hello World from F#!"
    0 // return an integer exit code
