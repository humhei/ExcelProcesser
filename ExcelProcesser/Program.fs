// Learn more about F# at http://fsharp.org

open System
open OfficeOpenXml
open System.IO
open Parser

[<EntryPoint>]
let main argv =
    let rps=IRunPaser "test3.xlsx"
    let a = Array2D.init 3 3 (fun x y -> (x,y))
    let b=a
    printfn "Hello World from F#!"
    0 // return an integer exit code
