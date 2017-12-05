// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open Tests.MyTests
let testConfig =  
    { defaultConfig with 
         parallelWorkers = 1
         verbosity = Debug }

let liteDbTests = 
    testList "All tests" [  
        MyTests
    ]

[<EntryPoint>]
let main _ = runTests testConfig liteDbTests
