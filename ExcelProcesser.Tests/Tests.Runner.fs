// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open Tests.ArrayParsers
open ExcelProcess
open Tests.MatrixParsers 
let testConfig =  
    { defaultConfig with 
         parallelWorkers = 1
         verbosity = Debug }

let tests = 
    testList "All tests" [  
        ArrayParserTests
        //MatrixParserTests
    ]

[<EntryPoint>]
let main _ = 
    let v = runTests testConfig tests
    #if NET462
    Tests.Types.app.Quit()
    #endif
    v
