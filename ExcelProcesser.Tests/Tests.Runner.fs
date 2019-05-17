// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open Tests.ArrayParsers
open ExcelProcess
open Tests.MatrixParsers 
open Tests.Math

let testConfig =  
    { defaultConfig with 
         parallelWorkers = 1
         verbosity = Debug }

let tests = 
    testList "All tests" [  
        ArrayParserTests
        MatrixParserTests
        //MathTests
    ]

[<EntryPoint>]
let main _ = 
    runTests testConfig tests
