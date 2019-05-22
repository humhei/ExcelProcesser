// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open System
open Tests.MatrixTests
open Tests.MathTests
open Tests.MatrixAstTests
let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let liteDbTests = 
    testList "All tests" [  
        matrixTests
        matrixAstTests
        mathTests
    ]


[<EntryPoint>]
let main argv = 
    runTests testConfig liteDbTests
