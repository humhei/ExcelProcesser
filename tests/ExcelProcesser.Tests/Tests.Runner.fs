// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open System
open Tests.MatrixTests
open Tests.MathTests
open Tests.SematicsParsers
open Tests.RealWorldSamples
open Tests

let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let liteDbTests = 
    testList "All tests" [  
        shiftTests
        matrixTests
        mathTests
        #if !TestVirtual
        SematicsParsers.Normal.tests
        SematicsParsers.RangeInHeader.tests
        #endif
        realWorldSamples
    ]


[<EntryPoint>]
let main argv = 
    runTests testConfig liteDbTests |> ignore
    Console.Read()
